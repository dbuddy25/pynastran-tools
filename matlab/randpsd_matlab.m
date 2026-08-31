function [T, curves] = randpsd_matlab(varargin)
%RANDPSD_MATLAB  Random-response PSD plots from a Nastran OP2 via pyNastran.
%
%   RANDPSD_MATLAB with no arguments runs the CONFIG block below: it reads the
%   OP2, pulls the acceleration PSD for each grid/DOF you listed, overlays the
%   enabled environments, and draws figure(s) ready to copy into PowerPoint.
%
%   [T, CURVES] = RANDPSD_MATLAB(...) also returns:
%       T       - summary table, one row per curve (Grid, Label, DOF, GRMS)
%       CURVES  - struct array carrying the raw .freqs / .psd / .cumrms vectors,
%                 so you can keep plotting without re-reading the OP2.
%
%   Any CONFIG field can be overridden per call without editing the file:
%       >> randpsd_matlab('SUBCASE', 20, 'PLOT_MODE', 'cumrms')
%
%   MATLAB port of the numeric core of
%       postprocessing/modules/asd_overlay.py + asd_common.py
%   Deliberately BATCH-oriented: hardcode the run once, get the same plots
%   every time. The interactive workflow stays in the Python ASD Overlay tool;
%   this is a companion to it, not a replacement.
%
%   ONE-TIME SETUP (per machine)
%   ----------------------------
%       >> pyenv('Version', '/path/to/your/python')   % a Python with pyNastran
%       >> py.importlib.import_module('pyNastran');   % should not error
%
%   REQUIREMENTS IN THE OP2
%   -----------------------
%   A SOL 111 random-response run with, in case control:
%       ACCELERATION(PLOT,PSDF) = ALL
%       RANDOM = <sid>
%
%   NOTE: do NOT also request EKE. pyNastran's OP2 reader mis-parses the
%   element kinetic energy table and the read fails with an unpack error.
%
%   See also MEFF_MATLAB.

% =========================================================================
%                                 CONFIG
%   Edit this block. Everything below it is machinery.
% =========================================================================
C = struct();

% --- input ---------------------------------------------------------------
C.OP2FILE = 'model_111.op2';
C.SUBCASE = [];         % [] = first subcase found; or an integer subcase ID

% --- grids to plot -------------------------------------------------------
%   {grid ID, legend label}.  Label '' falls back to "Grid <id>".
C.GRIDS = {
    101, 'Bracket top'
    205, 'Box corner'
    310, 'Panel center'
};

% --- DOFs ----------------------------------------------------------------
%   1=T1(X)  2=T2(Y)  3=T3(Z).  Any subset: [3] or [1 2 3].
C.DOFS = [1 2 3];

% --- environments (spec / input ASDs to overlay) -------------------------
%   freq [Hz] and asd [g^2/Hz] are BREAKPOINTS -- they get log-log
%   interpolated onto the response frequency grid when drawn.
%   Flip the 'on' flag to false to hide one without deleting it.
C.ENVIRONMENTS = {
%    name             on      freq [Hz]            asd [g^2/Hz]
    'Qual level',    true,   [20 50 800 2000],    [0.01   0.08 0.08 0.02 ]
    'Accept level',  true,   [20 50 800 2000],    [0.0064 0.05 0.05 0.013]
    'MPE',           false,  [20 100 1000 2000],  [0.005  0.04 0.04 0.01 ]
};

% --- what to draw --------------------------------------------------------
C.PLOT_MODE     = 'psd';    % 'psd' | 'cumrms' | 'both' (both = two figures)
C.ONE_FIG_PER   = 'dof';    % 'dof' = a figure per DOF | 'all' = one figure
C.SHOW_GRMS     = true;     % append "(GRMS = x.xx g)" to each legend entry
C.SHOW_ENVELOPE = true;     % draw the enabled ENVIRONMENTS
C.LOGLOG        = true;     % log-log axes for PSD (cumulative RMS forces linear Y)
C.COLOR_BY      = 'grid';   % 'grid' = colour per grid | 'dof' = colour per DOF
C.LINEWIDTH     = 1.5;
C.FREQ_LIMITS   = [];       % [] = auto, else [fmin fmax]
C.PSD_LIMITS    = [];       % [] = auto, else [ymin ymax]

% --- units ---------------------------------------------------------------
%   OP2 accelerations come out in model units; divide by this to reach g.
%   386.089 for in/s^2, 9.80665 for m/s^2.
C.UNIT_FACTOR = 386.089;
C.PSD_UNITS   = 'g^2/Hz';
C.RMS_UNITS   = 'g';

% --- output --------------------------------------------------------------
C.TITLE       = '';         % '' = auto from OP2 filename + subcase
C.EXPORT_PNG  = false;      % save a PNG beside the OP2
C.EXPORT_DPI  = 300;
C.PRINT_TABLE = true;       % echo the GRMS summary to the command window

% =========================================================================
%                       END CONFIG -- machinery below
% =========================================================================

C = apply_overrides(C, varargin);
validate_config(C);

[curves, sc_used] = read_op2_psd(C);
if isempty(curves)
    error('randpsd_matlab:noCurves', ...
        ['None of the requested grids were found in the PSD results.\n', ...
         'Check C.GRIDS against the model and confirm RANDOM output was requested.']);
end

T = summary_table(curves);
if C.PRINT_TABLE
    fprintf('\n'); disp(T); fprintf('\n');
end

draw_all(curves, C, sc_used);
end


% =========================================================================
function C = apply_overrides(C, args)
%APPLY_OVERRIDES  Fold name/value pairs into the CONFIG struct.
    if isempty(args), return; end
    if mod(numel(args), 2) ~= 0
        error('randpsd_matlab:badArgs', ...
              'Overrides must be name/value pairs, e.g. ''SUBCASE'', 20.');
    end
    for k = 1:2:numel(args)
        name = args{k};
        if ~ischar(name) && ~isstring(name)
            error('randpsd_matlab:badArgs', 'Override name must be a string.');
        end
        name = char(name);
        if ~isfield(C, name)
            error('randpsd_matlab:unknownOption', ...
                  'Unknown option "%s". Valid names are the CONFIG field names.', name);
        end
        C.(name) = args{k + 1};
    end
end


% =========================================================================
function validate_config(C)
%VALIDATE_CONFIG  Fail early and readably, rather than deep inside the Python call.
    if exist(C.OP2FILE, 'file') ~= 2
        error('randpsd_matlab:noFile', 'OP2 file not found: %s', C.OP2FILE);
    end
    if isempty(C.GRIDS)
        error('randpsd_matlab:noGrids', 'C.GRIDS is empty -- nothing to plot.');
    end
    if isempty(C.DOFS) || any(C.DOFS < 1) || any(C.DOFS > 6)
        error('randpsd_matlab:badDof', 'C.DOFS entries must lie in 1..6.');
    end
    if ~ismember(lower(C.PLOT_MODE), {'psd', 'cumrms', 'both'})
        error('randpsd_matlab:badMode', ...
              'C.PLOT_MODE must be ''psd'', ''cumrms'' or ''both''.');
    end
    if ~ismember(lower(C.ONE_FIG_PER), {'dof', 'all'})
        error('randpsd_matlab:badFigMode', ...
              'C.ONE_FIG_PER must be ''dof'' or ''all''.');
    end
    if C.UNIT_FACTOR <= 0
        error('randpsd_matlab:badUnits', 'C.UNIT_FACTOR must be positive.');
    end
    if ~isempty(C.ENVIRONMENTS) && size(C.ENVIRONMENTS, 2) ~= 4
        error('randpsd_matlab:badEnv', ...
              'C.ENVIRONMENTS needs 4 columns: name, on, freq, asd.');
    end
    for e = 1:size(C.ENVIRONMENTS, 1)
        if numel(C.ENVIRONMENTS{e, 3}) ~= numel(C.ENVIRONMENTS{e, 4})
            error('randpsd_matlab:badEnv', ...
                  'Environment "%s": freq and asd have different lengths.', ...
                  C.ENVIRONMENTS{e, 1});
        end
    end
end


% =========================================================================
function [curves, sc_used] = read_op2_psd(C)
%READ_OP2_PSD  Pull acceleration PSD per grid/DOF out of the OP2 via pyNastran.
%   Mirrors _get_response_psd() in asd_overlay.py. PSD is divided by
%   UNIT_FACTOR^2 (not UNIT_FACTOR) because PSD carries squared units.

    op2mod = py.importlib.import_module('pyNastran.op2.op2');
    op2 = op2mod.OP2(pyargs('mode', 'nx'));
    op2.read_op2(C.OP2FILE);

    psd_dict = op2.op2_results.psd.accelerations;
    if double(py.len(psd_dict)) == 0
        error('randpsd_matlab:noPsd', ...
            ['No acceleration PSD in this OP2.\n', ...
             'Add to case control:\n', ...
             '    ACCELERATION(PLOT,PSDF) = ALL\n', ...
             '    RANDOM = <sid>']);
    end

    [psd_tbl, sc_used] = lookup_subcase(psd_dict, C.SUBCASE);

    % _times holds the frequency vector; getattr because of the leading underscore.
    freqs = ndarray2mat(py.getattr(psd_tbl, '_times'));
    freqs = freqs(:);

    % node_gridtype is (nnodes, 2); column 1 is the grid ID.
    ids  = ndarray2mat(psd_tbl.node_gridtype);
    ids  = ids(:, 1);
    data = ndarray2mat(psd_tbl.data);      % (nfreq, nnodes, ndof)

    rms_tbl = try_rms_table(op2, sc_used);

    dof_names = {'T1 (X)', 'T2 (Y)', 'T3 (Z)', 'R1', 'R2', 'R3'};

    curves = struct('grid', {}, 'label', {}, 'idof', {}, 'dof_name', {}, ...
                    'freqs', {}, 'psd', {}, 'cumrms', {}, 'grms', {}, ...
                    'grms_src', {});

    for g = 1:size(C.GRIDS, 1)
        gid = C.GRIDS{g, 1};
        lbl = C.GRIDS{g, 2};
        if isempty(lbl), lbl = sprintf('Grid %d', gid); end

        row = find(ids == gid, 1);
        if isempty(row)
            warning('randpsd_matlab:gridMissing', ...
                    'Grid %d not found in the PSD results -- skipped.', gid);
            continue
        end

        for idof = C.DOFS(:)'
            if idof > size(data, 3)
                warning('randpsd_matlab:dofMissing', ...
                        'DOF %d not present in the results -- skipped.', idof);
                continue
            end
            psd = squeeze(data(:, row, idof)) / (C.UNIT_FACTOR ^ 2);
            psd = psd(:);

            cum  = cumulative_grms_loglog(freqs, psd);
            [grms, src] = overall_grms(rms_tbl, ids, row, idof, ...
                                       freqs, psd, C.UNIT_FACTOR);

            curves(end + 1) = struct( ...
                'grid', gid, 'label', lbl, 'idof', idof, ...
                'dof_name', dof_names{min(idof, numel(dof_names))}, ...
                'freqs', freqs, 'psd', psd, 'cumrms', cum, ...
                'grms', grms, 'grms_src', src);  %#ok<AGROW>
        end
    end
end


% =========================================================================
function rms_tbl = try_rms_table(op2, sc_used)
%TRY_RMS_TABLE  Nastran's own RMS table for this subcase, or [] if absent.
%   asd_overlay.py prefers this over integrating the PSD, so we do too.
    rms_tbl = [];
    try
        rms_dict = op2.op2_results.rms.accelerations;
        if double(py.len(rms_dict)) == 0
            return
        end
        rms_tbl = lookup_subcase(rms_dict, sc_used);
        if isa(rms_tbl, 'py.NoneType')
            rms_tbl = [];
        end
    catch
        rms_tbl = [];   % no RMS output in this OP2; integration covers it
    end
end


% =========================================================================
function [grms, src] = overall_grms(rms_tbl, ids, row, idof, freqs, psd, unit_factor)
%OVERALL_GRMS  Nastran's RMS value when available, else FEMCI log-log integration.
%   Mirrors _get_rms_scalar() in asd_overlay.py, including its fallback.
    if ~isempty(rms_tbl)
        try
            rids = ndarray2mat(rms_tbl.node_gridtype);
            rrow = find(rids(:, 1) == ids(row), 1);
            if ~isempty(rrow)
                rdata = ndarray2mat(rms_tbl.data);
                grms  = abs(rdata(1, rrow, idof)) / unit_factor;
                src   = 'OP2';
                return
            end
        catch
            % fall through to integration
        end
    end
    grms = sqrt(grms_loglog(freqs, psd));
    src  = 'integrated';
end


% =========================================================================
function [tbl, sc] = lookup_subcase(pydict, want)
%LOOKUP_SUBCASE  Fetch a result table by subcase ID regardless of key format.
%   pyNastran keys are sometimes a bare int, sometimes a tuple whose first
%   entry is the subcase -- mirrors sc_int()/lookup_subcase() in asd_common.py.
    keys = cell(py.list(pydict.keys()));
    if isempty(keys)
        error('randpsd_matlab:noSubcase', 'No subcases in the result dictionary.');
    end

    scs = zeros(1, numel(keys));
    for k = 1:numel(keys)
        scs(k) = key_to_int(keys{k});
    end
    [scs, order] = sort(scs);
    keys = keys(order);

    if isempty(want)
        idx = 1;                        % first subcase
    else
        idx = find(scs == want, 1);
        if isempty(idx)
            error('randpsd_matlab:noSubcase', ...
                  'Subcase %d not found. Available: %s', ...
                  want, mat2str(scs));
        end
    end
    tbl = pydict{keys{idx}};
    sc  = scs(idx);
end


% =========================================================================
function n = key_to_int(key)
%KEY_TO_INT  Normalise a pyNastran result-dict key to a plain integer.
    if isa(key, 'py.tuple')
        n = double(py.int(key{1}));
    else
        n = double(py.int(key));
    end
end


% =========================================================================
function out = interp_loglog(freqs_in, asd_in, query_freqs)
%INTERP_LOGLOG  Log-log interpolate an ASD onto query_freqs. Out of range -> 0.
%   Direct port of interp_loglog() in asd_common.py.
    freqs_in = freqs_in(:); asd_in = asd_in(:); query_freqs = query_freqs(:);
    out = zeros(numel(query_freqs), 1);

    for i = 1:numel(query_freqs)
        f = query_freqs(i);
        if f < freqs_in(1) || f > freqs_in(end)
            continue                                  % leave 0
        end
        idx = find(freqs_in <= f, 1, 'last');
        idx = min(idx, numel(freqs_in) - 1);
        fl = freqs_in(idx);   fh = freqs_in(idx + 1);
        al = asd_in(idx);     ah = asd_in(idx + 1);
        if fl <= 0 || fh <= 0 || al <= 0 || ah <= 0 || fl == fh
            out(i) = al;
        else
            b = log(ah / al) / log(fh / fl);
            out(i) = al * (f / fl) ^ b;
        end
    end
end


% =========================================================================
function area = grms_loglog(freqs, asd)
%GRMS_LOGLOG  Area under an ASD curve, analytical log-log segment integration.
%   Direct port of grms_loglog() in asd_common.py (the FEMCI formulation).
%   Returns AREA -- take sqrt() for RMS.
    area = 0.0;
    for i = 1:numel(freqs) - 1
        fl = freqs(i);   fh = freqs(i + 1);
        al = asd(i);     ah = asd(i + 1);
        if fl <= 0 || fh <= 0 || al <= 0 || ah <= 0
            continue
        end
        log_f = log(fh / fl);
        if log_f ~= 0
            b = log(ah / al) / log_f;
        else
            b = 0.0;
        end
        if abs(b + 1.0) < 1e-6
            area = area + al * fl * log_f;            % b == -1 limit
        else
            area = area + (ah * fh - al * fl) / (b + 1.0);
        end
    end
end


% =========================================================================
function cum = cumulative_grms_loglog(freqs, asd)
%CUMULATIVE_GRMS_LOGLOG  Running RMS up to each frequency. cum(1) = 0.
%   Direct port of cumulative_grms_loglog() in asd_common.py.
    n = numel(freqs);
    cum_area = zeros(n, 1);
    running  = 0.0;

    for i = 1:n - 1
        fl = freqs(i);   fh = freqs(i + 1);
        al = asd(i);     ah = asd(i + 1);
        if fl <= 0 || fh <= 0 || al <= 0 || ah <= 0
            cum_area(i + 1) = running;
            continue
        end
        log_f = log(fh / fl);
        if log_f ~= 0
            b = log(ah / al) / log_f;
        else
            b = 0.0;
        end
        if abs(b + 1.0) < 1e-6
            running = running + al * fl * log_f;
        else
            running = running + (ah * fh - al * fl) / (b + 1.0);
        end
        cum_area(i + 1) = running;
    end
    cum = sqrt(max(cum_area, 0.0));
end


% =========================================================================
function T = summary_table(curves)
%SUMMARY_TABLE  One row per curve: grid, label, DOF, GRMS and where it came from.
    Grid   = arrayfun(@(c) c.grid, curves)';
    Label  = string(arrayfun(@(c) string(c.label),    curves))';
    DOF    = string(arrayfun(@(c) string(c.dof_name), curves))';
    GRMS   = arrayfun(@(c) c.grms, curves)';
    Source = string(arrayfun(@(c) string(c.grms_src), curves))';
    T = table(Grid, Label, DOF, GRMS, Source);
end


% =========================================================================
function draw_all(curves, C, sc_used)
%DRAW_ALL  Fan out into the requested figures per PLOT_MODE / ONE_FIG_PER.
    mode = lower(C.PLOT_MODE);
    if strcmp(mode, 'both')
        modes = {'psd', 'cumrms'};
    else
        modes = {mode};
    end

    for m = 1:numel(modes)
        if strcmpi(C.ONE_FIG_PER, 'dof')
            for idof = C.DOFS(:)'
                sel = curves([curves.idof] == idof);
                if isempty(sel), continue; end
                draw_figure(sel, C, sc_used, modes{m}, sel(1).dof_name);
            end
        else
            draw_figure(curves, C, sc_used, modes{m}, '');
        end
    end
end


% =========================================================================
function draw_figure(sel, C, sc_used, mode, dof_tag)
%DRAW_FIGURE  One figure: the selected curves, optional envelopes, legend, export.
    fig = figure('Color', 'w');
    ax  = axes(fig); %#ok<LAXES>
    hold(ax, 'on'); grid(ax, 'on'); box(ax, 'on');

    is_cum = strcmpi(mode, 'cumrms');
    cols   = lines(max(numel(sel), 7));
    legend_txt = {};

    for k = 1:numel(sel)
        c = sel(k);
        if strcmpi(C.COLOR_BY, 'dof')
            colour = cols(mod(c.idof - 1, size(cols, 1)) + 1, :);
        else
            colour = cols(mod(k - 1, size(cols, 1)) + 1, :);
        end

        if is_cum
            y = c.cumrms;
        else
            y = c.psd;
        end
        plot(ax, c.freqs, y, 'LineWidth', C.LINEWIDTH, 'Color', colour);

        name = c.label;
        if isempty(dof_tag)
            name = sprintf('%s  %s', name, c.dof_name);
        end
        if C.SHOW_GRMS
            name = sprintf('%s  (GRMS = %.3g %s)', name, c.grms, C.RMS_UNITS);
        end
        legend_txt{end + 1} = name; %#ok<AGROW>
    end

    % --- environments -----------------------------------------------------
    %   Interpolated onto the response grid so they share the plotted band,
    %   and skipped for cumulative RMS where an input spec has no meaning.
    if C.SHOW_ENVELOPE && ~is_cum && ~isempty(C.ENVIRONMENTS)
        fgrid = sel(1).freqs;
        % Environments cycle through distinct dark shades and dash patterns so
        % they stay tellable apart as rows get added, while remaining visually
        % subordinate to the response curves.
        env_cols   = [0.20 0.20 0.20; 0.55 0.10 0.10; 0.10 0.30 0.55;
                      0.30 0.45 0.15; 0.45 0.20 0.50];
        env_styles = {'--', ':', '-.'};
        n_drawn = 0;
        for e = 1:size(C.ENVIRONMENTS, 1)
            if ~C.ENVIRONMENTS{e, 2}, continue; end
            ef = C.ENVIRONMENTS{e, 3};
            ea = C.ENVIRONMENTS{e, 4};
            ey = interp_loglog(ef, ea, fgrid);
            ey(ey <= 0) = NaN;              % keep gaps out of the log plot

            colour = env_cols(mod(n_drawn, size(env_cols, 1)) + 1, :);
            style  = env_styles{mod(n_drawn, numel(env_styles)) + 1};
            plot(ax, fgrid, ey, style, 'LineWidth', 1.2, 'Color', colour);
            n_drawn = n_drawn + 1;

            env_grms = sqrt(grms_loglog(ef(:), ea(:)));
            legend_txt{end + 1} = sprintf('%s  (%.3g %s)', ...
                C.ENVIRONMENTS{e, 1}, env_grms, C.RMS_UNITS); %#ok<AGROW>
        end
    end

    % --- axes -------------------------------------------------------------
    if is_cum
        % Linear Y: cumulative RMS is about where energy accumulates, and a
        % log Y axis flattens exactly the steps you are looking for.
        set(ax, 'YScale', 'linear');
        if C.LOGLOG, set(ax, 'XScale', 'log'); end
        ylabel(ax, sprintf('Cumulative RMS [%s]', C.RMS_UNITS));
    else
        if C.LOGLOG
            set(ax, 'XScale', 'log', 'YScale', 'log');
        end
        ylabel(ax, sprintf('Acceleration PSD [%s]', C.PSD_UNITS));
        if ~isempty(C.PSD_LIMITS), ylim(ax, C.PSD_LIMITS); end
    end
    xlabel(ax, 'Frequency [Hz]');
    if ~isempty(C.FREQ_LIMITS), xlim(ax, C.FREQ_LIMITS); end

    % --- title ------------------------------------------------------------
    if isempty(C.TITLE)
        [~, base, ext] = fileparts(C.OP2FILE);
        ttl = sprintf('%s%s  --  Subcase %d', base, ext, sc_used);
    else
        ttl = C.TITLE;
    end
    if ~isempty(dof_tag)
        ttl = sprintf('%s  --  %s', ttl, dof_tag);
    end
    if is_cum
        ttl = sprintf('%s  (cumulative RMS)', ttl);
    end
    title(ax, ttl, 'Interpreter', 'none');

    legend(ax, legend_txt, 'Location', 'best', 'Interpreter', 'none');

    % --- export -----------------------------------------------------------
    if C.EXPORT_PNG
        [dirname, base] = fileparts(C.OP2FILE);
        tag = mode;
        if ~isempty(dof_tag)
            tag = sprintf('%s_%s', mode, matlab.lang.makeValidName(dof_tag));
        end
        out = fullfile(dirname, sprintf('%s_sc%d_%s.png', base, sc_used, tag));
        exportgraphics(ax, out, 'Resolution', C.EXPORT_DPI);
        fprintf('Wrote %s\n', out);
    end
end


% =========================================================================
function A = ndarray2mat(nd)
%NDARRAY2MAT  numpy ndarray (or list) -> MATLAB double array, shape preserved.
%   numpy is row-major (C order), MATLAB column-major, so flatten in C order
%   then reshape + permute back. Same helper as meff_matlab.m.
    nd = py.numpy.asarray(nd);
    nd = py.numpy.ascontiguousarray(nd, pyargs('dtype', 'float64'));

    shp  = cellfun(@double, cell(nd.shape));
    flat = double(py.array.array('d', nd.flatten().tolist()));

    if isempty(shp)
        A = flat;
    elseif numel(shp) == 1
        A = flat(:);
    else
        A = permute(reshape(flat, fliplr(shp)), numel(shp):-1:1);
    end
end
