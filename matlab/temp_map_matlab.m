function [R, S, RM] = temp_map_matlab(varargin)
%TEMP_MAP_MATLAB  Map CSV temperature point clouds onto Nastran GRIDs, write TEMP cards.
%
%   Headless engine. TEMP_MAP_GUI is a front end for it; anything the GUI does
%   can be done from a script with this function.
%
%   QUICK START
%   -----------
%   Single part, every CSV in a folder, one TEMP file per time step:
%       >> temp_map_matlab('BDF_FILE', 'wing.bdf', 'CSV_DIR', 'clouds', ...
%                          'OUT_DIR', 'temp_cards', ...
%                          'BDF_LENGTH_UNITS', 'in', 'CSV_LENGTH_UNITS', 'm', ...
%                          'CSV_TEMP_UNITS', 'K',   'OUT_UNITS', 'F', ...
%                          'METHOD', 'linear', 'SID_START', 100)
%
%   Multi-part assembly (one BDF + one cloud folder per part; steps pair across
%   folders by the number before .csv -- wing_20.csv <-> fuse_20.csv -- or by
%   the whole file name when there is none; step k of every part gets the same SID):
%       >> temp_map_matlab('PARTS', {'wing', 'wing.bdf', 'clouds\wing'; ...
%                                    'fuse', 'fuse.bdf', 'clouds\fuse'}, ...
%                          'OUT_DIR', 'temp_cards', 'OUT_UNITS', 'F')
%
%   No cloud for a part?  Put a temperature in column 3 instead of a folder: a
%   constant (in OUT_UNITS, the model's units) or a 2-column CSV (time, T in
%   OUT_UNITS) for a uniform temperature that changes per step.  Uniform parts
%   follow the cloud parts' time steps, so they mix freely:
%       >> temp_map_matlab('PARTS', {'wing', 'wing.bdf', 'clouds\wing'; ...
%                                    'fuse', 'fuse.bdf', 70; ...
%                                    'tank', 'tank.bdf', 'tank_T.csv'}, 'OUT_UNITS', 'F')
%   Whole model isothermal (one step), or several uniform steps from a T(t)
%   table / an explicit list of times:
%       >> temp_map_matlab('BDF_FILE', 'model.bdf', 'ISOTHERMAL', 70, 'OUT_UNITS', 'F')
%       >> temp_map_matlab('BDF_FILE', 'model.bdf', 'ISOTHERMAL', 'model_T.csv')
%       >> temp_map_matlab('PARTS', P, 'TIMES', [0 100 200])       % no cloud parts in P
%
%   Or edit the CONFIG block below and run TEMP_MAP_MATLAB with no arguments.
%   Every name/value pair overrides the CONFIG field of the same name; unknown
%   names error out, so typos cannot silently do nothing.
%
%   OUTPUT FILES  (all under OUT_DIR)
%   -----------------------------------
%       <OUT_NAME>.bdf            TEMP cards for one time step, one SID, every
%                                 part in its own $ section; meant to be INCLUDEd.
%                                 Default temp_001.bdf, temp_002.bdf, ...
%       temp_subcases.dat         case control: global lines (CASE_EXTRA,
%                                 TEMPERATURE(INITIAL) if TREF) then one SUBCASE
%                                 per step with TEMPERATURE(LOAD) = SID.
%                                 Paste above BEGIN BULK.        (WRITE_CASE)
%       temp_includes.bdf         one INCLUDE per TEMP file (+ TEMPD for TREF).
%                                 Paste below BEGIN BULK.         (WRITE_CASE)
%       <OUT_NAME>.png            ISO check plot per step          (SAVE_PNG)
%       temp_map_report.html      run record: settings, coverage warnings,
%                                 min/max-vs-time chart, per-step table,
%                                 check plots                     (REPORT)
%
%   OUTPUTS
%   -------
%   [R, S, RM] = TEMP_MAP_MATLAB(...)
%       R   - results, one struct per (part, time step): grid ids / xyz / mapped
%             T (Kelvin), the cloud, hull / distance flags, coverage text.
%             Feed one to TEMP_MAP_PLOT.
%       S   - summary table, one row per part and step.
%       RM  - one merged, assembly-wide result per step (== R for one part).
%
%   REUSING WORK
%   ------------
%   Parse the BDF(s) once, map many folders:
%       >> G = temp_map_matlab('BDF_FILE', 'wing.bdf', 'READ_ONLY', true);
%       >> temp_map_matlab('GRIDS', G, 'CSV_DIR', 'clouds_hot');
%       >> temp_map_matlab('GRIDS', G, 'CSV_DIR', 'clouds_cold', 'METHOD', 'idw');
%   Re-write earlier results with different output options, no re-mapping:
%       >> [R, S] = temp_map_matlab(..., 'WRITE', false);
%       >> temp_map_matlab('RESULTS', R, 'OUT_UNITS', 'C', 'FIELD_SIZE', 16);
%   Consecutive CSVs with identical point locations (same thermal mesh, new
%   temperatures) reuse the previous mapping automatically: the geometry is
%   built once and each further step is a sparse matrix-vector product.
%   PARALLEL = true maps those later steps in a parfor (Parallel Computing
%   Toolbox) -- reading 500 CSVs is what takes the time once the map exists.
%   In parallel every step is compared against the FIRST step's cloud (serial
%   compares against the previous step), and Cancel takes effect between
%   steps rather than inside one.
%
%   CSV FORMAT  (one file = one time step; header row optional)
%   -----------------------------------------------------------
%       time_s, x, y, z, T
%   x/y/z in CSV_LENGTH_UNITS, basic frame; T in CSV_TEMP_UNITS. Columns
%   after the 5th are ignored; rows with a blank or non-numeric value are dropped.
%
%   MAPPING METHODS  (METHOD)
%   -------------------------
%   'linear'    Delaunay interpolation, nearest point outside the cloud's hull,
%               with a guard (SLIVER_FACTOR) that falls back to nearest inside
%               long "bridging" cells across concavities. Exact for volume
%               clouds. ~1 min and several GB at 1M points, built once.
%   'scattered' MATLAB's scatteredInterpolant(P, T, 'linear', 'nearest')
%               verbatim -- no guard -- for comparison with other scripts.
%   'nearest'   closest cloud point (kd-tree, seconds at 1M x 1M).
%   'idw'       inverse-distance mean of the IDW_K closest points (kd-tree).
%   'nearest' / 'idw' need knnsearch (Statistics Toolbox); without it they fall
%   back to the triangulation.
%
%   CHECKS
%   ------
%   Each step prints a coverage report (cloud vs mesh extents, grids outside
%   the hull, distance to the nearest cloud point, temperature ranges) and a
%   loud banner when the mesh overhangs the cloud (OVERHANG_WARN) -- the usual
%   sign of a length-units mismatch. R(k).extrap flags grids outside the hull,
%   R(k).far those beyond EXTRAP_WARN_DIST. GRIDs are assumed to be in the
%   basic coordinate system (CP = 0); a warning is issued otherwise.
%
%   Pure MATLAB -- no Python, no pyNastran.
%
%   See also TEMP_MAP_PLOT, TEMP_MAP_GUI.

% =========================================================================
%                                 CONFIG
%   Edit this block. Everything below it is machinery.
% =========================================================================
C = struct();

% --- input ---------------------------------------------------------------
C.BDF_FILE       = 'model.bdf';
C.CSV_DIR        = '.';         % folder holding the temperature clouds
C.PARTS          = {};          % multi-part: N x 3 cell {name, bdf_file, source};
                                % source = cloud folder | constant T | T(t) csv (see ISOTHERMAL)
                                % when given, BDF_FILE, CSV_DIR and ISOTHERMAL are ignored
C.ISOTHERMAL     = [];          % single part without a cloud: a constant temperature in
                                % OUT_UNITS, or a 2-column CSV (time, T in OUT_UNITS) giving
                                % one uniform temperature per step.  Replaces CSV_DIR.
C.TIMES          = [];          % runs with no cloud part: the time steps to write, one file
                                % each ([] = the T(t) table's times, or one step at t = 0)
C.CSV_FILES      = {};          % {} = every *.csv in the folder (natural-sorted),
                                % or an explicit cell list of file names (base
                                % names for multi-part: the same in every folder)
C.CSV_HAS_HEADER = true;        % first CSV row is a header

% --- units ---------------------------------------------------------------
%   Cloud coordinates are converted INTO the BDF's length units before mapping.
C.BDF_LENGTH_UNITS = 'in';      % 'in' | 'mm' | 'm'
C.CSV_LENGTH_UNITS = 'in';      % 'in' | 'mm' | 'm'
C.CSV_TEMP_UNITS   = 'K';       % 'K' | 'C' | 'F'   (what the CSV's 5th column is)

% --- output --------------------------------------------------------------
C.OUT_DIR        = 'temp_cards';    % created if missing
C.OUT_NAME       = 'temp_{index:03}';   % file name per step (no extension); tokens
                                    % {file} {time} {sid} {index}; {index:03} / {sid:04} zero-pad
C.OUT_UNITS      = 'K';             % 'K' | 'C' | 'F'  temperature units of the
                                    % structural model = units written on the TEMP cards
C.FIELD_SIZE     = 8;               % 8 = small field | 16 = large field
C.WRITE_TEMPD    = false;           % also emit TEMPD,SID,<mean T> for unlisted grids
C.WRITE          = true;            % false = map only, write nothing (GUI preview)
C.SAVE_PNG       = false;           % true = save an ISO-view PNG next to each .bdf
C.PNG_PLOT_ARGS  = {};              % extra temp_map_plot name/value pairs for those PNGs
                                    % (Style, CLim, Camera, Show*); the GUI passes its view
C.REPORT         = false;           % true = OUT_DIR/temp_map_report.html: settings,
                                    % per-step table, coverage, warnings, min/max chart,
                                    % and the PNGs when SAVE_PNG is on
C.PARALLEL       = false;           % true = map the steps of each part in a parfor
                                    % (Parallel Computing Toolbox). The first step of a
                                    % part builds the mapping; the rest reuse it.

% --- case control (optional, one file for the whole batch) ---------------------
C.WRITE_CASE     = true;            % write OUT_DIR/temp_subcases.dat + temp_includes.bdf
C.SUBCASE_OFFSET = 0;               % SUBCASE id = SID + this
C.SUBTITLE       = 't = {time} s';   % tokens: {file} {time} {sid} {index}; '' = none
C.LABEL          = '';              % same tokens; '' = none
C.CASE_EXTRA     = {};              % global case-control lines written once above the
                                    % first SUBCASE, e.g. {'SPC = 1', 'DISP(PLOT) = ALL'}
C.TREF           = [];              % reference temperature in OUT_UNITS: writes a TEMPD
                                    % card (SID = TREF_SID) and a global
                                    % TEMPERATURE(INITIAL) above the subcases.  [] = none
C.TREF_SID       = 99999;

% --- load set IDs --------------------------------------------------------
C.SID_START      = 1;           % SID = SID_START + (file index - 1) ...
C.SID_FROM_TIME  = false;       % ... or SID = round(time * TIME_SCALE)
C.TIME_SCALE     = 1;

% --- mapping -------------------------------------------------------------
C.METHOD         = 'linear';    % 'linear' | 'scattered' | 'nearest' | 'idw'   (see help)
C.IDW_K          = 8;           % neighbours used by 'idw'
C.IDW_POWER      = 2;           % 1/d^p weighting for 'idw'
C.SHELL_AVERAGE  = true;        % grids on shell elements (CQUAD4/CTRIA3/...): average the cloud
                                % through the plate thickness (PSHELL T / PCOMP total) along the
                                % grid normal instead of taking the mid-plane value
C.SHELL_LAYERS   = 5;           % sample levels across the thickness for that average
C.SLIVER_FACTOR  = 3;           % linear only: a grid whose Delaunay cell has an edge
                                % longer than this x the cloud's median spacing is in a
                                % "bridging" cell across a concavity -> nearest point is
                                % used instead of interpolating.  [] = off

% --- cloud surface for the check plot --------------------------------------
C.SURFACE_POINTS = 5e4;         % alpha-shape skin of the cloud from at most this
                                % many (random) cloud points; 0 = skip

% --- checks --------------------------------------------------------------
C.EXTRAP_WARN_DIST = [];        % [] = off; else warn when a grid's nearest cloud
                                % point is farther than this (model length units)
                                % and flag those grids in R(k).far
C.OVERHANG_WARN  = 0.10;        % loud warning when the mesh sticks out past the cloud
                                % by more than this fraction of the cloud's extent on
                                % any axis, or when >20% of grids are outside the hull

% --- GUI hooks -------------------------------------------------------------
C.RESULTS        = [];          % pass a previous R here to write it without
                                % re-mapping (SIDs stay as assigned at mapping time)
C.PROGRESS       = [];          % @(fraction, message) called at each stage
C.GRIDS          = [];          % grids struct from a previous READ_ONLY call:
                                % skips re-parsing the BDF (1M grids = seconds saved)
C.READ_ONLY      = false;       % true = parse the BDF and return the grids struct only:
                                %   G = temp_map_matlab('BDF_FILE', f, 'READ_ONLY', true)

% =========================================================================
%                               MACHINERY
% =========================================================================
C = apply_overrides(C, varargin);
validate_config(C);

parts = C.PARTS;
if isempty(parts), parts = {'', C.BDF_FILE, single_source(C)}; end   % single part = today's flow
np = size(parts, 1);

if isempty(C.RESULTS)
    % ---- grids per part ----------------------------------------------------------
    if ~isempty(C.GRIDS) && ~C.READ_ONLY
        G = C.GRIDS(:)';
        if numel(G) ~= np
            error('temp_map_matlab:badGrids', 'C.GRIDS has %d part(s) but PARTS has %d.', numel(G), np);
        end
        for i = 1:np
            G(i).csv_dir = parts{i, 3};      % folders may change between runs; grids do not
            if ~isfield(G, 'shell_t') || isempty(G(i).shell_t)          % cache from an older version
                G(i).shell_n = nan(numel(G(i).ids), 3); G(i).shell_t = nan(numel(G(i).ids), 1);
            end
            fprintf('Using %d cached grids from %s\n', numel(G(i).ids), G(i).bdf_file);
        end
    else
        G = repmat(struct('name', '', 'ids', [], 'xyz', [], 'faces', [], 'shell_n', [], 'shell_t', [], ...
                          'bdf_file', '', 'csv_dir', ''), 1, np);
        for i = 1:np
            lo = (i - 1) / np; w = 1 / np;
            tic;
            gstage = @(f, m) progress(C, lo + w * 0.45 * f, sprintf('%sGRIDs: %s', pfx(parts{i, 1}), m));
            [gid, gxyz] = read_grids(parts{i, 2}, gstage);
            fprintf('Read %d grids from %s  (%.1f s)\n', numel(gid), parts{i, 2}, toc);
            tic;
            estage = @(f, m) progress(C, lo + w * (0.45 + 0.50 * f), sprintf('%sElements: %s', pfx(parts{i, 1}), m));
            [gfaces, sh] = read_faces(parts{i, 2}, gid, gxyz, estage);
            fprintf('Read %d drawable element faces  (%.1f s)\n', size(gfaces, 1), toc);
            nsh = nnz(~isnan(sh.t)); nno = nnz(any(~isnan(sh.n), 2) & isnan(sh.t));
            if nsh + nno > 0
                fprintf('  %d grids on shell elements: thickness from PSHELL / PCOMP%s\n', nsh + nno, ...
                        ternary(nno > 0, sprintf(' -- %d WITHOUT a thickness (no property card): mid-plane value used', nno), ''));
            end
            G(i) = struct('name', parts{i, 1}, 'ids', gid, 'xyz', gxyz, 'faces', gfaces, ...
                          'shell_n', sh.n, 'shell_t', sh.t, 'bdf_file', parts{i, 2}, 'csv_dir', parts{i, 3});
        end
    end
    if C.READ_ONLY
        R = G; S = []; RM = [];
        progress(C, 1, 'Done.');
        return
    end
    if np > 1
        allids = vertcat(G.ids);
        ndup = numel(allids) - numel(unique(allids));
        if ndup > 0
            warning('temp_map_matlab:dupGrids', ...
                '%d grid id(s) appear in more than one part -- Nastran will reject the merged deck.', ndup);
        end
    end

    % ---- time steps: the first CLOUD part defines the list, every cloud part must
    %      have the same names; uniform (isothermal / T(t)) parts follow along -------
    kinds = cell(1, np); srcs = cell(1, np);
    for i = 1:np, [kinds{i}, srcs{i}] = part_source(G(i).csv_dir); end
    ic = find(strcmp(kinds, 'cloud'));
    iu = find(~strcmp(kinds, 'cloud'));
    files = cell(1, np);
    if ~isempty(ic)
        C1 = C; C1.CSV_DIR = srcs{ic(1)};
        files{ic(1)} = resolve_csv_files(C1);
        steps = cellfun(@shortname, files{ic(1)}, 'UniformOutput', false);
        ns = numel(steps);
        keys1 = step_keys(steps);
        for i = ic(2:end)
            d = dir(fullfile(srcs{i}, '*.csv'));
            names = {d.name};
            keysi = step_keys(names);
            [tf, loc] = ismember(keys1, keysi);
            if ~all(tf)
                miss = steps(~tf);
                error('temp_map_matlab:missingStep', ...
                    ['Part "%s" (%s) has no CSV matching %d of the %d time steps, e.g. %s\n' ...
                     '(steps pair across parts by the number before .csv, e.g. wing_20.csv <-> fuse_20.csv, ' ...
                     'or by the whole name when there is no trailing number).'], ...
                    G(i).name, srcs{i}, numel(miss), ns, strjoin(miss(1:min(3, end)), ', '));
            end
            [~, first] = unique(keysi, 'stable');
            amb = keys1(ismember(keys1, keysi(setdiff(1:numel(keysi), first))));
            if ~isempty(amb)
                error('temp_map_matlab:ambiguousStep', ...
                    'Part "%s" (%s) has more than one CSV ending in the same number, e.g. _%s.csv', ...
                    G(i).name, srcs{i}, amb{1});
            end
            files{i} = fullfile(srcs{i}, names(loc));
            files{i} = files{i}(:);
        end
    else
        times = uniform_times(C, kinds, srcs);
        ns = numel(times);
        if ns == 1 && isempty(C.TIMES) && ~any(strcmp(kinds, 'table'))
            steps = {'isothermal'};
        else
            steps = arrayfun(@(t) sprintf('t%g', t), times, 'UniformOutput', false);
        end
    end

    % ---- map every part at every step ------------------------------------------------
    %   Part-major: step 1 of a part builds the mapping (the slow geometry),
    %   later steps of the same cloud reuse it -- serially, or in a parfor
    %   with PARALLEL (the CSV read is what parallelises).
    R = repmat(empty_result(), np, ns);
    total = numel(ic) * ns;
    use_par = C.PARALLEL && ns > 1 && parallel_ok();
    for ci = 1:numel(ic)
        i = ic(ci);
        Gi = G(i);
        cache = struct('xyz', [], 'M', [], 'surface', []);
        fi = files{i};
        % step 1 always serial: it builds the cache
        st = @(k) @(frac, msg) progress(C, 0.05 + 0.95 * (((ci - 1) * ns + k - 1) + frac) / total, ...
                                        sprintf('[%d/%d] %s%s: %s', k, ns, pfx(Gi.name), steps{k}, msg));
        tic;
        [q, cache] = map_step(fi{1}, C, Gi, cache, st(1));
        report_step(q, Gi, steps{1}, toc);
        R(i, 1) = assemble_result(q, C, Gi, fi{1}, 1, cache);
        if ns > 1
            if use_par
                Cw = C; Cw.PROGRESS = []; Cw.RESULTS = []; Cw.GRIDS = [];   % nothing GUI-bound goes to workers
                dq = parallel.pool.DataQueue;
                cnt = containers.Map('KeyType', 'char', 'ValueType', 'double'); cnt('done') = 0;
                afterEach(dq, @(k) par_progress(C, cnt, ci, k, ns, total, Gi.name, steps{k}));
                t0 = tic;
                Qp = cell(1, ns - 1);
                parfor j = 1:ns - 1
                    Qp{j} = map_step(fi{j + 1}, Cw, Gi, cache, @(~, ~) []);
                    send(dq, j + 1);
                end
                fprintf('  %s%d steps mapped in parallel  (%.1f s)\n', pfx(Gi.name), ns - 1, toc(t0));
                for k = 2:ns
                    report_step(Qp{k - 1}, Gi, steps{k}, NaN);
                    R(i, k) = assemble_result(Qp{k - 1}, C, Gi, fi{k}, k, cache);
                end
                clear Qp
            else
                for k = 2:ns
                    tic;
                    [q, cache] = map_step(fi{k}, C, Gi, cache, st(k));
                    report_step(q, Gi, steps{k}, toc);
                    R(i, k) = assemble_result(q, C, Gi, fi{k}, k, cache);
                end
            end
        end
    end
    % uniform parts: the cloud steps' times when there are clouds, else the synthesized list
    for i = iu
        for k = 1:ns
            if ~isempty(ic), t = R(ic(1), k).time; else, t = times(k); end
            R(i, k) = uniform_result(G(i), kinds{i}, srcs{i}, t, steps{k}, k, C);
        end
    end
else
    R = C.RESULTS;                       % SIDs were assigned when these were mapped
    if isvector(R), R = R(:)'; end       % 1 x nSteps for a single part
end
RM = merge_parts(R);


if C.WRITE
    if exist(C.OUT_DIR, 'dir') ~= 7, mkdir(C.OUT_DIR); end
    ns = size(R, 2);
    bases = arrayfun(@(k) out_base(R(1, k), C, k), 1:ns, 'UniformOutput', false);
    if numel(unique(bases)) < ns
        error('temp_map_matlab:dupOutName', ...
              'OUT_NAME ''%s'' gives the same file name for more than one step -- add {index} or {sid}.', C.OUT_NAME);
    end
    wspan = ternary(C.SAVE_PNG, 0.5, 0.85);         % TEMP files, then PNGs, then case / report
    for k = 1:ns
        tic;
        base = bases{k};
        [R(:, k).out_file] = deal(fullfile(C.OUT_DIR, [base '.bdf']));
        progress(C, wspan * (k - 1) / ns, sprintf('writing TEMP cards %d / %d: %s.bdf ...', k, ns, base));
        write_temp_cards(R(:, k), C);
        fprintf('  wrote %s  (%.1f s)\n', R(1, k).out_file, toc);
    end
    if C.SAVE_PNG
        for k = 1:ns
            base = bases{k};
            progress(C, wspan + (0.85 - wspan) * (k - 1) / ns, sprintf('check plot %d / %d: %s.png ...', k, ns, base));
            h = temp_map_plot(RM(k), 'Visible', 'off', 'Units', C.OUT_UNITS, 'Buttons', false, C.PNG_PLOT_ARGS{:});
            png = fullfile(C.OUT_DIR, [base '.png']);
            try
                exportgraphics(h.ax, png, 'Resolution', 150);      % R2020a+
            catch
                print(h.fig, png, '-dpng', '-r150');
            end
            close(h.fig);
        end
    end
end

if C.WRITE && C.WRITE_CASE && ~isempty(R)
    progress(C, 0.88, 'writing case control + includes ...');
    write_case_control(R, C);
end

S = summary_table(R, C);
if C.WRITE && C.REPORT && ~isempty(R)
    progress(C, 0.94, 'writing HTML report ...');
    write_report(R, RM, S, C);
end
progress(C, 1, 'Done.');
if nargout == 0
    disp(S);
    clear R S RM
end
end


% =========================================================================
function [q, cache] = map_step(file, C, Gi, cache, stage)
%MAP_STEP  Read one CSV and map it onto part Gi.  Reuses cache when the cloud
%   points are unchanged, else builds a new mapping (returned in cache so the
%   caller can keep it for the next step).  q is a slim per-step result.
    stage(0, 'reading CSV ...');
    [t, cxyz, cT] = read_cloud(file, C.CSV_HAS_HEADER);
    cxyz = cxyz * length_factor(C.CSV_LENGTH_UNITS, C.BDF_LENGTH_UNITS);
    cT   = to_kelvin(cT, C.CSV_TEMP_UNITS);
    q = struct('t', t, 'cT', cT, 'gT', [], 'extrap', [], 'nndist', [], 'far', [], ...
               'method', '', 'reused', false, 'cxyz', [], 'srf', []);
    if same_cloud(cache.xyz, cxyz)
        stage(0.5, 'same cloud locations as the previous case: reusing the mapping ...');
        M = cache.M;
        q.reused = true;
    else
        [Q, sidx, L] = shell_samples(Gi, C);
        M = build_map(cxyz, Q, C, stage);
        n = size(Gi.xyz, 1);
        M.extrap = M.extrap(1:n); M.nndist = M.nndist(1:n);      % flags belong to the grid itself
        if ~isempty(sidx)
            M.method = sprintf('%s + through-thickness average on %d shell grids (%d layers)', M.method, numel(sidx), L);
        end
        stage(0.95, 'cloud surface (alpha shape) ...');
        srf = cloud_surface(cxyz, C.SURFACE_POINTS);
        cache = struct('xyz', cxyz, 'M', M, 'surface', srf);
        q.cxyz = cxyz; q.srf = srf;
    end
    q.gT     = collapse_layers(apply_map(M, cT), Gi, C);
    q.extrap = M.extrap;
    q.nndist = M.nndist;
    q.method = M.method;
    if q.reused, q.method = [q.method ' (reused)']; end
    q.far = false(size(q.nndist));
    if ~isempty(C.EXTRAP_WARN_DIST)
        q.far = q.nndist > C.EXTRAP_WARN_DIST;
    end
end


function r = assemble_result(q, C, Gi, file, k, cache)
%ASSEMBLE_RESULT  Full per-step result from the slim map_step output.  Shared
%   arrays (grids, a reused cloud) are assigned from one variable so MATLAB
%   keeps a single copy in memory across all the steps.
    r = empty_result();
    r.part      = Gi.name;
    r.csv_file  = file;
    r.bdf_file  = Gi.bdf_file;
    r.time      = q.t;
    r.sid       = assign_sid(C, k, q.t);
    r.grid_ids  = Gi.ids;
    r.grid_xyz  = Gi.xyz;
    r.faces     = Gi.faces;
    r.grid_T    = q.gT;                 % Kelvin, always
    if q.reused, r.cloud_xyz = cache.xyz; r.surface = cache.surface;
    else,        r.cloud_xyz = q.cxyz;   r.surface = q.srf; end
    r.cloud_T   = q.cT;                 % Kelvin, always
    r.extrap    = q.extrap;
    r.nn_dist   = q.nndist;
    r.far       = q.far;
    r.method    = q.method;
    r.coverage  = coverage_report(r);
    r.warnings  = coverage_warnings(r, C);
    if any(r.far)
        warning('temp_map_matlab:farGrids', ...
            '%s%s: %d grid(s) are farther than %g from any cloud point (max %.4g).', ...
            pfx(Gi.name), shortname(file), nnz(r.far), C.EXTRAP_WARN_DIST, max(r.nn_dist));
    end
    if ~isempty(r.warnings)
        fprintf('\n  ********** COVERAGE WARNING: %s%s **********\n', pfx(Gi.name), shortname(file));
        fprintf('  ** %s\n', r.warnings{:});
        fprintf('  ***********************************************************\n\n');
    end
end


function report_step(q, Gi, step, secs)
    if isnan(secs), tstr = ''; else, tstr = sprintf('  (%.1f s)', secs); end
    fprintf('  %-32s t=%-9.4g cloud=%-8d outside=%-6d far=%-6d T=[%.2f %.2f] K  %s%s\n', ...
        [pfx(Gi.name) step], q.t, numel(q.cT), nnz(q.extrap), nnz(q.far), min(q.gT), max(q.gT), q.method, tstr);
end


function ok = parallel_ok()
%PARALLEL_OK  Parallel Computing Toolbox present and a pool available?
    ok = license('test', 'Distrib_Computing_Toolbox') && exist('parpool', 'file') == 2;
    if ~ok
        warning('temp_map_matlab:noParallel', 'PARALLEL requested but the Parallel Computing Toolbox is not available; running serially.');
        return
    end
    try
        if isempty(gcp('nocreate')), parpool; end
    catch ME
        warning('temp_map_matlab:noPool', 'Could not start a parallel pool (%s); running serially.', ME.message);
        ok = false;
    end
end


function par_progress(C, cnt, i, k, ns, total, name, step) %#ok<INUSL>
    cnt('done') = cnt('done') + 1;
    progress(C, 0.05 + 0.95 * (((i - 1) * ns + cnt('done') + 1) / total), ...
             sprintf('%smapping steps in parallel  (%d of %d done, latest %s)', pfx(name), cnt('done'), ns - 1, step)); %#ok<NASGU>
end


% =========================================================================
function s = pfx(name)
%PFX  'part / ' prefix for messages, or '' for the single-part case.
    if isempty(name), s = ''; else, s = [char(name) ' / ']; end
end


% =========================================================================
function RM = merge_parts(R)
%MERGE_PARTS  One assembly-wide result per time step from the (parts x steps)
%   array: grids, faces (row offsets applied), clouds and coverage concatenated.
%   A single part comes back unchanged.
    [np, ns] = size(R);
    if np == 1, RM = R; return; end
    RM = repmat(empty_result(), 1, ns);
    for k = 1:ns
        rs = R(:, k);
        counts = arrayfun(@(r) numel(r.grid_ids), rs);
        offs = cumsum([0; counts(:)]);
        m = empty_result();
        m.part       = 'ALL';
        m.part_names = {rs.part};
        m.csv_file   = rs(1).csv_file;
        m.bdf_file   = strjoin({rs.bdf_file}, ' + ');
        m.time       = rs(1).time;
        m.sid        = rs(1).sid;
        m.grid_ids   = vertcat(rs.grid_ids);
        m.grid_xyz   = vertcat(rs.grid_xyz);
        m.grid_T     = vertcat(rs.grid_T);
        m.extrap     = vertcat(rs.extrap);
        m.nn_dist    = vertcat(rs.nn_dist);
        m.far        = vertcat(rs.far);
        m.grid_part  = repelem((1:np)', counts(:));
        F = zeros(0, 4);
        for i = 1:np
            if ~isempty(rs(i).faces), F = [F; rs(i).faces + offs(i)]; end   %#ok<AGROW>  NaN stays NaN
        end
        m.faces      = F;
        m.cloud_xyz  = vertcat(rs.cloud_xyz);
        m.cloud_T    = vertcat(rs.cloud_T);
        m.surface    = {rs.surface};
        m.method     = strjoin(unique({rs.method}, 'stable'), ' | ');
        cov = {}; wrn = {};
        for i = 1:np
            cov = [cov; {sprintf('--- %s ---', rs(i).part)}; rs(i).coverage(:)];        %#ok<AGROW>
            if ~isempty(rs(i).warnings)
                wrn = [wrn; strcat([rs(i).part ': '], rs(i).warnings(:))];             %#ok<AGROW>
            end
        end
        m.coverage   = cov;
        m.warnings   = wrn;
        RM(k) = m;
    end
end


% =========================================================================
function C = apply_overrides(C, args)
%APPLY_OVERRIDES  Fold name/value pairs into the CONFIG struct.
    if isempty(args), return; end
    if mod(numel(args), 2) ~= 0
        error('temp_map_matlab:badArgs', ...
              'Overrides must be name/value pairs, e.g. ''OUT_UNITS'', ''C''.');
    end
    for k = 1:2:numel(args)
        name = args{k};
        if ~ischar(name) && ~isstring(name)
            error('temp_map_matlab:badArgs', 'Override name must be a string.');
        end
        name = char(name);
        if ~isfield(C, name)
            error('temp_map_matlab:unknownOption', ...
                  'Unknown option "%s". Valid names are the CONFIG field names.', name);
        end
        C.(name) = args{k + 1};
    end
end


% =========================================================================
function validate_config(C)
%VALIDATE_CONFIG  Fail early and readably.
    if isempty(C.RESULTS)
        parts = C.PARTS;
        if isempty(parts), parts = {'', C.BDF_FILE, single_source(C)}; end
        if ~iscell(parts) || size(parts, 2) ~= 3
            error('temp_map_matlab:badParts', ...
                'C.PARTS must be an N x 3 cell: {name, bdf_file, cloud folder | temperature | T(t) csv}.');
        end
        if ~isempty(C.GRIDS) && ~all(isfield(C.GRIDS, {'ids', 'xyz', 'faces', 'bdf_file'}))
            error('temp_map_matlab:badGrids', 'C.GRIDS must come from a READ_ONLY call.');
        end
        for i = 1:size(parts, 1)
            if isempty(C.GRIDS) && exist(parts{i, 2}, 'file') ~= 2
                error('temp_map_matlab:noFile', '%sBDF file not found: %s', pfx(parts{i, 1}), parts{i, 2});
            end
            if size(parts, 1) > 1 && isempty(parts{i, 1})
                error('temp_map_matlab:badParts', 'Every part needs a name (row %d).', i);
            end
        end
        if C.READ_ONLY, return; end
        kinds = cell(1, size(parts, 1));
        for i = 1:size(parts, 1)
            [kinds{i}, src] = part_source(parts{i, 3});
            first_cloud = strcmp(kinds{i}, 'cloud') && ~any(strcmp(kinds(1:i-1), 'cloud'));
            if strcmp(kinds{i}, 'cloud') && ~(first_cloud && ~isempty(C.CSV_FILES)) && exist(src, 'dir') ~= 7
                error('temp_map_matlab:noDir', ...
                    '%ssource "%s" is not a cloud folder, a T(t) CSV file or a temperature.', pfx(parts{i, 1}), src);
            end
        end
        if ~any(strcmp(kinds, 'cloud')) && ~isempty(C.TIMES) && (~isnumeric(C.TIMES) || ~isvector(C.TIMES))
            error('temp_map_matlab:badTimes', 'C.TIMES must be a numeric vector of time steps.');
        end
    end
    length_factor(C.CSV_LENGTH_UNITS, C.BDF_LENGTH_UNITS);   % errors on a bad name
    to_kelvin(0, C.CSV_TEMP_UNITS);
    if ~ismember(upper(char(C.OUT_UNITS)), {'K', 'C', 'F'})
        error('temp_map_matlab:badUnits', 'C.OUT_UNITS must be ''K'', ''C'' or ''F''.');
    end
    if ~ismember(C.FIELD_SIZE, [8 16])
        error('temp_map_matlab:badField', 'C.FIELD_SIZE must be 8 or 16.');
    end
    if ~ismember(lower(char(C.METHOD)), {'linear', 'scattered', 'nearest', 'idw'})
        error('temp_map_matlab:badMethod', 'C.METHOD must be ''linear'', ''scattered'', ''nearest'' or ''idw''.');
    end
    if ~isnumeric(C.SHELL_LAYERS) || ~isscalar(C.SHELL_LAYERS) || C.SHELL_LAYERS < 2
        error('temp_map_matlab:badLayers', 'C.SHELL_LAYERS must be an integer >= 2.');
    end
    if C.SID_FROM_TIME && C.TIME_SCALE <= 0
        error('temp_map_matlab:badScale', 'C.TIME_SCALE must be positive.');
    end
end


% =========================================================================
function progress(C, frac, msg)
%PROGRESS  Forward a stage update to C.PROGRESS, if one was given.
    if ~isempty(C.PROGRESS)
        C.PROGRESS(min(max(frac, 0), 1), msg);
    end
end


% =========================================================================
function r = empty_result()
    r = struct('part', '', 'part_names', {{}}, 'grid_part', [], ...
               'csv_file', '', 'bdf_file', '', 'time', NaN, 'sid', NaN, ...
               'grid_ids', [], 'grid_xyz', [], 'grid_T', [], ...
               'cloud_xyz', [], 'cloud_T', [], 'extrap', [], 'nn_dist', [], ...
               'far', [], 'method', '', 'surface', [], 'coverage', {{}}, 'warnings', {{}}, ...
               'faces', [], 'out_file', '');
end


% =========================================================================
function sid = assign_sid(C, k, t)
    if C.SID_FROM_TIME
        sid = round(t * C.TIME_SCALE);
    else
        sid = C.SID_START + k - 1;
    end
    if sid < 1
        error('temp_map_matlab:badSid', 'Computed SID %d < 1 (file %d, t=%g).', sid, k, t);
    end
end


% =========================================================================
function src = single_source(C)
%SINGLE_SOURCE  Column 3 of the implicit single part: ISOTHERMAL when set, else CSV_DIR.
    if isempty(C.ISOTHERMAL), src = C.CSV_DIR; else, src = C.ISOTHERMAL; end
end


function [kind, src] = part_source(spec)
%PART_SOURCE  What column 3 of PARTS means.
%   'const'  a number (temperature in OUT_UNITS)         src = the number
%   'table'  an existing file: 2-column CSV (time, T)    src = the path
%   'cloud'  anything else: a folder of cloud CSVs       src = the path
    if isnumeric(spec) && isscalar(spec)
        kind = 'const'; src = double(spec); return
    end
    s = strtrim(char(spec));
    v = str2double(s);
    if ~isnan(v)
        kind = 'const'; src = v;
    elseif exist(s, 'dir') == 7
        kind = 'cloud'; src = s;
    elseif exist(s, 'file') == 2
        kind = 'table'; src = s;
    else
        kind = 'cloud'; src = s;
    end
end


function times = uniform_times(C, kinds, srcs)
%UNIFORM_TIMES  Time steps for a run with no cloud part: TIMES, else the first
%   T(t) table's times, else a single step at t = 0.
    if ~isempty(C.TIMES)
        times = double(C.TIMES(:)');
        return
    end
    it = find(strcmp(kinds, 'table'), 1);
    if isempty(it), times = 0; return; end
    times = read_temperature_table(srcs{it});
    times = times(:)';
end


function [tt, TT] = read_temperature_table(file)
%READ_TEMPERATURE_TABLE  (time, T) columns of a 2-column CSV; header rows dropped.
    M = readmatrix(file);
    if size(M, 2) < 2
        error('temp_map_matlab:badTable', '%s: expected 2 columns (time, T) but found %d.', file, size(M, 2));
    end
    M = M(:, 1:2);
    M(any(isnan(M), 2), :) = [];
    if isempty(M)
        error('temp_map_matlab:emptyTable', '%s: no numeric rows.', file);
    end
    [tt, order] = sort(M(:, 1));
    TT = M(order, 2);
    [tt, iu] = unique(tt, 'stable');
    TT = TT(iu);
end


function T = table_temperature(file, t)
%TABLE_TEMPERATURE  Uniform temperature at time t, linear in the table; clamped
%   to the end values outside it (with a warning).
    [tt, TT] = read_temperature_table(file);
    if numel(tt) == 1, T = TT; return; end
    if t < tt(1) - 1e-9 || t > tt(end) + 1e-9
        warning('temp_map_matlab:tableRange', ...
            '%s: t = %g is outside the table (%g .. %g); the end value is used.', shortname(file), t, tt(1), tt(end));
    end
    T = interp1(tt, TT, min(max(t, tt(1)), tt(end)), 'linear');
end


function r = uniform_result(Gi, kind, src, t, step, k, C)
%UNIFORM_RESULT  One part at one uniform temperature, no cloud: a constant, or
%   T(t) read from a table at the step's time.
    units = upper(char(C.OUT_UNITS));
    if strcmp(kind, 'const')
        Tout = src;
        how = sprintf('isothermal %g %s', Tout, units);
    else
        Tout = table_temperature(src, t);
        how = sprintf('isothermal %g %s from %s', Tout, units, shortname(src));
    end
    n = numel(Gi.ids);
    r = empty_result();
    r.part      = Gi.name;
    r.csv_file  = step;                  % names the output file like a cloud step would
    r.bdf_file  = Gi.bdf_file;
    r.time      = t;
    r.sid       = assign_sid(C, k, t);
    r.grid_ids  = Gi.ids;
    r.grid_xyz  = Gi.xyz;
    r.faces     = Gi.faces;
    r.grid_T    = repmat(to_kelvin(Tout, units), n, 1);
    r.cloud_xyz = zeros(0, 3);
    r.cloud_T   = zeros(0, 1);
    r.extrap    = false(n, 1);
    r.nn_dist   = zeros(n, 1);
    r.far       = false(n, 1);
    r.method    = how;
    r.coverage  = {sprintf('%s: all %d grids, no cloud', how, n)};
    fprintf('  %s%s: %s  (%d grids)\n', pfx(Gi.name), step, how, n);
end


function keys = step_keys(names)
%STEP_KEYS  What pairs a time step across parts: the number just before .csv
%   (leading zeros ignored: wing_020.csv and fuse_20.csv are the same step),
%   or the whole lower-case name when there is no trailing number.
    names = cellstr(names);
    keys = cell(size(names));
    for k = 1:numel(names)
        tok = regexp(names{k}, '(\d+)\.csv$', 'tokens', 'once', 'ignorecase');
        if isempty(tok)
            keys{k} = lower(names{k});
        else
            keys{k} = sprintf('%d', str2double(tok{1}));
        end
    end
end


function files = resolve_csv_files(C)
%RESOLVE_CSV_FILES  Explicit list, or every *.csv in CSV_DIR in natural order.
    if ~isempty(C.CSV_FILES)
        files = cellstr(C.CSV_FILES);
        files = files(:);
        for k = 1:numel(files)
            if exist(files{k}, 'file') ~= 2
                cand = fullfile(C.CSV_DIR, files{k});
                if exist(cand, 'file') == 2
                    files{k} = cand;
                else
                    error('temp_map_matlab:noCsv', 'CSV not found: %s', files{k});
                end
            end
        end
        return
    end
    d = dir(fullfile(C.CSV_DIR, '*.csv'));
    if isempty(d)
        error('temp_map_matlab:noCsv', 'No *.csv files in %s', C.CSV_DIR);
    end
    names = {d.name};
    % natural sort: zero-pad every digit run so t2 < t10
    keys = regexprep(names, '(\d+)', '${sprintf(''%012d'', str2double($1))}');
    [~, order] = sort(lower(keys));
    files = fullfile(C.CSV_DIR, names(order));
    files = files(:);
end


% =========================================================================
function [ids, xyz] = read_grids(bdf_file, stage)
%READ_GRIDS  GRID id and basic-frame xyz from a BDF, following INCLUDEs.
%   Handles small-field, large-field (GRID*) and free-field (comma) cards.
    if nargin < 2, stage = @(~, ~) []; end
    [ids, xyz, ncp] = read_grids_file(bdf_file, true, stage, [0 1]);
    stage(1, sprintf('%d grids read', numel(ids)));
    if isempty(ids)
        error('temp_map_matlab:noGrids', 'No GRID cards found in %s', bdf_file);
    end
    [ids, iu] = unique(ids, 'stable');
    xyz = xyz(iu, :);
    if ncp > 0
        warning('temp_map_matlab:nonBasicCP', ...
            '%d GRID(s) have CP <> 0. Their coordinates are taken AS WRITTEN (no transform).', ncp);
    end
end


function [ids, xyz, ncp] = read_grids_file(fname, is_top, stage, span)
%   Vectorised: the whole deck is split into lines once, GRID lines are picked
%   out with one regexp, and small/large-field cards are sliced as char
%   matrices.  Only free-field GRIDs and INCLUDEs are handled line by line.
%   stage(frac, msg) reports progress; span = [lo hi] is this file's share.
    sub = @(f, m) stage(span(1) + f * (span(2) - span(1)), sprintf('%s: %s', shortname(fname), m));
    sub(0.00, 'reading file ...');
    txt = fileread(fname);
    txt = strrep(txt, sprintf('\t'), ' ');
    sub(0.15, 'splitting lines ...');
    lines = splitlines(string(txt));
    base_dir = fileparts(fname);
    n = numel(lines);
    up = upper(strip(lines, 'left'));

    first = 1;
    if is_top
        ib = find(startsWith(up, "BEGIN"), 1);
        if ~isempty(ib) && ~isempty(regexp(char(up(ib)), '^BEGIN\s+BULK', 'once')), first = ib + 1; end
    end
    last = find(startsWith(up, "ENDDATA"), 1);
    if isempty(last), last = n; else, last = last - 1; end
    lines = lines(first:last);
    up    = up(first:last);
    n     = numel(lines);

    ids = zeros(0, 1); xyz = zeros(0, 3); ncp = 0; lineno = zeros(0, 1);

    % ---- GRID lines --------------------------------------------------------
    sub(0.35, sprintf('scanning %d lines for GRID cards ...', n));
    isg = ~cellfun('isempty', regexp(cellstr(up), '^GRID\*?\s*(,|\s|$)', 'once'));
    sub(0.55, sprintf('parsing %d GRID cards ...', nnz(isg)));
    gi  = find(isg);
    g   = lines(gi);
    isfree  = contains(g, ",");
    islarge = ~isfree & startsWith(up(gi), "GRID*");
    issmall = ~isfree & ~islarge;

    if any(issmall)
        M = pad_cols(char(g(issmall)), 48);
        [id, cp, x] = parse_fields(M(:, 9:16), M(:, 17:24), M(:, 25:32), M(:, 33:40), M(:, 41:48));
        ids = [ids; id]; xyz = [xyz; x]; ncp = ncp + cp; lineno = [lineno; gi(issmall)];
    end
    if any(islarge)
        li = gi(islarge);
        L1 = pad_cols(char(lines(li)), 72);
        L2 = pad_cols(char(lines(min(li + 1, n))), 72);   % continuation = next line
        [id, cp, x] = parse_fields(L1(:, 9:24), L1(:, 25:40), L1(:, 41:56), L1(:, 57:72), L2(:, 9:24));
        ids = [ids; id]; xyz = [xyz; x]; ncp = ncp + cp; lineno = [lineno; li];
    end
    if any(isfree)
        fi = gi(isfree);
        for i = fi(:)'
            f = strtrim(strsplit(char(lines(i)), ',', 'CollapseDelimiters', false));
            if isempty(f{end}), f(end) = []; end   % trailing comma = continues
            j = i + 1;
            while numel(f) < 6 && j <= n
                nxt = char(lines(j));
                if isempty(nxt) || nxt(1) == '$', break; end
                f2 = strtrim(strsplit(nxt, ',', 'CollapseDelimiters', false));
                if ~isempty(f2) && (isempty(f2{1}) || f2{1}(1) == '+' || f2{1}(1) == '*')
                    f2 = f2(2:end);               % drop continuation marker
                end
                f = [f f2];                       %#ok<AGROW>
                j = j + 1;
            end
            f(end+1:6) = {''};
            [id, cp, x] = parse_fields(f{2}, f{3}, f{4}, f{5}, f{6});
            ids = [ids; id]; xyz = [xyz; x]; ncp = ncp + cp; lineno = [lineno; i]; %#ok<AGROW>
        end
    end
    [~, order] = sort(lineno);      % keep deck order
    ids = ids(order); xyz = xyz(order, :);

    % ---- INCLUDE 'file'  (quoted path may span lines) ----------------------
    incs = find(startsWith(up, "INCLUDE"))';
    for q = 1:numel(incs)
        i = incs(q);
        rest = regexprep(char(lines(i)), '^\s*INCLUDE\s*', '', 'ignorecase');
        j = i + 1;
        while nnz(rest == '''') < 2 && j <= n   % unclosed quote: join next line
            rest = [rest strtrim(char(lines(j)))];    %#ok<AGROW>
            j = j + 1;
        end
        inc = strtrim(regexprep(strtrim(rest), '^''|''$', ''));
        % includes share the last 30% of this file's bar, evenly
        lo = span(1) + (0.70 + 0.30 * (q - 1) / numel(incs)) * (span(2) - span(1));
        hi = span(1) + (0.70 + 0.30 * q / numel(incs)) * (span(2) - span(1));
        [i2, x2, c2] = read_grids_file(resolve_include(inc, base_dir), false, stage, [lo hi]);
        ids = [ids; i2]; xyz = [xyz; x2]; ncp = ncp + c2;  %#ok<AGROW>
    end
end


function [id, ncp, xyz] = parse_fields(fid, fcp, fx, fy, fz)
%PARSE_FIELDS  Char-matrix (or single char) fields -> id, CP<>0 count, xyz.
    id  = fast_int(char(fid));
    cpv = fast_int(char(fcp));
    ncp = nnz(~isnan(cpv) & cpv ~= 0);
    xyz = [nas_real(cellstr(fx)) nas_real(cellstr(fy)) nas_real(cellstr(fz))];
    ok  = ~isnan(id);
    id  = id(ok); xyz = xyz(ok, :);
end


function p = resolve_include(inc, base_dir)
    inc = strrep(inc, '\', filesep);
    inc = strrep(inc, '/', filesep);
    cands = {fullfile(base_dir, inc), inc};
    for k = 1:numel(cands)
        if exist(cands{k}, 'file') == 2, p = cands{k}; return; end
    end
    error('temp_map_matlab:noInclude', 'INCLUDE file not found: %s (relative to %s)', inc, base_dir);
end


function M = pad_cols(M, w)
%PAD_COLS  Right-pad a char matrix with blanks to at least w columns.
    if isempty(M), M = repmat(' ', 0, w); end
    if size(M, 2) < w, M(:, end+1:w) = ' '; end
end


function v = nas_real(c)
%NAS_REAL  Parse Nastran reals: '1.5', '1.5E-3', '1.5-3', '.5', '1.', blank = 0.
%   c is a cellstr; returns a column.
    c = strtrim(c(:));
    blank = cellfun('isempty', c);
    c = regexprep(upper(c), 'D', 'E');
    c = regexprep(c, '^([+-]?[\d.]+)([+-]\d+)$', '$1E$2');   % 1.5-3 -> 1.5E-3
    v = str2double(c);
    v(blank) = 0;
    if any(isnan(v))
        bad = c(isnan(v));
        error('temp_map_matlab:badReal', 'Cannot parse "%s" as a real.', bad{1});
    end
end


% =========================================================================
function [F, SH] = read_faces(bdf_file, grid_ids, grid_xyz, stage)
%READ_FACES  Drawable faces (rows into grid_ids, NaN-padded to 4 columns)
%   from the shell and solid elements in the BDF: shells as-is, solids as
%   their free (outer) faces.  [] if the deck has no supported elements.
%   SH.n / SH.t: per-grid shell normal (unit, sign arbitrary) and thickness
%   (PSHELL T or PCOMP total), NaN where the grid is not on a shell / no property.
    if nargin < 4, stage = @(~, ~) []; end
    ng = numel(grid_ids);
    SH = struct('n', nan(ng, 3), 't', nan(ng, 1));
    try
        E = read_elements_file(bdf_file, true, stage, [0 0.8]);
    catch ME
        warning('temp_map_matlab:elements', 'Element read failed (%s); contour view unavailable.', ME.message);
        F = []; return
    end
    if isempty(E), F = []; return; end
    SH = shell_normals(E, grid_ids, grid_xyz);
    stage(0.85, 'building free faces of solids ...');
    faces = zeros(0, 4);
    % --- shells: one face each ---------------------------------------------
    for t = {'CQUAD4', 'CQUADR', 'CQUAD8'}
        g = E.(t{1});
        if ~isempty(g), faces = [faces; g(:, 1:4)]; end                 %#ok<AGROW>
    end
    for t = {'CTRIA3', 'CTRIAR', 'CTRIA6'}
        g = E.(t{1});
        if ~isempty(g), faces = [faces; g(:, 1:3) nan(size(g, 1), 1)]; end %#ok<AGROW>
    end
    % --- solids: free faces only ----------------------------------------------
    sf = zeros(0, 4);
    g = E.CTETRA;
    if ~isempty(g)
        sf = [sf; g(:, [1 2 3]) nan(size(g,1),1); g(:, [1 2 4]) nan(size(g,1),1); ...
                  g(:, [2 3 4]) nan(size(g,1),1); g(:, [1 3 4]) nan(size(g,1),1)];
    end
    g = E.CHEXA;
    if ~isempty(g)
        sf = [sf; g(:, [1 2 3 4]); g(:, [5 6 7 8]); g(:, [1 2 6 5]); ...
                  g(:, [2 3 7 6]); g(:, [3 4 8 7]); g(:, [4 1 5 8])];
    end
    g = E.CPENTA;
    if ~isempty(g)
        sf = [sf; g(:, [1 2 3]) nan(size(g,1),1); g(:, [4 5 6]) nan(size(g,1),1); ...
                  g(:, [1 2 5 4]); g(:, [2 3 6 5]); g(:, [3 1 4 6])];
    end
    if ~isempty(sf)
        key = sort(sf, 2);                     % NaN sorts last -> consistent key
        key(isnan(key)) = 0;
        [~, ~, ic] = unique(key, 'rows');
        cnt = accumarray(ic, 1);
        faces = [faces; sf(cnt(ic) == 1, :)];  % faces used by exactly one element
    end
    if isempty(faces), F = []; return; end
    stage(0.95, sprintf('indexing %d faces to grids ...', size(faces, 1)));
    % node ids -> rows of grid_ids; drop faces touching unknown grids
    [tf, loc] = ismember(faces, grid_ids);
    ok = all(tf | isnan(faces), 2);
    F = loc(ok, :);
    F(isnan(faces(ok, :))) = NaN;
    F = double(F);
end


function SH = shell_normals(E, grid_ids, gxyz)
%SHELL_NORMALS  Per-grid shell normal (dominant axis of the adjacent element
%   normals, so element orientation does not matter) and mean thickness.
    ng = numel(grid_ids);
    SH = struct('n', nan(ng, 3), 't', nan(ng, 1));
    props = [E.PSHELL; E.PCOMP];                       % [PID T]
    nodes = zeros(0, 4); enrm = zeros(0, 3); epid = zeros(0, 1);
    for t = {'CQUAD4', 'CQUADR', 'CQUAD8'}
        g = E.(t{1}); if isempty(g), continue; end
        [tf, loc] = ismember(g(:, 1:4), grid_ids); ok = all(tf, 2);
        loc = loc(ok, :);
        nrm = cross(gxyz(loc(:, 3), :) - gxyz(loc(:, 1), :), gxyz(loc(:, 4), :) - gxyz(loc(:, 2), :), 2);
        nodes = [nodes; loc]; enrm = [enrm; nrm]; epid = [epid; E.(['PID_' t{1}])(ok)];   %#ok<AGROW>
    end
    for t = {'CTRIA3', 'CTRIAR', 'CTRIA6'}
        g = E.(t{1}); if isempty(g), continue; end
        [tf, loc] = ismember(g(:, 1:3), grid_ids); ok = all(tf, 2);
        loc = loc(ok, :);
        nrm = cross(gxyz(loc(:, 2), :) - gxyz(loc(:, 1), :), gxyz(loc(:, 3), :) - gxyz(loc(:, 1), :), 2);
        nodes = [nodes; loc nan(size(loc, 1), 1)]; enrm = [enrm; nrm]; epid = [epid; E.(['PID_' t{1}])(ok)];   %#ok<AGROW>
    end
    if isempty(nodes), return; end
    enrm = enrm ./ max(vecnorm(enrm, 2, 2), eps);
    et = nan(size(epid));
    if ~isempty(props)
        [tf, loc] = ismember(epid, props(:, 1));
        et(tf) = props(loc(tf), 2);
    end
    % accumulate the orientation tensor n*n' per grid (sign-free), thickness mean
    Sxx = zeros(ng, 1); Syy = Sxx; Szz = Sxx; Sxy = Sxx; Sxz = Sxx; Syz = Sxx;
    cnt = Sxx; tsum = Sxx; tbad = Sxx; v0 = zeros(ng, 3);
    ref = [0.31 0.52 0.79];
    sgn = sign(enrm * ref'); sgn(sgn == 0) = 1;
    for k = 1:4
        idx = nodes(:, k); use = ~isnan(idx); idx = idx(use);
        n = enrm(use, :); tk = et(use);
        Sxx = Sxx + accumarray(idx, n(:, 1).^2, [ng 1]);
        Syy = Syy + accumarray(idx, n(:, 2).^2, [ng 1]);
        Szz = Szz + accumarray(idx, n(:, 3).^2, [ng 1]);
        Sxy = Sxy + accumarray(idx, n(:, 1) .* n(:, 2), [ng 1]);
        Sxz = Sxz + accumarray(idx, n(:, 1) .* n(:, 3), [ng 1]);
        Syz = Syz + accumarray(idx, n(:, 2) .* n(:, 3), [ng 1]);
        v0  = v0 + [accumarray(idx, n(:, 1) .* sgn(use), [ng 1]), accumarray(idx, n(:, 2) .* sgn(use), [ng 1]), ...
                    accumarray(idx, n(:, 3) .* sgn(use), [ng 1])];
        cnt = cnt + accumarray(idx, 1, [ng 1]);
        tsum = tsum + accumarray(idx, fillmissing_zero(tk), [ng 1]);
        tbad = tbad + accumarray(idx, double(isnan(tk)), [ng 1]);
    end
    on = cnt > 0;
    v = v0; small = vecnorm(v, 2, 2) < 1e-6; v(small, :) = repmat(ref, nnz(small), 1);
    for it = 1:8                                        % power iteration -> dominant axis
        v = [Sxx .* v(:, 1) + Sxy .* v(:, 2) + Sxz .* v(:, 3), ...
             Sxy .* v(:, 1) + Syy .* v(:, 2) + Syz .* v(:, 3), ...
             Sxz .* v(:, 1) + Syz .* v(:, 2) + Szz .* v(:, 3)];
        v = v ./ max(vecnorm(v, 2, 2), eps);
    end
    SH.n(on, :) = v(on, :);
    t = tsum ./ max(cnt, 1); t(tbad > 0 | ~on) = NaN;
    SH.t = t;
end


function x = fillmissing_zero(x)
    x(isnan(x)) = 0;
end


function [f, j] = card_fields(lines, i, n)
%CARD_FIELDS  Fields of the bulk-data card starting at line i (small, large or
%   free field, continuations included).  f{1} is the card name; j = next line.
    L = char(lines(i));
    f = {};
    if contains(L, ',')
        parts = strtrim(strsplit(L, ',', 'CollapseDelimiters', false));
        f = parts(1:min(9, end)); j = i + 1;                % field 10 is the continuation marker
        while j <= n
            nxt = char(lines(j));
            if isempty(strtrim(nxt)) || nxt(1) == '$' || ~contains(nxt, ','), break; end
            p2 = strtrim(strsplit(nxt, ',', 'CollapseDelimiters', false));
            if ~(isempty(p2{1}) || p2{1}(1) == '+' || p2{1}(1) == '*'), break; end
            p2 = p2(2:min(9, end));
            p2(end+1:8) = {''};
            f = [f p2]; j = j + 1;                            %#ok<AGROW>
        end
        f{1} = regexprep(f{1}, '\*$', '');
        return
    end
    large = numel(L) >= 8 && any(L(1:min(8, end)) == '*');
    w = 8; if large, w = 16; end
    L = pad_cols(L, 72); f = {strtrim(L(1:8))};
    body = L(9:72);
    f = [f cellstr(reshape(body, w, [])')'];
    j = i + 1;
    while j <= n
        nxt = char(lines(j));
        if isempty(strtrim(nxt)) || nxt(1) == '$' || contains(nxt, ','), break; end
        head = strtrim(nxt(1:min(8, end)));
        if ~(isempty(head) || head(1) == '+' || head(1) == '*'), break; end
        nxt = pad_cols(nxt, 72);
        f = [f cellstr(reshape(nxt(9:72), w, [])')']; j = j + 1;   %#ok<AGROW>
    end
    f{1} = regexprep(f{1}, '\*$', '');
end


function v = field_real(f, k)
    if k > numel(f) || isempty(strtrim(f{k})), v = 0; return; end
    v = nas_real(f(k));
end


function P = read_shell_props(lines, up, n)
%READ_SHELL_PROPS  [PID T] from PSHELL (T) and PCOMP (sum of ply T, doubled for SYM).
    P = zeros(0, 2);
    for i = find(startsWith(up, "PSHELL") | startsWith(up, "PCOMP"))'
        try
            f = card_fields(lines, i, n);
        catch
            continue
        end
        name = upper(regexprep(f{1}, '[^A-Z0-9]', ''));
        pid = str2double(f{2});
        if isnan(pid), continue; end
        switch name
            case 'PSHELL'
                P(end+1, :) = [pid field_real(f, 4)];                 %#ok<AGROW>
            case 'PCOMP'
                % plies from field 10 as (MID, T, THETA, SOUT); a blank T repeats the previous ply's
                t = 0; last = 0; k = 10;
                while k + 1 <= numel(f)
                    mid = strtrim(f{k}); tk = strtrim(f{k + 1});
                    if isempty(mid) && isempty(tk), break; end
                    if ~isempty(tk), last = nas_real(f(k + 1)); end
                    t = t + last; k = k + 4;
                end
                lam = ''; if numel(f) >= 9, lam = upper(strtrim(f{9})); end
                if startsWith(lam, 'SYM'), t = 2 * t; end
                P(end+1, :) = [pid t];                                %#ok<AGROW>
        end
    end
end


function E = read_elements_file(fname, is_top, stage, span)
%READ_ELEMENTS_FILE  Corner-node ids per supported element type, following
%   INCLUDEs.  Same vectorised slicing as read_grids_file; continuation lines
%   are assumed to directly follow their parent (true for every mesher).
    sub = @(f, m) stage(span(1) + f * (span(2) - span(1)), sprintf('%s: %s', shortname(fname), m));
    sub(0.00, 'reading file ...');
    types = {'CQUAD4', 4; 'CQUADR', 4; 'CQUAD8', 4; 'CTRIA3', 3; 'CTRIAR', 3; 'CTRIA6', 3; ...
             'CTETRA', 4; 'CPENTA', 6; 'CHEXA', 8};
    E = struct();
    for t = 1:size(types, 1), E.(types{t, 1}) = zeros(0, types{t, 2}); E.(['PID_' types{t, 1}]) = zeros(0, 1); end
    E.PSHELL = zeros(0, 2); E.PCOMP = zeros(0, 2);

    txt = fileread(fname);
    txt = strrep(txt, sprintf('\t'), ' ');
    lines = splitlines(string(txt));
    base_dir = fileparts(fname);
    n = numel(lines);
    up = upper(strip(lines, 'left'));
    first = 1;
    if is_top
        ib = find(startsWith(up, "BEGIN"), 1);
        if ~isempty(ib) && ~isempty(regexp(char(up(ib)), '^BEGIN\s+BULK', 'once')), first = ib + 1; end
    end
    last = find(startsWith(up, "ENDDATA"), 1);
    if isempty(last), last = n; else, last = last - 1; end
    lines = lines(first:last); up = up(first:last); n = numel(lines);

    % card name of every line (letters/digits before the first blank, comma or *)
    sub(0.20, sprintf('scanning %d lines for element cards ...', n));
    cand = startsWith(up, "C");                       % cheap pre-filter
    name_of = strings(n, 1);
    name_of(cand) = regexp(up(cand), '^[A-Z0-9]+', 'match', 'once');
    name_of(ismissing(name_of)) = "";
    star = cand & startsWith(extractAfter(up, strlength(name_of)), "*");
    for t = 1:size(types, 1)
        name = types{t, 1}; ng = types{t, 2}; nf = 2 + ng;      % EID PID G1..Gng
        mine = find(name_of == name);
        if isempty(mine), continue; end
        sub(0.35 + 0.35 * t / size(types, 1), sprintf('parsing %d %s ...', numel(mine), name));
        isfree  = contains(lines(mine), ",");
        islarge = ~isfree & star(mine);
        issmall = ~isfree & ~islarge;
        rows = zeros(0, nf);
        if any(issmall)
            li = mine(issmall);
            nl = ceil(nf / 8);
            M = zeros(numel(li), 0);  M = char(M);                 % grow columns
            for j = 1:nl
                Lj = pad_cols(char(lines(min(li + j - 1, n))), 72);
                M = [M Lj(:, 9:72)];                                 %#ok<AGROW>
            end
            rows = [rows; fields_to_num(M, 8, nf)];                  %#ok<AGROW>
        end
        if any(islarge)
            li = mine(islarge);
            nl = ceil(nf / 4);
            M = char(zeros(numel(li), 0));
            for j = 1:nl
                Lj = pad_cols(char(lines(min(li + j - 1, n))), 72);
                M = [M Lj(:, 9:72)];                                 %#ok<AGROW>
            end
            rows = [rows; fields_to_num(M, 16, nf)];                 %#ok<AGROW>
        end
        if any(isfree)
            for i = mine(isfree)'
                f = strtrim(strsplit(char(lines(i)), ',', 'CollapseDelimiters', false));
                if isempty(f{end}), f(end) = []; end
                j = i + 1;
                while numel(f) < nf + 1 && j <= n
                    nxt = char(lines(j));
                    if isempty(nxt) || nxt(1) == '$', break; end
                    f2 = strtrim(strsplit(nxt, ',', 'CollapseDelimiters', false));
                    if ~isempty(f2) && (isempty(f2{1}) || f2{1}(1) == '+' || f2{1}(1) == '*')
                        f2 = f2(2:end);
                    end
                    f = [f f2];                                      %#ok<AGROW>
                    j = j + 1;
                end
                f(end+1:nf+1) = {''};
                rows = [rows; str2double(f(2:nf+1))];                %#ok<AGROW>
            end
        end
        rows = rows(all(~isnan(rows(:, 3:end)), 2), :);
        E.(name) = [E.(name); rows(:, 3:end)];
        E.(['PID_' name]) = [E.(['PID_' name]); rows(:, 2)];
    end
    P = read_shell_props(lines, up, n);
    E.PSHELL = [E.PSHELL; P];

    % ---- INCLUDEs ------------------------------------------------------------
    incs = find(startsWith(up, "INCLUDE"))';
    for q = 1:numel(incs)
        i = incs(q);
        rest = regexprep(char(lines(i)), '^\s*INCLUDE\s*', '', 'ignorecase');
        j = i + 1;
        while nnz(rest == '''') < 2 && j <= n
            rest = [rest strtrim(char(lines(j)))];                   %#ok<AGROW>
            j = j + 1;
        end
        inc = strtrim(regexprep(strtrim(rest), '^''|''$', ''));
        lo = span(1) + (0.70 + 0.30 * (q - 1) / numel(incs)) * (span(2) - span(1));
        hi = span(1) + (0.70 + 0.30 * q / numel(incs)) * (span(2) - span(1));
        E2 = read_elements_file(resolve_include(inc, base_dir), false, stage, [lo hi]);
        for t = 1:size(types, 1)
            E.(types{t, 1}) = [E.(types{t, 1}); E2.(types{t, 1})];
            E.(['PID_' types{t, 1}]) = [E.(['PID_' types{t, 1}]); E2.(['PID_' types{t, 1}])];
        end
        E.PSHELL = [E.PSHELL; E2.PSHELL];
    end
end


function A = fields_to_num(M, w, nf)
%FIELDS_TO_NUM  First nf fixed-width integer fields of each row of a char matrix.
    A = zeros(size(M, 1), nf);
    for k = 1:nf
        A(:, k) = fast_int(M(:, (k-1)*w + 1 : k*w));
    end
end


function v = fast_int(M)
%FAST_INT  Non-negative integers from a char matrix, one per row, without
%   str2double: ~50x faster on a million rows.  Blank or non-digit -> NaN.
    isd = M >= '0' & M <= '9';
    D = double(M - '0');
    D(~isd) = 0;
    v = zeros(size(M, 1), 1);
    for c = 1:size(M, 2)
        v = v .* (1 + 9 * isd(:, c)) + D(:, c);      % x10 only when a digit lands
    end
    bad = ~any(isd, 2) | any(~isd & M ~= ' ', 2);
    v(bad) = NaN;
end


% =========================================================================
function f = length_factor(from, to)
%LENGTH_FACTOR  Multiply a length in FROM units to get TO units.
    m = struct('in', 0.0254, 'mm', 1e-3, 'm', 1);
    from = lower(char(from)); to = lower(char(to));
    if ~isfield(m, from) || ~isfield(m, to)
        error('temp_map_matlab:badLength', 'Length units must be in, mm or m (got %s -> %s).', from, to);
    end
    f = m.(from) / m.(to);
end


function T = to_kelvin(T, units)
    switch upper(char(units))
        case 'K', return
        case 'C', T = T + 273.15;
        case 'F', T = (T - 32) * 5/9 + 273.15;
        otherwise
            error('temp_map_matlab:badTemp', 'CSV_TEMP_UNITS must be K, C or F.');
    end
end


function T = from_kelvin(T, units)
    switch upper(char(units))
        case 'K', return
        case 'C', T = T - 273.15;
        case 'F', T = (T - 273.15) * 9/5 + 32;
        otherwise
            error('temp_map_matlab:badTemp', 'OUT_UNITS must be K, C or F.');
    end
end


% =========================================================================
function [t, xyz, T] = read_cloud(csv_file, has_header)
%READ_CLOUD  time, xyz, T from a 5-column CSV.  textscan fast path (a few
%   seconds per million rows); readmatrix as the tolerant fallback.
    M = [];
    fid = fopen(csv_file, 'r');
    if fid > 0
        cl = onCleanup(@() fclose(fid));
        try
            if has_header, fgetl(fid); end
            c = textscan(fid, '%f%f%f%f%f', 'Delimiter', ',', 'CollectOutput', true);
            M = c{1};
            if ~feof(fid), M = []; end       % stopped early on something odd: fall back
        catch
            M = [];
        end
        clear cl
    end
    if isempty(M)
        opts = detectImportOptions(csv_file, 'FileType', 'text');
        if ~has_header
            opts.DataLines = [1 Inf];
        end
        M = readmatrix(csv_file, opts);
    end
    if size(M, 2) < 5
        error('temp_map_matlab:badCsv', ...
            '%s: expected 5 columns (time,x,y,z,T) but found %d.', csv_file, size(M, 2));
    end
    M = M(:, 1:5);
    M(any(isnan(M), 2), :) = [];
    if isempty(M)
        error('temp_map_matlab:emptyCsv', '%s: no numeric rows.', csv_file);
    end
    t   = M(1, 1);
    xyz = M(:, 2:4);
    T   = M(:, 5);
    if any(abs(M(:, 1) - t) > 1e-9 * max(1, abs(t)))
        warning('temp_map_matlab:multiTime', ...
            '%s: time column is not constant (%g .. %g); using the first value.', ...
            csv_file, min(M(:, 1)), max(M(:, 1)));
    end
end


% =========================================================================
function [sidx, L] = shell_index(Gi, C)
%SHELL_INDEX  Rows of the grids that get a through-thickness average, and the layer count.
    sidx = []; L = 0;
    if ~C.SHELL_AVERAGE || ~isfield(Gi, 'shell_t') || isempty(Gi.shell_t), return; end
    sidx = find(~isnan(Gi.shell_t) & Gi.shell_t > 0 & all(~isnan(Gi.shell_n), 2));
    if ~isempty(sidx), L = max(2, round(C.SHELL_LAYERS)); end
end


function [Q, sidx, L] = shell_samples(Gi, C)
%SHELL_SAMPLES  Points to map: every grid, then L levels across +-t/2 along the
%   normal for each shell grid (appended, layer-major).
    Q = Gi.xyz;
    [sidx, L] = shell_index(Gi, C);
    if isempty(sidx), return; end
    z = linspace(-0.5, 0.5, L);
    m = numel(sidx);
    extra = zeros(m * L, 3);
    for j = 1:L
        extra((j - 1) * m + (1:m), :) = Gi.xyz(sidx, :) + Gi.shell_n(sidx, :) .* (z(j) * Gi.shell_t(sidx));
    end
    Q = [Q; extra];
end


function gT = collapse_layers(Tall, Gi, C)
%COLLAPSE_LAYERS  Grid temperatures from the mapped sample points: shell grids
%   take the mean of their layers, everything else its own value.
    n = size(Gi.xyz, 1);
    gT = Tall(1:n);
    [sidx, L] = shell_index(Gi, C);
    if isempty(sidx), return; end
    gT(sidx) = mean(reshape(Tall(n + 1:end), numel(sidx), L), 2);
end


function M = build_map(P, Q, C, stage)
%BUILD_MAP  Everything about mapping cloud POINTS P onto grids Q that does not
%   depend on the temperatures: returned as a mapping object M so that every
%   later time step with the same cloud locations is just APPLY_MAP(M, T).
%
%   M.W       sparse (nQ x nPu) weights: Tg = W * Tu, with Tu the temperatures
%             of the unique cloud points (duplicates averaged via M.ic).
%   M.F       (scattered method only) a scatteredInterpolant instead of W.
%   M.extrap  grids outside the cloud hull      M.nndist  distance to nearest pt
%   M.method  description                       M.ic / M.nPu  duplicate map
    if nargin < 4, stage = @(~, ~) []; end
    stage(0.10, sprintf('checking %d cloud points for duplicates ...', size(P, 1)));
    [P, ~, ic] = unique(P, 'rows');          % duplicates would break Delaunay
    nP = size(P, 1);
    nQ = size(Q, 1);
    meth = lower(char(C.METHOD));
    M = struct('W', [], 'F', [], 'Q', [], 'ic', ic, 'nPu', nP, ...
               'extrap', false(nQ, 1), 'nndist', [], 'method', '');

    if nP == 1
        M.W = sparse(ones(nQ, 1), 1, 1, nQ, 1);
        M.extrap = true(nQ, 1);
        M.nndist = vecnorm(Q - P, 2, 2);
        M.method = 'single point';
        return
    end

    % ---- kd-tree methods (fast at 1M x 1M) ---------------------------------
    have_knn = exist('knnsearch', 'file') == 2 && license('test', 'Statistics_Toolbox');
    if ismember(meth, {'nearest', 'idw'}) && ~have_knn
        warning('temp_map_matlab:noKnn', ...
            'knnsearch (Statistics Toolbox) not available; using the triangulation for METHOD=%s.', meth);
    end
    if strcmp(meth, 'scattered')
        stage(0.25, sprintf('scatteredInterpolant (linear/nearest) on %d cloud points ...', nP));
        M.F = scatteredInterpolant(P, zeros(nP, 1), 'linear', 'nearest');
        M.Q = Q;
        if have_knn
            [~, M.nndist] = knn_chunked(P, Q, 1, stage, 'nearest distance');
        else
            M.nndist = nan(nQ, 1);       % not worth a second triangulation
        end
        M.method = 'scatteredInterpolant linear/nearest';
        return
    end
    if strcmp(meth, 'nearest') && have_knn
        [nn, M.nndist] = knn_chunked(P, Q, 1, stage, 'nearest');
        M.W = sparse((1:nQ)', nn, 1, nQ, nP);
        M.method = 'nearest (kd-tree)';
        return
    elseif strcmp(meth, 'idw') && have_knn
        k = min(C.IDW_K, nP);
        [nn, d] = knn_chunked(P, Q, k, stage, sprintf('IDW k=%d', k));
        w  = 1 ./ max(d, eps) .^ C.IDW_POWER;
        hit = d(:, 1) == 0;                     % sitting on a cloud point: take it
        w(hit, :) = 0; w(hit, 1) = 1;
        w  = w ./ sum(w, 2);
        M.W = sparse(repmat((1:nQ)', k, 1), nn(:), w(:), nQ, nP);
        M.nndist = d(:, 1);
        M.method = sprintf('IDW (k=%d, p=%g)', k, C.IDW_POWER);
        return
    end

    % ---- triangulation path --------------------------------------------------
    mu = mean(P, 1);                          % effective dimension via SVD
    [~, Sv, V] = svd(P - mu, 'econ');
    sv = diag(Sv);
    dim = nnz(sv > 1e-9 * sv(1));
    Pp = (P - mu) * V(:, 1:dim);
    Qp = (Q - mu) * V(:, 1:dim);
    rows = (1:nQ)';

    switch dim
        case {2, 3}
            try
                stage(0.25, sprintf('Delaunay triangulation of %d cloud points (the slow step) ...', nP));
                DT = delaunayTriangulation(Pp);
            catch ME
                warning('temp_map_matlab:delaunay', ...
                    'Triangulation failed (%s); falling back to nearest point.', ME.message);
                DT = [];
            end
            if isempty(DT)
                nn = dsearchn(Pp, Qp);
                M.W = sparse(rows, nn, 1, nQ, nP);
                M.extrap = true(nQ, 1);
                M.method = 'nearest (brute force)';
            else
                stage(0.75, sprintf('locating %d grids in the triangulation ...', nQ));
                nn = nearestNeighbor(DT, Qp);
                if strcmp(meth, 'linear')
                    [ti, bc] = pointLocation(DT, Qp);
                    inside = ~isnan(ti);
                    tri = DT.ConnectivityList(ti(inside), :);    % (n_in x dim+1)
                    % bridging cells: Delaunay spans concavities of the body with
                    % long flat cells whose vertices are far-apart surface points;
                    % interpolating across those smears local hot/cold spots.
                    nbridge = 0;
                    if ~isempty(C.SLIVER_FACTOR) && any(inside)
                        stage(0.88, 'checking for bridging cells across concavities ...');
                        E = edges(DT);
                        h = median(vecnorm(Pp(E(:, 1), :) - Pp(E(:, 2), :), 2, 2));
                        maxedge = zeros(size(tri, 1), 1);
                        for a = 1:size(tri, 2) - 1
                            for b = a + 1:size(tri, 2)
                                maxedge = max(maxedge, vecnorm(Pp(tri(:, a), :) - Pp(tri(:, b), :), 2, 2));
                            end
                        end
                        bridge = maxedge > C.SLIVER_FACTOR * h;
                        nbridge = nnz(bridge);
                    else
                        bridge = false(size(tri, 1), 1);
                    end
                    % weights: barycentric inside well-formed cells, nearest elsewhere
                    lin  = find(inside); lin = lin(~bridge);
                    trl  = tri(~bridge, :); bcl = bc(inside, :); bcl = bcl(~bridge, :);
                    near = setdiff(rows, lin);
                    M.W = sparse([repmat(lin, size(trl, 2), 1); near], [trl(:); nn(near)], ...
                                 [bcl(:); ones(numel(near), 1)], nQ, nP);
                    M.extrap = ~inside;
                    if dim == 3, M.method = 'linear (3D Delaunay)';
                    else,        M.method = 'linear (planar cloud)'; end
                    if nbridge > 0
                        M.method = sprintf('%s, %d grids in bridging cells -> nearest', M.method, nbridge);
                    end
                else
                    M.W = sparse(rows, nn, 1, nQ, nP);
                    M.method = 'nearest (triangulation)';
                end
            end
        otherwise   % dim == 1: collinear cloud
            [s, order] = unique(Pp(:, 1));
            sq = min(max(Qp(:, 1), s(1)), s(end));
            nn = order(interp1(s, 1:numel(s), sq, 'nearest'));
            if strcmp(meth, 'linear') && numel(s) > 1
                seg = discretize(sq, s);                      % s(seg) <= sq <= s(seg+1)
                w = (sq - s(seg)) ./ (s(seg + 1) - s(seg));
                M.W = sparse([rows; rows], [order(seg); order(seg + 1)], [1 - w; w], nQ, nP);
                M.extrap = Qp(:, 1) < s(1) | Qp(:, 1) > s(end);
                M.method = 'linear (collinear cloud)';
            else
                M.W = sparse(rows, nn, 1, nQ, nP);
                M.method = 'nearest (collinear cloud)';
            end
    end
    M.nndist = vecnorm(Q - P(nn, :), 2, 2);
end


function Tg = apply_map(M, T)
%APPLY_MAP  Temperatures at the grids from cloud temperatures T (raw rows).
    Tu = accumarray(M.ic, T(:), [M.nPu 1], @mean);   % duplicates averaged
    if ~isempty(M.F)
        M.F.Values = Tu;
        Tg = M.F(M.Q);
    else
        Tg = M.W * Tu;
    end
    Tg = Tg(:);
end


function tf = same_cloud(A, B)
%SAME_CLOUD  True when two clouds have identical point locations.
    tf = ~isempty(A) && isequal(size(A), size(B)) && isequal(A, B);
end


% =========================================================================
function [nn, d] = knn_chunked(P, Q, k, stage, label)
%KNN_CHUNKED  k nearest cloud points for every query, built once and searched
%   in chunks so the progress bar moves and Cancel gets a chance to fire.
    nQ = size(Q, 1);
    stage(0.25, sprintf('%s: building kd-tree on %d cloud points ...', label, size(P, 1)));
    Mdl = KDTreeSearcher(P);
    nn = zeros(nQ, k); d = zeros(nQ, k);
    chunk = 50000;
    for a = 1:chunk:nQ
        b = min(a + chunk - 1, nQ);
        stage(0.3 + 0.6 * (a - 1) / nQ, sprintf('%s: searching grids %d-%d of %d ...', label, a, b, nQ));
        [nn(a:b, :), d(a:b, :)] = knnsearch(Mdl, Q(a:b, :), 'K', k);
    end
end


% =========================================================================
function lines = coverage_report(r)
%COVERAGE_REPORT  How well the cloud's extent covers the mesh, as text lines.
%   Per axis: cloud vs grid min/max and how far the mesh overhangs the cloud.
%   Plus hull / distance / temperature-range statistics.
    G = r.grid_xyz; P = r.cloud_xyz;
    ax = 'XYZ';
    lines = cell(0, 1);
    lines{end+1} = sprintf('%-2s %-19s %-19s %-19s', '', 'cloud [min max]', 'grid [min max]', 'overhang [lo hi]');
    for a = 1:3
        lo = max(0, min(P(:, a)) - min(G(:, a)));
        hi = max(0, max(G(:, a)) - max(P(:, a)));
        lines{end+1} = sprintf('%-2s [%8.4g %8.4g] [%8.4g %8.4g] [%8.4g %8.4g]', ...
            ax(a), min(P(:, a)), max(P(:, a)), min(G(:, a)), max(G(:, a)), lo, hi); %#ok<AGROW>
    end
    n = numel(r.grid_ids);
    d = r.nn_dist;
    lines{end+1} = sprintf('overhang = how far the mesh sticks out past the cloud on that side');
    lines{end+1} = sprintf('grids outside cloud hull : %d of %d (%.1f%%)', nnz(r.extrap), n, 100 * nnz(r.extrap) / n);
    lines{end+1} = sprintf('grid->nearest cloud pt   : median %.4g  95%% %.4g  max %.4g', ...
        median(d), prctile_plain(d, 95), max(d));
    lines{end+1} = sprintf('temperature K            : cloud [%.1f %.1f]  mapped [%.1f %.1f]', ...
        min(r.cloud_T), max(r.cloud_T), min(r.grid_T), max(r.grid_T));
end


function w = coverage_warnings(r, C)
%COVERAGE_WARNINGS  Loud, human-readable reasons to distrust this mapping.
    w = {};
    G = r.grid_xyz; P = r.cloud_xyz;
    ax = 'XYZ';
    ext = max(P) - min(P);
    for a = 1:3
        lo = min(P(:, a)) - min(G(:, a));
        hi = max(G(:, a)) - max(P(:, a));
        if ext(a) > 0 && max(lo, hi) > C.OVERHANG_WARN * ext(a)
            w{end+1} = sprintf('mesh overhangs the cloud in %s by %.4g (%.0f%% of the cloud''s %s extent)', ...
                ax(a), max(lo, hi), 100 * max(lo, hi) / ext(a), ax(a)); %#ok<AGROW>
        end
    end
    % bounding-box overlap: a units mismatch shows up as almost none
    lo = max(min(P), min(G)); hi = min(max(P), max(G));
    inter = prod(max(hi - lo, 0));
    union = prod(max(max(P), max(G)) - min(min(P), min(G)));
    if union > 0 && inter / union < 0.25
        w{end+1} = sprintf('cloud and mesh bounding boxes barely overlap (%.0f%%) -- length units mismatch?', ...
            100 * inter / union);
    end
    n = numel(r.grid_ids);
    if nnz(r.extrap) > 0.20 * n
        w{end+1} = sprintf('%.0f%% of grids are outside the cloud hull (nearest-point temps used)', ...
            100 * nnz(r.extrap) / n);
    end
end


function q = prctile_plain(x, p)
%PRCTILE_PLAIN  Percentile without the Statistics Toolbox.
    x = sort(x(:));
    q = x(max(1, min(numel(x), ceil(p / 100 * numel(x)))));
end


% =========================================================================
function srf = cloud_surface(P, nmax)
%CLOUD_SURFACE  Boundary of the cloud as an alpha shape (concave-aware),
%   built from a thinned subset so it is cheap.  Returns struct .F .V
%   (triangle faces / vertices) or [] if it cannot be built (planar cloud).
    srf = [];
    if nmax <= 0 || size(P, 1) < 4, return; end
    n = size(P, 1);
    if n > nmax
        rs = RandStream('mt19937ar', 'Seed', 0);
        P = P(sort(randperm(rs, n, round(nmax))), :);
    end
    P = unique(P, 'rows');                    % alphaShape warns on repeats
    if size(P, 1) < 4, return; end
    try
        shp = alphaShape(P(:, 1), P(:, 2), P(:, 3));
        shp.Alpha = 1.5 * criticalAlpha(shp, 'one-region');   % closed, but not convex
        [F, V] = boundaryFacets(shp);
        srf = struct('F', F, 'V', V, 'alpha', shp.Alpha, 'npts', size(P, 1));
    catch
        srf = [];
    end
end


% =========================================================================
function write_temp_cards(rs, C)
%WRITE_TEMP_CARDS  One bulk-data file of TEMP cards for one time step: every
%   part of rs (same SID) in its own $-headed section.  Each section is built
%   as a char matrix and written in one go -- a 1M-grid file takes seconds.
    units = upper(char(C.OUT_UNITS));
    w = C.FIELD_SIZE;
    out = rs(1).out_file;

    fid = fopen(out, 'w');
    if fid < 0
        error('temp_map_matlab:cantWrite', 'Cannot open %s for writing.', out);
    end
    cl = onCleanup(@() fclose(fid));

    fprintf(fid, '$ TEMP cards generated by temp_map_matlab  (%s)\n', datestr(now, 'yyyy-mm-dd HH:MM'));
    fprintf(fid, '$ Step   : %s   time = %g s   SID %d   %d part(s)\n', shortname(rs(1).csv_file), rs(1).time, rs(1).sid, numel(rs));
    fprintf(fid, '$ Lengths: BDF in %s, cloud read as %s%s\n', C.BDF_LENGTH_UNITS, C.CSV_LENGTH_UNITS, ...
            ternary(strcmpi(C.BDF_LENGTH_UNITS, C.CSV_LENGTH_UNITS), '', ' (converted)'));
    fprintf(fid, '$ Units  : deg %s\n', units);

    sidS = sprintf('%*d', w, rs(1).sid);
    if C.WRITE_TEMPD
        Tall = from_kelvin(vertcat(rs.grid_T), units);
        if w == 8, fprintf(fid, 'TEMPD   %s%s\n',  sidS, fmt_real(mean(Tall), w));
        else,      fprintf(fid, 'TEMPD*  %s%s\n', sidS, fmt_real(mean(Tall), w)); end
    end
    for i = 1:numel(rs)
        write_part_cards(fid, rs(i), C, sidS, numel(rs) > 1);
    end
end


function write_part_cards(fid, r, C, sidS, labelled)
    units = upper(char(C.OUT_UNITS));
    T = from_kelvin(r.grid_T(:), units);
    ids = r.grid_ids(:);
    n = numel(ids);
    w = C.FIELD_SIZE;

    if labelled, fprintf(fid, '$\n$ ---- part %s ----\n', r.part); end
    fprintf(fid, '$ BDF    : %s\n', r.bdf_file);
    if isempty(r.cloud_T)
        fprintf(fid, '$ Source : %s   (step %s, time = %g s)\n', r.method, r.csv_file, r.time);
    else
        fprintf(fid, '$ CSV    : %s   (time = %g s)\n', r.csv_file, r.time);
    end
    fprintf(fid, '$ Grids  : %d   outside cloud hull (nearest used): %d\n', n, nnz(r.extrap));
    fprintf(fid, '$ Method : %s\n', r.method);
    fprintf(fid, '$ Trange : %.3f .. %.3f\n', min(T), max(T));

    % fixed-width columns for every grid, padded to a multiple of 3 pairs
    idS = sprintf(sprintf('%%%dd', w), ids);
    if numel(idS) ~= w * n
        error('temp_map_matlab:idWidth', 'A grid id does not fit in %d characters.', w);
    end
    idS = reshape(idS, w, n)';
    tS  = fmt_real(T, w);
    pad = mod(-n, 3);
    idS = [idS; repmat(' ', pad, w)];
    tS  = [tS;  repmat(' ', pad, w)];
    m   = (n + pad) / 3;
    P1 = [idS(1:3:end, :) tS(1:3:end, :)];
    P2 = [idS(2:3:end, :) tS(2:3:end, :)];
    P3 = [idS(3:3:end, :) tS(3:3:end, :)];
    sidCol = repmat(sidS, m, 1);

    if w == 8
        % TEMP  SID  G1 T1  G2 T2  G3 T3   (64 columns)
        rows = [repmat('TEMP    ', m, 1) sidCol P1 P2 P3];
    else
        % TEMP* SID G1 T1 G2 *      /   * T2 G3 T3
        r1 = [repmat('TEMP*   ', m, 1) sidCol P1 idS(2:3:end, :) repmat('*', m, 1)];   % 73 cols
        r2 = [repmat('*       ', m, 1) tS(2:3:end, :) P3];                              % 56 cols
        r2(:, end+1:size(r1, 2)) = ' ';
        rows = [r1; r2];
        rows = rows(reshape([1:m; m+1:2*m], [], 1), :);   % interleave
    end
    rows(:, end+1) = newline;
    fwrite(fid, rows', 'char');
end


function S = fmt_real(v, w)
%FMT_REAL  Right-justified Nastran reals, one per row of an (n x w) char matrix.
%   Always carries a decimal point; drops the 'E' when that is what it takes
%   to fit ('1.2345-5'); falls back to fewer significant digits.  Vectorised.
    v = v(:);
    n = numel(v);
    S = repmat(' ', n, w);
    done = false(n, 1);
    z = v == 0;
    S(z, :) = repmat([blanks(w-2) '0.'], nnz(z), 1);
    done(z) = true;
    for p = (w-1):-1:1
        idx = find(~done);
        if isempty(idx), break; end
        c = compose(sprintf('%%.%dG', p), v(idx));                % "300.1235", "1.5E-05"
        c = regexprep(c, 'E([+-])0*(\d)', 'E$1$2');               % E-05 -> E-5
        c = regexprep(c, '^([+-]?\d+)(E|$)', '$1.$2');            % 300 -> 300.
        long = strlength(c) > w & contains(c, 'E');
        c(long) = strrep(c(long), 'E', '');                        % 1.5E-5 -> 1.5-5
        ok = strlength(c) <= w;
        if any(ok)
            S(idx(ok), :) = char(pad(c(ok), w, 'left'));
            done(idx(ok)) = true;
        end
    end
    if ~all(done)
        error('temp_map_matlab:fmt', 'Cannot format %g in %d characters.', v(find(~done, 1)), w);
    end
end


% =========================================================================
function write_case_control(R, C)
%WRITE_CASE_CONTROL  temp_subcases.dat (one SUBCASE per time step, ids matching
%   the TEMP SIDs) and temp_includes.bdf (one INCLUDE per step file, plus the
%   optional reference-temperature TEMPD), both in OUT_DIR.
    units = upper(char(C.OUT_UNITS));
    [np, ns] = size(R);

    % ---- case control -----------------------------------------------------------
    cf = fullfile(C.OUT_DIR, 'temp_subcases.dat');
    fid = fopen(cf, 'w');
    if fid < 0, error('temp_map_matlab:cantWrite', 'Cannot open %s for writing.', cf); end
    cl = onCleanup(@() fclose(fid));
    fprintf(fid, '$ Case control generated by temp_map_matlab  (%s)\n', datestr(now, 'yyyy-mm-dd HH:MM'));
    fprintf(fid, '$ %d subcases (%d part(s)), SUBCASE = TEMP SID + %d.  Paste above BEGIN BULK.\n', ...
            ns, np, C.SUBCASE_OFFSET);
    % global section: applies to every subcase below
    for e = 1:numel(C.CASE_EXTRA)
        fprintf(fid, '%s\n', strtrim(char(C.CASE_EXTRA{e})));
    end
    if ~isempty(C.TREF)
        fprintf(fid, '$ reference temperature %.3f %s (TEMPD SID %d in temp_includes.bdf)\n', C.TREF, units, C.TREF_SID);
        fprintf(fid, 'TEMPERATURE(INITIAL) = %d\n', C.TREF_SID);
    end
    for k = 1:ns
        r = R(1, k);
        [~, base] = fileparts(r.csv_file);
        tok = @(s) title72(fill_tokens(s, base, r.time, r.sid, k));
        fprintf(fid, 'SUBCASE %d\n', r.sid + C.SUBCASE_OFFSET);
        if ~isempty(C.SUBTITLE), fprintf(fid, '  SUBTITLE = %s\n', tok(C.SUBTITLE)); end
        if ~isempty(C.LABEL),    fprintf(fid, '  LABEL = %s\n',    tok(C.LABEL));    end
        fprintf(fid, '  TEMPERATURE(LOAD) = %d\n', r.sid);
    end
    clear cl
    fprintf('  wrote %s  (%d subcases)\n', cf, ns);

    % ---- bulk includes -----------------------------------------------------------
    inf_ = fullfile(C.OUT_DIR, 'temp_includes.bdf');
    fid = fopen(inf_, 'w');
    if fid < 0, error('temp_map_matlab:cantWrite', 'Cannot open %s for writing.', inf_); end
    cl = onCleanup(@() fclose(fid));
    fprintf(fid, '$ INCLUDE one line per TEMP file, generated by temp_map_matlab. Paste below BEGIN BULK.\n');
    if ~isempty(C.TREF)
        if C.FIELD_SIZE == 8
            fprintf(fid, 'TEMPD   %8d%s\n',  C.TREF_SID, fmt_real(C.TREF, 8));
        else
            fprintf(fid, 'TEMPD*  %16d%s\n', C.TREF_SID, fmt_real(C.TREF, 16));
        end
    end
    for k = 1:ns
        [~, b, e] = fileparts(R(1, k).out_file);
        fprintf(fid, 'INCLUDE ''%s''\n', [b e]);          % relative: sits next to this file
    end
    fprintf('  wrote %s  (%d includes)\n', inf_, ns);
end


function s = fill_tokens(s, file, t, sid, idx)
    s = strrep(char(s), '{file}', file);
    s = strrep(s, '{time}', sprintf('%g', t));
    s = pad_token(s, 'sid', sid);
    s = pad_token(s, 'index', idx);
end


function s = pad_token(s, name, v)
%PAD_TOKEN  {name} -> v, {name:03} -> zero-padded to 3 digits.
    ext = unique(regexp(s, ['\{' name '(:\d+)?\}'], 'match'));
    for i = 1:numel(ext)
        wdt = sscanf(ext{i}, ['{' name ':%d}']);
        if isempty(wdt), rep = sprintf('%d', v); else, rep = sprintf('%0*d', wdt, v); end
        s = strrep(s, ext{i}, rep);
    end
end


function s = title72(s)
    if numel(s) > 72, s = s(1:72); end          % Nastran title field limit
end


function base = out_base(r, C, k)
%OUT_BASE  File name (no extension) of step k from the OUT_NAME template.
    [~, csv] = fileparts(r.csv_file);
    base = fill_tokens(C.OUT_NAME, csv, r.time, r.sid, k);
    base = regexprep(strtrim(base), '[<>:"/\\|?*\s]+', '_');
    if isempty(base), base = sprintf('temp_%03d', k); end
end


% =========================================================================
function write_report(R, RM, S, C)
%WRITE_REPORT  Self-contained HTML run record in OUT_DIR/temp_map_report.html.
    f = fullfile(C.OUT_DIR, 'temp_map_report.html');
    fid = fopen(f, 'w');
    if fid < 0, error('temp_map_matlab:cantWrite', 'Cannot open %s for writing.', f); end
    cl = onCleanup(@() fclose(fid));
    [np, ns] = size(R);
    units = upper(char(C.OUT_UNITS));
    esc = @(t) regexprep(char(t), {'&', '<', '>'}, {'&amp;', '&lt;', '&gt;'});
    w = @(varargin) fprintf(fid, varargin{:});

    w('<!DOCTYPE html><html><head><meta charset="utf-8"><title>TEMP mapping report</title>\n');
    w(['<style>body{font:14px/1.4 Segoe UI,Arial,sans-serif;margin:24px;color:#222}h1{font-size:22px}' ...
       'h2{font-size:17px;margin-top:28px;border-bottom:1px solid #ccc}table{border-collapse:collapse;font-size:13px}' ...
       'th,td{border:1px solid #ccc;padding:3px 8px;text-align:right}th{background:#eee}td:first-child,td:nth-child(2){text-align:left}' ...
       '.warn{background:#fde8e8;border-left:4px solid #c00;padding:8px 12px;margin:6px 0}.ok{background:#e8f5e9;border-left:4px solid #2a2;padding:8px 12px}' ...
       'pre{background:#f6f6f6;padding:8px;font-size:12px;overflow-x:auto}details{margin:4px 0}img{max-width:100%%;border:1px solid #ccc;margin:4px 0}' ...
       'svg{background:#fff;border:1px solid #ccc}</style></head><body>\n']);
    w('<h1>TEMP mapping report</h1>\n');
    w('<p>%s &nbsp;&middot;&nbsp; %d part(s) &times; %d time step(s)</p>\n', datestr(now, 'yyyy-mm-dd HH:MM'), np, ns);

    % ---- settings ------------------------------------------------------------------
    w('<h2>Settings</h2><table>\n');
    for i = 1:np
        src = fileparts(R(i, 1).csv_file);
        if isempty(R(i, 1).cloud_T), src = R(i, 1).method; end
        w('<tr><td>part</td><td>%s</td><td style="text-align:left">%s &nbsp;&larr;&nbsp; %s</td></tr>\n', ...
          esc(R(i, 1).part), esc(R(i, 1).bdf_file), esc(src));
    end
    rows = {'method', C.METHOD; 'model length units', C.BDF_LENGTH_UNITS; 'cloud length units', C.CSV_LENGTH_UNITS; ...
            'cloud temperature units', C.CSV_TEMP_UNITS; 'output temperature units', units; ...
            'SID start', sprintf('%d', C.SID_START); 'card format', sprintf('%d', C.FIELD_SIZE); ...
            'output folder', C.OUT_DIR};
    for q = 1:size(rows, 1)
        w('<tr><td>%s</td><td colspan="2" style="text-align:left">%s</td></tr>\n', esc(rows{q, 1}), esc(rows{q, 2}));
    end
    w('</table>\n');

    % ---- warnings ------------------------------------------------------------------
    w('<h2>Coverage</h2>\n');
    nwarn = 0;
    for k = 1:ns
        if ~isempty(RM(k).warnings)
            nwarn = nwarn + 1;
            w('<div class="warn"><b>%s</b><br>%s</div>\n', esc(shortname(RM(k).csv_file)), ...
              strjoin(cellfun(esc, RM(k).warnings(:)', 'UniformOutput', false), '<br>'));
        end
    end
    if nwarn == 0
        w('<div class="ok">No coverage warnings: every step''s mesh lies within its cloud (overhang < %.0f%%, outside-hull < 20%%).</div>\n', ...
          100 * C.OVERHANG_WARN);
    end
    w('<details><summary>Coverage detail, first step</summary><pre>%s</pre></details>\n', ...
      esc(strjoin(RM(1).coverage(:)', newline)));

    % ---- min / max chart (inline SVG) -------------------------------------------------
    t = [RM.time]; [t, order] = sort(t);
    tmin = arrayfun(@(r) from_kelvin(min(r.grid_T), units), RM(order));
    tmax = arrayfun(@(r) from_kelvin(max(r.grid_T), units), RM(order));
    if ns > 1
        W = 900; Hh = 260; ml = 60; mr = 20; mt = 20; mb = 40;
        x = @(v) ml + (W - ml - mr) * (v - t(1)) / max(t(end) - t(1), eps);
        lo = min(tmin); hi = max(tmax); if hi == lo, hi = lo + 1; end
        y = @(v) mt + (Hh - mt - mb) * (1 - (v - lo) / (hi - lo));
        pts = @(v) strjoin(arrayfun(@(a, b) sprintf('%.1f,%.1f', x(a), y(b)), t, v, 'UniformOutput', false), ' ');
        w('<h2>Min / max mapped temperature vs time [%s]</h2>\n', units);
        w('<svg width="%d" height="%d">', W, Hh);
        for g = 0:4
            yy = y(lo + (hi - lo) * g / 4);
            w('<line x1="%d" y1="%.1f" x2="%d" y2="%.1f" stroke="#ddd"/><text x="%d" y="%.1f" font-size="11" text-anchor="end">%.1f</text>', ...
              ml, yy, W - mr, yy, ml - 6, yy + 4, lo + (hi - lo) * g / 4);
        end
        for g = 0:5
            tt = t(1) + (t(end) - t(1)) * g / 5;
            w('<text x="%.1f" y="%d" font-size="11" text-anchor="middle">%g</text>', x(tt), Hh - mb + 16, tt);
        end
        w('<polyline fill="none" stroke="#c22" stroke-width="2" points="%s"/>', pts(tmax));
        w('<polyline fill="none" stroke="#25c" stroke-width="2" points="%s"/>', pts(tmin));
        w('<text x="%d" y="%d" font-size="12" fill="#c22">max</text><text x="%d" y="%d" font-size="12" fill="#25c">min</text>', ...
          W - mr - 60, mt + 12, W - mr - 30, mt + 12);
        w('<text x="%.1f" y="%d" font-size="12" text-anchor="middle">time [s]</text></svg>\n', (ml + W - mr) / 2, Hh - 4);
    end

    % ---- per-step table --------------------------------------------------------------
    w('<h2>Time steps (whole assembly)</h2><table><tr><th>File</th><th>Time</th><th>SID</th><th>Grids</th><th>Outside hull</th><th>Far</th><th>Tmin [%s]</th><th>Tmax [%s]</th><th>Method</th></tr>\n', units, units);
    for k = 1:ns
        r = RM(k);
        w('<tr><td>%s</td><td>%g</td><td>%d</td><td>%d</td><td>%d</td><td>%d</td><td>%.2f</td><td>%.2f</td><td style="text-align:left">%s</td></tr>\n', ...
          esc(shortname(r.csv_file)), r.time, r.sid, numel(r.grid_ids), nnz(r.extrap), nnz(r.far), ...
          from_kelvin(min(r.grid_T), units), from_kelvin(max(r.grid_T), units), esc(r.method));
    end
    w('</table>\n');
    if np > 1
        w('<details><summary>Per part</summary><table><tr><th>Part</th><th>File</th><th>Time</th><th>SID</th><th>Grids</th><th>Outside</th><th>Tmin</th><th>Tmax</th><th>Output</th></tr>\n');
        for q = 1:height(S)
            w('<tr><td>%s</td><td>%s</td><td>%g</td><td>%d</td><td>%d</td><td>%d</td><td>%.2f</td><td>%.2f</td><td style="text-align:left">%s</td></tr>\n', ...
              esc(S.Part{q}), esc(S.File{q}), S.Time(q), S.SID(q), S.Grids(q), S.Extrap(q), S.Tmin(q), S.Tmax(q), esc(S.OutFile{q}));
        end
        w('</table></details>\n');
    end

    % ---- pictures ---------------------------------------------------------------------
    if C.SAVE_PNG
        w('<h2>Check plots</h2>\n');
        for k = 1:ns
            [~, base] = fileparts(R(1, k).out_file);
            w('<details%s><summary>%s &nbsp; t = %g s</summary><img src="%s.png"></details>\n', ...
              ternary(k == 1, ' open', ''), esc(shortname(RM(k).csv_file)), RM(k).time, esc(base));
        end
    end
    w('</body></html>\n');
    fprintf('  wrote %s\n', f);
end


% =========================================================================
function S = summary_table(R, C)
%SUMMARY_TABLE  One row per part and time step (step-major order).
    [np, ns] = size(R);
    n = np * ns;
    Part = cell(n, 1); File = cell(n, 1); Time = zeros(n, 1); SID = zeros(n, 1);
    CloudPts = zeros(n, 1); Grids = zeros(n, 1); Extrap = zeros(n, 1);
    Tmin = zeros(n, 1); Tmax = zeros(n, 1); OutFile = cell(n, 1);
    q = 0;
    for k = 1:ns
        for i = 1:np
            q = q + 1;
            r = R(i, k);
            Part{q}     = r.part;
            File{q}     = shortname(r.csv_file);
            Time(q)     = r.time;
            SID(q)      = r.sid;
            CloudPts(q) = numel(r.cloud_T);
            Grids(q)    = numel(r.grid_ids);
            Extrap(q)   = nnz(r.extrap);
            Tmin(q)     = from_kelvin(min(r.grid_T), C.OUT_UNITS);
            Tmax(q)     = from_kelvin(max(r.grid_T), C.OUT_UNITS);
            OutFile{q}  = r.out_file;
        end
    end
    S = table(Part, File, Time, SID, CloudPts, Grids, Extrap, Tmin, Tmax, OutFile);
end


function out = ternary(cond, a, b)
    if cond, out = a; else, out = b; end
end


function s = shortname(p)
    [~, b, e] = fileparts(p);
    s = [b e];
end
