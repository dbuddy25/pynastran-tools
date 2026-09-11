function [R, S] = temp_map_matlab(varargin)
%TEMP_MAP_MATLAB  Map CSV temperature point clouds onto Nastran GRIDs, write TEMP cards.
%
%   TEMP_MAP_MATLAB with no arguments runs the CONFIG block below: it reads the
%   GRID coordinates out of the BDF (INCLUDEs followed), then for every CSV
%   point cloud interpolates the cloud temperature onto each grid and writes a
%   bulk-data file of TEMP cards -- one file, one SID, per CSV.
%
%   [R, S] = TEMP_MAP_MATLAB(...) also returns:
%       R   - results struct array, one element per CSV, carrying the grid ids,
%             coordinates, mapped temperatures (Kelvin) and the cloud itself,
%             so TEMP_MAP_PLOT can draw them without re-reading anything.
%       S   - summary table, one row per CSV (file, time, SID, counts, T range).
%
%   Any CONFIG field can be overridden per call without editing the file:
%       >> temp_map_matlab('BDF_FILE', 'wing.bdf', 'CSV_DIR', 'clouds', 'OUT_UNITS', 'C')
%
%   Parse a big BDF once and reuse the grids across several runs:
%       >> G = temp_map_matlab('BDF_FILE', 'wing.bdf', 'READ_ONLY', true);
%       >> temp_map_matlab('GRIDS', G, 'CSV_DIR', 'clouds_run1');
%       >> temp_map_matlab('GRIDS', G, 'CSV_DIR', 'clouds_run2', 'METHOD', 'idw');
%
%   CSV FORMAT (one file = one time step)
%   -------------------------------------
%       time_s, x, y, z, T               (header row optional, see CSV_HAS_HEADER)
%   x/y/z in CSV_LENGTH_UNITS (converted to BDF_LENGTH_UNITS before mapping),
%   basic frame; T in CSV_TEMP_UNITS (converted to Kelvin internally).
%
%   MAPPING  (C.METHOD)
%   -------------------
%   'linear'  - linear interpolation on a Delaunay triangulation of the cloud
%               (what scatteredInterpolant does), nearest point outside the
%               cloud's convex hull.  Planar / collinear clouds are projected
%               first.  Grids outside the hull are flagged in R(k).extrap.
%               Exact for smooth fields, but the 3D triangulation costs ~1 min
%               and several GB at 1M cloud points.
%   'scattered' - MATLAB's scatteredInterpolant(P, T, 'linear', 'nearest'),
%               verbatim: same maths as 'linear' but without the bridging-cell
%               guard, for one-to-one comparison with other scripts.
%   'nearest' - each grid takes the temperature of its closest cloud point.
%   'idw'     - inverse-distance-weighted mean of the IDW_K closest points.
%   'nearest' and 'idw' use a kd-tree (knnsearch, Statistics Toolbox) and take
%   seconds at 1M x 1M; without that toolbox they fall back to the
%   triangulation.  R(k).nn_dist always holds each grid's distance to its
%   nearest cloud point; EXTRAP_WARN_DIST turns that into R(k).far.
%
%   Pure MATLAB -- no Python, no pyNastran.  GRIDs are assumed to be in the
%   basic coordinate system (CP = 0 or blank); a warning is issued otherwise.
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
C.CSV_FILES      = {};          % {} = every *.csv in CSV_DIR (natural-sorted),
                                % or an explicit cell list of file names
C.CSV_HAS_HEADER = true;        % first CSV row is a header

% --- units ---------------------------------------------------------------
%   Cloud coordinates are converted INTO the BDF's length units before mapping.
C.BDF_LENGTH_UNITS = 'in';      % 'in' | 'ft' | 'mm' | 'cm' | 'm'
C.CSV_LENGTH_UNITS = 'in';      % 'in' | 'ft' | 'mm' | 'cm' | 'm'
C.CSV_TEMP_UNITS   = 'K';       % 'K' | 'C' | 'F'   (what the CSV's 5th column is)

% --- output --------------------------------------------------------------
C.OUT_DIR        = 'temp_cards';    % created if missing
C.OUT_UNITS      = 'K';             % 'K' | 'C'   (C = K - 273.15)
C.FIELD_SIZE     = 8;               % 8 = small field | 16 = large field
C.WRITE_TEMPD    = false;           % also emit TEMPD,SID,<mean T> for unlisted grids
C.WRITE          = true;            % false = map only, write nothing (GUI preview)
C.SAVE_PNG       = false;           % true = save an ISO-view PNG next to each .bdf

% --- load set IDs --------------------------------------------------------
C.SID_START      = 1;           % SID = SID_START + (file index - 1) ...
C.SID_FROM_TIME  = false;       % ... or SID = round(time * TIME_SCALE)
C.TIME_SCALE     = 1;

% --- mapping -------------------------------------------------------------
C.METHOD         = 'linear';    % 'linear' | 'scattered' | 'nearest' | 'idw'   (see help)
C.IDW_K          = 8;           % neighbours used by 'idw'
C.IDW_POWER      = 2;           % 1/d^p weighting for 'idw'
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

if isempty(C.RESULTS)
    if ~isempty(C.GRIDS) && ~C.READ_ONLY
        gid = C.GRIDS.ids; gxyz = C.GRIDS.xyz; gfaces = C.GRIDS.faces;
        fprintf('Using %d cached grids from %s\n', numel(gid), C.GRIDS.bdf_file);
    else
        tic;
        progress(C, 0, sprintf('Reading GRIDs from %s ...', shortname(C.BDF_FILE)));
        [gid, gxyz] = read_grids(C.BDF_FILE);
        fprintf('Read %d grids from %s  (%.1f s)\n', numel(gid), C.BDF_FILE, toc);
        tic;
        progress(C, 0.03, 'Reading elements for the contour view ...');
        gfaces = read_faces(C.BDF_FILE, gid);
        fprintf('Read %d drawable element faces  (%.1f s)\n', size(gfaces, 1), toc);
    end
    if C.READ_ONLY
        R = struct('ids', gid, 'xyz', gxyz, 'faces', gfaces, 'bdf_file', C.BDF_FILE);
        S = [];
        progress(C, 1, 'Done.');
        return
    end

    files = resolve_csv_files(C);
    nf = numel(files);
    R = repmat(empty_result(), nf, 1);
    for k = 1:nf
        tic;
        f0 = 0.05 + 0.95 * (k - 1) / nf;          % this file's share of the bar
        fw = 0.95 / nf;
        stage = @(frac, msg) progress(C, f0 + fw * frac, ...
                    sprintf('[%d/%d] %s: %s', k, nf, shortname(files{k}), msg));
        stage(0, 'reading CSV ...');
        [t, cxyz, cT] = read_cloud(files{k}, C.CSV_HAS_HEADER);
        cxyz = cxyz * length_factor(C.CSV_LENGTH_UNITS, C.BDF_LENGTH_UNITS);
        cT   = to_kelvin(cT, C.CSV_TEMP_UNITS);
        [gT, extrap, nndist, method] = map_temps(cxyz, cT, gxyz, C, stage);
        r = empty_result();
        r.csv_file  = files{k};
        r.bdf_file  = C.BDF_FILE;
        if ~isempty(C.GRIDS), r.bdf_file = C.GRIDS.bdf_file; end
        r.time      = t;
        r.sid       = assign_sid(C, k, t);
        r.grid_ids  = gid;
        r.grid_xyz  = gxyz;
        r.faces     = gfaces;
        r.grid_T    = gT;         % Kelvin, always
        r.cloud_xyz = cxyz;
        r.cloud_T   = cT;         % Kelvin, always
        r.extrap    = extrap;
        r.nn_dist   = nndist;
        r.far       = false(size(nndist));
        r.method    = method;
        stage(0.95, 'cloud surface (alpha shape) ...');
        r.surface   = cloud_surface(cxyz, C.SURFACE_POINTS);
        if ~isempty(C.EXTRAP_WARN_DIST)
            r.far = nndist > C.EXTRAP_WARN_DIST;
            if any(r.far)
                warning('temp_map_matlab:farGrids', ...
                    '%s: %d grid(s) are farther than %g from any cloud point (max %.4g).', ...
                    files{k}, nnz(r.far), C.EXTRAP_WARN_DIST, max(nndist));
            end
        end
        r.coverage  = coverage_report(r);
        R(k) = r;
        fprintf('  %-32s t=%-9.4g cloud=%-8d outside=%-6d far=%-6d T=[%.2f %.2f] K  %s  (%.1f s)\n', ...
            shortname(files{k}), t, numel(cT), nnz(extrap), nnz(r.far), min(gT), max(gT), method, toc);
        fprintf('      %s\n', r.coverage{:});
    end
else
    R = C.RESULTS(:);       % SIDs were assigned when these were mapped
end

if C.WRITE
    if exist(C.OUT_DIR, 'dir') ~= 7, mkdir(C.OUT_DIR); end
    for k = 1:numel(R)
        tic;
        [~, base] = fileparts(R(k).csv_file);
        R(k).out_file = fullfile(C.OUT_DIR, [base '_temp.bdf']);
        write_temp_cards(R(k), C);
        fprintf('  wrote %s  (%.1f s)\n', R(k).out_file, toc);
        if C.SAVE_PNG
            h = temp_map_plot(R(k), 'Visible', 'off', 'Units', C.OUT_UNITS);
            saveas(h.fig, fullfile(C.OUT_DIR, [base '_temp.png']));
            close(h.fig);
        end
    end
end

progress(C, 1, 'Done.');
S = summary_table(R, C);
if nargout == 0
    disp(S);
    clear R S
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
        if isempty(C.GRIDS) && exist(C.BDF_FILE, 'file') ~= 2
            error('temp_map_matlab:noFile', 'BDF file not found: %s', C.BDF_FILE);
        end
        if ~isempty(C.GRIDS) && ~all(isfield(C.GRIDS, {'ids', 'xyz', 'faces', 'bdf_file'}))
            error('temp_map_matlab:badGrids', 'C.GRIDS must come from a READ_ONLY call.');
        end
        if C.READ_ONLY, return; end
        if isempty(C.CSV_FILES) && exist(C.CSV_DIR, 'dir') ~= 7
            error('temp_map_matlab:noDir', 'CSV folder not found: %s', C.CSV_DIR);
        end
    end
    length_factor(C.CSV_LENGTH_UNITS, C.BDF_LENGTH_UNITS);   % errors on a bad name
    to_kelvin(0, C.CSV_TEMP_UNITS);
    if ~ismember(upper(char(C.OUT_UNITS)), {'K', 'C'})
        error('temp_map_matlab:badUnits', 'C.OUT_UNITS must be ''K'' or ''C''.');
    end
    if ~ismember(C.FIELD_SIZE, [8 16])
        error('temp_map_matlab:badField', 'C.FIELD_SIZE must be 8 or 16.');
    end
    if ~ismember(lower(char(C.METHOD)), {'linear', 'scattered', 'nearest', 'idw'})
        error('temp_map_matlab:badMethod', 'C.METHOD must be ''linear'', ''scattered'', ''nearest'' or ''idw''.');
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
    r = struct('csv_file', '', 'bdf_file', '', 'time', NaN, 'sid', NaN, ...
               'grid_ids', [], 'grid_xyz', [], 'grid_T', [], ...
               'cloud_xyz', [], 'cloud_T', [], 'extrap', [], 'nn_dist', [], ...
               'far', [], 'method', '', 'surface', [], 'coverage', {{}}, 'faces', [], 'out_file', '');
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
function [ids, xyz] = read_grids(bdf_file)
%READ_GRIDS  GRID id and basic-frame xyz from a BDF, following INCLUDEs.
%   Handles small-field, large-field (GRID*) and free-field (comma) cards.
    [ids, xyz, ncp] = read_grids_file(bdf_file, true);
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


function [ids, xyz, ncp] = read_grids_file(fname, is_top)
%   Vectorised: the whole deck is split into lines once, GRID lines are picked
%   out with one regexp, and small/large-field cards are sliced as char
%   matrices.  Only free-field GRIDs and INCLUDEs are handled line by line.
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
    lines = lines(first:last);
    up    = up(first:last);
    n     = numel(lines);

    ids = zeros(0, 1); xyz = zeros(0, 3); ncp = 0; lineno = zeros(0, 1);

    % ---- GRID lines --------------------------------------------------------
    isg = ~cellfun('isempty', regexp(cellstr(up), '^GRID\*?\s*(,|\s|$)', 'once'));
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
    for i = find(startsWith(up, "INCLUDE"))'
        rest = regexprep(char(lines(i)), '^\s*INCLUDE\s*', '', 'ignorecase');
        j = i + 1;
        while nnz(rest == '''') < 2 && j <= n   % unclosed quote: join next line
            rest = [rest strtrim(char(lines(j)))];    %#ok<AGROW>
            j = j + 1;
        end
        inc = strtrim(regexprep(strtrim(rest), '^''|''$', ''));
        [i2, x2, c2] = read_grids_file(resolve_include(inc, base_dir), false);
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
function F = read_faces(bdf_file, grid_ids)
%READ_FACES  Drawable faces (rows into grid_ids, NaN-padded to 4 columns)
%   from the shell and solid elements in the BDF: shells as-is, solids as
%   their free (outer) faces.  [] if the deck has no supported elements.
    try
        E = read_elements_file(bdf_file, true);
    catch ME
        warning('temp_map_matlab:elements', 'Element read failed (%s); contour view unavailable.', ME.message);
        F = []; return
    end
    if isempty(E), F = []; return; end
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
    % node ids -> rows of grid_ids; drop faces touching unknown grids
    [tf, loc] = ismember(faces, grid_ids);
    ok = all(tf | isnan(faces), 2);
    F = loc(ok, :);
    F(isnan(faces(ok, :))) = NaN;
    F = double(F);
end


function E = read_elements_file(fname, is_top)
%READ_ELEMENTS_FILE  Corner-node ids per supported element type, following
%   INCLUDEs.  Same vectorised slicing as read_grids_file; continuation lines
%   are assumed to directly follow their parent (true for every mesher).
    types = {'CQUAD4', 4; 'CQUADR', 4; 'CQUAD8', 4; 'CTRIA3', 3; 'CTRIAR', 3; 'CTRIA6', 3; ...
             'CTETRA', 4; 'CPENTA', 6; 'CHEXA', 8};
    E = struct();
    for t = 1:size(types, 1), E.(types{t, 1}) = zeros(0, types{t, 2}); end

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
    cand = startsWith(up, "C");                       % cheap pre-filter
    name_of = strings(n, 1);
    name_of(cand) = regexp(up(cand), '^[A-Z0-9]+', 'match', 'once');
    name_of(ismissing(name_of)) = "";
    star = cand & startsWith(extractAfter(up, strlength(name_of)), "*");
    for t = 1:size(types, 1)
        name = types{t, 1}; ng = types{t, 2}; nf = 2 + ng;      % EID PID G1..Gng
        mine = find(name_of == name);
        if isempty(mine), continue; end
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
    end

    % ---- INCLUDEs ------------------------------------------------------------
    for i = find(startsWith(up, "INCLUDE"))'
        rest = regexprep(char(lines(i)), '^\s*INCLUDE\s*', '', 'ignorecase');
        j = i + 1;
        while nnz(rest == '''') < 2 && j <= n
            rest = [rest strtrim(char(lines(j)))];                   %#ok<AGROW>
            j = j + 1;
        end
        inc = strtrim(regexprep(strtrim(rest), '^''|''$', ''));
        E2 = read_elements_file(resolve_include(inc, base_dir), false);
        for t = 1:size(types, 1)
            E.(types{t, 1}) = [E.(types{t, 1}); E2.(types{t, 1})];
        end
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
    m = struct('in', 0.0254, 'ft', 0.3048, 'mm', 1e-3, 'cm', 1e-2, 'm', 1);
    from = lower(char(from)); to = lower(char(to));
    if ~isfield(m, from) || ~isfield(m, to)
        error('temp_map_matlab:badLength', 'Length units must be in, ft, mm, cm or m (got %s -> %s).', from, to);
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


% =========================================================================
function [t, xyz, T] = read_cloud(csv_file, has_header)
%READ_CLOUD  time, xyz, T(Kelvin) from a 5-column CSV.
    opts = detectImportOptions(csv_file, 'FileType', 'text');
    if ~has_header
        opts.DataLines = [1 Inf];
    end
    M = readmatrix(csv_file, opts);
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
function [Tg, extrap, nndist, method] = map_temps(P, T, Q, C, stage)
%MAP_TEMPS  Cloud (P,T) -> query points Q by C.METHOD.
%   linear : Delaunay barycentric interpolation, nearest outside the hull.
%            Planar / collinear clouds are projected down first.
%   nearest: closest cloud point (kd-tree if Statistics Toolbox, else the
%            triangulation's nearestNeighbor).
%   idw    : inverse-distance-weighted mean of the IDW_K closest points.
    if nargin < 5, stage = @(~, ~) []; end
    % duplicate cloud points would break the triangulation: average them
    [P, ~, ic] = unique(P, 'rows');
    T = accumarray(ic, T, [], @mean);
    nP = size(P, 1);
    nQ = size(Q, 1);
    meth = lower(char(C.METHOD));

    if nP == 1
        Tg = repmat(T, nQ, 1);
        extrap = true(nQ, 1);
        nndist = vecnorm(Q - P, 2, 2);
        method = 'single point';
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
        F = scatteredInterpolant(P, T, 'linear', 'nearest');
        stage(0.6, sprintf('evaluating at %d grids ...', nQ));
        Tg = F(Q);
        extrap = false(nQ, 1);
        if have_knn
            [~, nndist] = knn_chunked(P, Q, 1, stage, 'nearest distance');
        else
            nndist = nan(nQ, 1);       % not worth a second triangulation
        end
        method = 'scatteredInterpolant linear/nearest';
        Tg = Tg(:); nndist = nndist(:);
        return
    end
    if strcmp(meth, 'nearest') && have_knn
        [nn, nndist] = knn_chunked(P, Q, 1, stage, 'nearest');
        Tg = T(nn);
        extrap = false(nQ, 1);
        method = 'nearest (kd-tree)';
        return
    elseif strcmp(meth, 'idw') && have_knn
        k = min(C.IDW_K, nP);
        [nn, d] = knn_chunked(P, Q, k, stage, sprintf('IDW k=%d', k));
        w  = 1 ./ max(d, eps) .^ C.IDW_POWER;
        Tg = sum(w .* T(nn), 2) ./ sum(w, 2);
        hit = d(:, 1) == 0;                     % sitting on a cloud point
        Tg(hit) = T(nn(hit, 1));
        nndist = d(:, 1);
        extrap = false(nQ, 1);
        method = sprintf('IDW (k=%d, p=%g)', k, C.IDW_POWER);
        return
    end

    % ---- triangulation path --------------------------------------------------
    % effective dimension of the cloud via SVD of the centred points
    mu = mean(P, 1);
    [~, Sv, V] = svd(P - mu, 'econ');
    sv = diag(Sv);
    dim = nnz(sv > 1e-9 * sv(1));
    Pp = (P - mu) * V(:, 1:dim);
    Qp = (Q - mu) * V(:, 1:dim);

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
                Tg = T(nn);
                extrap = true(nQ, 1);
                method = 'nearest (brute force)';
            else
                stage(0.75, sprintf('locating %d grids in the triangulation ...', nQ));
                nn = nearestNeighbor(DT, Qp);
                Tg = T(nn);                                  % nearest everywhere ...
                if strcmp(meth, 'linear')
                    [ti, bc] = pointLocation(DT, Qp);
                    inside = ~isnan(ti);
                    tri = DT.ConnectivityList(ti(inside), :);    % (n_in x dim+1)
                    Tlin = sum(bc(inside, :) .* T(tri), 2);      % ... linear inside the hull
                    % bridging cells: Delaunay spans concavities of the body with
                    % long flat cells whose vertices are far-apart surface points;
                    % interpolating across those smears local hot/cold spots.
                    nbridge = 0;
                    if ~isempty(C.SLIVER_FACTOR) && any(inside)
                        E = edges(DT);
                        h = median(vecnorm(Pp(E(:, 1), :) - Pp(E(:, 2), :), 2, 2));
                        maxedge = zeros(size(tri, 1), 1);
                        for a = 1:size(tri, 2) - 1
                            for b = a + 1:size(tri, 2)
                                maxedge = max(maxedge, vecnorm(Pp(tri(:, a), :) - Pp(tri(:, b), :), 2, 2));
                            end
                        end
                        bridge = maxedge > C.SLIVER_FACTOR * h;
                        Tnear = Tg(inside);                  % nearest-point values
                        Tlin(bridge) = Tnear(bridge);
                        nbridge = nnz(bridge);
                    end
                    Tg(inside) = Tlin;
                    extrap = ~inside;
                    if dim == 3, method = 'linear (3D Delaunay)';
                    else,        method = 'linear (planar cloud)'; end
                    if nbridge > 0
                        method = sprintf('%s, %d grids in bridging cells -> nearest', method, nbridge);
                    end
                else
                    extrap = false(nQ, 1);
                    method = 'nearest (triangulation)';
                end
            end
        otherwise   % dim == 1: collinear cloud
            [s, order] = unique(Pp(:, 1));
            Ts = T(order);
            sq = min(max(Qp(:, 1), s(1)), s(end));
            nn = order(interp1(s, 1:numel(s), sq, 'nearest'));
            if strcmp(meth, 'linear')
                Tg = interp1(s, Ts, sq, 'linear');
                extrap = Qp(:, 1) < s(1) | Qp(:, 1) > s(end);
                method = 'linear (collinear cloud)';
            else
                Tg = T(nn);
                extrap = false(nQ, 1);
                method = 'nearest (collinear cloud)';
            end
    end
    nndist = vecnorm(Q - P(nn, :), 2, 2);
    Tg = Tg(:); extrap = extrap(:); nndist = nndist(:);
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
function write_temp_cards(r, C)
%WRITE_TEMP_CARDS  One bulk-data file of TEMP cards for one result.
%   Built as a char matrix and written in one go -- a 1M-grid file takes
%   seconds, not minutes.
    units = upper(char(C.OUT_UNITS));
    T = r.grid_T(:);
    if units == 'C', T = T - 273.15; end
    ids = r.grid_ids(:);
    n = numel(ids);
    w = C.FIELD_SIZE;

    fid = fopen(r.out_file, 'w');
    if fid < 0
        error('temp_map_matlab:cantWrite', 'Cannot open %s for writing.', r.out_file);
    end
    cl = onCleanup(@() fclose(fid));

    fprintf(fid, '$ TEMP cards generated by temp_map_matlab  (%s)\n', datestr(now, 'yyyy-mm-dd HH:MM'));
    fprintf(fid, '$ BDF    : %s\n', r.bdf_file);
    fprintf(fid, '$ CSV    : %s   (time = %g s)\n', r.csv_file, r.time);
    fprintf(fid, '$ Lengths: BDF in %s, cloud read as %s%s\n', C.BDF_LENGTH_UNITS, C.CSV_LENGTH_UNITS, ...
            ternary(strcmpi(C.BDF_LENGTH_UNITS, C.CSV_LENGTH_UNITS), '', ' (converted)'));
    fprintf(fid, '$ Units  : deg %s   grids: %d   outside cloud hull (nearest used): %d\n', ...
            units, n, nnz(r.extrap));
    fprintf(fid, '$ Method : %s\n', r.method);
    fprintf(fid, '$ Trange : %.3f .. %.3f\n', min(T), max(T));

    sidS = sprintf('%*d', w, r.sid);
    if C.WRITE_TEMPD
        if w == 8, fprintf(fid, 'TEMPD   %s%s\n',  sidS, fmt_real(mean(T), w));
        else,      fprintf(fid, 'TEMPD*  %s%s\n', sidS, fmt_real(mean(T), w)); end
    end

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
function S = summary_table(R, C)
    n = numel(R);
    File = cell(n, 1); Time = zeros(n, 1); SID = zeros(n, 1);
    CloudPts = zeros(n, 1); Grids = zeros(n, 1); Extrap = zeros(n, 1);
    Tmin = zeros(n, 1); Tmax = zeros(n, 1); OutFile = cell(n, 1);
    off = 0; if upper(char(C.OUT_UNITS)) == 'C', off = 273.15; end
    for k = 1:n
        File{k}     = shortname(R(k).csv_file);
        Time(k)     = R(k).time;
        SID(k)      = R(k).sid;
        CloudPts(k) = numel(R(k).cloud_T);
        Grids(k)    = numel(R(k).grid_ids);
        Extrap(k)   = nnz(R(k).extrap);
        Tmin(k)     = min(R(k).grid_T) - off;
        Tmax(k)     = max(R(k).grid_T) - off;
        OutFile{k}  = R(k).out_file;
    end
    S = table(File, Time, SID, CloudPts, Grids, Extrap, Tmin, Tmax, OutFile);
end


function out = ternary(cond, a, b)
    if cond, out = a; else, out = b; end
end


function s = shortname(p)
    [~, b, e] = fileparts(p);
    s = [b e];
end
