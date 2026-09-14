function [S, D, h] = temp_cloud_check(src, varargin)
%TEMP_CLOUD_CHECK  Is there a gradient worth meshing for?  Through-thickness vs in-plane, from the cloud alone.
%
%   Answers, before any BDF exists: isothermal, shell (in-plane gradient only),
%   or solid (through-thickness gradient too)?
%
%       >> temp_cloud_check('clouds\gasket')                       % every CSV in the folder
%       >> temp_cloud_check({'clouds\gasket\g_0.csv', ...})        % explicit files
%       >> [S, D, h] = temp_cloud_check(folder, 'CSV_LENGTH_UNITS', 'm', 'OUT_LENGTH_UNITS', 'in', ...
%                                       'CSV_TEMP_UNITS', 'K', 'OUT_UNITS', 'C')
%
%   METHOD
%     The part's thickness direction is the smallest principal axis of the point
%     cloud (PCA of the first step), or AXIS if you give one.  Points are binned
%     on an in-plane grid; within each cell the spread of T is the
%     THROUGH-THICKNESS delta, and the variation of the cell means across the
%     grid is the IN-PLANE delta.  A least-squares plane T = a + g.(x,y,z) gives
%     the global gradient split the same way (gfit_tt / gfit_ip) with its R^2.
%
%   OPTIONS (name/value)
%     CSV_LENGTH_UNITS  'in'      units of the CSV x/y/z          ('in' | 'mm' | 'm')
%     OUT_LENGTH_UNITS  'in'      units of the report (thickness, gradients per length)
%     CSV_TEMP_UNITS    'K'       units of the CSV T              ('K' | 'C' | 'F')
%     OUT_UNITS         'C'       units of the report
%     CSV_HAS_HEADER    true
%     BBOX              []        crop: [xmin ymin zmin; xmax ymax zmax] in CSV length units
%                                 (use when one cloud spans several parts)
%     PICK              ''        crop along one axis by clicking: 'x' | 'y' | 'z' | 'thickness'
%                                 (thickness = smallest PCA axis of the whole cloud).  Shows the
%                                 first step side-on plus a histogram along that axis; click
%                                 twice, everything between the clicks is kept for every step.
%     RANGE             []        [lo hi] along the PICK axis (OUT_LENGTH_UNITS): same crop, no
%                                 clicking -- the picked range is printed so you can reuse it
%     AXIS              []        thickness direction [x y z]; [] = PCA
%     CELL              []        in-plane bin size in OUT_LENGTH_UNITS; [] = auto (~60 pts / cell)
%     ISO_TOL           1         total delta T below this (OUT_UNITS) = "isothermal" in the verdict
%     PLOT              true      figure for the worst step (largest through-thickness delta)
%     MAX_POINTS        2e5       cloud points drawn
%
%   OUTPUTS
%     S   table, one row per step: File, Time, N, Thick, dT_total, dT_tt_95, dT_tt_max,
%         dT_tt_med, grad_tt, dT_ip, grad_ip, gfit_tt, gfit_ip, R2
%         (dT_tt_95 = 95th percentile of the per-cell through-thickness spread: the number
%         the verdict uses -- the max is usually a few edge cells; grad_tt = dT_tt_95 / Thick)
%     D   per-step detail (maps, axes, projected points) for your own plots
%     h   figure handle (PLOT)

% --- options --------------------------------------------------------------------
o = struct('CSV_LENGTH_UNITS', 'in', 'OUT_LENGTH_UNITS', 'in', 'CSV_TEMP_UNITS', 'K', 'OUT_UNITS', 'C', ...
           'CSV_HAS_HEADER', true, 'BBOX', [], 'PICK', '', 'RANGE', [], 'AXIS', [], 'CELL', [], ...
           'ISO_TOL', 1, 'PLOT', true, 'MAX_POINTS', 2e5);
for q = 1:2:numel(varargin)
    name = upper(char(varargin{q}));
    if ~isfield(o, name), error('temp_cloud_check:badOption', 'Unknown option "%s".', varargin{q}); end
    o.(name) = varargin{q + 1};
end
lunits = lower(char(o.OUT_LENGTH_UNITS));
tunits = upper(char(o.OUT_UNITS));
lf = length_factor(o.CSV_LENGTH_UNITS, lunits);

% --- files ------------------------------------------------------------------------
if iscell(src) || isstring(src) && numel(src) > 1
    files = cellstr(src); files = files(:);
else
    src = char(src);
    if exist(src, 'dir') == 7
        d = dir(fullfile(src, '*.csv'));
        if isempty(d), error('temp_cloud_check:noCsv', 'No *.csv files in %s', src); end
        names = {d.name};
        keys = regexprep(names, '(\d+)', '${sprintf(''%012d'', str2double($1))}');
        [~, order] = sort(lower(keys));
        files = fullfile(src, names(order))';
    elseif exist(src, 'file') == 2
        files = {src};
    else
        error('temp_cloud_check:noSrc', 'Not a folder or CSV file: %s', src);
    end
end
ns = numel(files);

% --- per step ---------------------------------------------------------------------
D = repmat(struct('file', '', 'time', NaN, 'xyz', [], 'T', [], 'axes', [], 'thick', NaN, ...
                  'cell', NaN, 'u', [], 'v', [], 'Tmean', [], 'dTtt', [], 'count', []), ns, 1);
File = cell(ns, 1); Time = zeros(ns, 1); N = zeros(ns, 1); Thick = zeros(ns, 1);
dT_total = zeros(ns, 1); dT_tt_95 = zeros(ns, 1); dT_tt_max = zeros(ns, 1); dT_tt_med = zeros(ns, 1); grad_tt = zeros(ns, 1);
dT_ip = zeros(ns, 1); grad_ip = zeros(ns, 1); gfit_tt = zeros(ns, 1); gfit_ip = zeros(ns, 1); R2 = zeros(ns, 1);
E = [];                                   % [e1 e2 e3] columns, fixed from the first step
crop = [];                                % {axis vector, [lo hi], centre} from PICK / RANGE
for k = 1:ns
    [t, xyz, T] = read_cloud(files{k}, o.CSV_HAS_HEADER);
    xyz = xyz * lf;
    T = from_kelvin(to_kelvin(T, o.CSV_TEMP_UNITS), tunits);
    if ~isempty(o.BBOX)
        bb = o.BBOX * lf;
        keep = all(xyz >= bb(1, :) & xyz <= bb(2, :), 2);
        xyz = xyz(keep, :); T = T(keep);
        if isempty(T), error('temp_cloud_check:emptyBox', '%s: no cloud points inside BBOX.', files{k}); end
    end
    if ~isempty(o.PICK)
        if isempty(crop), crop = pick_range(xyz, T, o, lunits, tunits); end
        w = (xyz - crop.ctr) * crop.dir;
        keep = w >= crop.range(1) & w <= crop.range(2);
        xyz = xyz(keep, :); T = T(keep);
        if isempty(T), error('temp_cloud_check:emptyRange', '%s: no cloud points inside the picked range.', files{k}); end
    end
    n = numel(T);
    if isempty(E), E = part_axes(xyz, o.AXIS); end
    ctr = mean(xyz, 1);
    P = (xyz - ctr) * E;                  % columns: in-plane u, in-plane v, thickness w
    thick = max(P(:, 3)) - min(P(:, 3));

    % in-plane bins
    ru = max(P(:, 1)) - min(P(:, 1)); rv = max(P(:, 2)) - min(P(:, 2));
    if isempty(o.CELL)
        cs = sqrt(60 * max(ru * rv, eps) / n);       % ~60 points per cell: a map, not speckle
        cs = max(cs, max(ru, rv) / 150);
    else
        cs = o.CELL;
    end
    iu = floor((P(:, 1) - min(P(:, 1))) / cs) + 1;  nu = max(iu);
    iv = floor((P(:, 2) - min(P(:, 2))) / cs) + 1;  nv = max(iv);
    cnt  = accumarray([iu iv], 1,  [nu nv]);
    Tsum = accumarray([iu iv], T,  [nu nv]);
    Tmax = accumarray([iu iv], T,  [nu nv], @max, -Inf);
    Tmin = accumarray([iu iv], T,  [nu nv], @min,  Inf);
    Tmean = Tsum ./ cnt; Tmean(cnt == 0) = NaN;
    dTtt = Tmax - Tmin; dTtt(cnt < 2) = NaN;

    % through-thickness: spread inside a cell (cells with a real sample only)
    good = cnt >= 8;
    if nnz(good) < 4, good = cnt >= 2; end
    tts = sort(dTtt(good));
    if isempty(tts)
        ttmax = 0; tt95 = 0; ttmed = 0;
    else
        ttmax = tts(end); ttmed = median(tts);
        tt95 = tts(max(1, ceil(0.95 * numel(tts))));
    end
    % in-plane: cell means, and the steepest step between neighbouring cells
    ipd = max(Tmean(:), [], 'omitnan') - min(Tmean(:), [], 'omitnan');
    gu = abs(diff(Tmean, 1, 1)) / cs; gv = abs(diff(Tmean, 1, 2)) / cs;
    gip = max([gu(:); gv(:); 0], [], 'omitnan');
    % global linear fit in the part frame
    A = [ones(n, 1) P];
    c = A \ T;
    res = T - A * c;
    r2 = 1 - sum(res.^2) / max(sum((T - mean(T)).^2), eps);

    File{k} = shortname(files{k}); Time(k) = t; N(k) = n; Thick(k) = thick;
    dT_total(k) = max(T) - min(T);
    dT_tt_95(k) = tt95; dT_tt_max(k) = ttmax; dT_tt_med(k) = ttmed; grad_tt(k) = tt95 / max(thick, eps);
    dT_ip(k) = ipd; grad_ip(k) = gip;
    gfit_tt(k) = abs(c(4)); gfit_ip(k) = norm(c(2:3)); R2(k) = r2;
    D(k) = struct('file', files{k}, 'time', t, 'xyz', xyz, 'T', T, 'axes', E, 'thick', thick, ...
                  'cell', cs, 'u', min(P(:, 1)) + (0.5:nu) * cs, 'v', min(P(:, 2)) + (0.5:nv) * cs, ...
                  'Tmean', Tmean, 'dTtt', dTtt, 'count', cnt);
    fprintf('  %-28s t=%-8g n=%-8d thick=%-8.4g dT=%-7.2f tt: 95%% %-6.2f max %-6.2f med %-6.2f (%.3g/%s)   ip: %-7.2f (%.3g/%s)  fit tt %.3g ip %.3g R2 %.2f\n', ...
        File{k}, t, n, thick, dT_total(k), tt95, ttmax, ttmed, grad_tt(k), lunits, ipd, gip, lunits, gfit_tt(k), gfit_ip(k), r2);
end
S = table(File, Time, N, Thick, dT_total, dT_tt_95, dT_tt_max, dT_tt_med, grad_tt, dT_ip, grad_ip, gfit_tt, gfit_ip, R2);

% --- verdict ------------------------------------------------------------------------
[~, kw] = max(dT_tt_95); [~, ki] = max(dT_ip); [~, kt] = max(dT_total);
fprintf('\nThickness axis [%.3f %.3f %.3f], thickness %.4g %s, in-plane cell %.4g %s\n', E(:, 3), Thick(1), lunits, D(1).cell, lunits);
fprintf('Worst total delta T      : %.2f %s  (%s)\n', dT_total(kt), tunits, File{kt});
fprintf('Worst through-thickness  : %.2f %s (95%% of cells; max %.2f, median %.2f) = %.3g %s/%s  (%s)\n', ...
        dT_tt_95(kw), tunits, dT_tt_max(kw), dT_tt_med(kw), grad_tt(kw), tunits, lunits, File{kw});
fprintf('Worst in-plane           : %.2f %s, steepest %.3g %s/%s  (%s)\n', dT_ip(ki), tunits, grad_ip(ki), tunits, lunits, File{ki});
if max(dT_total) < o.ISO_TOL
    verdict = sprintf('ISOTHERMAL: total delta T never exceeds %.2g %s -> one temperature per step.', o.ISO_TOL, tunits);
elseif max(dT_tt_95) < o.ISO_TOL
    verdict = sprintf('IN-PLANE ONLY: through-thickness delta < %.2g %s in 95%% of cells -> shells + TEMP (one T per grid).', o.ISO_TOL, tunits);
elseif max(dT_tt_95) < 0.25 * max(dT_ip)
    verdict = sprintf('IN-PLANE DOMINATES: through-thickness delta %.2f %s is < 1/4 of the in-plane %.2f %s -> shells + TEMP.', ...
                      max(dT_tt_95), tunits, max(dT_ip), tunits);
else
    verdict = sprintf('THROUGH-THICKNESS MATTERS: %.2f %s through the thickness vs %.2f %s in-plane -> solids.', ...
                      max(dT_tt_95), tunits, max(dT_ip), tunits);
end
fprintf('%s\n', verdict);
S.Properties.Description = verdict;

% --- figure: worst step ---------------------------------------------------------------
h = [];
if o.PLOT
    d = D(kw);
    h = figure('Name', sprintf('Cloud check: %s', File{kw}), 'NumberTitle', 'off', 'Color', 'w', ...
               'Position', [60 60 1500 560]);
    colormap(h, 'jet');
    tl = tiledlayout(h, 1, 3, 'TileSpacing', 'compact', 'Padding', 'compact');
    title(tl, {sprintf('%s    t = %g s    thickness %.4g %s    cell %.3g %s', File{kw}, d.time, d.thick, lunits, d.cell, lunits), ...
               verdict}, 'Interpreter', 'none', 'FontWeight', 'bold');

    ax1 = nexttile(tl);
    sel = thin(numel(d.T), o.MAX_POINTS);
    scatter3(ax1, d.xyz(sel, 1), d.xyz(sel, 2), d.xyz(sel, 3), 4, d.T(sel), 'filled');
    axis(ax1, 'equal'); axis(ax1, 'vis3d'); grid(ax1, 'on'); box(ax1, 'on'); view(ax1, 3);
    xlabel(ax1, 'X'); ylabel(ax1, 'Y'); zlabel(ax1, 'Z');
    cb = colorbar(ax1, 'eastoutside'); cb.Label.String = sprintf('T [%s]', tunits);
    hold(ax1, 'on');
    c0 = mean(d.xyz, 1); L = 0.5 * max(max(d.xyz) - min(d.xyz));
    quiver3(ax1, c0(1), c0(2), c0(3), L * E(1, 3), L * E(2, 3), L * E(3, 3), 0, 'k', 'LineWidth', 2, 'MaxHeadSize', 0.5);
    hold(ax1, 'off');
    title(ax1, {'cloud, arrow = thickness axis', sprintf('total delta T %.2f %s', dT_total(kw), tunits)});

    ax2 = nexttile(tl);
    imagesc(ax2, d.u, d.v, smooth_nan(d.Tmean)', 'AlphaData', ~isnan(d.Tmean'));
    set(ax2, 'YDir', 'normal'); axis(ax2, 'equal', 'tight'); box(ax2, 'on');
    cb = colorbar(ax2, 'southoutside'); cb.Label.String = sprintf('T [%s]', tunits);
    xlabel(ax2, sprintf('in-plane u [%s]', lunits)); ylabel(ax2, sprintf('in-plane v [%s]', lunits));
    title(ax2, {'IN-PLANE: mean T through the thickness', ...
                sprintf('delta %.2f %s, steepest %.3g %s/%s', dT_ip(kw), tunits, grad_ip(kw), tunits, lunits)});

    ax3 = nexttile(tl);
    imagesc(ax3, d.u, d.v, smooth_nan(d.dTtt)', 'AlphaData', ~isnan(d.dTtt'));
    set(ax3, 'YDir', 'normal'); axis(ax3, 'equal', 'tight'); box(ax3, 'on');
    caxis(ax3, [0 max(dT_tt_95(kw), eps)]);                       % edge cells saturate, the map stays readable
    cb = colorbar(ax3, 'southoutside'); cb.Label.String = sprintf('delta T [%s]  (clipped at the 95%% value)', tunits);
    xlabel(ax3, sprintf('in-plane u [%s]', lunits)); ylabel(ax3, sprintf('in-plane v [%s]', lunits));
    title(ax3, {'THROUGH-THICKNESS: spread of T within each cell', ...
                sprintf('95%% of cells < %.2f %s (median %.2f, max %.2f) = %.3g %s/%s', ...
                        dT_tt_95(kw), tunits, dT_tt_med(kw), dT_tt_max(kw), grad_tt(kw), tunits, lunits)});
end
if nargout == 0
    disp(S); clear S D h
end
end


% =========================================================================
function crop = pick_range(xyz, T, o, lunits, tunits)
%PICK_RANGE  Crop axis + [lo hi] along it, from RANGE or from two clicks on a side view.
    ax = lower(char(o.PICK));
    ctr = mean(xyz, 1);
    switch ax
        case 'x', dirn = [1 0 0]'; label = 'X';
        case 'y', dirn = [0 1 0]'; label = 'Y';
        case 'z', dirn = [0 0 1]'; label = 'Z';
        case {'thickness', 't', 'pca'}
            E = part_axes(xyz, []); dirn = E(:, 3); label = sprintf('thickness axis [%.2f %.2f %.2f]', dirn);
        otherwise
            error('temp_cloud_check:badPick', 'PICK must be ''x'', ''y'', ''z'' or ''thickness''.');
    end
    if strcmp(ax, 'x') || strcmp(ax, 'y') || strcmp(ax, 'z'), ctr = [0 0 0]; end   % world axes: absolute coords
    w = (xyz - ctr) * dirn;
    if ~isempty(o.RANGE)
        crop = struct('dir', dirn, 'ctr', ctr, 'range', sort(o.RANGE(:))');
        fprintf('Crop along %s: keeping %.4g .. %.4g %s (%d of %d points)\n', label, crop.range, lunits, ...
                nnz(w >= crop.range(1) & w <= crop.range(2)), numel(w));
        return
    end
    % side view: the crop axis horizontal, the widest perpendicular direction vertical
    [~, i] = min(abs(dirn)); a = zeros(3, 1); a(i) = 1;
    p1 = cross(dirn, a); p1 = p1 / norm(p1); p2 = cross(dirn, p1);
    Q = (xyz - ctr) * [p1 p2];
    [~, j] = max(max(Q) - min(Q)); v = Q(:, j);
    sel = thin(numel(T), o.MAX_POINTS);
    f = figure('Name', 'Cloud check: click twice to set the range to keep', 'NumberTitle', 'off', 'Color', 'w', ...
               'Position', [120 120 1100 620]);
    colormap(f, 'jet');
    a1 = subplot(3, 1, [1 2], 'Parent', f);
    scatter(a1, w(sel), v(sel), 5, T(sel), 'filled'); axis(a1, 'equal'); grid(a1, 'on'); colorbar(a1);
    xlabel(a1, sprintf('position along %s [%s]', label, lunits)); ylabel(a1, sprintf('perpendicular [%s]', lunits));
    prompt = @(k) title(a1, sprintf('Side view, T [%s].   Click %d of 2 on either panel: the range between the two clicks is kept.   (Esc / Enter = cancel)', tunits, k));
    prompt(1);
    a2 = subplot(3, 1, 3, 'Parent', f);
    histogram(a2, w, 200); grid(a2, 'on');
    xlabel(a2, sprintf('position along %s [%s]', label, lunits)); ylabel(a2, 'points');
    linkaxes([a1 a2], 'x');
    px = [];
    while numel(px) < 2
        figure(f);
        [x, ~, button] = ginput(1);
        if isempty(x) || ~ishandle(f) || (button ~= 1 && button ~= 2 && button ~= 3)
            if ishandle(f), close(f); end
            error('temp_cloud_check:noPick', 'Range not picked (need two mouse clicks).');
        end
        px(end+1) = x; %#ok<AGROW>
        for a = [a1 a2]
            hold(a, 'on'); yl = ylim(a);
            plot(a, [x x], yl, 'k--', 'LineWidth', 1.5);
            hold(a, 'off');
        end
        prompt(numel(px) + 1);
        drawnow
    end
    rng_ = sort(px(:))';
    title(a1, sprintf('Keeping %.4g .. %.4g %s along %s  (reuse: ''PICK'', ''%s'', ''RANGE'', [%.6g %.6g])', ...
                      rng_, lunits, label, ax, rng_));
    drawnow
    crop = struct('dir', dirn, 'ctr', ctr, 'range', rng_);
    fprintf('Crop along %s: keeping %.4g .. %.4g %s (%d of %d points).  Reuse with ''PICK'', ''%s'', ''RANGE'', [%.6g %.6g]\n', ...
            label, rng_, lunits, nnz(w >= rng_(1) & w <= rng_(2)), numel(w), ax, rng_);
end


function E = part_axes(xyz, ax3)
%PART_AXES  [e1 e2 e3]: e3 = thickness (given, or the smallest principal axis).
    if isempty(ax3)
        X = xyz - mean(xyz, 1);
        if size(X, 1) > 2e5, X = X(randperm(size(X, 1), 2e5), :); end
        [~, ~, V] = svd(X, 'econ');           % columns ordered by variance, largest first
        E = V;
    else
        e3 = ax3(:) / norm(ax3);
        [~, i] = min(abs(e3));                 % pick a world axis far from e3 to build the plane
        a = zeros(3, 1); a(i) = 1;
        e1 = cross(e3, a); e1 = e1 / norm(e1);
        e2 = cross(e3, e1);
        E = [e1 e2 e3];
    end
    if det(E) < 0, E(:, 2) = -E(:, 2); end    % right-handed
end


function M = smooth_nan(M)
%SMOOTH_NAN  3x3 mean of the defined neighbours (display only); NaN cells stay NaN.
    ok = ~isnan(M);
    V = M; V(~ok) = 0;
    k = ones(3);
    num = conv2(V, k, 'same'); den = conv2(double(ok), k, 'same');
    Ms = num ./ max(den, 1);
    M(ok) = Ms(ok);
end


function sel = thin(n, nmax)
    if n <= nmax, sel = (1:n)'; else, rng(0); sel = sort(randperm(n, nmax))'; end
end


function [t, xyz, T] = read_cloud(csv_file, has_header)
%READ_CLOUD  time, xyz, T from a 5-column CSV (same reader as temp_map_matlab).
    M = [];
    fid = fopen(csv_file, 'r');
    if fid > 0
        cl = onCleanup(@() fclose(fid));
        try
            if has_header, fgetl(fid); end
            c = textscan(fid, '%f%f%f%f%f', 'Delimiter', ',', 'CollectOutput', true);
            M = c{1};
            if ~feof(fid), M = []; end
        catch
            M = [];
        end
        clear cl
    end
    if isempty(M)
        opts = detectImportOptions(csv_file, 'FileType', 'text');
        if ~has_header, opts.DataLines = [1 Inf]; end
        M = readmatrix(csv_file, opts);
    end
    if size(M, 2) < 5
        error('temp_cloud_check:badCsv', '%s: expected 5 columns (time,x,y,z,T) but found %d.', csv_file, size(M, 2));
    end
    M = M(:, 1:5);
    M(any(isnan(M), 2), :) = [];
    if isempty(M), error('temp_cloud_check:emptyCsv', '%s: no numeric rows.', csv_file); end
    t = M(1, 1); xyz = M(:, 2:4); T = M(:, 5);
end


function f = length_factor(from, to)
    m = struct('in', 0.0254, 'mm', 1e-3, 'm', 1);
    from = lower(char(from)); to = lower(char(to));
    if ~isfield(m, from) || ~isfield(m, to)
        error('temp_cloud_check:badLength', 'Length units must be in, mm or m (got %s -> %s).', from, to);
    end
    f = m.(from) / m.(to);
end


function T = to_kelvin(T, units)
    switch upper(char(units))
        case 'K'
        case 'C', T = T + 273.15;
        case 'F', T = (T - 32) * 5/9 + 273.15;
        otherwise, error('temp_cloud_check:badTemp', 'Temperature units must be K, C or F.');
    end
end


function T = from_kelvin(T, units)
    switch upper(char(units))
        case 'K'
        case 'C', T = T - 273.15;
        case 'F', T = (T - 273.15) * 9/5 + 32;
        otherwise, error('temp_cloud_check:badTemp', 'Temperature units must be K, C or F.');
    end
end


function s = shortname(p)
    [~, b, e] = fileparts(p);
    s = [b e];
end
