function [S, D, h] = temp_cloud_check(src, varargin)
%TEMP_CLOUD_CHECK  Is there a gradient worth meshing for?  Through-thickness vs in-plane, from the cloud alone.
%
%   Answers, before any BDF exists: isothermal, shell (in-plane gradient only),
%   or solid (through-thickness gradient too)?
%
%       >> temp_cloud_check('clouds\gasket')                       % every CSV in the folder
%       >> temp_cloud_check({'clouds\gasket\g_0.csv', ...})        % explicit files
%       >> temp_cloud_check(folder, 'SPLIT', 0.05)                 % two plates in one cloud: split
%                                                                  % bodies more than 0.05 apart
%       >> [S, D, h] = temp_cloud_check(folder, 'CSV_LENGTH_UNITS', 'm', 'OUT_LENGTH_UNITS', 'in', ...
%                                       'CSV_TEMP_UNITS', 'K', 'OUT_UNITS', 'C')
%
%   METHOD
%     SPLIT (optional) cuts the cloud into bodies: points are voxelised at the
%     SPLIT distance and voxels touching (26-connectivity) belong to one body,
%     so two plates with a gap wider than ~SPLIT come apart.  Each body is then
%     checked on its own.  Its thickness direction is the smallest principal
%     axis of the body's points (first step), or AXIS if you give one.  Points
%     are binned on an in-plane grid; within each cell the spread of T is the
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
%     SPLIT             []        gap (OUT_LENGTH_UNITS) that separates bodies; [] = one body
%     MIN_BODY          0.01      bodies with fewer points than this fraction are dropped
%     AXIS              []        thickness direction [x y z] for every body; [] = PCA per body
%     CELL              []        in-plane bin size in OUT_LENGTH_UNITS; [] = auto (~8 pts / cell)
%     ISO_TOL           1         delta T below this (OUT_UNITS) counts as nothing
%     PLOT              true      figure for the worst body / step (largest through-thickness delta)
%     MAX_POINTS        2e5       cloud points drawn
%
%   OUTPUTS
%     S   table, one row per body and step: Body, File, Time, N, Thick, dT_total, dT_tt_max,
%         dT_tt_med, grad_tt, dT_ip, grad_ip, gfit_tt, gfit_ip, R2
%     D   per-row detail (maps, axes, projected points) for your own plots
%     h   figure handle (PLOT)

% --- options --------------------------------------------------------------------
o = struct('CSV_LENGTH_UNITS', 'in', 'OUT_LENGTH_UNITS', 'in', 'CSV_TEMP_UNITS', 'K', 'OUT_UNITS', 'C', ...
           'CSV_HAS_HEADER', true, 'BBOX', [], 'SPLIT', [], 'MIN_BODY', 0.01, 'AXIS', [], 'CELL', [], ...
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
if iscell(src) || (isstring(src) && numel(src) > 1)
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

% --- per step, per body ------------------------------------------------------------
rows = struct('Body', {}, 'File', {}, 'Time', {}, 'N', {}, 'Thick', {}, 'dT_total', {}, 'dT_tt_max', {}, ...
              'dT_tt_med', {}, 'grad_tt', {}, 'dT_ip', {}, 'grad_ip', {}, 'gfit_tt', {}, 'gfit_ip', {}, 'R2', {});
D = struct('body', {}, 'file', {}, 'time', {}, 'xyz', {}, 'T', {}, 'axes', {}, 'thick', {}, ...
           'cell', {}, 'u', {}, 'v', {}, 'Tmean', {}, 'dTtt', {}, 'count', {});
body = [];                                % per-point body label, fixed from the first step
E = {};                                   % per-body [e1 e2 e3]
nb = 1;
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
    if isempty(body) || numel(body) ~= numel(T)
        if ~isempty(o.SPLIT)
            body = split_bodies(xyz, o.SPLIT, o.MIN_BODY);
            nb = max(body);
            fprintf('%s: %d body(ies) more than %g %s apart', shortname(files{k}), nb, o.SPLIT, lunits);
            fprintf('  [%s pts]\n', strjoin(arrayfun(@(b) sprintf('%d', nnz(body == b)), 1:nb, 'UniformOutput', false), ', '));
        else
            body = ones(numel(T), 1); nb = 1;
        end
        E = cell(1, nb);
    end
    for b = 1:nb
        ib = body == b;
        if isempty(E{b}), E{b} = part_axes(xyz(ib, :), o.AXIS); end
        [r, d] = check_body(xyz(ib, :), T(ib), E{b}, o.CELL);
        r.Body = b; r.File = shortname(files{k}); r.Time = t;
        d.body = b; d.file = files{k}; d.time = t;
        rows(end+1) = orderfields(r, rows);   %#ok<AGROW>
        D(end+1) = orderfields(d, D);         %#ok<AGROW>
        fprintf('  body %d  %-24s t=%-8g n=%-8d thick=%-8.4g dT=%-7.2f tt: max %-6.2f med %-6.2f (%.3g/%s)   ip: %-7.2f (%.3g/%s)  fit tt %.3g ip %.3g R2 %.2f\n', ...
            b, r.File, t, r.N, r.Thick, r.dT_total, r.dT_tt_max, r.dT_tt_med, r.grad_tt, lunits, ...
            r.dT_ip, r.grad_ip, lunits, r.gfit_tt, r.gfit_ip, r.R2);
    end
end
S = struct2table(rows);
if nb == 1, S.Body = []; end

% --- verdicts, one per body ----------------------------------------------------------
verdict = cell(nb, 1);
fprintf('\n');
for b = 1:nb
    ib = [rows.Body] == b;
    Sb = S(ib, :); Db = D(ib);
    [~, kw] = max(Sb.dT_tt_max); [~, ki] = max(Sb.dT_ip); [~, kt] = max(Sb.dT_total);
    if nb > 1, fprintf('BODY %d  (%d points)\n', b, Sb.N(1)); end
    fprintf('Thickness axis [%.3f %.3f %.3f], thickness %.4g %s, in-plane cell %.4g %s\n', ...
            E{b}(:, 3), Sb.Thick(1), lunits, Db(1).cell, lunits);
    fprintf('Worst total delta T      : %.2f %s  (%s)\n', Sb.dT_total(kt), tunits, Sb.File{kt});
    fprintf('Worst through-thickness  : %.2f %s = %.3g %s/%s  (%s)\n', Sb.dT_tt_max(kw), tunits, Sb.grad_tt(kw), tunits, lunits, Sb.File{kw});
    fprintf('Worst in-plane           : %.2f %s, steepest %.3g %s/%s  (%s)\n', Sb.dT_ip(ki), tunits, Sb.grad_ip(ki), tunits, lunits, Sb.File{ki});
    if max(Sb.dT_total) < o.ISO_TOL
        v = sprintf('ISOTHERMAL: total delta T never exceeds %.2g %s -> one temperature per step.', o.ISO_TOL, tunits);
    elseif max(Sb.dT_tt_max) < o.ISO_TOL
        v = sprintf('IN-PLANE ONLY: through-thickness delta < %.2g %s everywhere -> shells + TEMP (one T per grid) are enough.', o.ISO_TOL, tunits);
    elseif max(Sb.dT_tt_max) < 0.25 * max(Sb.dT_ip)
        v = 'IN-PLANE DOMINATES: through-thickness delta < 1/4 of the in-plane delta -> shells + TEMP, or TEMPP1 for the small through-thickness part.';
    else
        v = 'THROUGH-THICKNESS MATTERS: comparable to or larger than the in-plane delta -> solids (or shells with TEMPP1 gradients).';
    end
    if nb > 1, v = sprintf('Body %d: %s', b, v); end
    verdict{b} = v;
    fprintf('%s\n\n', v);
end
S.Properties.Description = strjoin(verdict, newline);

% --- figure: worst body / step ----------------------------------------------------------
h = [];
if o.PLOT
    [~, iw] = max(S.dT_tt_max);
    d = D(iw); Eb = d.axes;
    ttl = sprintf('Cloud check: %s', S.File{iw});
    if nb > 1, ttl = sprintf('%s  body %d of %d', ttl, d.body, nb); end
    h = figure('Name', ttl, 'NumberTitle', 'off', 'Color', 'w', 'Position', [80 80 1400 500]);
    colormap(h, 'jet');
    ax1 = subplot(1, 3, 1, 'Parent', h);
    sel = thin(numel(d.T), o.MAX_POINTS);
    scatter3(ax1, d.xyz(sel, 1), d.xyz(sel, 2), d.xyz(sel, 3), 6, d.T(sel), 'filled');
    axis(ax1, 'equal'); axis(ax1, 'vis3d'); grid(ax1, 'on'); box(ax1, 'on'); view(ax1, 3);
    xlabel(ax1, 'X'); ylabel(ax1, 'Y'); zlabel(ax1, 'Z'); colorbar(ax1);
    hold(ax1, 'on');
    c0 = mean(d.xyz, 1); L = 0.5 * max(max(d.xyz) - min(d.xyz));
    quiver3(ax1, c0(1), c0(2), c0(3), L * Eb(1, 3), L * Eb(2, 3), L * Eb(3, 3), 0, 'k', 'LineWidth', 2, 'MaxHeadSize', 0.5);
    hold(ax1, 'off');
    title(ax1, sprintf('%s   t = %g s   T [%s]   arrow = thickness axis', ttl, d.time, tunits), 'Interpreter', 'none');

    ax2 = subplot(1, 3, 2, 'Parent', h);
    imagesc(ax2, d.u, d.v, d.Tmean', 'AlphaData', ~isnan(d.Tmean'));
    set(ax2, 'YDir', 'normal'); axis(ax2, 'equal', 'tight'); colorbar(ax2);
    xlabel(ax2, sprintf('in-plane u [%s]', lunits)); ylabel(ax2, sprintf('in-plane v [%s]', lunits));
    title(ax2, sprintf('mean T through the thickness [%s]   in-plane delta %.2f, steepest %.3g /%s', ...
                       tunits, S.dT_ip(iw), S.grad_ip(iw), lunits));

    ax3 = subplot(1, 3, 3, 'Parent', h);
    imagesc(ax3, d.u, d.v, d.dTtt', 'AlphaData', ~isnan(d.dTtt'));
    set(ax3, 'YDir', 'normal'); axis(ax3, 'equal', 'tight'); colorbar(ax3);
    xlabel(ax3, sprintf('in-plane u [%s]', lunits)); ylabel(ax3, sprintf('in-plane v [%s]', lunits));
    title(ax3, sprintf('through-thickness delta T [%s]   max %.2f over t = %.4g %s (%.3g /%s)', ...
                       tunits, S.dT_tt_max(iw), d.thick, lunits, S.grad_tt(iw), lunits));
    annotation(h, 'textbox', [0.01 0.005 0.98 0.06 + 0.03 * (nb - 1)], 'String', verdict, 'EdgeColor', 'none', ...
               'FontWeight', 'bold', 'HorizontalAlignment', 'center', 'Interpreter', 'none');
end
if nargout == 0
    disp(S); clear S D h
end
end


% =========================================================================
function [r, d] = check_body(xyz, T, E, cellsize)
%CHECK_BODY  Through-thickness vs in-plane statistics of one body at one step.
    n = numel(T);
    ctr = mean(xyz, 1);
    P = (xyz - ctr) * E;                  % columns: in-plane u, in-plane v, thickness w
    thick = max(P(:, 3)) - min(P(:, 3));

    ru = max(P(:, 1)) - min(P(:, 1)); rv = max(P(:, 2)) - min(P(:, 2));
    if isempty(cellsize)
        cs = sqrt(8 * max(ru * rv, eps) / n);
        cs = max(cs, max(ru, rv) / 400);
    else
        cs = cellsize;
    end
    iu = floor((P(:, 1) - min(P(:, 1))) / cs) + 1;  nu = max(iu);
    iv = floor((P(:, 2) - min(P(:, 2))) / cs) + 1;  nv = max(iv);
    cnt  = accumarray([iu iv], 1,  [nu nv]);
    Tsum = accumarray([iu iv], T,  [nu nv]);
    Tmax = accumarray([iu iv], T,  [nu nv], @max, -Inf);
    Tmin = accumarray([iu iv], T,  [nu nv], @min,  Inf);
    Tmean = Tsum ./ cnt; Tmean(cnt == 0) = NaN;
    dTtt = Tmax - Tmin; dTtt(cnt < 2) = NaN;

    good = cnt >= 3;
    if any(good(:))
        ttmax = max(dTtt(good)); ttmed = median(dTtt(good));
    else
        ttmax = max(dTtt(cnt >= 2)); ttmed = ttmax;
        if isempty(ttmax), ttmax = 0; ttmed = 0; end
    end
    ipd = max(Tmean(:), [], 'omitnan') - min(Tmean(:), [], 'omitnan');
    gu = abs(diff(Tmean, 1, 1)) / cs; gv = abs(diff(Tmean, 1, 2)) / cs;
    gip = max([gu(:); gv(:); 0], [], 'omitnan');
    A = [ones(n, 1) P];
    c = A \ T;
    res = T - A * c;
    r2 = 1 - sum(res.^2) / max(sum((T - mean(T)).^2), eps);

    r = struct('Body', 1, 'File', '', 'Time', NaN, 'N', n, 'Thick', thick, 'dT_total', max(T) - min(T), ...
               'dT_tt_max', ttmax, 'dT_tt_med', ttmed, 'grad_tt', ttmax / max(thick, eps), ...
               'dT_ip', ipd, 'grad_ip', gip, 'gfit_tt', abs(c(4)), 'gfit_ip', norm(c(2:3)), 'R2', r2);
    d = struct('body', 1, 'file', '', 'time', NaN, 'xyz', xyz, 'T', T, 'axes', E, 'thick', thick, ...
               'cell', cs, 'u', min(P(:, 1)) + (0.5:nu) * cs, 'v', min(P(:, 2)) + (0.5:nv) * cs, ...
               'Tmean', Tmean, 'dTtt', dTtt, 'count', cnt);
end


function body = split_bodies(xyz, gap, min_frac)
%SPLIT_BODIES  Label points by connected voxel (26-neighbour) component at voxel size GAP.
%   Bodies smaller than MIN_FRAC of the points are merged into the nearest big one.
    v = floor((xyz - min(xyz, [], 1)) / gap) + 1;
    dims = max(v, [], 1) + 1;
    lin = @(v) v(:, 1) + dims(1) * (v(:, 2) - 1) + dims(1) * dims(2) * (v(:, 3) - 1);
    [uid, ~, ptVox] = unique(lin(v));
    nv = numel(uid);
    % voxel ijk back from the linear id
    k3 = floor((uid - 1) / (dims(1) * dims(2))) + 1;
    k2 = floor((uid - 1 - (k3 - 1) * dims(1) * dims(2)) / dims(1)) + 1;
    k1 = uid - (k2 - 1) * dims(1) - (k3 - 1) * dims(1) * dims(2);
    uv = [k1 k2 k3];
    [o1, o2, o3] = ndgrid(-1:1, -1:1, -1:1);
    offs = [o1(:) o2(:) o3(:)];
    offs = offs(any(offs, 2), :);
    offs = offs(offs(:, 1) > 0 | (offs(:, 1) == 0 & offs(:, 2) > 0) | (offs(:, 1) == 0 & offs(:, 2) == 0 & offs(:, 3) > 0), :);
    src = []; dst = [];
    for q = 1:size(offs, 1)
        nb = uv + offs(q, :);
        ok = all(nb >= 1, 2) & all(nb <= dims, 2);
        [tf, loc] = ismember(lin(nb(ok, :)), uid);
        here = find(ok);
        src = [src; here(tf)]; dst = [dst; loc(tf)];   %#ok<AGROW>
    end
    Gr = graph(src, dst, [], nv);
    comp = conncomp(Gr)';
    body = comp(ptVox);
    % order by size, then fold tiny bodies into the nearest large one
    cnt = accumarray(body, 1);
    [cnt, order] = sort(cnt, 'descend');
    remap = zeros(size(order)); remap(order) = 1:numel(order);
    body = remap(body);
    big = find(cnt >= min_frac * numel(body));
    if isempty(big), big = 1; end
    small = body > numel(big);
    if any(small)
        ctrs = zeros(numel(big), 3);
        for b = 1:numel(big), ctrs(b, :) = mean(xyz(body == b, :), 1); end
        [~, nearest] = min(sum((permute(xyz(small, :), [1 3 2]) - permute(ctrs, [3 1 2])).^2, 3), [], 2);
        body(small) = nearest;
    end
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
