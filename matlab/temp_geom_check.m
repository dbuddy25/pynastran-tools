function [V, h] = temp_geom_check(Q, varargin)
%TEMP_GEOM_CHECK  Overlay the structural mesh and the temperature cloud to
%   confirm length units, orientation and position BEFORE mapping.
%
%   [V, h] = TEMP_GEOM_CHECK(Q)   Q from
%       Q = temp_map_matlab('PARTS', P, 'GEOMETRY_ONLY', true, <unit options>)
%   One figure: every part's element faces (grey) with its cloud points and
%   both bounding boxes (mesh solid, cloud dashed), view buttons, and a banner
%   that is GREEN when the mesh sits inside the cloud, AMBER when it partly
%   does, RED when it does not -- with the likely cause (unit factor, swapped
%   axes, offset origin) spelled out.
%
%   Options (name/value):
%       'LengthUnits'  label for the axes                (default '')
%       'MaxPoints'    cloud points drawn per part        (default 2e5)
%       'Visible'      'on' | 'off'                       (default 'on')
%
%   V(i): .part .level ('ok' | 'warn' | 'bad' | 'none') .cover (0..1)
%         .lines (cell of messages)      h: .fig .ax .banner
%
%   See also TEMP_MAP_MATLAB, TEMP_MAP_GUI.

o = struct('LengthUnits', '', 'MaxPoints', 2e5, 'Visible', 'on');
for k = 1:2:numel(varargin)
    name = char(varargin{k});
    if ~isfield(o, name), error('temp_geom_check:badOption', 'Unknown option "%s".', name); end
    o.(name) = varargin{k+1};
end

np = numel(Q);
V = repmat(struct('part', '', 'level', 'none', 'cover', NaN, 'lines', {{}}), 1, np);
for i = 1:np
    V(i) = verdict(Q(i), o.LengthUnits);
end

% ---- figure ---------------------------------------------------------------
fig = figure('Name', 'Geometry check: mesh vs temperature cloud', 'NumberTitle', 'off', ...
             'Color', 'w', 'Visible', o.Visible, 'Position', [80 80 1000 760]);
ax = axes('Parent', fig, 'Position', [0.07 0.06 0.86 0.74]);
hold(ax, 'on');
cols = lines(max(np, 1));
for i = 1:np
    q = Q(i); col = cols(i, :);
    tag = q.name; if isempty(tag), tag = 'model'; end
    if ~isempty(q.faces)
        big = size(q.faces, 1) > 2e5;
        patch('Parent', ax, 'Faces', q.faces, 'Vertices', q.grid_xyz, ...
              'FaceColor', [0.82 0.82 0.82], 'FaceAlpha', 0.55, ...
              'EdgeColor', ternary(big, 'none', [0.35 0.35 0.35]), 'EdgeAlpha', 0.35, ...
              'DisplayName', ['mesh: ' tag]);
    else
        gs = thin(size(q.grid_xyz, 1), o.MaxPoints);
        scatter3(ax, q.grid_xyz(gs, 1), q.grid_xyz(gs, 2), q.grid_xyz(gs, 3), 10, [0.4 0.4 0.4], ...
                 'filled', 'DisplayName', ['grids: ' tag]);
    end
    draw_box(ax, q.grid_xyz, [0.15 0.15 0.15], '-', 1.6, ['mesh box: ' tag]);
    if ~isempty(q.cloud_xyz)
        cs = thin(size(q.cloud_xyz, 1), o.MaxPoints);
        scatter3(ax, q.cloud_xyz(cs, 1), q.cloud_xyz(cs, 2), q.cloud_xyz(cs, 3), 6, col, '.', ...
                 'DisplayName', sprintf('cloud: %s  (%s)', tag, shortname(q.csv_file)));
        draw_box(ax, q.cloud_xyz, col, '--', 1.8, ['cloud box: ' tag]);
    end
end
hold(ax, 'off');
axis(ax, 'equal'); axis(ax, 'vis3d'); grid(ax, 'on'); box(ax, 'on');
u = ''; if ~isempty(o.LengthUnits), u = sprintf(' [%s]', o.LengthUnits); end
xlabel(ax, ['X' u]); ylabel(ax, ['Y' u]); zlabel(ax, ['Z' u]);
legend(ax, 'Location', 'northeastoutside', 'Interpreter', 'none');
set_view(ax, 'ISO');
try
    ax.Interactions = [rotateInteraction zoomInteraction panInteraction];
catch
    try rotate3d(ax, 'on'); catch, end
end

% ---- banner + buttons -------------------------------------------------------
lv = {V.level};
if any(strcmp(lv, 'bad')),      level = 'bad';
elseif any(strcmp(lv, 'warn')), level = 'warn';
elseif any(strcmp(lv, 'ok')),   level = 'ok';
else,                           level = 'none';
end
[bg, fgc, head] = banner_style(level);
msg = {};
for i = 1:np
    if ~strcmp(V(i).level, 'none'), msg = [msg, V(i).lines(1)]; end   %#ok<AGROW>  headline per part
end
if isempty(msg), msg = {'no cloud parts to compare'}; end
banner = uicontrol(fig, 'Style', 'text', 'Units', 'normalized', 'Position', [0.02 0.895 0.96 0.085], ...
                   'String', [{head}, msg], 'BackgroundColor', bg, 'ForegroundColor', fgc, ...
                   'FontSize', 11, 'FontWeight', 'bold', 'HorizontalAlignment', 'left');
labels = {'+X', '-X', '+Y', '-Y', '+Z', '-Z', 'ISO'};
w = 0.055; x0 = 0.07;
for k = 1:numel(labels)
    uicontrol(fig, 'Style', 'pushbutton', 'String', labels{k}, ...
              'Units', 'normalized', 'Position', [x0 + (k-1)*(w+0.005), 0.835, w, 0.04], ...
              'Callback', @(~, ~) set_view(ax, labels{k}));
end
uicontrol(fig, 'Style', 'text', 'Units', 'normalized', 'Position', [0.52 0.835 0.46 0.04], ...
          'String', 'mesh box: solid black     cloud box: dashed, part colour', ...
          'BackgroundColor', 'w', 'HorizontalAlignment', 'right', 'FontSize', 9);

for i = 1:np
    if strcmp(V(i).level, 'none'), continue; end
    fprintf('%s\n', V(i).lines{:});
end
h = struct('fig', fig, 'ax', ax, 'banner', banner);
if nargout == 0, clear V h; end
end


% =========================================================================
function v = verdict(q, lu)
    tag = q.name; if isempty(tag), tag = 'model'; end
    v = struct('part', tag, 'level', 'none', 'cover', NaN, 'lines', {{}});
    if isempty(q.cloud_xyz), return; end
    M = q.grid_xyz; P = q.cloud_xyz;
    lom = min(M); him = max(M); loc = min(P); hic = max(P);
    em = him - lom; ec = hic - loc;
    scale = max([em ec]);
    tiny = 1e-3 * scale;

    % per-axis: how much of the mesh extent lies inside the cloud extent
    cov = zeros(1, 3);
    for a = 1:3
        if em(a) <= tiny
            cm = 0.5 * (lom(a) + him(a));
            cov(a) = double(cm >= loc(a) - 0.05 * scale && cm <= hic(a) + 0.05 * scale);
        else
            cov(a) = max(min(him(a), hic(a)) - max(lom(a), loc(a)), 0) / em(a);
        end
    end
    cover = min(cov);
    v.cover = cover;

    ax = 'XYZ';
    fmt = @(lo, hi) sprintf('%s', strjoin(arrayfun(@(a) sprintf('%c %.4g..%.4g', ax(a), lo(a), hi(a)), 1:3, ...
                                                  'UniformOutput', false), '   '));
    if ~isempty(lu), lu = [' ' lu]; end
    L = {};
    L{end+1} = sprintf('  mesh  extents%s: %s', lu, fmt(lom, him));
    L{end+1} = sprintf('  cloud extents%s: %s', lu, fmt(loc, hic));
    L{end+1} = sprintf('  mesh inside cloud: X %.0f%%  Y %.0f%%  Z %.0f%%', 100 * cov);

    % --- likely causes ---------------------------------------------------------
    causes = {};
    units_bad = false;
    live = em > tiny & ec > tiny;
    emx = max(em, tiny); ecx = max(ec, tiny);
    if any(live)
        med = exp(median(log(ec(live) ./ em(live))));
        known = {25.4, 'mm / in'; 1000, 'm / mm'; 39.37, 'm / in'; 12, 'ft / in'; 1e6, 'm / um'};
        for k = 1:size(known, 1)
            if abs(log(med / known{k, 1})) < log(1.12) || abs(log(med * known{k, 1})) < log(1.12)
                causes{end+1} = sprintf('cloud extents are ~%.4g x the mesh -- a %s factor (%g): check the CSV vs model length units', ...
                                        med, known{k, 2}, known{k, 1}); %#ok<AGROW>
                units_bad = true;
                break
            end
        end
        if isempty(causes) && (med > 2.5 || med < 0.4)
            causes{end+1} = sprintf('cloud extents are ~%.3g x the mesh -- length units or a scaled export?', med);
            units_bad = true;
        end
    end
    % axis permutation: does the cloud's shape match the mesh better with axes swapped?
    perms3 = perms(1:3);
    err = zeros(size(perms3, 1), 1);
    for p = 1:size(perms3, 1)
        r = log(ecx(perms3(p, :)) ./ emx);
        err(p) = std(r);                              % a pure unit factor is a constant shift
    end
    [~, best] = min(err);
    ident = find(all(perms3 == [1 2 3], 2));
    if best ~= ident && cover < 0.9 && err(ident) > log(1.15) && err(best) < 0.5 * err(ident)
        pb = perms3(best, :);
        causes{end+1} = sprintf('axes look swapped: cloud (%c, %c, %c) matches mesh (X, Y, Z)', ...
                                ax(pb(1)), ax(pb(2)), ax(pb(3)));
    end
    d = 0.5 * (loc + hic) - 0.5 * (lom + him);
    if cover < 0.9 && norm(d) > 0.5 * norm(emx) && isempty(causes)
        causes{end+1} = sprintf('cloud centre is offset from the mesh by [%.4g %.4g %.4g]%s -- different origin / coordinate system?', ...
                                d(1), d(2), d(3), lu);
    end

    if units_bad
        % a cloud 25x too big still "contains" the mesh -- that is still wrong
        v.level = 'bad';
        head = sprintf('%s: MISMATCH -- cloud and mesh are different sizes', tag);
    elseif cover >= 0.9 && isempty(causes)
        v.level = 'ok';
        head = sprintf('%s: OK -- mesh sits inside the cloud (%.0f%% on the tightest axis)', tag, 100 * cover);
    elseif cover >= 0.5
        v.level = 'warn';
        head = sprintf('%s: CHECK -- only %.0f%% of the mesh is inside the cloud on one axis', tag, 100 * cover);
    else
        v.level = 'bad';
        head = sprintf('%s: MISMATCH -- mesh and cloud barely overlap (%.0f%%)', tag, 100 * cover);
    end
    if ~isempty(causes), head = [head '.  ' strjoin(causes, '; ')]; end
    v.lines = [{head}, L];
end


function [bg, fg, head] = banner_style(level)
    switch level
        case 'ok',   bg = [0.80 0.94 0.80]; fg = [0.05 0.40 0.10]; head = 'GEOMETRY OK -- units, orientation and position agree';
        case 'warn', bg = [1.00 0.93 0.72]; fg = [0.55 0.35 0.00]; head = 'GEOMETRY: CHECK -- partial overlap';
        case 'bad',  bg = [0.98 0.80 0.80]; fg = [0.65 0.05 0.05]; head = 'GEOMETRY MISMATCH -- fix units / orientation before mapping';
        otherwise,   bg = [0.92 0.92 0.92]; fg = [0.30 0.30 0.30]; head = 'GEOMETRY: nothing to compare';
    end
end


function draw_box(ax, X, col, ls, lw, name)
    lo = min(X); hi = max(X);
    c = [lo(1) lo(2) lo(3); hi(1) lo(2) lo(3); hi(1) hi(2) lo(3); lo(1) hi(2) lo(3); ...
         lo(1) lo(2) hi(3); hi(1) lo(2) hi(3); hi(1) hi(2) hi(3); lo(1) hi(2) hi(3)];
    E = [1 2; 2 3; 3 4; 4 1; 5 6; 6 7; 7 8; 8 5; 1 5; 2 6; 3 7; 4 8];
    xs = [c(E(:, 1), 1) c(E(:, 2), 1) nan(12, 1)]'; ys = [c(E(:, 1), 2) c(E(:, 2), 2) nan(12, 1)]';
    zs = [c(E(:, 1), 3) c(E(:, 2), 3) nan(12, 1)]';
    plot3(ax, xs(:), ys(:), zs(:), 'Color', col, 'LineStyle', ls, 'LineWidth', lw, 'DisplayName', name);
end


function set_view(ax, name)
    camva(ax, 'auto'); camtarget(ax, 'auto'); campos(ax, 'auto'); camup(ax, 'auto');
    axis(ax, 'auto'); axis(ax, 'equal'); axis(ax, 'tight');
    set(ax, 'Projection', 'orthographic');
    switch upper(name)
        case '+X',  dirn = [ 1  0  0]; up = [0 0 1];
        case '-X',  dirn = [-1  0  0]; up = [0 0 1];
        case '+Y',  dirn = [ 0  1  0]; up = [0 0 1];
        case '-Y',  dirn = [ 0 -1  0]; up = [0 0 1];
        case '+Z',  dirn = [ 0  0  1]; up = [0 1 0];
        case '-Z',  dirn = [ 0  0 -1]; up = [0 1 0];
        otherwise,  dirn = [-1 -1  1] / sqrt(3); up = [0 0 1];
    end
    tgt = mean([ax.XLim; ax.YLim; ax.ZLim], 2)';
    span = norm([diff(ax.XLim) diff(ax.YLim) diff(ax.ZLim)]);
    camtarget(ax, tgt); campos(ax, tgt + dirn * max(span, eps) * 3); camup(ax, up); camva(ax, 'auto');
    axis(ax, 'vis3d');
end


function sel = thin(n, nmax)
    if n <= nmax, sel = (1:n)'; return; end
    rs = RandStream('mt19937ar', 'Seed', 0);
    sel = sort(randperm(rs, n, round(nmax)))';
end


function out = ternary(cond, a, b)
    if cond, out = a; else, out = b; end
end


function s = shortname(p)
    [~, n, e] = fileparts(char(p));
    s = [n e];
end
