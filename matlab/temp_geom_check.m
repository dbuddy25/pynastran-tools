function [Vout, hout] = temp_geom_check(Q, varargin)
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
%       'LengthUnits'  model length units, axes label     (default '')
%       'CsvUnits'     length units the CSV was READ as; with LengthUnits the
%                      verdict names the unit the CSV is probably really in
%       'MaxPoints'    cloud points drawn per part        (default 2e5)
%       'Visible'      'on' | 'off'                       (default 'on')
%       'Tiled'        true = one panel per part instead of one overlay (default false)
%       'OnApply'      @(axesStr, offset, partIdx): when given, an 'Apply proposed
%                      transform' button hands the proposal back (the GUI fills its
%                      Cloud xform fields). Standalone use just prints it.
%
%   V(i): .part .level ('ok' | 'warn' | 'bad' | 'none') .cover (0..1)
%         .proposed  [] or struct .axes ('Y X Z' etc, model axis <- cloud axis)
%                    .offset [dx dy dz] (model units) that would align the cloud
%         .lines (cell of messages)      h: .fig .ax .banner .show_part(i)
%   With several parts a dropdown shows one part at a time (0 = all).
%
%   See also TEMP_MAP_MATLAB, TEMP_MAP_GUI.

o = struct('LengthUnits', '', 'CsvUnits', '', 'MaxPoints', 2e5, 'Visible', 'on', 'Tiled', false, 'OnApply', []);
for k = 1:2:numel(varargin)
    name = char(varargin{k});
    if ~isfield(o, name), error('temp_geom_check:badOption', 'Unknown option "%s".', name); end
    o.(name) = varargin{k+1};
end

np = numel(Q);
V = repmat(struct('part', '', 'level', 'none', 'cover', NaN, 'lines', {{}}, 'proposed', []), 1, np);
for i = 1:np
    V(i) = verdict(Q(i), o.LengthUnits, o.CsvUnits);
end

% ---- overall verdict --------------------------------------------------------
lv = {V.level};
if any(strcmp(lv, 'bad')),      level = 'bad';
elseif any(strcmp(lv, 'warn')), level = 'warn';
elseif any(strcmp(lv, 'ok')),   level = 'ok';
else,                           level = 'none';
end
names = arrayfun(@(i) ternary(isempty(Q(i).name), 'model', Q(i).name), 1:np, 'UniformOutput', false);
cols = lines(max(np, 1));
u = ''; if ~isempty(o.LengthUnits), u = sprintf(' [%s]', o.LengthUnits); end
tiled = o.Tiled && np > 1;

% ---- figure ---------------------------------------------------------------
fig = figure('Name', 'Geometry check: mesh vs temperature cloud', 'NumberTitle', 'off', ...
             'Color', 'w', 'Visible', o.Visible, 'Position', [80 80 1000 760]);
HP = cell(1, np);                                  % graphics per part, for the filter
if tiled
    % one panel per part, its own verdict as a coloured title
    tl = tiledlayout(fig, 'flow', 'Padding', 'compact', 'TileSpacing', 'compact');
    tl.OuterPosition = [0 0 1 0.82];
    AX = gobjects(1, np);
    for i = 1:np
        AX(i) = nexttile(tl);
        hold(AX(i), 'on');
        HP{i} = draw_part(AX(i), Q(i), cols(i, :), o);
        hold(AX(i), 'off');
        dress(AX(i), u);
        [bgi, fgi] = banner_style(V(i).level);
        if isempty(V(i).lines), ttl = {sprintf('%s: no cloud (uniform part)', names{i})};
        else, ttl = strsplit(V(i).lines{1}, '.  '); end      % headline, then the causes
        title(AX(i), ttl, 'Color', fgi, 'BackgroundColor', bgi, 'Interpreter', 'none', ...
              'FontSize', 9, 'FontWeight', 'bold');
        set_view(AX(i), 'ISO');
    end
else
    AX = axes('Parent', fig, 'Position', [0.07 0.06 0.86 0.74]);
    hold(AX, 'on');
    for i = 1:np
        HP{i} = draw_part(AX, Q(i), cols(i, :), o);
    end
    hold(AX, 'off');
    dress(AX, u);
    legend(AX, 'Location', 'northeastoutside', 'Interpreter', 'none');
    set_view(AX, 'ISO');
end
for a = AX
    try
        a.Interactions = [rotateInteraction zoomInteraction panInteraction];
    catch
        try rotate3d(a, 'on'); catch, end
    end
end
ax = AX(1);

% ---- banner + buttons -------------------------------------------------------
[bg, fgc, head] = banner_style(level);
msg = {};
for i = 1:np
    if ~strcmp(V(i).level, 'none'), msg{end+1} = short_line(V(i)); end   %#ok<AGROW>
end
if isempty(msg), msg = {'no cloud parts to compare'}; end
if numel(msg) > 3
    nb = nnz(strcmp(lv, 'bad')); nw = nnz(strcmp(lv, 'warn')); nk = nnz(strcmp(lv, 'ok'));
    msg = [{sprintf('%d parts: %d mismatch, %d check, %d ok -- per-part detail in the command window%s', ...
                    np, nb, nw, nk, ternary(tiled, ' and above each panel', '; Tile parts shows one panel each'))}, ...
           msg(~strcmp(lv(~strcmp(lv, 'none')), 'ok'))];
    msg = msg(1:min(3, end));
end
bw = ternary(isempty(o.OnApply), 0.96, 0.76);
banner = uicontrol(fig, 'Style', 'text', 'Units', 'normalized', 'Position', [0.02 0.885 bw 0.10], ...
                   'String', [{head}, msg], 'BackgroundColor', bg, 'ForegroundColor', fgc, ...
                   'FontSize', 10, 'FontWeight', 'bold', 'HorizontalAlignment', 'left');
labels = {'+X', '-X', '+Y', '-Y', '+Z', '-Z', 'ISO'};
w = 0.055; x0 = 0.07;
for k = 1:numel(labels)
    uicontrol(fig, 'Style', 'pushbutton', 'String', labels{k}, ...
              'Units', 'normalized', 'Position', [x0 + (k-1)*(w+0.005), 0.835, w, 0.04], ...
              'Callback', @(~, ~) view_all(labels{k}));
end
uicontrol(fig, 'Style', 'text', 'Units', 'normalized', 'Position', [0.70 0.835 0.28 0.04], ...
          'String', 'mesh box: solid black   cloud box: dashed, part colour', ...
          'BackgroundColor', 'w', 'HorizontalAlignment', 'right', 'FontSize', 9);
allStr = banner.String;
cur = 0;                                            % 0 = all parts shown
applyB = [];
if ~isempty(o.OnApply)
    applyB = uicontrol(fig, 'Style', 'pushbutton', 'Units', 'normalized', 'Position', [0.80 0.905 0.18 0.06], ...
                       'String', 'Apply proposed transform', 'FontSize', 10, 'FontWeight', 'bold', ...
                       'Tooltip', 'Fill the GUI''s Cloud xform fields with the axes / offset proposed for the shown part', ...
                       'Callback', @(~, ~) apply_proposal());
    refresh_apply();
end
if np > 1
    uicontrol(fig, 'Style', 'pushbutton', 'Units', 'normalized', 'Position', [0.52 0.835 0.09 0.04], ...
              'String', ternary(tiled, 'Overlay', 'Tile parts'), 'FontSize', 10, ...
              'Tooltip', 'Open the other layout in a new window', ...
              'Callback', @(~, ~) temp_geom_check(Q, 'Tiled', ~tiled, 'LengthUnits', o.LengthUnits, 'CsvUnits', o.CsvUnits, ...
                                                  'MaxPoints', o.MaxPoints));
    if ~tiled
        uicontrol(fig, 'Style', 'popupmenu', 'Units', 'normalized', 'Position', [0.615 0.835 0.08 0.04], ...
                  'String', [{'All parts'}, names], 'Value', 1, 'FontSize', 10, ...
                  'Callback', @(src, ~) show_part(src.Value - 1));
    end
end

for i = 1:np
    if strcmp(V(i).level, 'none'), continue; end
    fprintf('%s\n', V(i).lines{:});
end
Vout = V;
hout = struct('fig', fig, 'ax', AX, 'banner', banner, 'show_part', @show_part);
if nargout == 0, clear Vout hout; end

    function show_part(i)
        % 0 = all parts; otherwise only part i, with its own verdict in the banner
        cur = i;
        for j = 1:np
            set(HP{j}, 'Visible', ternary(i == 0 || i == j, 'on', 'off'));
        end
        if i == 0
            [bg_, fg_] = banner_style(level);
            set(banner, 'String', allStr, 'BackgroundColor', bg_, 'ForegroundColor', fg_);
        else
            [bg_, fg_, head_] = banner_style(V(i).level);
            if strcmp(V(i).level, 'none'), txt = {head_; sprintf('%s: no cloud (uniform part)', names{i})};
            else, txt = [{head_}, V(i).lines(1)]; end
            set(banner, 'String', txt, 'BackgroundColor', bg_, 'ForegroundColor', fg_);
        end
        set_view(ax, 'FIT');
        refresh_apply();
    end

    function view_all(name)
        for a_ = AX, set_view(a_, name); end
    end

    function i = proposal_part()
        % the shown part if it has a proposal, else the first part that does
        i = 0;
        if cur > 0 && ~isempty(V(cur).proposed), i = cur; return; end
        if cur == 0
            k_ = find(arrayfun(@(v_) ~isempty(v_.proposed), V), 1);
            if ~isempty(k_), i = k_; end
        end
    end

    function refresh_apply()
        if isempty(applyB), return; end
        i = proposal_part();
        if i == 0
            set(applyB, 'Enable', 'off', 'String', 'No transform to propose');
        else
            set(applyB, 'Enable', 'on', 'String', sprintf('Apply: axes %s  (%s)', V(i).proposed.axes, names{i}));
        end
    end

    function apply_proposal()
        i = proposal_part();
        if i == 0, return; end
        o.OnApply(V(i).proposed.axes, V(i).proposed.offset, i);
    end
end


% =========================================================================
function v = verdict(q, lu, cu)
    tag = q.name; if isempty(tag), tag = 'model'; end
    v = struct('part', tag, 'level', 'none', 'cover', NaN, 'lines', {{}}, 'proposed', []);
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
        % which unit would the CSV have to be in for the sizes to agree?
        to_m = struct('in', 0.0254, 'mm', 1e-3, 'm', 1, 'ft', 0.3048, 'cm', 0.01);
        if ~isempty(cu) && isfield(to_m, lower(strtrim(cu)))
            cu = lower(strtrim(cu));
            cands = fieldnames(to_m);
            for k = 1:numel(cands)
                c = cands{k};
                if strcmp(c, cu), continue; end
                if abs(log(med * to_m.(c) / to_m.(cu))) < log(1.12)
                    causes{end+1} = sprintf('cloud is %.4g x the mesh: the CSV looks like it is in %s, not %s -- set CSV length units to %s', ...
                                            med, c, cu, c); %#ok<AGROW>
                    units_bad = true;
                    break
                end
            end
        end
        known = {25.4, 'mm / in'; 1000, 'm / mm'; 39.37, 'm / in'; 12, 'ft / in'; 1e6, 'm / um'};
        for k = 1:size(known, 1)
            if isempty(causes) && (abs(log(med / known{k, 1})) < log(1.12) || abs(log(med * known{k, 1})) < log(1.12))
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
    pb = [1 2 3]; swapped = false;
    if best ~= ident && cover < 0.9 && err(ident) > log(1.15) && err(best) < 0.5 * err(ident)
        pb = perms3(best, :); swapped = true;
        causes{end+1} = sprintf('axes look swapped: cloud (%c, %c, %c) matches mesh (X, Y, Z)', ...
                                ax(pb(1)), ax(pb(2)), ax(pb(3)));
    end
    d = 0.5 * (loc + hic) - 0.5 * (lom + him);
    offset_bad = cover < 0.9 && norm(d) > 0.5 * norm(emx) && ~units_bad;
    if offset_bad && ~swapped
        causes{end+1} = sprintf('cloud centre is offset from the mesh by [%.4g %.4g %.4g]%s -- different origin / coordinate system?', ...
                                d(1), d(2), d(3), lu);
    end

    % --- proposed rigid fix: model axis i <- sign * cloud axis pb(i), then a shift ----
    if ~units_bad && (swapped || offset_bad)
        sgn = ones(1, 3);
        skew = @(x) mean((x - mean(x)).^3) / max(std(x), eps)^3;
        for i = 1:3
            sm = skew(M(:, i)); sc = skew(P(:, pb(i)));
            if abs(sm) > 0.15 && abs(sc) > 0.15 && sign(sm) ~= sign(sc), sgn(i) = -1; end
        end
        A = zeros(3); for i = 1:3, A(i, pb(i)) = sgn(i); end
        off = 0.5 * (lom + him) - 0.5 * (loc + hic) * A';
        off(abs(off) < 1e-6 * max(scale, eps)) = 0;
        if isequal(A, eye(3)) && norm(off) <= 0.05 * norm(emx), off(:) = 0; end
        if ~isequal(A, eye(3)) || any(off)
            names = arrayfun(@(i) [ternary(sgn(i) < 0, '-', '') ax(pb(i))], 1:3, 'UniformOutput', false);
            v.proposed = struct('axes', strjoin(names, ' '), 'offset', off);
            L{end+1} = sprintf('  proposed cloud transform: axes ''%s'' (model <- cloud), offset [%.4g %.4g %.4g]%s%s', ...
                               v.proposed.axes, off(1), off(2), off(3), lu, ...
                               ternary(any(sgn < 0), '  (sign flips from shape skew -- best effort)', ''));
        end
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


function hp = draw_part(ax, q, col, o)
    tag = q.name; if isempty(tag), tag = 'model'; end
    hp = gobjects(0);
    if ~isempty(q.faces)
        big = size(q.faces, 1) > 2e5;
        hp(end+1) = patch('Parent', ax, 'Faces', q.faces, 'Vertices', q.grid_xyz, ...
              'FaceColor', [0.82 0.82 0.82], 'FaceAlpha', 0.55, ...
              'EdgeColor', ternary(big, 'none', [0.35 0.35 0.35]), 'EdgeAlpha', 0.35, ...
              'DisplayName', ['mesh: ' tag]);
    else
        gs = thin(size(q.grid_xyz, 1), o.MaxPoints);
        hp(end+1) = scatter3(ax, q.grid_xyz(gs, 1), q.grid_xyz(gs, 2), q.grid_xyz(gs, 3), 10, [0.4 0.4 0.4], ...
                 'filled', 'DisplayName', ['grids: ' tag]);
    end
    hp(end+1) = draw_box(ax, q.grid_xyz, [0.15 0.15 0.15], '-', 1.6, ['mesh box: ' tag]);
    if ~isempty(q.cloud_xyz)
        cs = thin(size(q.cloud_xyz, 1), o.MaxPoints);
        hp(end+1) = scatter3(ax, q.cloud_xyz(cs, 1), q.cloud_xyz(cs, 2), q.cloud_xyz(cs, 3), 6, col, '.', ...
                 'DisplayName', sprintf('cloud: %s  (%s)', tag, shortname(q.csv_file)));
        hp(end+1) = draw_box(ax, q.cloud_xyz, col, '--', 1.8, ['cloud box: ' tag]);
    end
end


function s = short_line(v)
%SHORT_LINE  One banner line per part: name, verdict word, the first cause.
    head = v.lines{1};
    parts = strsplit(head, '.  ');
    word = regexp(parts{1}, ':\s*(\w+)', 'tokens', 'once');
    if isempty(word), word = {'?'}; end
    s = sprintf('%s: %s', v.part, word{1});
    if numel(parts) > 1, s = [s ' -- ' parts{2}]; end
end


function dress(ax, u)
    axis(ax, 'equal'); axis(ax, 'vis3d'); grid(ax, 'on'); box(ax, 'on');
    xlabel(ax, ['X' u]); ylabel(ax, ['Y' u]); zlabel(ax, ['Z' u]);
end


function [bg, fg, head] = banner_style(level)
    switch level
        case 'ok',   bg = [0.80 0.94 0.80]; fg = [0.05 0.40 0.10]; head = 'GEOMETRY OK -- units, orientation and position agree';
        case 'warn', bg = [1.00 0.93 0.72]; fg = [0.55 0.35 0.00]; head = 'GEOMETRY: CHECK -- partial overlap';
        case 'bad',  bg = [0.98 0.80 0.80]; fg = [0.65 0.05 0.05]; head = 'GEOMETRY MISMATCH -- fix units / orientation before mapping';
        otherwise,   bg = [0.92 0.92 0.92]; fg = [0.30 0.30 0.30]; head = 'GEOMETRY: nothing to compare';
    end
end


function hl = draw_box(ax, X, col, ls, lw, name)
    lo = min(X); hi = max(X);
    c = [lo(1) lo(2) lo(3); hi(1) lo(2) lo(3); hi(1) hi(2) lo(3); lo(1) hi(2) lo(3); ...
         lo(1) lo(2) hi(3); hi(1) lo(2) hi(3); hi(1) hi(2) hi(3); lo(1) hi(2) hi(3)];
    E = [1 2; 2 3; 3 4; 4 1; 5 6; 6 7; 7 8; 8 5; 1 5; 2 6; 3 7; 4 8];
    xs = [c(E(:, 1), 1) c(E(:, 2), 1) nan(12, 1)]'; ys = [c(E(:, 1), 2) c(E(:, 2), 2) nan(12, 1)]';
    zs = [c(E(:, 1), 3) c(E(:, 2), 3) nan(12, 1)]';
    hl = plot3(ax, xs(:), ys(:), zs(:), 'Color', col, 'LineStyle', ls, 'LineWidth', lw, 'DisplayName', name);
end


function set_view(ax, name)
    cur = campos(ax) - camtarget(ax); cur_up = camup(ax);
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
        case 'FIT', dirn = cur / max(norm(cur), eps); up = cur_up;
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
