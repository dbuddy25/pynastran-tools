function h = temp_map_plot(r, varargin)
%TEMP_MAP_PLOT  Rotatable 3D check plot of one TEMP_MAP_MATLAB result.
%
%   H = TEMP_MAP_PLOT(R(k)) opens a figure showing the model grids as filled
%   markers coloured by their mapped temperature, the source cloud as small
%   grey points, and any grid that fell outside the cloud's convex hull ringed
%   in black.  Drag to rotate; the buttons along the top snap to the six axis
%   views and the isometric view.
%
%   H = TEMP_MAP_PLOT(R(k), 'Name', value, ...) options:
%       'Parent'      axes to draw into (e.g. a uiaxes) -- no buttons are added
%       'Style'       'points' (default) | 'contour'  -- contour paints the
%                     element faces read from the BDF with the temperature
%                     interpolated across each face (needs shell/solid elements)
%       'Units'       'K' (default) | 'C' | 'F'  -- display units
%       'ShowCloud'   true | false           (default true)
%       'ShowSurface' true | false           (default true)  translucent skin of
%                     the cloud (alpha shape) so you can see where the thermal
%                     volume actually ends
%       'ShowExtrap'  true | false           (default true)
%       'Colormap'    'jet' (default), 'parula', 'turbo', ...
%       'CLim'        [] = auto from the grid temps, else [lo hi]
%       'MarkerSize'  grid marker area      (default 36)
%       'CloudSize'   cloud marker area     (default 8)
%       'ShowMinMax'  true | false (default true)  label the hottest and coldest
%                     grid with its ID and temperature
%       'MaxPoints'   draw at most this many grids / cloud points (default 2e5);
%                     larger sets are randomly thinned for a responsive view
%       'View'        '+X' '-X' '+Y' '-Y' '+Z' '-Z' 'ISO'   (default 'ISO');
%                     H.set_view('FIT') re-frames without changing direction
%       'Visible'     'on' | 'off'  (figure visibility, standalone only)
%       'Buttons'     true | false   add the view buttons / toggles to a
%                     standalone figure (false for figures you will export)
%
%   H is a struct of handles: .fig .ax .grids .cloud .extrap .cbar, plus
%   H.set_view(name) to snap the view programmatically.
%
%   See also TEMP_MAP_MATLAB, TEMP_MAP_GUI.

o = struct('Parent', [], 'Style', 'points', 'Units', 'K', 'ShowCloud', true, 'ShowSurface', true, 'ShowExtrap', true, ...
           'Colormap', 'jet', 'CLim', [], 'MarkerSize', 36, 'CloudSize', 8, ...
           'ShowMinMax', true, 'MaxPoints', 2e5, 'View', 'ISO', 'Visible', 'on', 'Buttons', true);
for k = 1:2:numel(varargin)
    name = char(varargin{k});
    if ~isfield(o, name)
        error('temp_map_plot:badOption', 'Unknown option "%s".', name);
    end
    o.(name) = varargin{k+1};
end

units = upper(char(o.Units));
switch units
    case 'K', Tg = r.grid_T;
    case 'C', Tg = r.grid_T - 273.15;
    case 'F', Tg = (r.grid_T - 273.15) * 9/5 + 32;
    otherwise, error('temp_map_plot:badUnits', 'Units must be K, C or F.');
end
ex = r.extrap;
if isfield(r, 'far') && ~isempty(r.far), ex = ex | r.far; end

% thin huge sets for display only (temps/limits still use the full set)
ng = numel(Tg); nc = size(r.cloud_xyz, 1);
gsel = thin(ng, o.MaxPoints); csel = thin(nc, o.MaxPoints);
gxyz = r.grid_xyz(gsel, :); cxyz = r.cloud_xyz(csel, :);
note = '';
if numel(gsel) < ng || numel(csel) < nc
    note = sprintf('   [showing %d of %d grids, %d of %d cloud pts]', numel(gsel), ng, numel(csel), nc);
end

% --- figure / axes -------------------------------------------------------
standalone = isempty(o.Parent);
if standalone
    fig = figure('Name', ['TEMP map: ' shortname(r.csv_file)], 'NumberTitle', 'off', ...
                 'Color', 'w', 'Visible', o.Visible, 'Position', [100 100 900 700]);
    ax = axes('Parent', fig, 'Position', [0.08 0.08 0.80 0.80]);
else
    ax = o.Parent;
    fig = ancestor(ax, 'figure');
    cla(ax, 'reset');          % drop stale zoom/camera state from the previous plot
end
hold(ax, 'on');

% --- data ----------------------------------------------------------------
h = struct('fig', fig, 'ax', ax, 'grids', [], 'mesh', [], 'cloud', [], 'surface', [], 'extrap', [], ...
           'minmax', [], 'cbar', []);

contour = strcmpi(o.Style, 'contour');
if contour && (~isfield(r, 'faces') || isempty(r.faces))
    warning('temp_map_plot:noFaces', 'No shell/solid elements were read from the BDF; showing points.');
    contour = false;
end
if contour
    % full mesh, never thinned: the faces index every grid
    h.mesh = patch('Parent', ax, 'Faces', r.faces, 'Vertices', r.grid_xyz, ...
                   'FaceVertexCData', Tg, 'FaceColor', 'interp', 'EdgeColor', 'none', ...
                   'DisplayName', 'mesh (interpolated T)');
end

% one translucent skin per part (a merged multi-part result carries a cell)
srfs = {};
if isfield(r, 'surface') && ~isempty(r.surface)
    if iscell(r.surface), srfs = r.surface; else, srfs = {r.surface}; end
end
srfs = srfs(~cellfun('isempty', srfs));
h.surface = hggroup('Parent', ax, 'DisplayName', 'cloud surface (alpha shape)');
for q = 1:numel(srfs)
    patch('Parent', h.surface, 'Faces', srfs{q}.F, 'Vertices', srfs{q}.V, ...
          'FaceColor', [0.6 0.6 0.6], 'FaceAlpha', 0.18, 'EdgeColor', 'none');
end
h.surface.Visible = onoff(o.ShowSurface && ~isempty(srfs));

h.cloud = scatter3(ax, cxyz(:,1), cxyz(:,2), cxyz(:,3), ...
                   o.CloudSize, [0.55 0.55 0.55], '.', 'DisplayName', 'cloud points');
h.cloud.Visible = onoff(o.ShowCloud);

h.grids = scatter3(ax, gxyz(:,1), gxyz(:,2), gxyz(:,3), ...
                   o.MarkerSize, Tg(gsel), 'filled', 'DisplayName', 'grids (mapped T)');
h.grids.Visible = onoff(~contour);

exs = ex(gsel);
h.extrap = scatter3(ax, gxyz(exs,1), gxyz(exs,2), gxyz(exs,3), ...
                    o.MarkerSize * 2.2, 'Marker', 'o', 'MarkerEdgeColor', 'k', ...
                    'MarkerFaceColor', 'none', 'LineWidth', 1.0, ...
                    'DisplayName', sprintf('outside hull / far (%d)', nnz(ex)));
h.extrap.Visible = onoff(o.ShowExtrap && any(ex));

% --- min / max markers -------------------------------------------------------
[Tmax, imax] = max(Tg); [Tmin, imin] = min(Tg);
pm = r.grid_xyz([imax imin], :);
h.minmax = hggroup('Parent', ax, 'DisplayName', 'min / max grid');
scatter3(ax, pm(:,1), pm(:,2), pm(:,3), o.MarkerSize * 4, 'Marker', 'p', ...
         'MarkerEdgeColor', 'k', 'MarkerFaceColor', 'w', 'LineWidth', 1.2, 'Parent', h.minmax);
text(pm(1,1), pm(1,2), pm(1,3), sprintf('  MAX %.1f %s  (grid %d)', Tmax, units, r.grid_ids(imax)), ...
     'Parent', h.minmax, 'FontWeight', 'bold', 'BackgroundColor', [1 1 1 0.75], 'Margin', 2, ...
     'Interpreter', 'none', 'Clipping', 'off');
text(pm(2,1), pm(2,2), pm(2,3), sprintf('  MIN %.1f %s  (grid %d)', Tmin, units, r.grid_ids(imin)), ...
     'Parent', h.minmax, 'FontWeight', 'bold', 'BackgroundColor', [1 1 1 0.75], 'Margin', 2, ...
     'Interpreter', 'none', 'Clipping', 'off');
h.minmax.Visible = onoff(o.ShowMinMax);

% --- colour ----------------------------------------------------------------
colormap(ax, o.Colormap);
lims = o.CLim;
if isempty(lims)
    lims = [min(Tg) max(Tg)];
    if lims(1) == lims(2), lims = lims + [-0.5 0.5]; end
end
set(ax, 'CLim', lims);
h.cbar = colorbar(ax);
% ticks run exactly from the coldest to the hottest grid; the min / max values
% go in the bar's title and side label so long tick text is never clipped
tk = linspace(lims(1), lims(2), 9);
h.cbar.Ticks = tk;
h.cbar.TickLabels = arrayfun(@(v) sprintf('%.1f', v), tk, 'UniformOutput', false);
h.cbar.Title.String = sprintf('max %.1f', lims(2));
h.cbar.Title.FontWeight = 'bold';
h.cbar.Label.String = sprintf('T [deg %s]      min %.1f  /  max %.1f', units, lims(1), lims(2));

% --- dressing --------------------------------------------------------------
axis(ax, 'equal'); axis(ax, 'vis3d'); grid(ax, 'on'); box(ax, 'on');
xlabel(ax, 'X'); ylabel(ax, 'Y'); zlabel(ax, 'Z');
head = sprintf('%s    t = %g s', shortname(r.csv_file), r.time);
if isfield(r, 'part_names') && numel(r.part_names) > 1
    head = sprintf('%s    (%d parts: %s)', head, numel(r.part_names), strjoin(r.part_names, ', '));
elseif isfield(r, 'part') && ~isempty(r.part) && ~strcmp(r.part, 'ALL')
    head = sprintf('%s    [%s]', head, r.part);
end
if ~standalone
    % embedded in a GUI (a uiaxes in a grid): a one-line title stays inside the cell
    title(ax, sprintf('%s    SID %d    T = %.1f .. %.1f %s%s', head, r.sid, min(Tg), max(Tg), units, note), ...
          'Interpreter', 'none');
else
title(ax, {head, ...
           sprintf('SID %d    T = %.1f .. %.1f %s    %s%s', r.sid, min(Tg), max(Tg), units, r.method, note)}, ...
      'Interpreter', 'none');
end
legend(ax, 'Location', 'northeast');
hold(ax, 'off');
h.set_view = @(name) set_view(ax, name);
h.set_view(o.View);
% drag = rotate, scroll = zoom, shift-drag / right-drag = pan (R2019a+ interactions);
% older releases fall back to rotate3d mode (zoom via the figure toolbar there).
try
    ax.Interactions = [rotateInteraction zoomInteraction panInteraction];
catch
    try rotate3d(ax, 'on'); catch, end
end

% --- view buttons (standalone figure only) ---------------------------------
if standalone && o.Buttons
    labels = {'+X', '-X', '+Y', '-Y', '+Z', '-Z', 'ISO'};
    w = 0.055; x0 = 0.08;
    for k = 1:numel(labels)
        uicontrol(fig, 'Style', 'pushbutton', 'String', labels{k}, ...
                  'Units', 'normalized', 'Position', [x0 + (k-1)*(w+0.005), 0.935, w, 0.045], ...
                  'Callback', @(~, ~) set_view(ax, labels{k}));
    end
    uicontrol(fig, 'Style', 'checkbox', 'String', 'cloud', 'Value', o.ShowCloud, ...
              'Units', 'normalized', 'Position', [0.56 0.935 0.10 0.045], ...
              'BackgroundColor', 'w', ...
              'Callback', @(src, ~) set(h.cloud, 'Visible', onoff(src.Value)));
    uicontrol(fig, 'Style', 'checkbox', 'String', 'surface', 'Value', o.ShowSurface, ...
              'Units', 'normalized', 'Position', [0.66 0.935 0.10 0.045], ...
              'BackgroundColor', 'w', ...
              'Callback', @(src, ~) set(h.surface, 'Visible', onoff(src.Value)));
    uicontrol(fig, 'Style', 'checkbox', 'String', 'outside hull', 'Value', o.ShowExtrap, ...
              'Units', 'normalized', 'Position', [0.76 0.935 0.14 0.045], ...
              'BackgroundColor', 'w', ...
              'Callback', @(src, ~) set(h.extrap, 'Visible', onoff(src.Value && any(ex))));
end
end


% =========================================================================
function set_view(ax, name)
%SET_VIEW  Snap to a named view.  '+X' = looking at the model FROM +X.
%   Camera is placed explicitly (direction + up vector) so every axis view is
%   upright and un-mirrored, then the model is re-framed to fill the axes.
%   'FIT' keeps the current direction and only re-frames (undoes zoom / pan).
    cur_dir = campos(ax) - camtarget(ax);
    cur_up  = camup(ax);
    camva(ax, 'auto'); camtarget(ax, 'auto'); campos(ax, 'auto'); camup(ax, 'auto');
    axis(ax, 'auto'); axis(ax, 'equal'); axis(ax, 'tight');
    set(ax, 'Projection', 'orthographic');
    switch upper(strtrim(char(name)))
        case '+X',  dirn = [ 1  0  0]; up = [0 0 1];
        case '-X',  dirn = [-1  0  0]; up = [0 0 1];
        case '+Y',  dirn = [ 0  1  0]; up = [0 0 1];
        case '-Y',  dirn = [ 0 -1  0]; up = [0 0 1];
        case '+Z',  dirn = [ 0  0  1]; up = [0 1 0];
        case '-Z',  dirn = [ 0  0 -1]; up = [0 1 0];
        case 'ISO', dirn = [-1 -1  1] / sqrt(3); up = [0 0 1];   % MATLAB's classic view(3) feel
        case 'FIT', dirn = cur_dir / max(norm(cur_dir), eps); up = cur_up;
        otherwise
            error('temp_map_plot:badView', 'Unknown view "%s".', name);
    end
    lim = [xlim(ax); ylim(ax); zlim(ax)];
    ctr = mean(lim, 2)';
    span = max(diff(lim, 1, 2));
    camtarget(ax, ctr);
    campos(ax, ctr + dirn * span * 4);
    camup(ax, up);
    camva(ax, 'auto');
    axis(ax, 'vis3d');                      % freeze aspect so dragging stays rigid
end


function sel = thin(n, nmax)
%THIN  Indices of a reproducible random subset when n > nmax.
    if n <= nmax
        sel = (1:n)';
    else
        rs = RandStream('mt19937ar', 'Seed', 0);
        sel = sort(randperm(rs, n, nmax))';
    end
end


function s = onoff(tf)
    if tf, s = 'on'; else, s = 'off'; end
end


function s = shortname(p)
    [~, b, e] = fileparts(p);
    s = [b e];
end
