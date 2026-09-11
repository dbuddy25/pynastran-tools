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
%       'Units'       'K' (default) | 'C'  -- display units
%       'ShowCloud'   true | false           (default true)
%       'ShowSurface' true | false           (default true)  translucent skin of
%                     the cloud (alpha shape) so you can see where the thermal
%                     volume actually ends
%       'ShowExtrap'  true | false           (default true)
%       'Colormap'    'jet' (default), 'parula', 'turbo', ...
%       'CLim'        [] = auto from the grid temps, else [lo hi]
%       'MarkerSize'  grid marker area      (default 36)
%       'CloudSize'   cloud marker area     (default 8)
%       'MaxPoints'   draw at most this many grids / cloud points (default 2e5);
%                     larger sets are randomly thinned for a responsive view
%       'View'        '+X' '-X' '+Y' '-Y' '+Z' '-Z' 'ISO'   (default 'ISO')
%       'Visible'     'on' | 'off'  (figure visibility, standalone only)
%
%   H is a struct of handles: .fig .ax .grids .cloud .extrap .cbar, plus
%   H.set_view(name) to snap the view programmatically.
%
%   See also TEMP_MAP_MATLAB, TEMP_MAP_GUI.

o = struct('Parent', [], 'Units', 'K', 'ShowCloud', true, 'ShowSurface', true, 'ShowExtrap', true, ...
           'Colormap', 'jet', 'CLim', [], 'MarkerSize', 36, 'CloudSize', 8, ...
           'MaxPoints', 2e5, 'View', 'ISO', 'Visible', 'on');
for k = 1:2:numel(varargin)
    name = char(varargin{k});
    if ~isfield(o, name)
        error('temp_map_plot:badOption', 'Unknown option "%s".', name);
    end
    o.(name) = varargin{k+1};
end

units = upper(char(o.Units));
off = 0; if units == 'C', off = 273.15; end
Tg = r.grid_T - off;
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
    cla(ax);
end
hold(ax, 'on');

% --- data ----------------------------------------------------------------
h = struct('fig', fig, 'ax', ax, 'grids', [], 'cloud', [], 'surface', [], 'extrap', [], 'cbar', []);

if isfield(r, 'surface') && ~isempty(r.surface)
    h.surface = patch('Parent', ax, 'Faces', r.surface.F, 'Vertices', r.surface.V, ...
                      'FaceColor', [0.6 0.6 0.6], 'FaceAlpha', 0.18, 'EdgeColor', 'none', ...
                      'DisplayName', 'cloud surface (alpha shape)');
    h.surface.Visible = onoff(o.ShowSurface);
else
    h.surface = patch('Parent', ax, 'Faces', [], 'Vertices', zeros(0, 3), ...
                      'DisplayName', 'cloud surface (n/a)', 'Visible', 'off');
end

h.cloud = scatter3(ax, cxyz(:,1), cxyz(:,2), cxyz(:,3), ...
                   o.CloudSize, [0.55 0.55 0.55], '.', 'DisplayName', 'cloud points');
h.cloud.Visible = onoff(o.ShowCloud);

h.grids = scatter3(ax, gxyz(:,1), gxyz(:,2), gxyz(:,3), ...
                   o.MarkerSize, Tg(gsel), 'filled', 'DisplayName', 'grids (mapped T)');

exs = ex(gsel);
h.extrap = scatter3(ax, gxyz(exs,1), gxyz(exs,2), gxyz(exs,3), ...
                    o.MarkerSize * 2.2, 'Marker', 'o', 'MarkerEdgeColor', 'k', ...
                    'MarkerFaceColor', 'none', 'LineWidth', 1.0, ...
                    'DisplayName', sprintf('outside hull / far (%d)', nnz(ex)));
h.extrap.Visible = onoff(o.ShowExtrap && any(ex));

% --- colour ----------------------------------------------------------------
colormap(ax, o.Colormap);
lims = o.CLim;
if isempty(lims)
    lims = [min(Tg) max(Tg)];
    if lims(1) == lims(2), lims = lims + [-0.5 0.5]; end
end
set(ax, 'CLim', lims);
h.cbar = colorbar(ax);
h.cbar.Label.String = sprintf('T [deg %s]', units);

% --- dressing --------------------------------------------------------------
axis(ax, 'equal'); axis(ax, 'vis3d'); grid(ax, 'on'); box(ax, 'on');
xlabel(ax, 'X'); ylabel(ax, 'Y'); zlabel(ax, 'Z');
title(ax, sprintf('%s    t = %g s    SID %d    T = %.1f .. %.1f %s%s', ...
      shortname(r.csv_file), r.time, r.sid, min(Tg), max(Tg), units, note), ...
      'Interpreter', 'none');
legend(ax, 'Location', 'northeast');
hold(ax, 'off');
h.set_view = @(name) set_view(ax, name);
h.set_view(o.View);
try rotate3d(ax, 'on'); catch, end   % uiaxes on older releases: use the axes toolbar

% --- view buttons (standalone figure only) ---------------------------------
if standalone
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
    switch upper(strtrim(char(name)))
        case '+X',  view(ax,  90,   0);
        case '-X',  view(ax, -90,   0);
        case '+Y',  view(ax, 180,   0);
        case '-Y',  view(ax,   0,   0);
        case '+Z',  view(ax,   0,  90);
        case '-Z',  view(ax,   0, -90);
        case 'ISO', view(ax, -37.5, 30);
        otherwise
            error('temp_map_plot:badView', 'Unknown view "%s".', name);
    end
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
