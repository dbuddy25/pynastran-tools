function fig = temp_map_gui()
%TEMP_MAP_GUI  Point-and-click front end for TEMP_MAP_MATLAB.
%
%   TEMP_MAP_GUI opens a window: pick the BDF and the CSV folder, tick the
%   clouds you want, press "Load & Map" to see the mapped temperatures on a
%   rotatable 3D view of the grids, then "Write" to emit the TEMP card files.
%
%   All the work is done by TEMP_MAP_MATLAB (headless engine) and
%   TEMP_MAP_PLOT (the picture); this file only wires up the controls.
%   Anything you can do here you can do from the command line with those two.
%
%   See also TEMP_MAP_MATLAB, TEMP_MAP_PLOT.

R = [];          % results from the last Load & Map
G = [];          % cached grids from Read BDF (ids, xyz, bdf_file)
H = [];          % plot handles from temp_map_plot
VIEWS = {'+X', '-X', '+Y', '-Y', '+Z', '-Z', 'ISO'};

% =========================================================================
%                                 LAYOUT
% =========================================================================
fig = uifigure('Name', 'TEMP Mapper  --  CSV temperature cloud -> Nastran TEMP cards', ...
               'Position', [80 80 1280 760]);
root = uigridlayout(fig, [1 2]);
root.ColumnWidth = {520, '1x'};
root.Padding = [8 8 8 8];

% --- left column -----------------------------------------------------------
L = uigridlayout(root, [20 3]);
L.ColumnWidth = {90, '1x', 34};
L.RowHeight   = {24, 30, 24, 150, 24, 24, 24, 24, 24, 24, 24, 24, 30, 30, 30, 24, 24, 24, 150, '1x'};
L.RowSpacing  = 6;
L.Padding     = [0 0 0 0];

row = 1;
lbl(L, row, 'BDF file');
bdfE = uieditfield(L, 'text', 'Placeholder', 'model.bdf', ...
                   'ValueChangedFcn', @(~, ~) invalidate_grids());
bdfE.Layout.Row = row; bdfE.Layout.Column = 2;
b = uibutton(L, 'Text', '...', 'ButtonPushedFcn', @on_browse_bdf);
b.Layout.Row = row; b.Layout.Column = 3;

row = 2;
readB = uibutton(L, 'Text', 'Read BDF', 'ButtonPushedFcn', @on_read_bdf);
readB.Layout.Row = row; readB.Layout.Column = [1 3];

row = 3;
lbl(L, row, 'CSV folder');
csvE = uieditfield(L, 'text', 'Placeholder', 'folder of *.csv clouds', ...
                   'ValueChangedFcn', @(~, ~) refresh_list());
csvE.Layout.Row = row; csvE.Layout.Column = 2;
b = uibutton(L, 'Text', '...', 'ButtonPushedFcn', @on_browse_csv);
b.Layout.Row = row; b.Layout.Column = 3;

row = 4;
csvLB = uilistbox(L, 'Items', {}, 'Multiselect', 'on');
csvLB.Layout.Row = row; csvLB.Layout.Column = [1 3];

row = 5;
lbl(L, row, 'Output dir');
outE = uieditfield(L, 'text', 'Value', 'temp_cards');
outE.Layout.Row = row; outE.Layout.Column = 2;
b = uibutton(L, 'Text', '...', 'ButtonPushedFcn', @on_browse_out);
b.Layout.Row = row; b.Layout.Column = 3;

row = 6;
lbl(L, row, 'Units out');
unitsDD = uidropdown(L, 'Items', {'K', 'C'}, 'Value', 'K', ...
                     'ValueChangedFcn', @(~, ~) replot());
unitsDD.Layout.Row = row; unitsDD.Layout.Column = [2 3];

row = 7;
lbl(L, row, 'SID start');
sidE = uieditfield(L, 'numeric', 'Value', 1, 'Limits', [1 Inf], 'RoundFractionalValues', 'on');
sidE.Layout.Row = row; sidE.Layout.Column = [2 3];

row = 8;
lbl(L, row, 'Field size');
fieldDD = uidropdown(L, 'Items', {'8 (small field)', '16 (large field)'}, ...
                     'ItemsData', [8 16], 'Value', 8);
fieldDD.Layout.Row = row; fieldDD.Layout.Column = [2 3];

row = 9;
lbl(L, row, 'Method');
methodDD = uidropdown(L, 'Items', {'linear (Delaunay) - exact, slow at 1M', ...
                                   'nearest (kd-tree) - fast', ...
                                   'IDW k=8 (kd-tree) - fast, smooth'}, ...
                      'ItemsData', {'linear', 'nearest', 'idw'}, 'Value', 'linear');
methodDD.Layout.Row = row; methodDD.Layout.Column = [2 3];

row = 10;
lbl(L, row, 'Warn dist');
warnE = uieditfield(L, 'numeric', 'Value', 0, 'Limits', [0 Inf], ...
                    'Tooltip', 'Warn when a grid is farther than this from any cloud point. 0 = off.');
warnE.Layout.Row = row; warnE.Layout.Column = [2 3];

row = 11;
tempdCB = uicheckbox(L, 'Text', 'Also write TEMPD (mean T)', 'Value', false);
tempdCB.Layout.Row = row; tempdCB.Layout.Column = [1 3];

row = 12;
headerCB = uicheckbox(L, 'Text', 'CSV has header row', 'Value', true);
headerCB.Layout.Row = row; headerCB.Layout.Column = [1 3];

row = 13;
mapB = uibutton(L, 'Text', 'Load & Map', 'FontWeight', 'bold', ...
                'ButtonPushedFcn', @on_map);
mapB.Layout.Row = row; mapB.Layout.Column = [1 3];

row = 14;
wselB = uibutton(L, 'Text', 'Write selected', 'Enable', 'off', ...
                 'ButtonPushedFcn', @(~, ~) on_write(false));
wselB.Layout.Row = row; wselB.Layout.Column = [1 3];

row = 15;
wallB = uibutton(L, 'Text', 'Write all', 'Enable', 'off', ...
                 'ButtonPushedFcn', @(~, ~) on_write(true));
wallB.Layout.Row = row; wallB.Layout.Column = [1 3];

row = 16;
cloudCB = uicheckbox(L, 'Text', 'Show cloud', 'Value', true, ...
                     'ValueChangedFcn', @(src, ~) toggle('cloud', src.Value));
cloudCB.Layout.Row = row; cloudCB.Layout.Column = [1 2];
extrapCB = uicheckbox(L, 'Text', 'Ring outside/far', 'Value', true, ...
                      'ValueChangedFcn', @(src, ~) toggle('extrap', src.Value));
extrapCB.Layout.Row = row; extrapCB.Layout.Column = [2 3];

row = 17;
surfCB = uicheckbox(L, 'Text', 'Show cloud surface (alpha shape)', 'Value', true, ...
                    'ValueChangedFcn', @(src, ~) toggle('surface', src.Value));
surfCB.Layout.Row = row; surfCB.Layout.Column = [1 3];

row = 18;
lbl(L, row, 'Colormap');
cmapDD = uidropdown(L, 'Items', {'jet', 'parula', 'turbo', 'hot', 'cool'}, 'Value', 'jet', ...
                    'ValueChangedFcn', @(src, ~) colormap_now(src.Value));
cmapDD.Layout.Row = row; cmapDD.Layout.Column = [2 3];

row = 19;
sumT = uitable(L, 'ColumnName', {'File', 'Time', 'SID', 'Cloud', 'Grids', 'Outside', 'Far', 'Tmin', 'Tmax'}, ...
               'ColumnWidth', {120, 55, 40, 65, 65, 55, 45, 60, 60}, 'RowName', [], 'Data', {});
sumT.Layout.Row = row; sumT.Layout.Column = [1 3];

row = 20;
statusTA = uitextarea(L, 'Editable', 'off', 'FontName', 'Courier New', 'FontSize', 11, ...
                      'Value', {'Pick a BDF and a CSV folder, then Load & Map.'});
statusTA.Layout.Row = row; statusTA.Layout.Column = [1 3];

% --- right column ----------------------------------------------------------
Rg = uigridlayout(root, [2 1]);
Rg.RowHeight = {30, '1x'};
Rg.Padding = [0 0 0 0];

bar = uigridlayout(Rg, [1 9]);
bar.ColumnWidth = [{50, '1x'}, repmat({48}, 1, 7)];
bar.Padding = [0 0 0 0];
t = uilabel(bar, 'Text', 'View:');
t.Layout.Column = 1;
viewDD = uidropdown(bar, 'Items', {'(nothing mapped yet)'}, ...
                    'ValueChangedFcn', @(~, ~) on_view_change());
viewDD.Layout.Column = 2;
for k = 1:numel(VIEWS)
    b = uibutton(bar, 'Text', VIEWS{k}, 'ButtonPushedFcn', @(src, ~) snap(src.Text));
    b.Layout.Column = 2 + k;
end

ax = uiaxes(Rg);
ax.Layout.Row = 2;
title(ax, 'Load & Map to see the grids coloured by temperature');

% =========================================================================
%                                CALLBACKS
% =========================================================================
    function on_browse_bdf(~, ~)
        [f, p] = uigetfile({'*.bdf;*.dat;*.nas;*.blk;*.inc', 'Nastran decks'; '*.*', 'All files'}, ...
                           'Pick the BDF');
        figure(fig);
        if isequal(f, 0), return; end
        bdfE.Value = fullfile(p, f);
        invalidate_grids();
        if isempty(csvE.Value), csvE.Value = p; refresh_list(); end
    end

    function invalidate_grids()
        G = [];
        readB.Text = 'Read BDF';
    end

    function ok = on_read_bdf(~, ~)
        ok = false;
        dlg = uiprogressdlg(fig, 'Title', 'Reading BDF', 'Indeterminate', 'on', ...
                            'Message', sprintf('Parsing GRIDs from %s ...', shortname(bdfE.Value)));
        t0 = tic;
        try
            G = temp_map_matlab('BDF_FILE', bdfE.Value, 'READ_ONLY', true);
        catch ME
            close(dlg);
            uialert(fig, ME.message, 'Read failed');
            return
        end
        close(dlg);
        readB.Text = sprintf('BDF loaded: %d grids  (re-read)', numel(G.ids));
        status(sprintf('Read %d grids from %s in %.1f s', numel(G.ids), shortname(G.bdf_file), toc(t0)));
        ok = true;
    end

    function on_browse_csv(~, ~)
        start = csvE.Value; if isempty(start), start = pwd; end
        p = uigetdir(start, 'Folder holding the temperature CSVs');
        figure(fig);
        if isequal(p, 0), return; end
        csvE.Value = p;
        refresh_list();
    end

    function on_browse_out(~, ~)
        p = uigetdir(pwd, 'Output folder for TEMP card files');
        figure(fig);
        if isequal(p, 0), return; end
        outE.Value = p;
    end

    function refresh_list()
        d = dir(fullfile(csvE.Value, '*.csv'));
        names = {d.name};
        keys = regexprep(names, '(\d+)', '${sprintf(''%012d'', str2double($1))}');
        [~, order] = sort(lower(keys));
        names = names(order);
        csvLB.Items = names;
        csvLB.Value = names;                 % everything selected by default
        status(sprintf('%d CSV file(s) found in %s', numel(names), csvE.Value));
    end

    function on_map(~, ~)
        sel = cellstr(csvLB.Value);
        if isempty(sel)
            uialert(fig, 'No CSV files selected.', 'Nothing to map'); return
        end
        if isempty(G) && ~on_read_bdf(), return; end   % first run: read it now
        dlg = uiprogressdlg(fig, 'Title', 'Mapping', 'Value', 0, 'Cancelable', 'on', ...
                            'Message', 'Starting ...');
        t0 = tic;
        prog = @(frac, msg) set_progress(dlg, frac, msg, t0);
        try
            [R, S] = temp_map_matlab( ...
                'GRIDS',             G, ...
                'CSV_DIR',           csvE.Value, ...
                'CSV_FILES',         fullfile(csvE.Value, sel), ...
                'CSV_HAS_HEADER',    headerCB.Value, ...
                'SID_START',         sidE.Value, ...
                'METHOD',            methodDD.Value, ...
                'EXTRAP_WARN_DIST',  warn_dist(), ...
                'OUT_UNITS',         unitsDD.Value, ...
                'PROGRESS',          prog, ...
                'WRITE',             false);
        catch ME
            close(dlg);
            R = [];
            if strcmp(ME.identifier, 'temp_map_gui:cancelled')
                status('Cancelled.');
            else
                uialert(fig, ME.message, 'Mapping failed');
            end
            return
        end
        close(dlg);
        viewDD.Items = cellfun(@shortname, {R.csv_file}, 'UniformOutput', false);
        viewDD.ItemsData = 1:numel(R);
        viewDD.Value = 1;
        wselB.Enable = 'on'; wallB.Enable = 'on';
        fill_summary(S);
        replot();
        show_coverage();
    end

    function on_write(all_of_them)
        if isempty(R), return; end
        if all_of_them
            sub = R;
        else
            want = cellfun(@shortname, cellstr(csvLB.Value), 'UniformOutput', false);
            have = cellfun(@shortname, {R.csv_file}, 'UniformOutput', false);
            sub = R(ismember(have, want));
            if isempty(sub)
                uialert(fig, 'None of the highlighted CSVs have been mapped yet.', 'Nothing to write');
                return
            end
        end
        try
            [~, S] = temp_map_matlab( ...
                'RESULTS',      sub, ...
                'OUT_DIR',      outE.Value, ...
                'OUT_UNITS',    unitsDD.Value, ...
                'FIELD_SIZE',   fieldDD.Value, ...
                'WRITE_TEMPD',  tempdCB.Value, ...
                'WRITE',        true);
        catch ME
            uialert(fig, ME.message, 'Write failed');
            return
        end
        fill_summary(S);
        status([{sprintf('Wrote %d file(s) to %s:', height(S), outE.Value)}; S.OutFile(:)]);
    end

    function show_coverage()
        k = viewDD.Value;
        status([{sprintf('Coverage: %s', shortname(R(k).csv_file))}; R(k).coverage(:)]);
    end

    function fill_summary(S)
        n = height(S);
        far = zeros(n, 1);
        for k = 1:n, far(k) = nnz(R(k).far); end
        sumT.Data = [S.File, num2cell(S.Time), num2cell(S.SID), num2cell(S.CloudPts), ...
                     num2cell(S.Grids), num2cell(S.Extrap), num2cell(far), ...
                     num2cell(round(S.Tmin, 2)), num2cell(round(S.Tmax, 2))];
    end

    function replot()
        if isempty(R) || isempty(viewDD.ItemsData), return; end
        k = viewDD.Value;
        H = temp_map_plot(R(k), 'Parent', ax, ...
                          'Units',      unitsDD.Value, ...
                          'ShowCloud',  cloudCB.Value, ...
                          'ShowSurface', surfCB.Value, ...
                          'ShowExtrap', extrapCB.Value, ...
                          'Colormap',   cmapDD.Value);
    end

    function on_view_change()
        replot();
        show_coverage();
    end

    function snap(name)
        if isempty(H), return; end
        H.set_view(name);
    end

    function toggle(what, on)
        if isempty(H), return; end
        if strcmp(what, 'extrap'), on = on && any(R(viewDD.Value).extrap | R(viewDD.Value).far); end
        if on, v = 'on'; else, v = 'off'; end
        H.(what).Visible = v;
    end

    function colormap_now(name)
        if isempty(H), return; end
        colormap(ax, name);
    end

    function d = warn_dist()
        d = warnE.Value;
        if d <= 0, d = []; end
    end

    function status(lines)
        statusTA.Value = cellstr(lines);
    end
end


% =========================================================================
function set_progress(dlg, frac, msg, t0)
%   Called by the engine between stages; a running triangulation cannot be
%   interrupted, so Cancel takes effect at the next stage boundary.
    if dlg.CancelRequested
        error('temp_map_gui:cancelled', 'Cancelled by user.');
    end
    dlg.Value = frac;
    dlg.Message = sprintf('%s   (%.0f s elapsed)', msg, toc(t0));
    drawnow limitrate
end


function lbl(parent, row, text)
    t = uilabel(parent, 'Text', text);
    t.Layout.Row = row; t.Layout.Column = 1;
end


function s = shortname(p)
    [~, b, e] = fileparts(p);
    s = [b e];
end
