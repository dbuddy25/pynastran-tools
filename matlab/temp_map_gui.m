function fig = temp_map_gui()
%TEMP_MAP_GUI  Point-and-click front end for TEMP_MAP_MATLAB.
%
%   TEMP_MAP_GUI opens a window laid out as the four steps of the job:
%       1  Model    - pick the BDF, read it once (grids + element faces cached)
%       2  Clouds   - pick the folder of CSV temperature clouds, tick the ones to map
%       3  Map      - choose the method, Load & Map -> 3D preview + coverage check
%       4  Write    - emit one bulk-data file of TEMP cards per CSV
%   Paths, units and method are remembered between sessions (setpref).
%
%   All the work is done by TEMP_MAP_MATLAB (headless engine) and
%   TEMP_MAP_PLOT (the picture); this file only wires up the controls.
%   Anything you can do here you can do from the command line with those two.
%
%   See also TEMP_MAP_MATLAB, TEMP_MAP_PLOT.

R = [];          % results from the last Load & Map
G = [];          % cached grids from Read BDF (ids, xyz, faces, bdf_file)
H = [];          % plot handles from temp_map_plot
hNow = [];       % current-case marker line on the time strip
VIEWS = {'+X', '-X', '+Y', '-Y', '+Z', '-Z', 'ISO'};
PREF = 'temp_map_gui';
ACCENT = [0.13 0.45 0.80];
LEN_M  = struct('in', 0.0254, 'mm', 1e-3, 'm', 1);   % mirrors the engine's table

% =========================================================================
%                                 WINDOW
% =========================================================================
scr = get(0, 'ScreenSize');
W = max(1100, min(1400, round(0.85 * scr(3))));
Hh = max(720, min(980, round(0.85 * scr(4))));
fig = uifigure('Name', 'TEMP Mapper  --  CSV temperature cloud -> Nastran TEMP cards', ...
               'Position', [round((scr(3) - W) / 2), round((scr(4) - Hh) / 2) - 20, W, Hh]);
root = uigridlayout(fig, [2 2]);
root.ColumnWidth = {540, '1x'};
root.RowHeight   = {'1x', 200};
root.Padding = [8 8 8 8];

% =========================================================================
%                              LEFT: STEPS
% =========================================================================
Lg = uigridlayout(root, [7 1]);
Lg.RowHeight = {26, 122, '1x', 96, 126, 150, 128};
Lg.Padding = [0 0 0 0]; Lg.RowSpacing = 6;

% --- setup file: everything below except the CSV list ----------------------------
sb = uigridlayout(Lg, [1 4]);
sb.ColumnWidth = {'1x', 90, 90, 90}; sb.Padding = [0 0 0 0]; sb.ColumnSpacing = 6;
setupLbl = put(uilabel(sb, 'Text', 'Setup: (last session)', 'FontColor', [0.45 0.45 0.45]), 1, 1);
put(uibutton(sb, 'Text', 'Load setup', 'Tooltip', 'Restore BDF, units, method, output and SID settings from a .json file', ...
             'ButtonPushedFcn', @on_load_setup), 1, 2);
put(uibutton(sb, 'Text', 'Save setup', 'Tooltip', 'Save the current settings (not the CSV list) to a .json file', ...
             'ButtonPushedFcn', @on_save_setup), 1, 3);
put(uibutton(sb, 'Text', 'Reset', 'Tooltip', 'Back to defaults', ...
             'ButtonPushedFcn', @on_reset_setup), 1, 4);

% --- 1  Model -----------------------------------------------------------------
p1 = uipanel(Lg, 'Title', '1  Model', 'FontWeight', 'bold');
g1 = uigridlayout(p1, [3 4]);
g1.ColumnWidth = {60, '1x', 64, 64}; g1.RowHeight = {24, 26, 24};
g1.Padding = [8 4 8 4]; g1.RowSpacing = 5;

put(uilabel(g1, 'Text', 'BDF'), 1, 1);
bdfE = put(uieditfield(g1, 'text', 'Placeholder', 'model.bdf', ...
           'Tooltip', 'Nastran bulk data deck. INCLUDEs are followed.', ...
           'ValueChangedFcn', @(~, ~) on_bdf_edit()), 1, [2 3]);
put(uibutton(g1, 'Text', 'Browse', 'ButtonPushedFcn', @on_browse_bdf), 1, 4);

readB = put(uibutton(g1, 'Text', 'Read BDF', 'Enable', 'off', ...
            'Tooltip', 'Parse GRIDs and element faces once; cached until the path changes.', ...
            'ButtonPushedFcn', @on_read_bdf), 2, [1 2]);
modelLbl = put(uilabel(g1, 'Text', 'not read yet', 'FontColor', [0.45 0.45 0.45]), 2, [3 4]);

put(uilabel(g1, 'Text', 'Units'), 3, 1);
su = uigridlayout(g1, [1 5]); put(su, 3, [2 4]);
su.ColumnWidth = {44, 60, 38, 60, '1x'}; su.Padding = [0 0 0 0]; su.ColumnSpacing = 4;
uilabel(su, 'Text', 'length');
bdfUnitsDD = uidropdown(su, 'Items', {'in', 'mm', 'm'}, 'Value', 'in', ...
                        'Tooltip', 'Length units of the structural model (BDF)', ...
                        'ValueChangedFcn', @(~, ~) on_len_units());
uilabel(su, 'Text', 'temp');
unitsDD = uidropdown(su, 'Items', {'K', 'C', 'F'}, 'Value', 'K', ...
                     'Tooltip', 'Temperature units of the structural model = what is written on the TEMP cards', ...
                     'ValueChangedFcn', @(~, ~) on_units_change());
uilabel(su, 'Text', '= TEMP card units', 'FontColor', [0.45 0.45 0.45]);

% --- 2  Temperature clouds ---------------------------------------------------
p2 = uipanel(Lg, 'Title', '2  Temp clouds', 'FontWeight', 'bold');
g2 = uigridlayout(p2, [4 4]);
g2.ColumnWidth = {60, '1x', 64, 64}; g2.RowHeight = {24, '1x', 24, 22};
g2.Padding = [8 4 8 4]; g2.RowSpacing = 5;

put(uilabel(g2, 'Text', 'Folder'), 1, 1);
csvE = put(uieditfield(g2, 'text', 'Placeholder', 'folder of *.csv clouds', ...
           'Tooltip', 'Every *.csv here is listed, natural-sorted (t2 before t10).', ...
           'ValueChangedFcn', @(~, ~) refresh_list()), 1, [2 3]);
put(uibutton(g2, 'Text', 'Browse', 'ButtonPushedFcn', @on_browse_csv), 1, 4);

csvLB = put(uilistbox(g2, 'Items', {}, 'Multiselect', 'on', ...
            'Tooltip', 'Ctrl / Shift-click to choose which time steps to map.', ...
            'ValueChangedFcn', @(~, ~) update_state()), 2, [1 4]);

put(uilabel(g2, 'Text', 'Units'), 3, 1);
cu = uigridlayout(g2, [1 5]); put(cu, 3, [2 4]);
cu.ColumnWidth = {44, 60, 38, 60, '1x'}; cu.Padding = [0 0 0 0]; cu.ColumnSpacing = 4;
uilabel(cu, 'Text', 'length');
csvUnitsDD = uidropdown(cu, 'Items', {'in', 'mm', 'm'}, 'Value', 'in', ...
                        'Tooltip', 'Length units of the CSV x / y / z columns', ...
                        'ValueChangedFcn', @(~, ~) on_len_units());
uilabel(cu, 'Text', 'temp');
csvTempDD = uidropdown(cu, 'Items', {'K', 'C', 'F'}, 'Value', 'K', ...
                       'Tooltip', 'Temperature units of the CSV 5th column', ...
                       'ValueChangedFcn', @(~, ~) save_prefs());
factorLbl = uilabel(cu, 'Text', '', 'FontColor', [0.45 0.45 0.45]);

headerCB = put(uicheckbox(g2, 'Text', 'First row is a header', 'Value', true, ...
               'Tooltip', 'Untick if the CSV starts straight with numbers.', ...
               'ValueChangedFcn', @(~, ~) save_prefs()), 4, [1 2]);
put(uilabel(g2, 'Text', 'columns: time, x, y, z, T', 'FontColor', [0.45 0.45 0.45], ...
            'HorizontalAlignment', 'right'), 4, [3 4]);

% --- 3  Map ------------------------------------------------------------------
p3 = uipanel(Lg, 'Title', '3  Map', 'FontWeight', 'bold');
g3 = uigridlayout(p3, [2 4]);
g3.ColumnWidth = {60, '1x', 90, 64}; g3.RowHeight = {24, 30};
g3.Padding = [8 4 8 4]; g3.RowSpacing = 5;

put(uilabel(g3, 'Text', 'Method'), 1, 1);
methodDD = put(uidropdown(g3, ...
    'Items', {'Linear (recommended)', 'Nearest', 'IDW (k = 8)', 'scatteredInterpolant (reference)'}, ...
    'ItemsData', {'linear', 'nearest', 'idw', 'scattered'}, 'Value', 'linear', ...
    'Tooltip', sprintf(['Linear: Delaunay interpolation, nearest outside the hull, with a guard\n' ...
                        '   against smearing across concavities. Exact for volume clouds.\n' ...
                        'Nearest: closest cloud point (kd-tree, seconds at 1M).\n' ...
                        'IDW: inverse-distance mean of the 8 closest points (kd-tree).\n' ...
                        'scatteredInterpolant: MATLAB''s linear/nearest verbatim, no guard,\n' ...
                        '   for one-to-one comparison with other scripts.\n' ...
                        'The mapping is built once and reused while the cloud points do not move.']), ...
    'ValueChangedFcn', @(~, ~) save_prefs()), 1, 2);
put(uilabel(g3, 'Text', 'Warn distance', 'HorizontalAlignment', 'right'), 1, 3);
warnE = put(uieditfield(g3, 'numeric', 'Value', 0, 'Limits', [0 Inf], ...
            'Tooltip', 'Flag grids farther than this (model length units) from any cloud point. 0 = off.'), 1, 4);

mapB = put(uibutton(g3, 'Text', 'Load & Map', 'FontWeight', 'bold', 'FontSize', 13, ...
           'BackgroundColor', ACCENT, 'FontColor', 'w', 'Enable', 'off', ...
           'Tooltip', 'Map every ticked CSV onto the grids and show the result. Nothing is written yet.', ...
           'ButtonPushedFcn', @on_map), 2, [1 4]);

% --- 4  Write ----------------------------------------------------------------
p4 = uipanel(Lg, 'Title', '4  Write TEMP cards', 'FontWeight', 'bold');
g4 = uigridlayout(p4, [3 4]);
g4.ColumnWidth = {60, '1x', 90, 64}; g4.RowHeight = {24, 24, 28};
g4.Padding = [8 4 8 4]; g4.RowSpacing = 5;

put(uilabel(g4, 'Text', 'Output'), 1, 1);
outE = put(uieditfield(g4, 'text', 'Value', 'temp_cards', ...
           'Tooltip', 'One <csvname>_temp.bdf per CSV lands here (created if missing).', ...
           'ValueChangedFcn', @(~, ~) save_prefs()), 1, [2 3]);
put(uibutton(g4, 'Text', 'Browse', 'ButtonPushedFcn', @on_browse_out), 1, 4);

put(uilabel(g4, 'Text', 'First SID'), 2, 1);
sidE = put(uieditfield(g4, 'numeric', 'Value', 1, 'Limits', [1 Inf], 'RoundFractionalValues', 'on', ...
           'Tooltip', 'Load set ID of the first CSV; the rest count up in list order. Applied at Load & Map.'), 2, 2);
put(uilabel(g4, 'Text', 'Card format', 'HorizontalAlignment', 'right'), 2, 3);
fieldDD = put(uidropdown(g4, 'Items', {'small (8)', 'large (16)'}, 'ItemsData', [8 16], 'Value', 8, ...
              'Tooltip', 'Small field: 8-character columns. Large field (TEMP*): 16, for ids > 99,999,999 or more digits.', ...
              'ValueChangedFcn', @(~, ~) save_prefs()), 2, 4);

wr = uigridlayout(g4, [1 3]); put(wr, 3, [1 4]);
wr.ColumnWidth = {'1x', '1x', '1x'}; wr.Padding = [0 0 0 0]; wr.ColumnSpacing = 6;
wselB = uibutton(wr, 'Text', 'Write selected', 'Enable', 'off', ...
                 'Tooltip', 'Write the CSVs highlighted in the list (they must have been mapped).', ...
                 'ButtonPushedFcn', @(~, ~) on_write(false));
wallB = uibutton(wr, 'Text', 'Write all', 'Enable', 'off', ...
                 'Tooltip', 'Write every mapped CSV.', ...
                 'ButtonPushedFcn', @(~, ~) on_write(true));
tempdCB = uicheckbox(wr, 'Text', 'Add TEMPD (mean)', 'Value', false, ...
                     'Tooltip', 'Also write a TEMPD card with the mean temperature, the default for grids not listed.');

% --- results table + coverage -------------------------------------------------
sumT = uitable(Lg, 'ColumnName', {'File', 'Time', 'SID', 'Grids', 'Outside', 'Far', 'Tmin', 'Tmax'}, ...
               'ColumnWidth', {'1x', 55, 40, 62, 60, 44, 60, 60}, 'RowName', [], 'Data', {}, ...
               'Tooltip', 'One row per mapped CSV. Click a row to view it.', ...
               'CellSelectionCallback', @(~, ev) on_table_click(ev));
statusTA = uitextarea(Lg, 'Editable', 'off', 'FontName', 'Courier New', 'FontSize', 11, ...
                      'Value', {'1  pick a BDF and Read it    2  pick the CSV folder    3  Load & Map    4  Write'});

% =========================================================================
%                              RIGHT: VIEW
% =========================================================================
Rg = uigridlayout(root, [4 1]);
Rg.RowHeight = {30, 26, 24, '1x'};
Rg.Padding = [0 0 0 0]; Rg.RowSpacing = 4;

% --- row 1: case navigation + view presets -------------------------------------
bar = uigridlayout(Rg, [1 13]);
bar.ColumnWidth = [{44, 30, '1x', 30}, repmat({44}, 1, 7), {8, 50}];
bar.Padding = [0 0 0 0]; bar.ColumnSpacing = 4;
put(uilabel(bar, 'Text', 'Case'), 1, 1);
put(uibutton(bar, 'Text', '<', 'Tooltip', 'Previous case  (Left arrow)', ...
             'ButtonPushedFcn', @(~, ~) step_case(-1)), 1, 2);
viewDD = put(uidropdown(bar, 'Items', {'(nothing mapped yet)'}, ...
             'ValueChangedFcn', @(~, ~) on_view_change()), 1, 3);
put(uibutton(bar, 'Text', '>', 'Tooltip', 'Next case  (Right arrow)', ...
             'ButtonPushedFcn', @(~, ~) step_case(+1)), 1, 4);
for k = 1:numel(VIEWS)
    put(uibutton(bar, 'Text', VIEWS{k}, 'Tooltip', ['Look at the model from ' VIEWS{k} ' and re-frame'], ...
                 'ButtonPushedFcn', @(src, ~) snap(src.Text)), 1, 4 + k);
end
put(uibutton(bar, 'Text', 'Fit', 'Tooltip', 'Re-frame the model without changing the view direction  (Home)', ...
             'ButtonPushedFcn', @(~, ~) snap('FIT')), 1, 13);

% --- row 2: display options ----------------------------------------------------
db = uigridlayout(Rg, [1 8]);
db.ColumnWidth = {200, 90, 100, 110, 110, 70, 90, '1x'};
db.Padding = [0 0 0 0]; db.ColumnSpacing = 6;
styleDD = put(uidropdown(db, 'Items', {'Points (grids)', 'Smooth contour (element faces)'}, ...
              'ItemsData', {'points', 'contour'}, 'Value', 'points', ...
              'Tooltip', 'Contour paints the element faces read from the BDF with the temperature interpolated across each face.', ...
              'ValueChangedFcn', @(~, ~) replot()), 1, 1);
cloudCB = put(uicheckbox(db, 'Text', 'Cloud pts', 'Value', true, ...
              'ValueChangedFcn', @(src, ~) toggle('cloud', src.Value)), 1, 2);
surfCB = put(uicheckbox(db, 'Text', 'Cloud skin', 'Value', true, ...
             'Tooltip', 'Translucent alpha-shape surface of the cloud: where the thermal volume ends.', ...
             'ValueChangedFcn', @(src, ~) toggle('surface', src.Value)), 1, 3);
extrapCB = put(uicheckbox(db, 'Text', 'Ring outside', 'Value', true, ...
               'Tooltip', 'Black rings on grids outside the cloud hull or beyond the warn distance.', ...
               'ValueChangedFcn', @(src, ~) toggle('extrap', src.Value)), 1, 4);
mmCB = put(uicheckbox(db, 'Text', 'Min / max', 'Value', true, ...
           'Tooltip', 'Star + label on the hottest and coldest grid.', ...
           'ValueChangedFcn', @(src, ~) toggle('minmax', src.Value)), 1, 5);
put(uilabel(db, 'Text', 'Colormap', 'HorizontalAlignment', 'right'), 1, 6);
cmapDD = put(uidropdown(db, 'Items', {'jet', 'turbo', 'parula', 'hot', 'cool'}, 'Value', 'jet', ...
             'ValueChangedFcn', @(src, ~) colormap_now(src.Value)), 1, 7);

% --- row 3: coverage banner -----------------------------------------------------
banner = uilabel(Rg, 'Text', 'Coverage check appears here after Load & Map', ...
                 'FontWeight', 'bold', 'HorizontalAlignment', 'center', ...
                 'BackgroundColor', [0.94 0.94 0.94], 'FontColor', [0.35 0.35 0.35]);

% --- row 4: 3D axes -----------------------------------------------------------------
ax = uiaxes(Rg);
title(ax, 'Load & Map to see the grids coloured by temperature');

% --- bottom strip: min / max temperature vs time across all cases ------------
tax = uiaxes(root);
tax.Layout.Row = 2; tax.Layout.Column = [1 2];
title(tax, 'Min / max mapped temperature vs time');
xlabel(tax, 'time [s]'); grid(tax, 'on'); box(tax, 'on');
tax.ButtonDownFcn = @(~, ev) jump_to_time(ev.IntersectionPoint(1));

fig.KeyPressFcn = @(~, ev) on_key(ev);
load_prefs();
on_len_units();
if ~isempty(csvE.Value) && exist(csvE.Value, 'dir') == 7, refresh_list(); end
update_state();

% =========================================================================
%                                CALLBACKS
% =========================================================================
    function on_bdf_edit()
        invalidate_grids();
        save_prefs();
        update_state();
    end

    function on_browse_bdf(~, ~)
        start = bdfE.Value; if isempty(start) || exist(start, 'file') ~= 2, start = pwd; end
        [f, p] = uigetfile({'*.bdf;*.dat;*.nas;*.blk;*.inc', 'Nastran decks'; '*.*', 'All files'}, ...
                           'Pick the BDF', start);
        figure(fig);
        if isequal(f, 0), return; end
        bdfE.Value = fullfile(p, f);
        on_bdf_edit();
    end

    function invalidate_grids()
        G = [];
        modelLbl.Text = 'not read yet';
        modelLbl.FontColor = [0.45 0.45 0.45];
        p1.Title = '1  Model';
    end

    function ok = on_read_bdf(~, ~)
        ok = false;
        dlg = uiprogressdlg(fig, 'Title', 'Reading BDF', 'Value', 0, 'Cancelable', 'on', ...
                            'Message', 'Starting ...');
        t0 = tic;
        try
            G = temp_map_matlab('BDF_FILE', bdfE.Value, 'READ_ONLY', true, ...
                                'PROGRESS', @(frac, msg) set_progress(dlg, frac, msg, t0));
        catch ME
            close(dlg);
            if ~strcmp(ME.identifier, 'temp_map_gui:cancelled')
                uialert(fig, ME.message, 'Read failed');
            else
                status('Cancelled.');
            end
            return
        end
        close(dlg);
        modelLbl.Text = sprintf('%s grids, %s faces', fmtn(numel(G.ids)), fmtn(size(G.faces, 1)));
        modelLbl.FontColor = [0.1 0.5 0.2];
        p1.Title = sprintf('1  Model  --  %s grids', fmtn(numel(G.ids)));
        status(sprintf('Read %s grids, %s faces from %s in %.1f s', fmtn(numel(G.ids)), ...
                       fmtn(size(G.faces, 1)), shortname(G.bdf_file), toc(t0)));
        update_state();
        ok = true;
    end

    function on_browse_csv(~, ~)
        start = csvE.Value; if isempty(start) || exist(start, 'dir') ~= 7, start = pwd; end
        p = uigetdir(start, 'Folder holding the temperature CSVs');
        figure(fig);
        if isequal(p, 0), return; end
        csvE.Value = p;
        refresh_list();
    end

    function on_browse_out(~, ~)
        start = outE.Value; if isempty(start) || exist(start, 'dir') ~= 7, start = pwd; end
        p = uigetdir(start, 'Output folder for TEMP card files');
        figure(fig);
        if isequal(p, 0), return; end
        outE.Value = p;
        save_prefs();
    end

    function refresh_list()
        d = dir(fullfile(csvE.Value, '*.csv'));
        names = {d.name};
        keys = regexprep(names, '(\d+)', '${sprintf(''%012d'', str2double($1))}');
        [~, order] = sort(lower(keys));
        names = names(order);
        csvLB.Items = names;
        csvLB.Value = names;                 % everything selected by default
        save_prefs();
        update_state();
    end

    function on_len_units()
        f = LEN_M.(csvUnitsDD.Value) / LEN_M.(bdfUnitsDD.Value);
        if abs(f - 1) < 1e-12
            factorLbl.Text = 'same as model';
            factorLbl.FontColor = [0.45 0.45 0.45];
        else
            factorLbl.Text = sprintf('x %.6g into model %s', f, bdfUnitsDD.Value);
            factorLbl.FontColor = [0.80 0.45 0.05];
        end
        save_prefs();
    end

    function update_state()
        have_bdf = exist(bdfE.Value, 'file') == 2;
        nsel = numel(cellstr(csvLB.Value));
        readB.Enable = onoff(have_bdf);
        mapB.Enable  = onoff(have_bdf && nsel > 0);
        wselB.Enable = onoff(~isempty(R));
        wallB.Enable = onoff(~isempty(R));
        if isempty(csvLB.Items)
            p2.Title = '2  Temp clouds';
        else
            p2.Title = sprintf('2  Temp clouds  --  %d files, %d selected', numel(csvLB.Items), nsel);
        end
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
                'BDF_LENGTH_UNITS',  bdfUnitsDD.Value, ...
                'CSV_LENGTH_UNITS',  csvUnitsDD.Value, ...
                'CSV_TEMP_UNITS',    csvTempDD.Value, ...
                'SID_START',         sidE.Value, ...
                'METHOD',            methodDD.Value, ...
                'EXTRAP_WARN_DIST',  warn_dist(), ...
                'OUT_UNITS',         unitsDD.Value, ...
                'PROGRESS',          prog, ...
                'WRITE',             false);
        catch ME
            close(dlg);
            R = [];
            update_state();
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
        mname = methodDD.Items{strcmp(methodDD.ItemsData, methodDD.Value)};
        p3.Title = sprintf('3  Map  --  %d cases, %s, %s', numel(R), mname, datestr(now, 'HH:MM'));
        fill_summary(S);
        update_state();
        on_view_change();
        plot_history();
        bad = find(arrayfun(@(r) ~isempty(r.warnings), R));
        if ~isempty(bad)
            msg = {};
            for k = bad(:)'
                msg{end+1} = sprintf('%s:', shortname(R(k).csv_file));         %#ok<AGROW>
                msg = [msg, strcat({'    - '}, R(k).warnings(:)')];           %#ok<AGROW>
            end
            uialert(fig, strjoin(msg, newline), 'COVERAGE WARNING -- check units and extents', ...
                    'Icon', 'warning');
        end
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
                'BDF_LENGTH_UNITS',  bdfUnitsDD.Value, ...
                'CSV_LENGTH_UNITS',  csvUnitsDD.Value, ...
                'FIELD_SIZE',   fieldDD.Value, ...
                'WRITE_TEMPD',  tempdCB.Value, ...
                'WRITE',        true);
        catch ME
            uialert(fig, ME.message, 'Write failed');
            return
        end
        p4.Title = sprintf('4  Write TEMP cards  --  %d written %s', height(S), datestr(now, 'HH:MM'));
        status([{sprintf('Wrote %d file(s) to %s:', height(S), outE.Value)}; S.OutFile(:)]);
    end

    function show_coverage()
        k = viewDD.Value;
        lines = [{sprintf('Coverage: %s', shortname(R(k).csv_file))}; R(k).coverage(:)];
        if ~isempty(R(k).warnings)
            lines = [lines; {''; '!!!!! WARNING !!!!!'}; strcat({'!! '}, R(k).warnings(:))];
            banner.Text = ['WARNING  ' strjoin(R(k).warnings, '   |   ')];
            banner.BackgroundColor = [0.98 0.85 0.85];
            banner.FontColor = [0.65 0.05 0.05];
        else
            banner.Text = sprintf('Coverage OK  --  %.1f%% of grids outside the cloud hull, max distance to a cloud point %.3g', ...
                100 * nnz(R(k).extrap) / numel(R(k).grid_ids), max(R(k).nn_dist));
            banner.BackgroundColor = [0.86 0.95 0.86];
            banner.FontColor = [0.05 0.40 0.10];
        end
        status(lines);
    end

    function fill_summary(S)
        n = height(S);
        far = zeros(n, 1);
        for k = 1:n, far(k) = nnz(R(k).far); end
        sumT.Data = [S.File, num2cell(S.Time), num2cell(S.SID), ...
                     num2cell(S.Grids), num2cell(S.Extrap), num2cell(far), ...
                     num2cell(round(S.Tmin, 2)), num2cell(round(S.Tmax, 2))];
        highlight_row();
    end

    function highlight_row()
        try
            removeStyle(sumT);
            if ~isempty(R) && ~isempty(viewDD.ItemsData)
                addStyle(sumT, uistyle('BackgroundColor', [0.85 0.91 0.98], 'FontWeight', 'bold'), ...
                         'row', viewDD.Value);
            end
        catch
        end
    end

    function on_table_click(ev)
        if isempty(R) || isempty(ev.Indices), return; end
        k = ev.Indices(1);
        if k < 1 || k > numel(R) || k == viewDD.Value, return; end
        viewDD.Value = k;
        on_view_change();
    end

    function on_units_change()
        save_prefs();
        replot();
        if ~isempty(R), plot_history(); end
    end

    function replot()
        if isempty(R) || isempty(viewDD.ItemsData), return; end
        k = viewDD.Value;
        H = temp_map_plot(R(k), 'Parent', ax, 'Style', styleDD.Value, ...
                          'Units',      unitsDD.Value, ...
                          'ShowCloud',  cloudCB.Value, ...
                          'ShowSurface', surfCB.Value, ...
                          'ShowMinMax', mmCB.Value, ...
                          'ShowExtrap', extrapCB.Value, ...
                          'Colormap',   cmapDD.Value);
    end

    function on_view_change()
        replot();
        show_coverage();
        mark_history();
        highlight_row();
    end

    function plot_history()
        if isempty(R), return; end
        t = [R.time];
        [t, order] = sort(t);
        u = unitsDD.Value;
        tmin = arrayfun(@(r) conv_out(min(r.grid_T), u), R(order));
        tmax = arrayfun(@(r) conv_out(max(r.grid_T), u), R(order));
        cla(tax);
        hold(tax, 'on');
        plot(tax, t, tmax, '-o', 'Color', [0.85 0.1 0.1], 'LineWidth', 1.6, 'MarkerFaceColor', [0.85 0.1 0.1], ...
             'MarkerSize', 4, 'DisplayName', 'max', 'HitTest', 'off');
        plot(tax, t, tmin, '-o', 'Color', [0.1 0.3 0.85], 'LineWidth', 1.6, 'MarkerFaceColor', [0.1 0.3 0.85], ...
             'MarkerSize', 4, 'DisplayName', 'min', 'HitTest', 'off');
        hold(tax, 'off');
        ylabel(tax, sprintf('T [deg %s]', u));
        title(tax, sprintf('Min / max mapped temperature vs time  (%d cases; click to jump)', numel(R)));
        legend(tax, 'Location', 'northwest', 'Orientation', 'horizontal');
        grid(tax, 'on'); box(tax, 'on');
        if numel(t) > 1
            pad = 0.03 * (max(t) - min(t));
            xlim(tax, [min(t) - pad, max(t) + pad]);
        else
            xlim(tax, t + [-1 1] * max(1, abs(t)) * 0.1);
        end
        mark_history();
    end

    function mark_history()
        if ~isempty(hNow) && isvalid(hNow), delete(hNow); end
        hNow = [];
        if isempty(R) || isempty(viewDD.ItemsData), return; end
        tk = R(viewDD.Value).time;
        hNow = xline(tax, tk, '-', 'Color', [0.2 0.2 0.2], 'LineWidth', 1.5, ...
                     'HandleVisibility', 'off', 'HitTest', 'off');
    end

    function jump_to_time(tclick)
        if isempty(R), return; end
        [~, k] = min(abs([R.time] - tclick));
        viewDD.Value = k;
        on_view_change();
    end

    function v = conv_out(v, u)
        switch upper(u)
            case 'C', v = v - 273.15;
            case 'F', v = (v - 273.15) * 9/5 + 32;
        end
    end

    function step_case(delta)
        if isempty(R) || isempty(viewDD.ItemsData), return; end
        k = viewDD.Value + delta;
        if k < 1 || k > numel(R), return; end
        viewDD.Value = k;
        on_view_change();
    end

    function on_key(ev)
        switch ev.Key
            case 'leftarrow',  step_case(-1);
            case 'rightarrow', step_case(+1);
            case 'home',       snap('FIT');
        end
    end

    function snap(name)
        if isempty(H), return; end
        H.set_view(name);
    end

    function toggle(what, on)
        if isempty(H), return; end
        if strcmp(what, 'extrap'), on = on && any(R(viewDD.Value).extrap | R(viewDD.Value).far); end
        H.(what).Visible = onoff(on);
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

% --- setup files ---------------------------------------------------------------
    function st = collect_setup()
        st = struct('bdf', bdfE.Value, 'csvdir', csvE.Value, 'outdir', outE.Value, ...
                    'bdf_len', bdfUnitsDD.Value, 'bdf_temp', unitsDD.Value, ...
                    'csv_len', csvUnitsDD.Value, 'csv_temp', csvTempDD.Value, ...
                    'method', methodDD.Value, 'field', fieldDD.Value, 'header', headerCB.Value, ...
                    'sid_start', sidE.Value, 'warn_dist', warnE.Value, 'tempd', tempdCB.Value, ...
                    'style', styleDD.Value, 'colormap', cmapDD.Value);
    end

    function apply_setup(st)
        f = @(name, dflt) getfield_or(st, name, dflt);
        bdfE.Value       = f('bdf', '');
        csvE.Value       = f('csvdir', '');
        outE.Value       = f('outdir', 'temp_cards');
        bdfUnitsDD.Value = f('bdf_len', 'in');
        unitsDD.Value    = f('bdf_temp', 'K');
        csvUnitsDD.Value = f('csv_len', 'in');
        csvTempDD.Value  = f('csv_temp', 'K');
        methodDD.Value   = f('method', 'linear');
        fieldDD.Value    = f('field', 8);
        headerCB.Value   = f('header', true);
        sidE.Value       = f('sid_start', 1);
        warnE.Value      = f('warn_dist', 0);
        tempdCB.Value    = f('tempd', false);
        styleDD.Value    = f('style', 'points');
        cmapDD.Value     = f('colormap', 'jet');
        invalidate_grids();
        on_len_units();
        if ~isempty(csvE.Value) && exist(csvE.Value, 'dir') == 7, refresh_list(); else, csvLB.Items = {}; end
        update_state();
    end

    function on_save_setup(~, ~)
        [f, p] = uiputfile({'*.json', 'TEMP Mapper setup'}, 'Save setup as', 'temp_map_setup.json');
        figure(fig);
        if isequal(f, 0), return; end
        txt = jsonencode(collect_setup(), 'PrettyPrint', true);
        fid = fopen(fullfile(p, f), 'w'); fwrite(fid, txt, 'char'); fclose(fid);
        setupLbl.Text = ['Setup: ' f];
        status(sprintf('Saved setup to %s', fullfile(p, f)));
    end

    function on_load_setup(~, ~)
        [f, p] = uigetfile({'*.json', 'TEMP Mapper setup'}, 'Load setup');
        figure(fig);
        if isequal(f, 0), return; end
        try
            st = jsondecode(fileread(fullfile(p, f)));
        catch ME
            uialert(fig, ME.message, 'Could not read setup'); return
        end
        apply_setup(st);
        save_prefs();
        setupLbl.Text = ['Setup: ' f];
        status(sprintf('Loaded setup %s -- pick the new cloud folder if it changed, then Load & Map.', f));
    end

    function on_reset_setup(~, ~)
        apply_setup(struct());
        save_prefs();
        setupLbl.Text = 'Setup: defaults';
    end

% --- preferences -------------------------------------------------------------
    function load_prefs()
        bdfE.Value        = getpref(PREF, 'bdf',      '');
        csvE.Value        = getpref(PREF, 'csvdir',   '');
        outE.Value        = getpref(PREF, 'outdir',   'temp_cards');
        bdfUnitsDD.Value  = getpref(PREF, 'bdf_len',  'in');
        unitsDD.Value     = getpref(PREF, 'bdf_temp', 'K');
        csvUnitsDD.Value  = getpref(PREF, 'csv_len',  'in');
        csvTempDD.Value   = getpref(PREF, 'csv_temp', 'K');
        methodDD.Value    = getpref(PREF, 'method',   'linear');
        fieldDD.Value     = getpref(PREF, 'field',    8);
        headerCB.Value    = getpref(PREF, 'header',   true);
    end

    function save_prefs()
        setpref(PREF, {'bdf', 'csvdir', 'outdir', 'bdf_len', 'bdf_temp', 'csv_len', 'csv_temp', ...
                       'method', 'field', 'header'}, ...
                      {bdfE.Value, csvE.Value, outE.Value, bdfUnitsDD.Value, unitsDD.Value, ...
                       csvUnitsDD.Value, csvTempDD.Value, methodDD.Value, fieldDD.Value, headerCB.Value});
    end
end


% =========================================================================
function set_progress(dlg, frac, msg, t0)
%   Called by the engine between stages; a running triangulation cannot be
%   interrupted, so Cancel takes effect at the next stage boundary.
    drawnow                                   % register a Cancel click made during the last step
    if dlg.CancelRequested
        error('temp_map_gui:cancelled', 'Cancelled by user.');
    end
    dlg.Value = frac;
    dlg.Message = sprintf('%s   (%.0f s elapsed)', msg, toc(t0));
    drawnow
end


function c = put(c, row, col)
%PUT  Place a component in its parent grid and hand it back.
    c.Layout.Row = row; c.Layout.Column = col;
end


function v = getfield_or(st, name, dflt)
    if isstruct(st) && isfield(st, name) && ~isempty(st.(name)), v = st.(name); else, v = dflt; end
    if ischar(dflt) && isstring(v), v = char(v); end
    if islogical(dflt), v = logical(v); end
end


function s = onoff(tf)
    if tf, s = 'on'; else, s = 'off'; end
end


function s = fmtn(n)
%FMTN  Thousands separators: 1204331 -> '1,204,331'.
    s = sprintf('%d', round(n));
    s = fliplr(regexprep(fliplr(s), '(\d{3})(?=\d)', '$1,'));
end


function s = shortname(p)
    [~, b, e] = fileparts(p);
    s = [b e];
end
