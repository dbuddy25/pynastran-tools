function fig = temp_map_gui()
%TEMP_MAP_GUI  Point-and-click front end for TEMP_MAP_MATLAB.
%
%   TEMP_MAP_GUI opens a window laid out as the four steps of the job:
%       1  Parts    - one row per part: its BDF and its folder of CSV clouds;
%                     Read BDFs parses them once (grids + element faces cached)
%       2  Clouds   - the time steps (CSV names, taken from part 1); tick the ones to map
%       3  Map      - choose the method, Load & Map -> 3D preview + coverage check
%       4  Write    - emit one bulk-data file of TEMP cards per CSV
%   Paths, units and method are remembered between sessions (setpref).
%
%   All the work is done by TEMP_MAP_MATLAB (headless engine) and
%   TEMP_MAP_PLOT (the picture); this file only wires up the controls.
%   Anything you can do here you can do from the command line with those two.
%
%   See also TEMP_MAP_MATLAB, TEMP_MAP_PLOT.

R = [];          % results from the last Load & Map  (parts x steps)
RM = [];         % merged assembly view per step
G = [];          % cached grids from Read BDFs, one struct per part
stepTimes = [];  % uniform step times when no part has a cloud (from a T(t) table)
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
Lg.RowHeight = {26, 196, 230, 124, 210, 150, 128};
Lg.Padding = [0 0 0 0]; Lg.RowSpacing = 6;
Lg.Scrollable = 'on';                  % small screens: scroll the step column instead of squashing it

% --- setup file: everything below except the CSV list ----------------------------
sb = uigridlayout(Lg, [1 5]);
sb.ColumnWidth = {'1x', 84, 84, 98, 64}; sb.Padding = [0 0 0 0]; sb.ColumnSpacing = 6;
setupLbl = put(uilabel(sb, 'Text', 'Setup: (last session)', 'FontColor', [0.45 0.45 0.45]), 1, 1);
put(uibutton(sb, 'Text', 'Load setup', 'Tooltip', 'Restore BDF, units, method, output and SID settings from a .json file', ...
             'ButtonPushedFcn', @on_load_setup), 1, 2);
put(uibutton(sb, 'Text', 'Save setup', 'Tooltip', 'Save the current settings (not the CSV list) to a .json file', ...
             'ButtonPushedFcn', @on_save_setup), 1, 3);
put(uibutton(sb, 'Text', 'Export script', ...
             'Tooltip', sprintf(['Write a .m file that runs the engine headless with exactly these settings\n' ...
                                 '(parts, ticked steps, units, method, output options) -- no GUI needed.']), ...
             'ButtonPushedFcn', @on_export_script), 1, 4);
put(uibutton(sb, 'Text', 'Reset', 'Tooltip', 'Back to defaults', ...
             'ButtonPushedFcn', @on_reset_setup), 1, 5);

% --- 1  Parts ----------------------------------------------------------------
p1 = uipanel(Lg, 'Title', '1  Parts  (BDF + cloud folder, or a temperature)', 'FontWeight', 'bold');
g1 = uigridlayout(p1, [3 5]);
g1.ColumnWidth = {'1x', 72, 104, 72, 84}; g1.RowHeight = {'1x', 26, 24};
g1.Padding = [8 4 8 4]; g1.RowSpacing = 5;

partsT = put(uitable(g1, 'ColumnName', {'Part', 'BDF', 'Cloud folder  |  T  |  T(t).csv'}, ...
             'ColumnWidth', {90, '1x', '1x'}, 'ColumnEditable', [true false true], ...
             'RowName', [], 'Data', cell(0, 3), ...
             'Tooltip', sprintf(['One row per part = one BDF and its temperature source:\n' ...
                                 '   a folder holding THAT part''s cloud CSVs (Add part), or\n' ...
                                 '   a constant temperature in model units, or a 2-column CSV of time, T (Add isothermal).\n' ...
                                 'Isothermal parts follow the cloud parts'' time steps; with no cloud part at all you get\n' ...
                                 'one step (or the T(t) table''s steps).  Double-click a cell to rename a part or type a temperature.']), ...
             'CellEditCallback', @(~, ~) on_parts_edited()), 1, [1 5]);
modelLbl = put(uilabel(g1, 'Text', 'Add part: BDF + cloud folder.  Add isothermal: BDF + one temperature', ...
               'FontColor', [0.45 0.45 0.45]), 2, 1);
put(uibutton(g1, 'Text', 'Add part', 'Tooltip', 'Pick a BDF, then the folder holding its temperature clouds', ...
             'ButtonPushedFcn', @on_add_part), 2, 2);
put(uibutton(g1, 'Text', 'Add isothermal', ...
             'Tooltip', sprintf(['Pick a BDF, then type one temperature (model units) for the whole part,\n' ...
                                 'or leave it blank to pick a 2-column CSV (time, T) for a uniform T per step.']), ...
             'ButtonPushedFcn', @on_add_iso), 2, 3);
put(uibutton(g1, 'Text', 'Remove', 'Tooltip', 'Remove the highlighted part', ...
             'ButtonPushedFcn', @on_remove_part), 2, 4);
readB = put(uibutton(g1, 'Text', 'Read BDFs', 'Enable', 'off', ...
            'Tooltip', 'Parse GRIDs and element faces of every part once; cached until a part changes.', ...
            'ButtonPushedFcn', @on_read_bdf), 2, 5);

put(uilabel(g1, 'Text', 'Structural units'), 3, 1);
su = uigridlayout(g1, [1 5]); put(su, 3, [2 5]);
su.ColumnWidth = {44, 60, 38, 60, '1x'}; su.Padding = [0 0 0 0]; su.ColumnSpacing = 4;
uilabel(su, 'Text', 'length');
bdfUnitsDD = uidropdown(su, 'Items', {'in', 'mm', 'm'}, 'Value', 'in', ...
                        'Tooltip', 'Length units of the structural models (all parts)', ...
                        'ValueChangedFcn', @(~, ~) on_len_units());
uilabel(su, 'Text', 'temp');
unitsDD = uidropdown(su, 'Items', {'K', 'C', 'F'}, 'Value', 'K', ...
                     'Tooltip', 'Temperature units of the structural model = what is written on the TEMP cards', ...
                     'ValueChangedFcn', @(~, ~) on_units_change());
uilabel(su, 'Text', '= TEMP card units = isothermal T units', 'FontColor', [0.45 0.45 0.45]);

% --- 2  Temperature clouds ---------------------------------------------------
p2 = uipanel(Lg, 'Title', '2  Temp clouds', 'FontWeight', 'bold');
g2 = uigridlayout(p2, [4 4]);
g2.ColumnWidth = {60, '1x', 64, 92}; g2.RowHeight = {22, '1x', 24, 22};
g2.Padding = [8 4 8 4]; g2.RowSpacing = 5;

stepsLbl = put(uilabel(g2, 'Text', 'Time steps (CSV files) -- add a part first', ...
               'FontColor', [0.45 0.45 0.45]), 1, [1 3]);
put(uibutton(g2, 'Text', 'Cloud check', ...
             'Tooltip', sprintf(['Before meshing: is there a gradient worth modelling?  Reads the clouds alone (no BDF)\n' ...
                                 'and splits the temperature variation into THROUGH-THICKNESS vs IN-PLANE, with a\n' ...
                                 'verdict: isothermal / shells + TEMP / solids.  Uses the first cloud part''s folder and\n' ...
                                 'the ticked steps, or asks for a folder when no part has a cloud.']), ...
             'ButtonPushedFcn', @on_cloud_check), 1, 4);

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
g3 = uigridlayout(p3, [3 4]);
g3.ColumnWidth = {60, '1x', 90, 64}; g3.RowHeight = {24, 22, 30};
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

parCB = put(uicheckbox(g3, 'Text', 'Parallel steps (Parallel Computing Toolbox)', 'Value', false, ...
            'Tooltip', sprintf(['Map the steps after the first in a parfor. The first step of each part builds the\n' ...
                                'mapping; the rest reuse it, so the CSV read is what gets parallelised.']), ...
            'ValueChangedFcn', @(~, ~) save_prefs()), 2, [1 4]);
mapB = put(uibutton(g3, 'Text', 'Load & Map', 'FontWeight', 'bold', 'FontSize', 13, ...
           'BackgroundColor', ACCENT, 'FontColor', 'w', 'Enable', 'off', ...
           'Tooltip', 'Map every ticked CSV onto the grids and show the result. Nothing is written yet.', ...
           'ButtonPushedFcn', @on_map), 3, [1 4]);

% --- 4  Write ----------------------------------------------------------------
p4 = uipanel(Lg, 'Title', '4  Write TEMP cards', 'FontWeight', 'bold');
g4 = uigridlayout(p4, [6 4]);
g4.ColumnWidth = {60, '1x', 90, 64}; g4.RowHeight = {24, 24, 24, 24, 22, 28};
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

caseCB = put(uicheckbox(g4, 'Text', 'Case control', 'Value', true, ...
             'Tooltip', sprintf(['Also write temp_subcases.dat (SUBCASE id = TEMP SID, with\n' ...
                                 'TEMPERATURE(LOAD)) and temp_includes.bdf (one INCLUDE per TEMP file).']), ...
             'ValueChangedFcn', @(~, ~) save_prefs()), 3, 1);
subE = put(uieditfield(g4, 'text', 'Value', '{file}  t = {time} s', ...
           'Tooltip', 'SUBTITLE template. Tokens: {file} {time} {sid} {index}. Empty = no SUBTITLE.', ...
           'ValueChangedFcn', @(~, ~) save_prefs()), 3, [2 4]);
put(uilabel(g4, 'Text', 'Global lines'), 4, 1);
extraE = put(uieditfield(g4, 'text', 'Value', '', 'Placeholder', 'SPC = 1 ; DISP(PLOT) = ALL', ...
             'Tooltip', 'Case-control lines written once above the first SUBCASE (apply to all), separated by ;', ...
             'ValueChangedFcn', @(~, ~) save_prefs()), 4, 2);
put(uilabel(g4, 'Text', 'T ref', 'HorizontalAlignment', 'right', ...
            'Tooltip', 'Reference (stress-free) temperature in model units -> TEMPD + global TEMPERATURE(INITIAL). Blank = none.'), 4, 3);
trefE = put(uieditfield(g4, 'text', 'Value', '', 'Placeholder', 'none', ...
            'Tooltip', 'Reference (stress-free) temperature in model units -> TEMPD card + global TEMPERATURE(INITIAL) above the subcases. Blank = none.', ...
            'ValueChangedFcn', @(~, ~) save_prefs()), 4, 4);

reportCB = put(uicheckbox(g4, 'Text', 'HTML report', 'Value', true, ...
               'Tooltip', 'temp_map_report.html: settings, coverage warnings, min/max chart, per-step table, check plots.', ...
               'ValueChangedFcn', @(~, ~) save_prefs()), 5, [1 2]);
pngCB = put(uicheckbox(g4, 'Text', 'PNG per step (slow at 1M)', 'Value', false, ...
            'Tooltip', 'Save an ISO-view check plot per time step next to the TEMP files (and into the report).', ...
            'ValueChangedFcn', @(~, ~) save_prefs()), 5, [3 4]);

wr = uigridlayout(g4, [1 3]); put(wr, 6, [1 4]);
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
                      'Value', {'How to use:', ...
                                '1  Add part -- pick a BDF, then the folder holding that part''s CSVs (t000.csv, t001.csv ...).', ...
                                '   Add isothermal -- pick a BDF, then one temperature (model units) or a time,T CSV. No cloud needed.', ...
                                '   One row per part / assembly. Read BDFs parses them once.', ...
                                '2  Tick the time steps to map (names come from the first cloud part; other cloud folders pair by the number before .csv, e.g. wing_20.csv <-> fuse_20.csv).', ...
                                '3  Set units and method, Load & Map -> check the 3D view and the coverage banner.', ...
                                '4  Write -> one TEMP file per part per step, plus subcases + includes.'});

% =========================================================================
%                              RIGHT: VIEW
% =========================================================================
Rg = uigridlayout(root, [5 1]);
Rg.RowHeight = {34, 30, 26, 20, '1x'};     % banner / case + views / display / plot title / axes
Rg.Padding = [0 0 0 0]; Rg.RowSpacing = 6;

% --- row 1: coverage banner (top of the column, clear of the plot) ----------------
banner = uilabel(Rg, 'Text', 'Coverage check appears here after Load & Map', ...
                 'FontWeight', 'bold', 'HorizontalAlignment', 'center', 'WordWrap', 'on', ...
                 'BackgroundColor', [0.94 0.94 0.94], 'FontColor', [0.35 0.35 0.35]);

% --- row 2: case navigation + view presets -------------------------------------
bar = uigridlayout(Rg, [1 15]);
bar.ColumnWidth = [{44, 30, '1x', 30, 36, 130}, repmat({44}, 1, 7), {8, 50}];
bar.Padding = [0 0 0 0]; bar.ColumnSpacing = 4;
put(uilabel(bar, 'Text', 'Case'), 1, 1);
put(uibutton(bar, 'Text', '<', 'Tooltip', 'Previous case  (Left arrow)', ...
             'ButtonPushedFcn', @(~, ~) step_case(-1)), 1, 2);
viewDD = put(uidropdown(bar, 'Items', {'(nothing mapped yet)'}, ...
             'ValueChangedFcn', @(~, ~) on_view_change()), 1, 3);
put(uibutton(bar, 'Text', '>', 'Tooltip', 'Next case  (Right arrow)', ...
             'ButtonPushedFcn', @(~, ~) step_case(+1)), 1, 4);
put(uilabel(bar, 'Text', 'Part', 'HorizontalAlignment', 'right'), 1, 5);
partDD = put(uidropdown(bar, 'Items', {'All parts'}, 'ItemsData', 0, 'Value', 0, ...
             'Tooltip', 'Show the whole assembly or a single part', ...
             'ValueChangedFcn', @(~, ~) on_view_change()), 1, 6);
for k = 1:numel(VIEWS)
    put(uibutton(bar, 'Text', VIEWS{k}, 'Tooltip', ['Look at the model from ' VIEWS{k} ' and re-frame'], ...
                 'ButtonPushedFcn', @(src, ~) snap(src.Text)), 1, 6 + k);
end
put(uibutton(bar, 'Text', 'Fit', 'Tooltip', 'Re-frame the model without changing the view direction  (Home)', ...
             'ButtonPushedFcn', @(~, ~) snap('FIT')), 1, 15);

% --- row 3: display options ----------------------------------------------------
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

% --- row 4: plot title as a label: a 3D uiaxes title drifts above the axes box
%     with the camera framing and gets hidden under the toolbar ------------------------
plotTitle = uilabel(Rg, 'Text', 'Load & Map to see the grids coloured by temperature', ...
                    'FontWeight', 'bold', 'HorizontalAlignment', 'center', 'Interpreter', 'none');

% --- row 5: 3D axes -----------------------------------------------------------------
ax = uiaxes(Rg);

% --- bottom strip: min / max temperature vs time across all cases ------------
tax = uiaxes(root);
tax.Layout.Row = 2; tax.Layout.Column = [1 2];
title(tax, 'Min / max mapped temperature vs time');
xlabel(tax, 'time [s]'); grid(tax, 'on'); box(tax, 'on');
tax.ButtonDownFcn = @(~, ev) jump_to_time(ev.IntersectionPoint(1));

fig.KeyPressFcn = @(~, ev) on_key(ev);
load_prefs();
on_len_units();
refresh_list();
update_state();

% =========================================================================
%                                CALLBACKS
% =========================================================================
    function P = parts()
        P = partsT.Data;                       % {name, bdf, folder | temperature | T(t) csv}
        if isempty(P), P = cell(0, 3); end
    end

    function [kind, src] = source_kind(spec)
        % mirrors the engine's part_source: 'const' | 'table' | 'cloud'
        if isnumeric(spec) && isscalar(spec), kind = 'const'; src = double(spec); return; end
        s = strtrim(char(spec)); v = str2double(s);
        if ~isnan(v),                 kind = 'const'; src = v;
        elseif exist(s, 'dir') == 7,  kind = 'cloud'; src = s;
        elseif exist(s, 'file') == 2, kind = 'table'; src = s;
        else,                         kind = 'cloud'; src = s;
        end
    end

    function [ic, it] = source_rows()
        % row of the first cloud part and of the first T(t) table part (0 = none)
        P = parts(); ic = 0; it = 0;
        for i = 1:size(P, 1)
            k = source_kind(P{i, 3});
            if ic == 0 && strcmp(k, 'cloud'), ic = i; end
            if it == 0 && strcmp(k, 'table'), it = i; end
        end
    end

    function set_parts(P)
        partsT.Data = P;
        invalidate_grids();
        save_prefs();
        refresh_list();
        update_state();
    end

    function on_add_part(~, ~)
        P = parts();
        start = pwd;
        if ~isempty(P), start = fileparts(P{end, 2}); end
        [f, p] = uigetfile({'*.bdf;*.dat;*.nas;*.blk;*.inc', 'Nastran decks'; '*.*', 'All files'}, ...
                           'Pick the part''s BDF', start);
        figure(fig);
        if isequal(f, 0), return; end
        [~, name] = fileparts(f);
        d = uigetdir(p, sprintf('Folder holding the temperature clouds for "%s"', name));
        figure(fig);
        if isequal(d, 0), return; end
        set_parts([P; {name, fullfile(p, f), d}]);
    end

    function on_add_iso(~, ~)
        P = parts();
        start = pwd;
        if ~isempty(P), start = fileparts(P{end, 2}); end
        [f, p] = uigetfile({'*.bdf;*.dat;*.nas;*.blk;*.inc', 'Nastran decks'; '*.*', 'All files'}, ...
                           'Pick the isothermal part''s BDF', start);
        figure(fig);
        if isequal(f, 0), return; end
        [~, name] = fileparts(f);
        a = inputdlg(sprintf(['Temperature of "%s" in model units (deg %s), written to every grid.\n' ...
                              'Leave blank to pick a 2-column CSV (time, T) instead.'], name, unitsDD.Value), ...
                     'Isothermal part', 1, {'70'});
        figure(fig);
        if isempty(a), return; end
        v = str2double(strtrim(a{1}));
        if ~isnan(v)
            spec = v;
        else
            [tf, tp] = uigetfile({'*.csv', 'T(t) table: time, T'}, sprintf('T(t) table for "%s"', name), p);
            figure(fig);
            if isequal(tf, 0), return; end
            spec = fullfile(tp, tf);
        end
        set_parts([P; {name, fullfile(p, f), spec}]);
    end

    function on_remove_part(~, ~)
        P = parts();
        sel = partsT.Selection;
        if isempty(P) || isempty(sel), return; end
        P(unique(sel(:, 1)), :) = [];
        set_parts(P);
    end

    function on_parts_edited()
        invalidate_grids();
        save_prefs();
        refresh_list();
        update_state();
    end

    function invalidate_grids()
        G = [];
        P = parts();
        if isempty(P)
            modelLbl.Text = 'Add part: BDF + cloud folder.  Add isothermal: BDF + one temperature';
        else
            modelLbl.Text = sprintf('%d part(s), not read yet', size(P, 1));
        end
        modelLbl.FontColor = [0.45 0.45 0.45];
        p1.Title = '1  Parts  (BDF + cloud folder, or a temperature)';
    end

    function ok = on_read_bdf(~, ~)
        ok = false;
        P = parts();
        if isempty(P), return; end
        dlg = uiprogressdlg(fig, 'Title', 'Reading BDFs', 'Value', 0, 'Cancelable', 'on', ...
                            'Message', 'Starting ...');
        t0 = tic;
        try
            G = temp_map_matlab('PARTS', P, 'READ_ONLY', true, ...
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
        ng = sum(arrayfun(@(g) numel(g.ids), G));
        nf = sum(arrayfun(@(g) size(g.faces, 1), G));
        modelLbl.Text = sprintf('%s grids, %s faces', fmtn(ng), fmtn(nf));
        modelLbl.FontColor = [0.1 0.5 0.2];
        p1.Title = sprintf('1  Parts  --  %d part(s), %s grids', numel(G), fmtn(ng));
        lines = cell(numel(G), 1);
        for i = 1:numel(G)
            lines{i} = sprintf('%-14s %s grids  %s faces   %s', G(i).name, fmtn(numel(G(i).ids)), ...
                               fmtn(size(G(i).faces, 1)), shortname(G(i).bdf_file));
        end
        status([{sprintf('Read %d part(s) in %.1f s', numel(G), toc(t0))}; lines]);
        update_state();
        ok = true;
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
        P = parts();
        stepTimes = [];
        [ic, it] = source_rows();
        if isempty(P)
            csvLB.Items = {}; csvLB.Value = {};
            stepsLbl.Text = 'Time steps (CSV files) -- add a part first';
            update_state();
            return
        end
        if ic > 0
            [~, folder] = source_kind(P{ic, 3});
            if exist(folder, 'dir') ~= 7
                csvLB.Items = {}; csvLB.Value = {};
                stepsLbl.Text = sprintf('Cloud folder not found: %s', folder);
                update_state();
                return
            end
            stepsLbl.Text = sprintf('Time steps in %s -- other cloud folders pair by the number before .csv (wing_20 <-> fuse_20)', folder);
            d = dir(fullfile(folder, '*.csv'));
            names = {d.name};
            keys = regexprep(names, '(\d+)', '${sprintf(''%012d'', str2double($1))}');
            [~, order] = sort(lower(keys));
            names = names(order);
        elseif it > 0
            [~, tab] = source_kind(P{it, 3});
            try
                M = readmatrix(tab); M = M(:, 1:2); M(any(isnan(M), 2), :) = [];
                stepTimes = unique(M(:, 1))';
            catch
                stepTimes = [];
            end
            names = arrayfun(@(t) sprintf('t%g', t), stepTimes, 'UniformOutput', false);
            stepsLbl.Text = sprintf('No cloud part: one uniform step per time in %s', shortname(tab));
        else
            names = {'isothermal'};
            stepsLbl.Text = 'No cloud part: one isothermal step (every grid at the part''s temperature)';
        end
        csvLB.Items = names;
        csvLB.Value = names;                 % everything selected by default
        save_prefs();
        update_state();
    end

    function on_cloud_check(~, ~)
        P = parts();
        ic = source_rows();
        if ic > 0
            [~, folder] = source_kind(P{ic, 3});
            sel = cellstr(csvLB.Value);
            if isempty(sel) || numel(sel) == numel(csvLB.Items), src = folder; else, src = fullfile(folder, sel); end
        else
            src = uigetdir(pwd, 'Folder holding the temperature clouds to check');
            figure(fig);
            if isequal(src, 0), return; end
        end
        status('Cloud check running ... (details in the command window)'); drawnow
        try
            S = temp_cloud_check(src, 'CSV_LENGTH_UNITS', csvUnitsDD.Value, 'OUT_LENGTH_UNITS', bdfUnitsDD.Value, ...
                                 'CSV_TEMP_UNITS', csvTempDD.Value, 'OUT_UNITS', unitsDD.Value, ...
                                 'CSV_HAS_HEADER', headerCB.Value);
        catch ME
            uialert(fig, ME.message, 'Cloud check failed'); status('Cloud check failed.'); return
        end
        lu = bdfUnitsDD.Value;
        lines = {S.Properties.Description; ''; ...
                 sprintf('%-22s %7s %8s %9s %8s %9s %5s', 'step', 'dT', 'tt max', ['tt/' lu], 'ip', ['ip/' lu], 'R2')};
        for k = 1:height(S)
            lines{end+1} = sprintf('%-22s %7.2f %8.2f %9.3g %8.2f %9.3g %5.2f', S.File{k}, S.dT_total(k), ...
                                   S.dT_tt_max(k), S.grad_tt(k), S.dT_ip(k), S.grad_ip(k), S.R2(k)); %#ok<AGROW>
        end
        lines{end+1} = sprintf('(thickness %.4g %s; tt = through-thickness delta T, ip = in-plane delta T, deg %s)', ...
                               S.Thick(1), lu, unitsDD.Value);
        status(lines);
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
        P = parts();
        have_bdf = ~isempty(P) && all(cellfun(@(f) exist(f, 'file') == 2, P(:, 2)));
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
            uialert(fig, 'No time steps selected.', 'Nothing to map'); return
        end
        if isempty(G) && ~on_read_bdf(), return; end   % first run: read it now
        ic = source_rows();
        if ic > 0
            stepArgs = {'CSV_FILES', sel};
        elseif ~isempty(stepTimes)
            stepArgs = {'TIMES', stepTimes(ismember(csvLB.Items, sel))};
        else
            stepArgs = {};
        end
        dlg = uiprogressdlg(fig, 'Title', 'Mapping', 'Value', 0, 'Cancelable', 'on', ...
                            'Message', 'Starting ...');
        t0 = tic;
        prog = @(frac, msg) set_progress(dlg, frac, msg, t0);
        try
            [R, S, RM] = temp_map_matlab( ...
                'PARTS',             parts(), ...
                'GRIDS',             G, ...
                stepArgs{:}, ...
                'CSV_HAS_HEADER',    headerCB.Value, ...
                'BDF_LENGTH_UNITS',  bdfUnitsDD.Value, ...
                'CSV_LENGTH_UNITS',  csvUnitsDD.Value, ...
                'CSV_TEMP_UNITS',    csvTempDD.Value, ...
                'SID_START',         sidE.Value, ...
                'METHOD',            methodDD.Value, ...
                'EXTRAP_WARN_DIST',  warn_dist(), ...
                'OUT_UNITS',         unitsDD.Value, ...
                'PARALLEL',          parCB.Value, ...
                'PROGRESS',          prog, ...
                'WRITE',             false);
        catch ME
            close(dlg);
            R = []; RM = [];
            update_state();
            if strcmp(ME.identifier, 'temp_map_gui:cancelled')
                status('Cancelled.');
            else
                uialert(fig, ME.message, 'Mapping failed');
            end
            return
        end
        close(dlg);
        viewDD.Items = cellfun(@shortname, {RM.csv_file}, 'UniformOutput', false);
        viewDD.ItemsData = 1:numel(RM);
        viewDD.Value = 1;
        if size(R, 1) > 1
            partDD.Items = [{'All parts'}, {R(:, 1).part}];
            partDD.ItemsData = 0:size(R, 1);
        else
            partDD.Items = {'All parts'}; partDD.ItemsData = 0;
        end
        partDD.Value = 0;
        mname = methodDD.Items{strcmp(methodDD.ItemsData, methodDD.Value)};
        p3.Title = sprintf('3  Map  --  %d steps x %d part(s), %s, %s', size(R, 2), size(R, 1), mname, datestr(now, 'HH:MM'));
        fill_summary(S);
        update_state();
        on_view_change();
        plot_history();
        bad = find(arrayfun(@(r) ~isempty(r.warnings), RM));
        if ~isempty(bad)
            msg = {};
            for k = bad(:)'
                msg{end+1} = sprintf('%s:', shortname(RM(k).csv_file));        %#ok<AGROW>
                msg = [msg, strcat({'    - '}, RM(k).warnings(:)')];          %#ok<AGROW>
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
            want = cellstr(csvLB.Value);
            have = cellfun(@shortname, {RM.csv_file}, 'UniformOutput', false);
            sub = R(:, ismember(have, want));
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
                'WRITE_CASE',   caseCB.Value, ...
                'SUBTITLE',     subE.Value, ...
                'CASE_EXTRA',   split_lines(extraE.Value), ...
                'TREF',         tref_value(), ...
                'REPORT',       reportCB.Value && all_of_them, ...   % a partial write must not overwrite the run report
                'SAVE_PNG',     pngCB.Value, ...
                'WRITE',        true);
        catch ME
            uialert(fig, ME.message, 'Write failed');
            return
        end
        p4.Title = sprintf('4  Write TEMP cards  --  %d file(s) written %s', height(S), datestr(now, 'HH:MM'));
        extra = {};
        if caseCB.Value
            extra = {fullfile(outE.Value, 'temp_subcases.dat'); fullfile(outE.Value, 'temp_includes.bdf')};
        end
        if reportCB.Value && all_of_them
            extra = [extra; {fullfile(outE.Value, 'temp_map_report.html')}];
        end
        status([{sprintf('Wrote %d file(s) to %s:', height(S) + numel(extra), outE.Value)}; extra; S.OutFile(:)]);
    end

    function r = current()
        % the result being viewed: merged assembly or one part, at the chosen step
        k = viewDD.Value;
        if partDD.Value == 0, r = RM(k); else, r = R(partDD.Value, k); end
    end

    function show_coverage()
        r = current();
        lines = [{sprintf('Coverage: %s   [%s]', shortname(r.csv_file), partDD.Items{partDD.Value + 1})}; r.coverage(:)];
        if ~isempty(r.warnings)
            lines = [lines; {''; '!!!!! WARNING !!!!!'}; strcat({'!! '}, r.warnings(:))];
            banner.Text = ['WARNING  ' strjoin(r.warnings, '   |   ')];
            banner.BackgroundColor = [0.98 0.85 0.85];
            banner.FontColor = [0.65 0.05 0.05];
        else
            banner.Text = sprintf('Coverage OK  --  %.1f%% of grids outside the cloud hull, max distance to a cloud point %.3g', ...
                100 * nnz(r.extrap) / numel(r.grid_ids), max(r.nn_dist));
            banner.BackgroundColor = [0.86 0.95 0.86];
            banner.FontColor = [0.05 0.40 0.10];
        end
        status(lines);
    end

    function fill_summary(~)
        % one row per time step, whole assembly (per-part numbers live in the coverage box)
        n = numel(RM);
        D = cell(n, 8);
        for k = 1:n
            r = RM(k);
            D(k, :) = {shortname(r.csv_file), r.time, r.sid, numel(r.grid_ids), nnz(r.extrap), nnz(r.far), ...
                       round(conv_out(min(r.grid_T), unitsDD.Value), 2), round(conv_out(max(r.grid_T), unitsDD.Value), 2)};
        end
        sumT.Data = D;
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
        if k < 1 || k > numel(RM) || k == viewDD.Value, return; end
        viewDD.Value = k;
        on_view_change();
    end

    function on_units_change()
        save_prefs();
        replot();
        if ~isempty(R), plot_history(); fill_summary([]); end
    end

    function replot()
        if isempty(R) || isempty(viewDD.ItemsData), return; end
        H = temp_map_plot(current(), 'Parent', ax, 'Style', styleDD.Value, ...
                          'Units',      unitsDD.Value, ...
                          'ShowCloud',  cloudCB.Value, ...
                          'ShowSurface', surfCB.Value, ...
                          'ShowMinMax', mmCB.Value, ...
                          'ShowExtrap', extrapCB.Value, ...
                          'Colormap',   cmapDD.Value);
        plotTitle.Text = char(ax.Title.String);
        title(ax, '');
    end

    function on_view_change()
        replot();
        show_coverage();
        mark_history();
        highlight_row();
    end

    function plot_history()
        if isempty(RM), return; end
        t = [RM.time];
        [t, order] = sort(t);
        u = unitsDD.Value;
        tmin = arrayfun(@(r) conv_out(min(r.grid_T), u), RM(order));
        tmax = arrayfun(@(r) conv_out(max(r.grid_T), u), RM(order));
        cla(tax);
        hold(tax, 'on');
        plot(tax, t, tmax, '-o', 'Color', [0.85 0.1 0.1], 'LineWidth', 1.6, 'MarkerFaceColor', [0.85 0.1 0.1], ...
             'MarkerSize', 4, 'DisplayName', 'max', 'HitTest', 'off');
        plot(tax, t, tmin, '-o', 'Color', [0.1 0.3 0.85], 'LineWidth', 1.6, 'MarkerFaceColor', [0.1 0.3 0.85], ...
             'MarkerSize', 4, 'DisplayName', 'min', 'HitTest', 'off');
        hold(tax, 'off');
        ylabel(tax, sprintf('T [deg %s]', u));
        ytickformat(tax, '%.1f');
        span = max(tmax) - min(tmin); if span <= 0, span = 1; end
        ylim(tax, [min(tmin) - 0.05 * span, max(tmax) + 0.05 * span]);
        title(tax, sprintf('Min / max mapped temperature vs time, whole assembly  (%d steps; click to jump)', numel(RM)));
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
        if isempty(RM) || isempty(viewDD.ItemsData), return; end
        tk = RM(viewDD.Value).time;
        hNow = xline(tax, tk, '-', 'Color', [0.2 0.2 0.2], 'LineWidth', 1.5, ...
                     'HandleVisibility', 'off', 'HitTest', 'off');
    end

    function jump_to_time(tclick)
        if isempty(RM), return; end
        [~, k] = min(abs([RM.time] - tclick));
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
        if k < 1 || k > numel(RM), return; end
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
        if strcmp(what, 'extrap'), r = current(); on = on && any(r.extrap | r.far); end
        H.(what).Visible = onoff(on);
    end

    function colormap_now(name)
        if isempty(H), return; end
        colormap(ax, name);
    end

    function v = tref_value()
        v = str2double(strtrim(trefE.Value));
        if isnan(v), v = []; end
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
        P = parts();
        st = struct('parts', struct('name', P(:, 1), 'bdf', P(:, 2), 'csv_dir', P(:, 3)), ...
                    'outdir', outE.Value, ...
                    'bdf_len', bdfUnitsDD.Value, 'bdf_temp', unitsDD.Value, ...
                    'csv_len', csvUnitsDD.Value, 'csv_temp', csvTempDD.Value, ...
                    'method', methodDD.Value, 'field', fieldDD.Value, 'header', headerCB.Value, ...
                    'sid_start', sidE.Value, 'warn_dist', warnE.Value, 'tempd', tempdCB.Value, ...
                    'case', caseCB.Value, 'subtitle', subE.Value, 'extra', extraE.Value, 'tref', trefE.Value, ...
                    'parallel', parCB.Value, 'report', reportCB.Value, 'png', pngCB.Value, ...
                    'style', styleDD.Value, 'colormap', cmapDD.Value);
    end

    function apply_setup(st)
        f = @(name, dflt) getfield_or(st, name, dflt);
        P = parts_from_setup(st);
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
        caseCB.Value     = f('case', true);
        subE.Value       = f('subtitle', '{file}  t = {time} s');
        extraE.Value     = f('extra', '');
        trefE.Value      = f('tref', '');
        parCB.Value      = f('parallel', false);
        reportCB.Value   = f('report', true);
        pngCB.Value      = f('png', false);
        styleDD.Value    = f('style', 'points');
        cmapDD.Value     = f('colormap', 'jet');
        partsT.Data      = P;
        invalidate_grids();
        on_len_units();
        refresh_list();
        update_state();
    end

    function P = parts_from_setup(st)
        P = cell(0, 3);
        if isstruct(st) && isfield(st, 'parts') && ~isempty(st.parts)
            ps = st.parts;
            for i = 1:numel(ps)
                src = ps(i).csv_dir;
                if ~isnumeric(src), src = char(src); end
                P(end+1, :) = {char(ps(i).name), char(ps(i).bdf), src}; %#ok<AGROW>
            end
        elseif isstruct(st) && isfield(st, 'bdf') && ~isempty(st.bdf)     % old single-part setup
            [~, nm] = fileparts(char(st.bdf));
            P = {nm, char(st.bdf), char(getfield_or(st, 'csvdir', ''))};
        end
    end

    function on_save_setup(~, ~)
        [f, p] = uiputfile({'*.json', 'TEMP Mapper setup'}, 'Save setup as', 'temp_map_setup.json');
        figure(fig);
        if isequal(f, 0), return; end
        try
            txt = jsonencode(collect_setup(), 'PrettyPrint', true);   % R2021a+
        catch
            txt = jsonencode(collect_setup());
        end
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
        status(sprintf('Loaded setup %s -- %d part(s). Read BDFs, then Load & Map.', f, size(parts(), 1)));
    end

    function on_export_script(~, ~)
        % a runnable .m file: the same name/value call Load & Map + Write would make
        P = parts();
        if isempty(P)
            uialert(fig, 'Add at least one part first.', 'Nothing to export'); return
        end
        [f, p] = uiputfile({'*.m', 'MATLAB script'}, 'Export headless run as', 'run_temp_map.m');
        figure(fig);
        if isequal(f, 0), return; end
        sel = cellstr(csvLB.Value);
        ic = source_rows();
        stepPair = {};
        if ic > 0 && ~isempty(sel) && numel(sel) < numel(csvLB.Items)
            stepPair = {'CSV_FILES', sel};
        elseif ic == 0 && ~isempty(stepTimes) && numel(sel) < numel(csvLB.Items)
            stepPair = {'TIMES', stepTimes(ismember(csvLB.Items, sel))};
        end
        args = [stepPair, { ...
            'CSV_HAS_HEADER',   headerCB.Value, ...
            'BDF_LENGTH_UNITS', bdfUnitsDD.Value, ...
            'CSV_LENGTH_UNITS', csvUnitsDD.Value, ...
            'CSV_TEMP_UNITS',   csvTempDD.Value, ...
            'OUT_UNITS',        unitsDD.Value, ...
            'METHOD',           methodDD.Value, ...
            'EXTRAP_WARN_DIST', warn_dist(), ...
            'SID_START',        sidE.Value, ...
            'PARALLEL',         parCB.Value, ...
            'OUT_DIR',          outE.Value, ...
            'FIELD_SIZE',       fieldDD.Value, ...
            'WRITE_TEMPD',      tempdCB.Value, ...
            'WRITE_CASE',       caseCB.Value, ...
            'SUBTITLE',         subE.Value, ...
            'CASE_EXTRA',       split_lines(extraE.Value), ...
            'TREF',             tref_value(), ...
            'REPORT',           reportCB.Value, ...
            'SAVE_PNG',         pngCB.Value, ...
            'WRITE',            true}];
        L = {};
        L{end+1} = sprintf('%% %s -- headless TEMP mapping run exported from temp_map_gui  (%s)', f, datestr(now, 'yyyy-mm-dd HH:MM'));
        L{end+1} = '%   Run it as a script; edit any name/value pair freely (see "help temp_map_matlab").';
        L{end+1} = '%   Column 3 of PARTS: cloud folder | constant temperature (model units) | time,T csv.';
        L{end+1} = sprintf('addpath(%s);   %% where temp_map_matlab.m lives', mlit(fileparts(mfilename('fullpath'))));
        L{end+1} = '';
        L{end+1} = 'PARTS = { ...';
        for i = 1:size(P, 1)
            L{end+1} = sprintf('    %s, %s, %s; ...', mlit(P{i, 1}), mlit(P{i, 2}), mlit(P{i, 3})); %#ok<AGROW>
        end
        L{end+1} = '    };';
        L{end+1} = '';
        L{end+1} = '[R, S, RM] = temp_map_matlab( ...';
        L{end+1} = '    ''PARTS'',            PARTS, ...';
        for q = 1:2:numel(args)
            L{end+1} = sprintf('    %-19s %s, ...', ['''' args{q} ''','], mlit(args{q + 1})); %#ok<AGROW>
        end
        L{end} = regexprep(L{end}, ', \.\.\.$', ');');
        L{end+1} = 'disp(S);';
        L{end+1} = '% temp_map_plot(RM(1));      % 3D check of the first step';
        fid = fopen(fullfile(p, f), 'w');
        if fid < 0, uialert(fig, sprintf('Cannot write %s', fullfile(p, f)), 'Export failed'); return; end
        fprintf(fid, '%s\n', L{:});
        fclose(fid);
        status([{sprintf('Exported headless run to %s', fullfile(p, f))}; L(:)]);
    end

    function on_reset_setup(~, ~)
        apply_setup(struct());
        save_prefs();
        setupLbl.Text = 'Setup: defaults';
    end

% --- preferences -------------------------------------------------------------
    function load_prefs()
        try
            partsT.Data   = parts_from_setup(jsondecode(getpref(PREF, 'parts_json', '{}')));
        catch
            partsT.Data   = cell(0, 3);
        end
        outE.Value        = getpref(PREF, 'outdir',   'temp_cards');
        bdfUnitsDD.Value  = getpref(PREF, 'bdf_len',  'in');
        unitsDD.Value     = getpref(PREF, 'bdf_temp', 'K');
        csvUnitsDD.Value  = getpref(PREF, 'csv_len',  'in');
        csvTempDD.Value   = getpref(PREF, 'csv_temp', 'K');
        methodDD.Value    = getpref(PREF, 'method',   'linear');
        fieldDD.Value     = getpref(PREF, 'field',    8);
        headerCB.Value    = getpref(PREF, 'header',   true);
        caseCB.Value      = getpref(PREF, 'case',     true);
        subE.Value        = getpref(PREF, 'subtitle', '{file}  t = {time} s');
        extraE.Value      = getpref(PREF, 'extra',    '');
        trefE.Value       = getpref(PREF, 'tref',     '');
        parCB.Value       = getpref(PREF, 'parallel', false);
        reportCB.Value    = getpref(PREF, 'report',   true);
        pngCB.Value       = getpref(PREF, 'png',      false);
    end

    function save_prefs()
        P = parts();
        pj = jsonencode(struct('parts', struct('name', P(:, 1), 'bdf', P(:, 2), 'csv_dir', P(:, 3))));
        setpref(PREF, {'parts_json', 'outdir', 'bdf_len', 'bdf_temp', 'csv_len', 'csv_temp', ...
                       'method', 'field', 'header', 'case', 'subtitle', 'extra', 'tref', ...
                       'parallel', 'report', 'png'}, ...
                      {pj, outE.Value, bdfUnitsDD.Value, unitsDD.Value, ...
                       csvUnitsDD.Value, csvTempDD.Value, methodDD.Value, fieldDD.Value, headerCB.Value, ...
                       caseCB.Value, subE.Value, extraE.Value, trefE.Value, ...
                       parCB.Value, reportCB.Value, pngCB.Value});
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


function s = mlit(v)
%MLIT  MATLAB source literal for a value the GUI hands the engine.
    if ischar(v) || isstring(v)
        s = ['''' strrep(char(v), '''', '''''') ''''];
    elseif islogical(v) && isscalar(v)
        if v, s = 'true'; else, s = 'false'; end
    elseif isnumeric(v) && isempty(v)
        s = '[]';
    elseif isnumeric(v) && isscalar(v)
        s = sprintf('%.15g', v);
    elseif isnumeric(v)
        s = ['[' strjoin(arrayfun(@(x) sprintf('%.15g', x), v(:)', 'UniformOutput', false), ' ') ']'];
    elseif iscell(v)
        if isempty(v), s = '{}'; return; end
        s = ['{' strjoin(cellfun(@mlit, v(:)', 'UniformOutput', false), ', ') '}'];
    else
        error('temp_map_gui:literal', 'Cannot write a %s as a script literal.', class(v));
    end
end

function c = split_lines(txt)
%SPLIT_LINES  'SPC = 1 ; DISP = ALL' -> {'SPC = 1', 'DISP = ALL'}
    c = strtrim(strsplit(char(txt), ';'));
    c = c(~cellfun('isempty', c));
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
