function run_temp_map_test()
%RUN_TEMP_MAP_TEST  Self-check for temp_map_matlab against an analytic field.
%   The fixture clouds carry T = 300 + 10x + 5y - 2z (t000) and 350 + ... (t100),
%   so linear interpolation inside the hull must reproduce them exactly.
%   Grid 999 sits outside the hull and must take the nearest cloud point's T.
here = fileparts(mfilename('fullpath'));
addpath(fileparts(here));
out = fullfile(tempdir, 'temp_map_test_out');
if exist(out, 'dir'), rmdir(out, 's'); end

bdf = fullfile(here, 'test_temp_map.bdf');
[R, S] = temp_map_matlab('BDF_FILE', bdf, 'CSV_DIR', here, 'OUT_DIR', out, ...
                         'SID_START', 10, 'EXTRAP_WARN_DIST', 5);
disp(S);

% --- grids parsed from every card format + the INCLUDE -------------------
want_ids = [1 2 3 4 5 6 7 8 9 999 11 12]';
want_xyz = [0 0 0; 1 0 0; 2 1 .5; 3 2 1; 1.5 .5 1; 2.5 1.5 .25; ...
            .5 1 .75; 1.25 1.75 .5; 2.5 .1 .5; 8 8 8; 1.5 .25 .75; 2.75 1.25 .9];
check(isequal(R(1).grid_ids, want_ids), 'grid ids / order');
check(max(abs(R(1).grid_xyz - want_xyz), [], 'all') < 1e-12, 'grid coordinates');

% --- mapping --------------------------------------------------------------
base = [300 350];
for k = 1:2
    x = R(k).grid_xyz;
    Texp = base(k) + 10*x(:,1) + 5*x(:,2) - 2*x(:,3);
    inside = R(k).grid_ids ~= 999;
    err = max(abs(R(k).grid_T(inside) - Texp(inside)));
    check(err < 1e-6, sprintf('%s: linear field reproduced (max err %.2e)', S.File{k}, err));
    check(isequal(R(k).extrap, ~inside), sprintf('%s: only grid 999 flagged outside hull', S.File{k}));
    [~, nn] = min(vecnorm(R(k).cloud_xyz - [8 8 8], 2, 2));
    check(abs(R(k).grid_T(~inside) - R(k).cloud_T(nn)) < 1e-9, ...
          sprintf('%s: outside-hull grid took nearest cloud T', S.File{k}));
    check(R(k).sid == 9 + k, 'SID numbering from SID_START');
end
check(abs(R(2).time - 100) < 1e-12, 'time read from CSV');
check(contains(R(2).method, '(reused)') && ~contains(R(1).method, '(reused)'), ...
      'second CSV with identical points reused the first mapping');
check(isequal(R(1).far, R(1).grid_ids == 999), 'EXTRAP_WARN_DIST flags only grid 999 as far');

% --- cached grids -----------------------------------------------------------
G = temp_map_matlab('BDF_FILE', bdf, 'READ_ONLY', true);
Rg = temp_map_matlab('GRIDS', G, 'CSV_FILES', {fullfile(here, 't000.csv')}, 'WRITE', false);
check(isequal(G.ids, want_ids) && isequal(Rg.grid_T, R(1).grid_T), 'READ_ONLY + GRIDS reuse gives the same result');

check(~isempty(R(1).warnings) && any(contains(R(1).warnings, 'overhangs')), ...
      'coverage warning raised (grid 999 at 8,8,8 overhangs the cloud)');

% --- unit conversion --------------------------------------------------------
Ru = temp_map_matlab('BDF_FILE', bdf, 'CSV_FILES', {fullfile(here, 't000.csv')}, 'WRITE', false, ...
                     'BDF_LENGTH_UNITS', 'mm', 'CSV_LENGTH_UNITS', 'in', 'CSV_TEMP_UNITS', 'C');
check(max(abs(Ru.cloud_xyz - R(1).cloud_xyz * 25.4), [], 'all') < 1e-9, 'cloud x/y/z converted in -> mm');
check(max(abs(Ru.cloud_T - (R(1).cloud_T + 273.15))) < 1e-9, 'cloud T converted C -> K');

outF = fullfile(tempdir, 'temp_map_test_F');
if exist(outF, 'dir'), rmdir(outF, 's'); end
Rf = temp_map_matlab('RESULTS', R(1), 'OUT_DIR', outF, 'OUT_UNITS', 'F');   % separate folder: the
txt = fileread(Rf(1).out_file);                                                % Kelvin files are re-read below
lines = regexp(txt, '\r?\n', 'split');
L = lines{find(startsWith(lines, 'TEMP    '), 1)};
check(abs(str2double(L(25:32)) - ((R(1).grid_T(1) - 273.15) * 9/5 + 32)) < 5e-3, 'Fahrenheit output');

% --- nearest / idw methods against brute force ------------------------------
Rn = temp_map_matlab('BDF_FILE', bdf, 'CSV_FILES', {fullfile(here, 't000.csv')}, ...
                     'METHOD', 'nearest', 'WRITE', false, 'SHELL_AVERAGE', false);
D = pdist2_plain(Rn.grid_xyz, Rn.cloud_xyz);
[dmin, nn] = min(D, [], 2);
check(max(abs(Rn.grid_T - Rn.cloud_T(nn))) < 1e-9, 'nearest: matches brute-force nearest point');
check(max(abs(Rn.nn_dist - dmin)) < 1e-9, 'nearest: nn_dist matches brute force');
Rs = temp_map_matlab('BDF_FILE', bdf, 'CSV_FILES', {fullfile(here, 't000.csv')}, ...
                     'METHOD', 'scattered', 'WRITE', false);
check(max(abs(Rs.grid_T - R(1).grid_T)) < 1e-6, 'scattered: matches linear on the fixture');
Ri = temp_map_matlab('BDF_FILE', bdf, 'CSV_FILES', {fullfile(here, 't000.csv')}, ...
                     'METHOD', 'idw', 'WRITE', false, 'SHELL_AVERAGE', false);
check(all(Ri.grid_T >= min(Ri.cloud_T) - 1e-9 & Ri.grid_T <= max(Ri.cloud_T) + 1e-9), ...
      'idw: within cloud temperature range');

% --- written cards parse back ---------------------------------------------
txt = fileread(R(1).out_file);
lines = regexp(txt, '\r?\n', 'split');
cards = lines(startsWith(lines, 'TEMP    '));
check(numel(cards) == 4, '12 grids -> 4 small-field TEMP cards');
got = zeros(0, 2);
for c = cards
    L = c{1};
    check(str2double(L(9:16)) == 10, 'SID on every card');
    for a = 17:16:65
        if a + 15 > numel(L), break; end
        g = str2double(L(a:a+7)); t = str2double(L(a+8:a+15));
        if ~isnan(g), got(end+1, :) = [g t]; end %#ok<AGROW>
    end
end
check(isequal(got(:,1), want_ids), 'all grids written in order');
check(max(abs(got(:,2) - R(1).grid_T)) < 5e-4, 'written temps match (8-char rounding)');

% --- case control -----------------------------------------------------------
txt = fileread(fullfile(out, 'temp_subcases.dat'));
check(contains(txt, 'SUBCASE 10') && contains(txt, 'SUBCASE 11') && ...
      contains(txt, 'TEMPERATURE(LOAD) = 10') && contains(txt, 'SUBTITLE = t100  t = 100 s'), ...
      'temp_subcases.dat: SUBCASE ids match SIDs, subtitle tokens filled');
txt = fileread(fullfile(out, 'temp_includes.bdf'));
check(contains(txt, 'INCLUDE ''t000_temp.bdf''') && contains(txt, 'INCLUDE ''t100_temp.bdf'''), ...
      'temp_includes.bdf lists both TEMP files');
temp_map_matlab('RESULTS', R, 'OUT_DIR', out, 'TREF', 20, 'OUT_UNITS', 'C', 'CASE_EXTRA', {'SPC = 1'});
txt = fileread(fullfile(out, 'temp_subcases.dat'));
i_glob = regexp(txt, 'TEMPERATURE\(INITIAL\) = 99999', 'once'); i_spc = regexp(txt, '^SPC = 1', 'once', 'lineanchors');
i_sub = regexp(txt, 'SUBCASE 10', 'once');
check(~isempty(i_glob) && ~isempty(i_spc) && i_glob < i_sub && i_spc < i_sub, ...
      'TREF + global lines sit above the first SUBCASE');
txt = fileread(fullfile(out, 'temp_includes.bdf'));
check(~isempty(regexp(txt, 'TEMPD\s+99999\s+20\.', 'once')), 'TEMPD reference card written');

% --- Celsius + large field + TEMPD ----------------------------------------
temp_map_matlab('RESULTS', R(1), 'OUT_DIR', out, 'OUT_UNITS', 'C', ...
                'FIELD_SIZE', 16, 'WRITE_TEMPD', true);
txt = fileread(R(1).out_file);
check(contains(txt, 'TEMPD*') && contains(txt, 'TEMP*   '), 'large-field TEMP* + TEMPD* written');
lines = regexp(txt, '\r?\n', 'split');
L = lines{find(startsWith(lines, 'TEMP*   '), 1)};
check(str2double(L(9:24)) == 10 && str2double(L(25:40)) == R(1).grid_ids(1) && ...
      abs(str2double(L(41:56)) - (R(1).grid_T(1) - 273.15)) < 1e-6 && ...
      str2double(L(57:72)) == R(1).grid_ids(2) && L(73) == '*', ...
      'large-field SID / G1 / T1(Celsius) / G2 / continuation columns');

check(~isempty(R(1).surface) && size(R(1).surface.F, 2) == 3, 'cloud surface (alpha shape) built');

% --- element faces for the contour view --------------------------------------
check(size(R(1).faces, 1) == 12 && size(R(1).faces, 2) == 4, ...
      'quad + tri + hexa(6) + tetra(4) -> 12 free faces');
check(nnz(isnan(R(1).faces(:, 4))) == 1 + 4, 'tri faces NaN-padded (1 tri + 4 tet faces)');
check(all(R(1).faces(~isnan(R(1).faces)) >= 1 & R(1).faces(~isnan(R(1).faces)) <= 12), 'faces index grid rows');

% --- shells: through-thickness average -------------------------------------------------
check(contains(R(1).method, 'shell grids') && nnz(~isnan(G.shell_t)) == 7 && all(abs(G.shell_t(~isnan(G.shell_t)) - 0.2) < 1e-12), ...
      'fixture: 7 grids on the quad + tri get t = 0.2 from the PSHELL in the INCLUDE; linear field unchanged by averaging');
shdir = fullfile(tempdir, 'temp_map_test_shell');
if exist(shdir, 'dir'), rmdir(shdir, 's'); end
mkdir(shdir);
[gx, gy, gz] = ndgrid(-2:2, -2:2, -1:0.5:1);
M = [zeros(numel(gx), 1) gx(:) gy(:) gz(:) 300 + 100 * gz(:).^2];      % T = 300 + 100 z^2: mid-plane 300
shcsv = fullfile(shdir, 'quad_0.csv');
fid = fopen(shcsv, 'w'); fprintf(fid, 'time,x,y,z,T\n'); fprintf(fid, '%g,%g,%g,%g,%g\n', M'); fclose(fid);
shbdf = fullfile(shdir, 'quad.bdf');
fid = fopen(shbdf, 'w');
fprintf(fid, 'BEGIN BULK\nGRID,1,,-1.,-1.,0.\nGRID,2,,1.,-1.,0.\nGRID,3,,1.,1.,0.\nGRID,4,,-1.,1.,0.\n');
fprintf(fid, 'CQUAD4,1,1,1,2,3,4\nPSHELL,1,1,2.0\nENDDATA\n'); fclose(fid);
Rsh = temp_map_matlab('BDF_FILE', shbdf, 'CSV_FILES', {shcsv}, 'WRITE', false, 'SHELL_LAYERS', 5);
check(all(abs(Rsh.grid_T - 350) < 1e-6), 'shell grids: mean of 5 levels across t = 2 along the normal (z^2 field -> 350, mid-plane 300)');
Roff = temp_map_matlab('BDF_FILE', shbdf, 'CSV_FILES', {shcsv}, 'WRITE', false, 'SHELL_AVERAGE', false);
check(all(abs(Roff.grid_T - 300) < 1e-6) && ~contains(Roff.method, 'shell'), 'SHELL_AVERAGE off: mid-plane value');
fid = fopen(shbdf, 'w');
fprintf(fid, 'BEGIN BULK\nGRID,1,,-1.,-1.,0.\nGRID,2,,1.,-1.,0.\nGRID,3,,1.,1.,0.\nGRID,4,,-1.,1.,0.\n');
fprintf(fid, 'CQUAD4,1,1,1,2,3,4\nPCOMP,1,,,,,,,,\n,1,0.5,0.,YES,1,0.5,0.,YES\nENDDATA\n'); fclose(fid);
Rpc = temp_map_matlab('BDF_FILE', shbdf, 'CSV_FILES', {shcsv}, 'WRITE', false, 'SHELL_LAYERS', 5);
check(all(abs(Rpc.grid_T - 315) < 1e-6), 'PCOMP total thickness 1.0: levels at 0, +-0.25, +-0.5 -> 315');
fid = fopen(shbdf, 'w');
fprintf(fid, 'BEGIN BULK\nGRID,1,,-1.,-1.,0.\nGRID,2,,1.,-1.,0.\nGRID,3,,1.,1.,0.\nGRID,4,,-1.,1.,0.\n');
fprintf(fid, 'CQUAD4,1,7,1,2,3,4\nENDDATA\n'); fclose(fid);
Rnp = temp_map_matlab('BDF_FILE', shbdf, 'CSV_FILES', {shcsv}, 'WRITE', false);
check(all(abs(Rnp.grid_T - 300) < 1e-6), 'shell without a property card: falls back to the mid-plane value');

% --- multi-part assembly ----------------------------------------------------
outM = fullfile(tempdir, 'temp_map_test_multi');
if exist(outM, 'dir'), rmdir(outM, 's'); end
PARTS = {'A', bdf, fullfile(here, 'partA'); 'B', fullfile(here, 'test_temp_map_B.bdf'), fullfile(here, 'partB')};
[R2, S2, RM] = temp_map_matlab('PARTS', PARTS, 'OUT_DIR', outM, 'SID_START', 50, 'EXTRAP_WARN_DIST', 5);
check(isequal(size(R2), [2 2]) && numel(RM) == 2 && height(S2) == 4, 'parts x steps result shapes');
check(numel(RM(1).grid_ids) == 24 && max(RM(1).faces(:), [], 'omitnan') == 24, 'merged view: 24 grids, faces re-indexed');
check(R2(1, 2).sid == R2(2, 2).sid && R2(1, 2).sid == 51, 'same SID across parts for a step');
Bexp = 300 + 10 * (R2(2, 1).grid_xyz(:, 1) - 10) + 5 * R2(2, 1).grid_xyz(:, 2) - 2 * R2(2, 1).grid_xyz(:, 3);
inB = R2(2, 1).grid_ids ~= 1999;
check(max(abs(R2(2, 1).grid_T(inB) - Bexp(inB))) < 1e-6, 'part B mapped against its own cloud');
txt = fileread(fullfile(outM, 't100_temp.bdf'));
check(exist(fullfile(outM, 'A'), 'dir') ~= 7 && strcmp(R2(1, 2).out_file, R2(2, 2).out_file) && ...
      contains(txt, '$ ---- part A ----') && contains(txt, '$ ---- part B ----') && ...
      numel(regexp(txt, '^TEMP ', 'lineanchors')) == ceil(numel(R2(1, 2).grid_ids) / 3) + ceil(numel(R2(2, 2).grid_ids) / 3), ...
      'one file per step holds every part');
txt = fileread(fullfile(outM, 'temp_includes.bdf'));
check(numel(regexp(txt, 'INCLUDE ''t\d+_temp\.bdf''')) == 2, 'includes: one per step');
txt = fileread(fullfile(outM, 'temp_subcases.dat'));
check(numel(regexp(txt, 'SUBCASE \d+')) == 2, 'one SUBCASE per step for the whole assembly');
check(iscell(RM(1).surface) && numel(RM(1).surface) == 2 && any(contains(RM(1).warnings, 'B:')), ...
      'merged surfaces per part, warnings prefixed by part');
% steps pair across parts by the trailing number, not the file name
outR = fullfile(tempdir, 'temp_map_test_renamed'); if exist(outR, 'dir'), rmdir(outR, 's'); end
mkdir(outR);
copyfile(fullfile(here, 'partB', 't000.csv'), fullfile(outR, 'fuse_case_0.csv'));
copyfile(fullfile(here, 'partB', 't100.csv'), fullfile(outR, 'fuse_case_0100.csv'));
[Rr, Sr] = temp_map_matlab('PARTS', {'A', bdf, fullfile(here, 'partA'); 'B', fullfile(here, 'test_temp_map_B.bdf'), outR}, ...
                           'WRITE', false);
check(max(abs(Rr(2, 2).grid_T - R2(2, 2).grid_T)) < 1e-12 && contains(Rr(2, 2).csv_file, 'fuse_case_0100.csv') && ...
      strcmp(Sr.File{1}, 't000.csv'), 'differently named CSVs paired by trailing number (leading zeros ignored)');
copyfile(fullfile(here, 'partB', 't100.csv'), fullfile(outR, 'fuse_other_100.csv'));
try
    temp_map_matlab('PARTS', {'A', bdf, fullfile(here, 'partA'); 'B', fullfile(here, 'test_temp_map_B.bdf'), outR}, 'WRITE', false);
    check(false, 'two CSVs with the same number should error');
catch ME
    check(strcmp(ME.identifier, 'temp_map_matlab:ambiguousStep'), 'ambiguous step number errors clearly');
end
delete(fullfile(outR, 'fuse_other_100.csv')); delete(fullfile(outR, 'fuse_case_0.csv'));
try
    temp_map_matlab('PARTS', {'A', bdf, fullfile(here, 'partA'); 'B', fullfile(here, 'test_temp_map_B.bdf'), outR}, 'WRITE', false);
    check(false, 'missing step number should error');
catch ME
    check(strcmp(ME.identifier, 'temp_map_matlab:missingStep') && contains(ME.message, 't000.csv'), 'missing step number errors clearly');
end
G2 = temp_map_matlab('PARTS', PARTS, 'READ_ONLY', true);
check(numel(G2) == 2 && strcmp(G2(2).name, 'B'), 'READ_ONLY returns one grids struct per part');
try
    temp_map_matlab('PARTS', PARTS, 'CSV_FILES', {'t000.csv', 'nope.csv'}, 'WRITE', false);
    check(false, 'missing step should error');
catch ME
    check(contains(ME.message, 'nope.csv') || contains(ME.message, 'not found'), 'missing step errors clearly');
end
h = temp_map_plot(RM(1), 'Visible', 'off', 'Style', 'contour');
check(~isempty(h.mesh), 'merged assembly contour plot');
close(h.fig);

% --- isothermal / uniform parts (no cloud) -------------------------------------------
toF = @(K) (K - 273.15) * 9/5 + 32;
outI = fullfile(tempdir, 'temp_map_test_iso');
if exist(outI, 'dir'), rmdir(outI, 's'); end
[Ri, Si] = temp_map_matlab('BDF_FILE', bdf, 'ISOTHERMAL', 70, 'OUT_UNITS', 'F', 'OUT_DIR', outI, 'SID_START', 7);
check(isequal(size(Ri), [1 1]) && max(abs(toF(Ri.grid_T) - 70)) < 1e-9 && numel(Ri.grid_T) == 12, ...
      'ISOTHERMAL constant: every grid at 70 F (stored in K)');
check(Ri.sid == 7 && isempty(Ri.cloud_T) && ~any(Ri.extrap) && isempty(Ri.warnings) && Ri.time == 0, ...
      'isothermal: one step at t = 0, no cloud, nothing flagged');
check(strcmp(Si.File{1}, 'isothermal') && exist(fullfile(outI, 'isothermal_temp.bdf'), 'file') == 2, ...
      'isothermal -> isothermal_temp.bdf');
txt = fileread(Ri.out_file);
lines = regexp(txt, '\r?\n', 'split');
L = lines{find(startsWith(lines, 'TEMP    '), 1)};
check(abs(str2double(L(25:32)) - 70) < 5e-3 && contains(txt, '$ Source : isothermal 70 F'), ...
      'isothermal card value 70 F and header names the source');
tab = fullfile(outI, 'T_of_t.csv');
fid = fopen(tab, 'w'); fprintf(fid, 'time_s,T_F\n100,150\n0,50\n'); fclose(fid);   % unsorted on purpose
[Rt, St] = temp_map_matlab('BDF_FILE', bdf, 'ISOTHERMAL', tab, 'OUT_UNITS', 'F', 'OUT_DIR', outI);
check(isequal(size(Rt), [1 2]) && isequal([Rt.time], [0 100]) && isequal(St.File', {'t0', 't100'}), ...
      'T(t) table: one step per table time, sorted');
check(max(abs(toF(Rt(1).grid_T) - 50)) < 1e-9 && max(abs(toF(Rt(2).grid_T) - 150)) < 1e-9 && ...
      exist(fullfile(outI, 't100_temp.bdf'), 'file') == 2, 'T(t) table: temperatures per step, t<time>_temp.bdf');
Rx = temp_map_matlab('BDF_FILE', bdf, 'ISOTHERMAL', tab, 'OUT_UNITS', 'F', 'TIMES', [25 75], 'WRITE', false);
check(isequal([Rx.time], [25 75]) && abs(toF(Rx(1).grid_T(1)) - 75) < 1e-9 && abs(toF(Rx(2).grid_T(1)) - 125) < 1e-9, ...
      'TIMES: explicit steps, T interpolated in the table');
% mixed assembly: cloud part + constant part, uniform part listed first
outX = fullfile(tempdir, 'temp_map_test_mixed');
if exist(outX, 'dir'), rmdir(outX, 's'); end
PX = {'B', fullfile(here, 'test_temp_map_B.bdf'), 20; 'A', bdf, fullfile(here, 'partA')};
[Rm, Sm, RMm] = temp_map_matlab('PARTS', PX, 'OUT_DIR', outX, 'OUT_UNITS', 'C', 'SID_START', 50);
check(isequal(size(Rm), [2 2]) && isequal([Rm(1, :).time], [Rm(2, :).time]) && isequal([Rm(1, :).time], [0 100]), ...
      'mixed: steps and times come from the cloud part even when listed second');
check(max(abs(Rm(1, 2).grid_T - 293.15)) < 1e-9 && max(abs(Rm(2, 1).grid_T - R2(1, 1).grid_T)) < 1e-9, ...
      'mixed: part B uniform 20 C, part A mapped from its cloud');
check(Rm(1, 2).sid == Rm(2, 2).sid && Rm(1, 2).sid == 51 && strcmp(Sm.File{1}, 't000.csv'), ...
      'mixed: shared SIDs and step names');
check(numel(RMm(1).grid_ids) == 24 && size(RMm(1).cloud_xyz, 1) == size(R2(1, 1).cloud_xyz, 1), ...
      'mixed: merged view carries both parts and only the real cloud');
txt = fileread(fullfile(outX, 't000_temp.bdf'));
check(contains(txt, '$ ---- part B ----') && contains(txt, '$ ---- part A ----') && contains(txt, '$ Source : '), ...
      'mixed: one step file holds the uniform and the cloud part');
txt = fileread(fullfile(outX, 'temp_includes.bdf'));
check(numel(regexp(txt, 'INCLUDE ''t\d+_temp\.bdf''')) == 2, 'mixed: includes one per step');
h = temp_map_plot(RMm(1), 'Visible', 'off', 'Style', 'contour');
check(~isempty(h.mesh), 'mixed: merged plot with a cloud-less part');
close(h.fig);
h = temp_map_plot(Ri, 'Visible', 'off');
close(h.fig);
check(true, 'isothermal-only plot ran');
% typed as text (the GUI table hands the engine a string)
Rs2 = temp_map_matlab('PARTS', {'B', fullfile(here, 'test_temp_map_B.bdf'), ' 25 '}, 'OUT_UNITS', 'C', 'WRITE', false);
check(max(abs(Rs2.grid_T - 298.15)) < 1e-9, 'a temperature typed as text is accepted');
try
    temp_map_matlab('PARTS', {'B', fullfile(here, 'test_temp_map_B.bdf'), 'no_such_source'}, 'WRITE', false);
    check(false, 'bad source should error');
catch ME
    check(strcmp(ME.identifier, 'temp_map_matlab:noDir'), 'bad source errors clearly');
end

% --- parallel path (falls back to serial without the toolbox) + report ---------
outP = fullfile(tempdir, 'temp_map_test_par');
if exist(outP, 'dir'), rmdir(outP, 's'); end
[Rp, ~, ~] = temp_map_matlab('BDF_FILE', bdf, 'CSV_DIR', here, 'OUT_DIR', outP, 'SID_START', 10, ...
                             'PARALLEL', true, 'REPORT', true, 'SAVE_PNG', true);
check(isequal(size(Rp), [1 2]) && max(abs(Rp(2).grid_T - R(2).grid_T)) < 1e-12, ...
      'PARALLEL gives the same temperatures as serial');
check(contains(Rp(2).method, '(reused)'), 'PARALLEL step 2 reused the step-1 mapping');
rep = fullfile(outP, 'temp_map_report.html');
check(exist(rep, 'file') == 2, 'HTML report written');
txt = fileread(rep);
check(contains(txt, 't100.csv') && contains(txt, '<svg') && contains(txt, 'overhangs') && contains(txt, 't000_temp.png'), ...
      'report has the step table, chart, warnings and plot links');
check(exist(fullfile(outP, 't100_temp.png'), 'file') == 2, 'PNG per step written');

% --- cloud check (no BDF) ----------------------------------------------------------
[Sc, Dc, hc] = temp_cloud_check(here, 'OUT_UNITS', 'K', 'OUT_LENGTH_UNITS', 'in');
check(height(Sc) == 2 && abs(Sc.R2(1) - 1) < 1e-9, 'cloud check: linear fixture field fits with R2 = 1');
check(abs(Sc.gfit_ip(1)^2 + Sc.gfit_tt(1)^2 - 129) < 1e-6, 'cloud check: fit gradient sqrt(129) split into in-plane + through-thickness');
check(abs(Sc.dT_total(1) - (max(R(1).cloud_T) - min(R(1).cloud_T))) < 1e-9 && numel(Dc) == 2 && ishandle(hc), ...
      'cloud check: total delta T, detail struct, figure');
close(hc);
Sr = temp_cloud_check({fullfile(here, 't000.csv')}, 'PICK', 'z', 'RANGE', [-0.5 0.5], 'AXIS', [0 0 1], 'PLOT', false, 'OUT_UNITS', 'K');
check(Sr.N < Sc.N(1) && Sr.Thick <= 1 + 1e-9 && abs(Sr.R2 - 1) < 1e-6, ...
      'cloud check: RANGE along z keeps only the middle layer');
Sa = temp_cloud_check({fullfile(here, 't000.csv')}, 'AXIS', [0 0 1], 'PLOT', false, 'OUT_UNITS', 'K');
check(abs(Sa.gfit_tt - 2) < 1e-6 && abs(Sa.gfit_ip - sqrt(125)) < 1e-6 && abs(Sa.Thick - 3) < 1e-9, ...
      'cloud check: AXIS [0 0 1] -> through-thickness gradient 2, in-plane sqrt(125), thickness 3');

% --- plot smoke test --------------------------------------------------------
h = temp_map_plot(R(1), 'Visible', 'off', 'Units', 'C');
h.set_view('+Z'); h.set_view('ISO');
close(h.fig);
h = temp_map_plot(R(1), 'Visible', 'off', 'Style', 'contour');
check(~isempty(h.mesh), 'contour style drew the mesh patch');
close(h.fig);
check(true, 'temp_map_plot ran');

fprintf('\nALL CHECKS PASSED\n');
end

function D = pdist2_plain(A, B)
%PDIST2_PLAIN  Euclidean distance matrix without the Statistics Toolbox.
    D = sqrt(max(sum(A.^2, 2) + sum(B.^2, 2)' - 2 * (A * B'), 0));
end

function check(ok, what)
    if ok
        fprintf('  PASS  %s\n', what);
    else
        error('run_temp_map_test:fail', 'FAIL  %s', what);
    end
end
