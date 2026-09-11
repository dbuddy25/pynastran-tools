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

temp_map_matlab('RESULTS', R(1), 'OUT_DIR', out, 'OUT_UNITS', 'F');
txt = fileread(R(1).out_file);
lines = regexp(txt, '\r?\n', 'split');
L = lines{find(startsWith(lines, 'TEMP    '), 1)};
check(abs(str2double(L(25:32)) - ((R(1).grid_T(1) - 273.15) * 9/5 + 32)) < 5e-3, 'Fahrenheit output');

% --- nearest / idw methods against brute force ------------------------------
Rn = temp_map_matlab('BDF_FILE', bdf, 'CSV_FILES', {fullfile(here, 't000.csv')}, ...
                     'METHOD', 'nearest', 'WRITE', false);
D = pdist2_plain(Rn.grid_xyz, Rn.cloud_xyz);
[dmin, nn] = min(D, [], 2);
check(max(abs(Rn.grid_T - Rn.cloud_T(nn))) < 1e-9, 'nearest: matches brute-force nearest point');
check(max(abs(Rn.nn_dist - dmin)) < 1e-9, 'nearest: nn_dist matches brute force');
Rs = temp_map_matlab('BDF_FILE', bdf, 'CSV_FILES', {fullfile(here, 't000.csv')}, ...
                     'METHOD', 'scattered', 'WRITE', false);
check(max(abs(Rs.grid_T - R(1).grid_T)) < 1e-6, 'scattered: matches linear on the fixture');
Ri = temp_map_matlab('BDF_FILE', bdf, 'CSV_FILES', {fullfile(here, 't000.csv')}, ...
                     'METHOD', 'idw', 'WRITE', false);
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
