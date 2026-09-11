function T = meff_matlab(op2file)
%MEFF_MATLAB  Modal effective mass fractions from a Nastran OP2 via pyNastran.
%
%   T = MEFF_MATLAB(OP2FILE) reads OP2FILE with pyNastran and returns a MATLAB
%   table with columns:
%       Mode, Freq_Hz, then Frac/Sum for each of the 6 directions
%       (Tx, Ty, Tz, Rx, Ry, Rz).
%   "Frac" is the per-mode modal effective mass fraction; "Sum" is the
%   cumulative sum down the modes for that direction.
%
%   This is a MATLAB port of the numeric core of
%       postprocessing/modules/meff.py
%   (the GUI / Excel styling is intentionally left out).
%
%   ONE-TIME SETUP (per machine)
%   ----------------------------
%   Point MATLAB at a Python that has pyNastran installed:
%       >> pyenv('Version', '/path/to/your/python')   % e.g. your venv python
%   Confirm it took:
%       >> pyenv                                       % Status should be 'Loaded'
%       >> py.importlib.import_module('pyNastran');    % should not error
%
%   USAGE
%   -----
%       >> T = meff_matlab('model_103.op2');
%       >> head(T)
%       >> writetable(T, 'meff.xlsx')                  % styled export, if wanted
%
%   To see pyNastran's console output in the MATLAB window, run redirectstdout
%   once first (the patched copy that ships alongside this file):
%       >> redirectstdout
%
%   REQUIREMENTS IN THE OP2
%   -----------------------
%   A SOL 103 run with  MEFFMASS(PLOT) = ALL  in the case control deck, so the
%   EFMFACS matrix is written to the OP2.

    directions = {'Tx','Ty','Tz','Rx','Ry','Rz'};

    % --- read the OP2 -------------------------------------------------------
    % import_module is used instead of a MATLAB `import py....` statement to
    % sidestep the namespace-caching quirks called out in the pyNastran docs.
    op2mod = py.importlib.import_module('pyNastran.op2.op2');
    op2 = op2mod.OP2(pyargs('mode', 'nx'));
    op2.read_op2(op2file);

    % --- eigenvalues: mode numbers + frequencies (Hz) -----------------------
    if double(py.len(op2.eigenvalues)) == 0
        error('meff_matlab:noEigenvalues', ...
              'No eigenvalue data in this OP2. Is it a SOL 103 run?');
    end
    evals  = cell(py.list(op2.eigenvalues.values()));
    eigtab = evals{1};                        % first subcase, mirrors next(iter(...))
    modes  = ndarray2mat(eigtab.mode);        % mode numbers
    freqs  = ndarray2mat(eigtab.cycles);      % cycles == frequency in Hz

    % --- EFMFACS modal-effective-mass-fraction matrix -----------------------
    % op2.matrices is a Python dict; .get returns None (py.NoneType) if absent.
    meff = op2.matrices.get('EFMFACS');
    if isa(meff, 'py.NoneType')
        error('meff_matlab:noMeffmass', ...
            ['No MEFFMASS matrices in this OP2.\n', ...
             'Add to your Nastran case control:  MEFFMASS(PLOT) = ALL']);
    end

    raw  = matrix_to_dense(meff.data);        % (6, nmodes), matching numpy
    frac = raw.';                             % (nmodes, 6)
    csum = cumsum(frac, 1);                    % cumulative sum down the modes

    n = min(size(frac, 1), numel(modes));
    modes = modes(1:n);
    freqs = freqs(1:n);
    frac  = frac(1:n, :);
    csum  = csum(1:n, :);

    % --- assemble the output table ------------------------------------------
    varNames = {'Mode', 'Freq_Hz'};
    M = [modes(:), freqs(:)];
    for k = 1:6
        varNames{end+1} = [directions{k} '_Frac']; %#ok<AGROW>
        varNames{end+1} = [directions{k} '_Sum'];  %#ok<AGROW>
        M = [M, frac(:, k), csum(:, k)];            %#ok<AGROW>
    end
    T = array2table(M, 'VariableNames', varNames);
end


% =========================================================================
function A = matrix_to_dense(data)
%MATRIX_TO_DENSE  pyNastran Matrix .data (numpy or scipy.sparse) -> MATLAB.
%   Mirrors _matrix_to_dense() in meff.py: densify scipy sparse first.
    if logical(py.scipy.sparse.issparse(data))
        data = data.toarray();
    end
    A = ndarray2mat(data);
end


% =========================================================================
function A = ndarray2mat(nd)
%NDARRAY2MAT  Convert a numpy ndarray (or list) to a MATLAB double array,
%   preserving shape. numpy is row-major (C order); MATLAB is column-major,
%   so we flatten in C order then reshape + permute back.
    nd = py.numpy.asarray(nd);
    nd = py.numpy.ascontiguousarray(nd, pyargs('dtype', 'float64'));

    shp = cellfun(@double, cell(nd.shape));         % size vector
    flat = double(py.array.array('d', nd.flatten().tolist()));

    if isempty(shp)                 % 0-D scalar
        A = flat;
    elseif numel(shp) == 1          % 1-D vector
        A = flat(:);
    else                            % N-D: undo C-order flatten
        A = permute(reshape(flat, fliplr(shp)), numel(shp):-1:1);
    end
end
