% void h5toPLX(string filePath)
% Created by Dan Vicarel around 9pm on 5/25/2016
% Opens the HDF5 file with the provided path and stores the spike
% timestamps that it contains into a matrix.

function allData = hdf2arr()
    % Get a folder with HDF5 files from the user
    folderPath = uigetdir(matlabroot, 'Select a folder with HDF5 files');
    if folderPath == 0
        error('You have to select a folder with HDF5 files, yo!');
    end

    % Get all HDF5 files in that folder
    hdf5Files = dir([folderPath, '/', '*.hdf5']);
    h5Files   = dir([folderPath, '/', '*.h5']);
    he5Files  = dir([folderPath, '/', '*.he5']);
    files     = [hdf5Files, h5Files, he5Files];
    clear('hdf5Files', 'h5Files', 'he5Files');

    % For each file, read its spike timestamps into a 2D array
    numFiles = numel(files);
    data = cell(numFiles, 1);
    for f = 1 : numFiles
        filePath = [folderPath, '\', files(f).name];
        fileData = getSpikeData(filePath);
        data(f) = {fileData};
        numSpikes = h5read(filePath, '/summary/totalspikes');
        msg = [num2str(numSpikes) ' spikes loaded from ' ...
               num2str(numel(fileData.UnitSpikes)) ' units in ' ...
               '"' files(f).name, '"'];
        disp(msg);
    end
    
    % Show a success message and return all loaded data
    disp(' ');
    disp(['Successfully loaded spike timestamps from all ' num2str(numFiles) ' files!']);
    allData = data;
end

function data = getSpikeData(h5Path)
    % Read data from the h5 file
    dur          = h5read(h5Path, '/summary/duration');
    age          = h5read(h5Path, '/meta/age');
    genotype     = h5read(h5Path, '/meta/genotype');
    array        = h5read(h5Path, '/array');
    counts       = h5read(h5Path, '/sCount');
    epos         = h5read(h5Path, '/epos');
    spikesLinear = h5read(h5Path, '/spikes');

    % Create the spike timestamp matrix
    numUnits = numel(counts);
    unitSpikes = cell(numUnits, 1);

    % Populate the matrix with timestamps
    offset = 0;
    for u = 1 : numUnits
        spikes = zeros(counts(u), 1);
        for s = 1 : counts(u)
            spikes(s) = spikesLinear(offset + s);
        end
        unitSpikes{u} = spikes;
        offset = offset + counts(u);
    end
    
    % Create unit names in the Multi Channel Sytems format
    unitNames = getUnitNames(numUnits, array, epos);
    
    % Return all necessary data wrapped in a struct
    data = struct('FullName', h5Path, ...
                  'Duration', dur, ...
                  'Age', age, ...
                  'Genotype', genotype, ...
                  'UnitSpikes', {unitSpikes}, ...
                  'Names', {unitNames});
end

function names = getUnitNames(N, array, epos)
    % Determine the MEA's interelectrode distance
    array = array{1};
    underScoreIndices = strfind(array, '_');
    umIndices = strfind(array, 'um');
    iStart = underScoreIndices(length(underScoreIndices)) + 1;
    iEnd = umIndices(length(umIndices)) - 1;
    interElecDist = array(iStart : iEnd);
    
    % Divide all electrode positions by that distance
    epos = epos / str2num(interElecDist);
    
    % For each electrode, store a name in the form 'adch_{row}{column}{unit}'
    % {unit} is a letter to distinguish multiple units from the same electrode
    names = cell(N, 1);
    for u = 1 : N
        name = ['adch_' num2str(epos(u,2)) num2str(epos(u,1))];
        letter = 'a';
        if u > 1
            prevName = names(u - 1);
            prevName = prevName{1};
            diffChannel = isempty(strfind(prevName, name));
            if ~diffChannel
                unitLetterAscii = uint8(prevName(length(prevName)));
                letter = char(unitLetterAscii + 1);
            end
        end
        names(u) = {[ name letter ]};        
    end
end

