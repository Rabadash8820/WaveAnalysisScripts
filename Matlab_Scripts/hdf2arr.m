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
        msg = [num2str(fileData.NumSpikes) ' spikes loaded from ' ...
               num2str(fileData.NumUnits) ' units in ' ...
               '"' files(f).name, '"'];
        disp(msg);
    end
    
    % Show a success message and return all loaded data
    disp(' ');
    disp('All spike timestamps succesfully loaded!');
    allData = data;    
end

function data = getSpikeData(h5Path)
    % Read data from the h5 file
    names        = h5read(h5Path, '/names');
    counts       = h5read(h5Path, '/sCount');
    spikesLinear = h5read(h5Path, '/spikes');
    N            = h5read(h5Path, '/summary/N');
    numSpikes    = h5read(h5Path, '/summary/totalspikes');

    % Create the spike timestamp matrix
    unitSpikes = cell(N, 1);

    % Populate the matrix with timestamps
    offset = 0;
    for u = 1 : N
        spikes = zeros(counts(u), 1);
        for s = 1 : counts(u)
            spikes(s) = spikesLinear(offset + s);
        end
        unitSpikes{u} = spikes;
        offset = offset + counts(u);
    end
    
    % Return all necessary data wrapped in a struct
    data = struct('NumUnits', N, 'NumSpikes', numSpikes, 'Names', {names}, 'Spikes', {unitSpikes});
end


