% void h5toPLX(string filePath)
% Created by Dan Vicarel around 9pm on 5/25/2016
% Opens the HDF5 file with the provided path and stores the spike
% timestamps that it contains into a matrix.

% Read data from the h5 file
counts       = h5read(filePath,'/sCount');
spikesLinear = h5read(filePath,'/spikes');
N            = h5read(filePath,'/summary/N');
numSpikes    = h5read(filePath,'/summary/totalspikes');

% Create the spike timestamp matrix
spikesMatrix = zeros(max(counts), N);

% Populate the matrix with timestamps
offset = 0;
for unit = 1 : N
    for spike = 1 : counts(unit)
        spikesMatrix(unit, spike) = spikesLinear(offset + spike);
    end
    offset = offset + counts(unit);
end

% Display a success message and free up memory
disp(['Successfully loaded ',num2str(numSpikes),' spikes from ',num2str(N),' units!']);
clear('counts', 'spikesLinear', 'N', 'numSpikes');
clear('offset', 'unit', 'spike');
