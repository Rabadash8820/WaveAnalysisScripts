#pragma once

public ref class Sttc {
	// INTERFACE
public:
	double Run(array<double>^ spikeTrain1, array<double>^ spikeTrain2, double startTime, double endTime, double dt);

	// HELPERS
private:
	double getP(array<double>^ spikeTrain1, array<double>^ spikeTrain2, double dt);
	double getT(array<double>^ spikeTrain, double startTime, double endTime, double dt);
};
