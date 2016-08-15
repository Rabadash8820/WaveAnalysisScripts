#include "Sttc.h"

#include <cmath>

using namespace System;

// INTERFACE
double Sttc::getP(array<double>^ spikeTrain1, array<double>^ spikeTrain2, double dt) {
	int Nab = 0;

	// For every spike in train 1, increment Nab if that spike is within dt of a spike on train 2 (don't count spike pairs)
	// We don't need to search all of train 2 each iteration
	int s2 = 0;
	for (int s1 = 0; s1 <= (spikeTrain1->Length - 1); s1++) {
		while (s2 < spikeTrain2->Length) {
			if (fabs(spikeTrain1[s1] - spikeTrain2[s2]) <= dt) {
				Nab = Nab + 1;
				break;
			}
			else if (spikeTrain2[s2] > spikeTrain1[s1])
				break;
			else
				s2 = s2 + 1;
		}
	}

	return (double)Nab / (double)spikeTrain1->Length;
}
double Sttc::getT(array<double>^ spikeTrain, double startTime, double endTime, double dt) {
	// maximum
	int n = spikeTrain->Length;
	double time = 2 * (double)n * dt;

	// if just one spike in train 
	if (n == 1) {
		if (spikeTrain[0] - startTime < dt)
			time += -startTime + spikeTrain[0] - dt;
		else if (spikeTrain[0] + dt > endTime)
			time += -spikeTrain[0] - dt + endTime;
	}

	// if more than one spike in train
	else {
		int i = 0;
		while (i < n - 1) {
			double diff = spikeTrain[i + 1] - spikeTrain[i];
			//subtract overlap
			if (diff < 2 * dt)
				time += -2 * dt + diff;
			i++;
		}

		//check if spikes are within dt of the start and/or end, if so just need to subract
		//overlap of first and/or last spike as all within-train overlaps have been accounted for
		if (spikeTrain[0] - startTime < dt)
			time += -startTime + spikeTrain[0] - dt;
		if (endTime - spikeTrain[n - 1] < dt)
			time += -spikeTrain[n - 1] - dt + endTime;
	}

	return time / (endTime - startTime);
}
double Sttc::Run(array<double>^ spikeTrain1, array<double>^ spikeTrain2, double startTime, double endTime, double dt) {
	// If either spike train has zero spikes then throw an exception
	if (spikeTrain1->Length == 0)
		throw gcnew DivideByZeroException("No spikes in spike train 1!");
	else if (spikeTrain2->Length == 0)
		throw gcnew DivideByZeroException("No spikes in spike train 2!");

	// Otherwise, compute the STTC!
	else {
		double ta = getT(spikeTrain1, startTime, endTime, dt);
		double tb = getT(spikeTrain2, startTime, endTime, dt);
		double pa = getP(spikeTrain1, spikeTrain2, dt);
		double pb = getP(spikeTrain2, spikeTrain1, dt);
		return 0.5 * ((pa - tb)/(1 - tb*pa) + (pb - ta)/(1 - ta*pb));
	}

}
