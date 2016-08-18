#pragma once

namespace CuttsEglen {

	public ref class Sttc {
	public:
		static double Calculate(array<double>^ spikeTrain1, array<double>^ spikeTrain2, double startTime, double endTime, double dt);
	};

}
