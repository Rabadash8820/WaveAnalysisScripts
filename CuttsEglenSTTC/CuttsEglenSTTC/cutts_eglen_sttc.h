namespace CuttsEglen {

	extern double run_P(int N1, int N2, double dt, double* spike_times_1, double* spike_times_2);
	extern double run_T(int N1v, double dtv, double startv, double endv, double* spike_times_1);
	extern void run_sttc(int* N1v, int* N2v, double* dtv, double* Time, double* index, double* spike_times_1, double* spike_times_2);

}
