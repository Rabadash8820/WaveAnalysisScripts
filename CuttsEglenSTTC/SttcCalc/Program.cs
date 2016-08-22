using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;

using CuttsEglen;

namespace SttcCalc {

    internal class Program {

        private const double CORRELATION_DT = 0.05d;

        private static void Main(string[] args) {
            CliParams cp;
            try {
                cp = CliParser.Parse(args);
				if (cp.ShowHelp)
					CliParser.ShowDescription();
				else
					run(cp);
            }
            catch (CliException e) {
                Console.WriteLine($"Error: {e.Message}");
				Console.WriteLine();
                CliParser.ShowUsage();
            }
            catch (Exception e) {
                Console.WriteLine("Wo, an error occurred!");
                Console.WriteLine($"Message: {e.Message}");
				Console.WriteLine();
                CliParser.ShowUsage();
            }

            //Console.WriteLine("\nPress any key to finish...");
            //Console.ReadKey();
        }
        
        private static void run(CliParams cp) {
            // Analyze all provided recordings in that directory
            Console.Write($"Reading data from the provided recordings...  ");
            IEnumerable<FileInfo> files = cp.InputDir.EnumerateFiles("*.txt");
            if (files.Count() == 0) {
                Console.WriteLine("Complete!");
                Console.WriteLine("No files found.");
                return;
            }
            Recording[] recs = files.Select(f => RecordingWrapper.FromText(f)).ToArray();
            Console.WriteLine("Complete!");

            // Get STTC vs distance for all Recordings
            Console.Write("Calculating STTCs...  ");
            IDictionary<Recording, IDictionary<double, double[]>> results = new Dictionary<Recording, IDictionary<double, double[]>>();
            foreach (Recording rec in recs) {
                IDictionary<double, double[]> recResults = RecordingWrapper.STTCvsDistance(rec, CORRELATION_DT);
                results.Add(rec, recResults);
            }
            Console.WriteLine("Complete!");

            // Output values to a file
            Console.Write($"Saving values to {cp.OutputFile.FullName}...  ");
            using (FileStream fs = cp.OutputFile.Create())
            using (BufferedStream bs = new BufferedStream(fs))
            using (StreamWriter sw = new StreamWriter(bs)) {
				if (!cp.ValuesOnly) {
					sw.WriteLine($"STTC results generated at {DateTime.Now.ToShortTimeString()} on {DateTime.Now.ToShortDateString()}:");
					sw.WriteLine("Unit-Distance\tSTTC");
				}
                foreach (Recording rec in results.Keys) {
					if (!cp.ValuesOnly) {
						int minutes = (int)(rec.Duration / 60d);
						double seconds = rec.Duration - 60d * minutes;
						sw.WriteLine($"\nRecording from {rec.TextFile} ({minutes}m {seconds}s long):");
					}
                    foreach (double dist in results[rec].Keys)
                        foreach (double sttc in results[rec][dist])
                            sw.WriteLine($"{dist}\t{sttc}");
                }
            }
            Console.WriteLine("Complete!");
        }
        private static void runTests() {
            // Create test variables
            Console.WriteLine($"Testing with two 60s spike trains, each with one spike at 30s, and a dt of 0.05s ...");
            double dt = 0.05;
            double startTime = 0d;
            double endTime = 60d;
            double[] emptyTrain = new double[0];
            double[] spikeTrain1 = new double[1] { 30.0d };
            double[] spikeTrain2 = new double[1] { 30.0d };

            // Test #1
            Console.WriteLine($"\nTEST #1");
            Console.WriteLine($"Calculating STTC with the {nameof(Sttc)} C++/CLI class...");
            double result = Sttc.Calculate(spikeTrain1, spikeTrain2, startTime, endTime, dt);
            Console.WriteLine($"Result: {result}");

            // Test #2
            Console.WriteLine($"\nTEST #2");
            Console.WriteLine($"Same, but with empty first spike train...");
            try {
                result = Sttc.Calculate(emptyTrain, spikeTrain2, startTime, endTime, dt);
            }
            catch (DivideByZeroException e) {
                Console.WriteLine($"Caught {nameof(DivideByZeroException)} with message: \n\t{e.Message}");
            }

            // Test #3
            Console.WriteLine($"\nTEST #3");
            Console.WriteLine($"Same, but with empty second spike train...");
            try {
                result = Sttc.Calculate(spikeTrain1, emptyTrain, startTime, endTime, dt);
            }
            catch (DivideByZeroException e) {
                Console.WriteLine($"Caught {nameof(DivideByZeroException)} with message: \n\t{e.Message}");
            }
        }

    }

}
