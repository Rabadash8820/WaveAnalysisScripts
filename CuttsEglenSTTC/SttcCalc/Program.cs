using System;
using System.IO;
using System.Linq;

using CuttsEglen;

namespace SttcCalc {

    class Program {

        private static void Main(string[] args) {
            // Make sure only one directory path is provided
            if (args.Length == 0) {
                showUsage("You must provide a directory path");
                return;
            }
            else if (args.Length > 1) {
                showUsage("Too many arguments provided.");
                return;
            }

            // Analyze all provided recordings in that directory
            Console.WriteLine("Analyzing provided recordings...");
            DirectoryInfo dir = new DirectoryInfo(args[0]);
            Recording[] recs = dir.EnumerateFiles("*.txt")
                                  .Select(f => Recording.FromText(f))
                                  .ToArray();
            Console.WriteLine("Complete!");

            Console.ReadKey();
        }
        private static void showUsage(string errorText = "") {
            // Display any provided error messages
            if (errorText != "")
                Console.WriteLine($"ERROR: {errorText}");

            // Display correct syntax
            Console.WriteLine();
            Console.WriteLine("Calculates the STTC for all unit pairs in all recordings in the provided directory");
            Console.WriteLine();
            Console.WriteLine(@"SttcCalc drive:\path");
            Console.WriteLine();
            Console.WriteLine(string.Join("\n",
                "Recording files should be ASCII .txt files in the format exported by",
                "NeuroExplorer.  I.e., one column per unit, with the unit's name on the",
                "first row, and spike times sorted ascending."
            ));
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
