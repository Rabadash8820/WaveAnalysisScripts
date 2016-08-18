using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;

namespace SttcCalc {

    public class Recording {
        // CONSTRUCTORS
        public Recording(FileInfo textFile, int meaRows, int meaCols) {
            TextFile = textFile;
            UnitGrid = new Unit[meaRows, meaCols];
        }

        // INTERFACE
        public FileInfo TextFile { get; }
        public Tuple<short, short> Dimensions { get; }
        public Unit[,] UnitGrid { get; }
        public static Recording FromText(string filePath) {
            return fromText(new FileInfo(filePath));
        }
        public static Recording FromText(FileInfo file) {
            return fromText(file);
        }

        // HELPERS
        private static Recording fromText(FileInfo file) {
            // Get all unit names from the file
            IEnumerable<string> lines = File.ReadLines(file.FullName);
            long numLines = lines.LongCount();
            string[] unitNames = lines.First().Split('\t');
            int numUnits = (unitNames.Length % 3 == 0 ? unitNames.Length / 3 : (unitNames.Length - 2) / 3);
            Unit[] units = new Unit[numUnits];

            // Define all units while determining the dimensions of the MEA
            short maxRows = 0;
            short maxCols = 0;
            for (int u = 0; u < numUnits; ++u) {
                Unit unit = new Unit(unitNames[u]);
                maxRows = Math.Max(maxRows, unit.MeaCoordinates.Item1);
                maxCols = Math.Max(maxCols, unit.MeaCoordinates.Item2);
                units[u] = unit;
            }

            // Populate each unit's spike train
            using (FileStream fs = file.OpenRead())
            using (BufferedStream bs = new BufferedStream(fs))
            using (StreamReader sr = new StreamReader(bs)) {
                sr.ReadLine();  // Skip the unit names line
                string line;
                while ((line = sr.ReadLine()) != null) {
                    string[] strs = line.Split('\t').Take(numUnits).ToArray();
                    double[] spikes = Array.ConvertAll(strs, (s) =>
                        s == " " ? -1d : Convert.ToDouble(s));
                    for (int u = 0; u < numUnits; ++u) {
                        if (spikes[u] != -1d)
                            units[u].SpikeTrain.Add(spikes[u]);
                    }
                }
            }

            // Create the Recording object with these units and MEA dimensions
            Recording rec = new Recording(file, maxRows + 1, maxCols + 1);
            foreach (Unit unit in units)
                rec.UnitGrid[unit.MeaCoordinates.Item1, unit.MeaCoordinates.Item2] = unit;

            return rec;
        }
    }

}
