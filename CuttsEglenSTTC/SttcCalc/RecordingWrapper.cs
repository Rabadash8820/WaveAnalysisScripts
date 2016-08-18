using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using CuttsEglen;

namespace SttcCalc {

    class RecordingWrapper {
        // INTERFACE
        public static IDictionary<double, double[]> STTCvsDistance(Recording rec, double dt) {
            IDictionary<double, IList<double>> sqrResults = new Dictionary<double, IList<double>>();

            // For every pair of units...
            for (short r1 = 0; r1 < rec.Dimensions.Item1; ++r1) {
                for (short c1 = 0; c1 < rec.Dimensions.Item2; ++c1) {
                    Unit u1 = rec.UnitGrid[r1, c1];
                    if (u1 == null)
                        continue;
                    for (short r2 = 0; r2 < rec.Dimensions.Item1; ++r2) {
                        for (short c2 = 0; c2 < rec.Dimensions.Item2; ++c2) {
                            Unit u2 = rec.UnitGrid[r2, c2];
                            if (u2 == null)
                                continue;

                            // Add their STTC value to the array for their square-intervening-distance
                            double sqrDist = Unit.SqrDistance(u1, u2);
                            double sttc = Sttc.Calculate(u1.SpikeTrain.ToArray(), u2.SpikeTrain.ToArray(), rec.StartTime, rec.EndTime, dt);
                            if (!sqrResults.ContainsKey(sqrDist))
                                sqrResults.Add(sqrDist, new List<double>());
                            sqrResults[sqrDist].Add(sttc);
                        }
                    }
                }
            }

            // Get the STTC arrays versus actual distance by sqrting distances from above
            IDictionary<double, double[]> results = sqrResults.ToDictionary(
                pair => Math.Sqrt(pair.Key),
                pair => pair.Value.ToArray());
            return results;
        }
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
            ++maxRows;  // because the highest row index is one less than the NUMBER of rows
            ++maxCols;  // same for columns
            double approxStart = units.Min(u => u.SpikeTrain.Min());
            double approxEnd = units.Max(u => u.SpikeTrain.Max());
            Recording rec = new Recording(file, maxRows, maxCols, approxStart, approxEnd);
            foreach (Unit unit in units)
                rec.UnitGrid[unit.MeaCoordinates.Item1, unit.MeaCoordinates.Item2] = unit;

            return rec;
        }
    }

}
