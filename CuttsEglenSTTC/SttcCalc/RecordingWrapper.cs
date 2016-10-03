using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using CuttsEglen;

namespace SttcCalc {

    class RecordingWrapper {
        // INTERFACE
        public static IDictionary<double, double[]> STTCvsDistance(Recording rec, double dt) {
            // For every unique pair of units, calculate the STTC, then
            // group these values by the distance between the units' channels
            IEnumerable<Unit> units = rec.EnumerateChannels().SelectMany(ch => ch.Units);
            IDictionary<double, double[]> results =
                units.SelectMany(
                        (u, index) => units.Skip(index + 1),
                        (u1, u2) => Tuple.Create(u1, u2))
                     .GroupBy(
                        pair => Unit.Distance(pair.Item1, pair.Item2),
                        pair => Sttc.Calculate(pair.Item1.SpikeTrain.ToArray(), pair.Item2.SpikeTrain.ToArray(), rec.StartTime, rec.EndTime, dt))
                     .OrderBy(pair => pair.Key)
                     .ToDictionary(
                        pair => pair.Key,
                        pair => pair.ToArray()
                     );
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
            bool validFile = unitNames.Any(n => n.Contains("adch_"));
            if (!validFile)
                throw new ApplicationException($"{file.FullName} is not in the valid Recording format!");
            int numUnits = (unitNames.Length % 3 == 0 ? unitNames.Length / 3 : (unitNames.Length - 2) / 3);
            Unit[] units = new Unit[numUnits];
            IList<Channel> channels = new List<Channel>();

            // Define all units while determining the dimensions of the MEA
            short maxRows = 0;
            short maxCols = 0;
            Channel prevChannel = null;
            for (int u = 0; u < numUnits; ++u) {
                // Get or create this unit's Channel based on the MEA coordinates in its name
                string chName = new string(unitNames[u].Take(unitNames[u].Length - 1).ToArray());
                Tuple<short, short> coords = Channel.CoordsFromName(chName);
                Channel ch;
                if (coords.Equals(prevChannel?.MeaCoordinates))
                    ch = prevChannel;
                else {
                    ch = new Channel(chName, coords);
                    channels.Add(ch);
                    prevChannel = ch;
                }

                // Create the unit itself
                Unit unit = new Unit(ch, unitNames[u].Last());
                units[u] = unit;

                // Adjust the number of rows/columns on this MEA
                maxRows = Math.Max(maxRows, ch.MeaCoordinates.Item1);
                maxCols = Math.Max(maxCols, ch.MeaCoordinates.Item2);
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
            foreach (Unit u in units) {
                Channel ch = u.Channel;
                rec.Channels[ch.MeaCoordinates.Item1, ch.MeaCoordinates.Item2] = ch;
            }

            return rec;
        }
    }

}
