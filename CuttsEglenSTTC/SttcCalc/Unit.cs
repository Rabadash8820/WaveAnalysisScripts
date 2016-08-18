using System;
using System.Collections.Generic;

namespace SttcCalc {

    public class Unit {
        // CONSTRUCTORS
        public Unit(string name) {
            reset(name);
        }

        // INTERFACE
        public string Name { get; private set; }
        public Tuple<short, short> MeaCoordinates { get; private set; }
        public IList<double> SpikeTrain { get; } = new List<double>();
        public static double SqrDistance(Unit unit1, Unit unit2, double interChannelDist = 1d) {
            int dr = unit1.MeaCoordinates.Item1 - unit2.MeaCoordinates.Item1;
            int dc = unit1.MeaCoordinates.Item2 - unit2.MeaCoordinates.Item2;
            double sqrDist = interChannelDist * interChannelDist * (dr * dr + dc * dc);
            return sqrDist;
        }
        public static double Distance(Unit unit1, Unit unit2, double interChannelDist = 1d) {
            double dist = Math.Sqrt(SqrDistance(unit1, unit2, interChannelDist));
            return dist;
        }

        // HELPERS
        private void reset(string name) {
            Name = name;
            MeaCoordinates = coordsFromName(name);
        }
        private static Tuple<short, short> coordsFromName(string name) {
            // Assumes unit name is in format "adch_{row}{col}"
            // Where {row} and {col} are 1-based (and must be converted to 0-based)
            string coordsStr = name.Substring("adch_".Length, 2);
            int row = Convert.ToInt16(coordsStr[0]) - 49;
            int col = Convert.ToInt16(coordsStr[1]) - 49;
            return new Tuple<short, short>((short)row, (short)col);
        }
    }

}
