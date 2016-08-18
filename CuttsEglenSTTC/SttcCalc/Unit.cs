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
