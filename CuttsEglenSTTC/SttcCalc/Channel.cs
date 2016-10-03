using System;
using System.Linq;
using System.Collections.Generic;

namespace SttcCalc {

    public class Channel {

        // HIDDEN FIELDS
        private ISet<Unit> _units = new HashSet<Unit>();

        // CONSTRUCTORS
        public Channel(string name) {
            Name = name;
            MeaCoordinates = CoordsFromName(name);
        }
        public Channel(string name, Tuple<short, short> coords) {
            Name = name;
            MeaCoordinates = coords;
        }

        // INTERFACE
        public string Name { get; }
        public Tuple<short, short> MeaCoordinates { get; }
        public Unit[] Units => _units.ToArray();
        public void AddUnit(Unit unit) {
            _units.Add(unit);
            unit.Channel = this;
        }
        public override string ToString() {
            return Name;
        }
        public static double SqrDistance(Channel ch1, Channel ch2, double interChannelDist = 1d) {
            int dRow = ch1.MeaCoordinates.Item1 - ch2.MeaCoordinates.Item1;
            int dCol = ch1.MeaCoordinates.Item2 - ch2.MeaCoordinates.Item2;
            double sqrDist = interChannelDist * interChannelDist * (dRow * dRow + dCol * dCol);
            return sqrDist;
        }
        public static double Distance(Channel ch1, Channel ch2, double interChannelDist = 1d) {
            double dist = Math.Sqrt(SqrDistance(ch1, ch2, interChannelDist));
            return dist;
        }
        public static Tuple<short, short> CoordsFromName(string name) {
            // Assumes unit name is in format "adch_{row}{col}{rest}"
            // Where {row} and {col} are 1-based (and will be converted to 0-based)
            string coordsStr = name.Substring("adch_".Length, 2);
            int row = Convert.ToInt16(coordsStr[0]) - 49;
            int col = Convert.ToInt16(coordsStr[1]) - 49;
            return Tuple.Create((short)row, (short)col);
        }

        // HELPERS
    }

}
