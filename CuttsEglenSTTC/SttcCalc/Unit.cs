using System;
using System.Collections.Generic;

namespace SttcCalc {

    public class Unit {

        // HIDDEN FIELDS
        private Channel _ch;

        // CONSTRUCTORS
        public Unit(char letter) {
            Letter = letter;
        }
        public Unit(Channel channel, char letter) {
            Channel = channel;
            Letter = letter;
        }

        // INTERFACE
        public Channel Channel {
            get { return _ch; }
            set { setChannel(value); }
        }
        public char Letter { get; }
        public IList<double> SpikeTrain { get; } = new List<double>();
        public static double SqrDistance(Unit unit1, Unit unit2, double interChannelDist = 1d) {
            return Channel.SqrDistance(unit1.Channel, unit2.Channel, interChannelDist);
        }
        public static double Distance(Unit unit1, Unit unit2, double interChannelDist = 1d) {
            return Channel.Distance(unit1.Channel, unit2.Channel, interChannelDist);
        }
        public override string ToString() {
            return Channel?.Name + Letter;
        }

        // HELPERS
        private void setChannel(Channel channel) {
            // Only allow the Channel to be set once
            if (_ch != null) {
                if (channel != _ch)
                    throw new InvalidOperationException($"A {nameof(Unit)} cannot have its {nameof(Channel)} set more than once!");
                return;
            }

            // Create a parent-child relationship with the provided Channel
            _ch = channel;
            _ch.AddUnit(this);
        }

    }

}
