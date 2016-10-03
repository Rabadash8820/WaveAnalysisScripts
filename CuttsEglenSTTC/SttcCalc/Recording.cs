using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;

namespace SttcCalc {

    public class Recording {
        // CONSTRUCTORS
        public Recording(FileInfo textFile, short meaRows, short meaCols, double startTime, double endTime) {
            TextFile = textFile;
            Dimensions = new Tuple<short, short>(meaRows, meaCols);
            Channels = new Channel[meaRows, meaCols];
            StartTime = startTime;
            EndTime = endTime;
        }

        // INTERFACE
        public FileInfo TextFile { get; }
        public Tuple<short, short> Dimensions { get; }
        public double StartTime { get; }
        public double EndTime { get; }
        public Channel[,] Channels { get; }
        public IEnumerable<Channel> EnumerateChannels() {
            return Channels.Cast<Channel>().Where(ch => ch != null);
        }
        public double Duration => EndTime - StartTime;
    }

}
