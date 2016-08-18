using System;
using System.IO;

namespace SttcCalc {

    public class Recording {
        // CONSTRUCTORS
        public Recording(FileInfo textFile, short meaRows, short meaCols, double startTime, double endTime) {
            TextFile = textFile;
            Dimensions = new Tuple<short, short>(meaRows, meaCols);
            UnitGrid = new Unit[meaRows, meaCols];
            StartTime = startTime;
            EndTime = endTime;
        }

        // INTERFACE
        public FileInfo TextFile { get; }
        public Tuple<short, short> Dimensions { get; }
        public double StartTime { get; }
        public double EndTime { get; }
        public Unit[,] UnitGrid { get; }
        public double Duration => EndTime - StartTime;
    }

}
