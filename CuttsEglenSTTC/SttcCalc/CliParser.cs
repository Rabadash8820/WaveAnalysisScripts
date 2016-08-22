using System;
using System.IO;

namespace SttcCalc {

    internal struct CliParams {
		public bool ShowHelp { get; set; }
        public DirectoryInfo InputDir { get; set; }
        public FileInfo OutputFile { get; set; }
		public bool ValuesOnly { get; set; }
    }
    internal class CliException : ArgumentException {
        public string ValueProvided { get; }

        public CliException(string message) :
            base(message) { }
        public CliException(string message, string paramName) :
            base(message, paramName) { }
        public CliException(string message, string paramName, string value) :
            base(message, paramName) {
            ValueProvided = value;
        }
        public CliException(string message, string paramName, string value, Exception innerException) :
            base(message, paramName, innerException) {
            ValueProvided = value;
        }
    }

    internal class CliParser {
        // HIDDEN FIELDS
        private static int _argIndex = 0;

        // INTERFACE
        public static CliParams Parse(string[] args) {
            // Set default command line values
			CliParams cp = defaultParams();

			// Parse options/arguments provided by the user
            for (int a = 0; a < args.Length; ++a) {
                if (isCliOption(args[a]))
                    a += parseCliOption(args, a, ref cp);
                else
                    a += parseCliArgument(args[a], ref cp);
            }

			// Provide final high-level validation
			validateParams(ref cp);

            return cp;
        }
        public static void ShowUsage() {
            // Display correct syntax
            Console.WriteLine(@"SttcCalc drive:\path [/H] [/V] [/O outputFilename]");
            Console.WriteLine();
            Console.WriteLine("  /O\tSet the file to which STTC values will be redirected.");
            Console.WriteLine("  /H\tShow this help text.");
            Console.WriteLine("  /V\tOutput STTC values only, i.e. no separators to make output more human-readable.");
        }
        public static void ShowDescription() {
            // Display a brief summary of this program's functionality
            Console.WriteLine("Calculates the STTC for all unit pairs in all recordings ");
            Console.WriteLine("in the provided directory");
            Console.WriteLine();
			
            // Display correct syntax
            Console.WriteLine();
			ShowUsage();
			
            // Display an extended description
            Console.WriteLine();
            Console.WriteLine(string.Join("\n",
                "Recording files should be ASCII .txt files in the format exported by",
                "NeuroExplorer.  I.e., one column per unit, with the unit's name on the",
                "first row, and spike times sorted ascending."
            ));
        }

        // HELPERS
        private static CliParams defaultParams() {
            CliParams cp = new CliParams() {
				ValuesOnly = false,
				ShowHelp = false,
            };
            return cp;
        }
        private static bool isCliOption(string token) {
			char t0 = token[0];
            bool isOption = (t0 == '/' || t0 == '-');
            return isOption;
        }
        private static int parseCliOption(string[] args, int index, ref CliParams cliParams) {
			// If there are no other tokens in the option string, then throw an Exception
			if (args[index].Length == 1) {
                string errMsg = $"'{args[index]}' is not a valid option.";
                throw new CliException(errMsg);
			}

            char opt = args[index][1];
			int indexDelta = 0;

            // Parse help switch
            if (opt == 'h' || opt == 'H' || opt == '?')
                cliParams.ShowHelp = true;

            // Parse output file option
            else if (opt == 'o' || opt == 'O') {
				string paramName = "outputFile";
                if (args.Length <= index + 1) {
                    string errMsg = "You must provide a filename with the /O flag.";
                    throw new CliException(errMsg, paramName);
                }
                else {
                    string fileName = args[index + 1];
                    if (isCliOption(fileName)) {
                        string errMsg = $"{fileName} is not a valid filename.";
                        throw new CliException(errMsg, paramName, fileName);
                    }
                    else {
                        try {
                            FileInfo fi = new FileInfo(fileName);
                            cliParams.OutputFile = fi;
							indexDelta = 1;
                        }
                        catch (Exception e) {
                            string errMsg = $"{fileName} is not a valid filename or could not be created.";
                            throw new CliException(errMsg, paramName, fileName, e);
                        }
                    }
                }
            }

            // Parse values-only switch
            else if (opt == 'v' || opt == 'V')
                cliParams.ValuesOnly = true;

			// If this is not a valid option, then throw an exception
			else {
                string errMsg = $"'{args[index]}' is not a valid option.";
                throw new CliException(errMsg);
			}

			return indexDelta;
        }
        private static int parseCliArgument(string arg, ref CliParams cliParams) {
			int indexDelta = 0;

            // Parse input directory
            if (_argIndex == 0) {
                try {
                    DirectoryInfo di = new DirectoryInfo(arg);
                    cliParams.InputDir = di;
                }
                catch (Exception e) {
                    string errMsg = $"{arg} is not a valid directory path or could not be opened.";
                    throw new CliException(errMsg, "inputDirectory", arg, e);
                }
            }

            else if (_argIndex >= 1) {
                string errMsg = $"Too many arguments provided.";
                throw new CliException(errMsg);
            }

            ++_argIndex;
			return indexDelta;
        }
		private static void validateParams(ref CliParams cliParams) {
			// If we're just showing the help text then don't bother validating further
			if (cliParams.ShowHelp)
				return;

			// Make sure an input directory has been provided
			if (cliParams.InputDir == null) {
				string errMsg = $"You must provide the path to a directory containing recording text files!";
				throw new CliException(errMsg, "inputDirectory");				
			}

			// If so, default the output file to a file in that directory (if one wasn't provided explicitly)
			else if (cliParams.OutputFile == null) {
				cliParams.OutputFile = new FileInfo(cliParams.InputDir.FullName + @"\sttc.txt");
			}
		}

    }

}
