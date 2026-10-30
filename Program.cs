using System;
using System.Diagnostics;

namespace TiaExportBlocks
{
    class Program
    {
        /// <returns>0 done, 1 export location could not be emptied, 2 invalid arguments.</returns>
        static int Main(string[] args)
        {
            CommandLineOptions options;
            string error;
            if (!CommandLineOptions.TryParse(args, out options, out error))
            {
                Console.WriteLine("ERROR: " + error);
                Console.WriteLine(CommandLineOptions.Usage);
                return 2;
            }

            Console.WriteLine("Export location is " + options.ExportLocation
                + (options.Incremental ? " (incremental export)" : " (full export)")
                + ", batch size " + options.BatchSize);
            var settings = new ExporterSettings { Incremental = options.Incremental, BatchSize = options.BatchSize };
            var exporter = new Exporter(options.ExportLocation, Console.Out, settings);
            try
            {
                exporter.PrepareExportLocation();
            }
            catch (ExportLocationNotCleanException exc)
            {
                Console.WriteLine("ERROR: " + exc.Message);
                return 1;
            }

            var watch = Stopwatch.StartNew();
            exporter.Run(SiemensTia.GetProcesses());
            watch.Stop();

            Console.WriteLine($"Execution Time: {watch.ElapsedMilliseconds} ms");
            Console.WriteLine("Done");
            return 0;
        }
    }
}
