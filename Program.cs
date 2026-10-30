using System;
using System.Diagnostics;

namespace TiaExportBlocks
{
    class Program
    {
        static int Main(string[] args)
        {
            string exportLocation;
            if (!ExportRules.TryParseArgs(args, out exportLocation))
            {
                Console.WriteLine("No arguments provided");
                Console.WriteLine(@"Usage TiaExportBlocks.exe C:\Test");
                Console.ReadLine();
                return 2;
            }

            string error;
            if (!ExportRules.TryNormalizeExportLocation(exportLocation, out exportLocation, out error))
            {
                Console.WriteLine("ERROR: " + error);
                return 2;
            }

            Console.WriteLine("Export location is " + exportLocation);
            var exporter = new Exporter(exportLocation, Console.Out);
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
