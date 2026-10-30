using System;
using System.Diagnostics;

namespace TiaExportBlocks
{
    class Program
    {
        static void Main(string[] args)
        {
            string exportLocation;
            if (!ExportRules.TryParseArgs(args, out exportLocation))
            {
                Console.WriteLine("No arguments provided");
                Console.WriteLine(@"Usage TiaExportBlocks.exe C:\Test");
                Console.ReadLine();
                return;
            }

            Console.WriteLine("Export location is " + exportLocation);
            var exporter = new Exporter(exportLocation, Console.Out);
            exporter.PrepareExportLocation();

            var watch = Stopwatch.StartNew();
            exporter.Run(SiemensTia.GetProcesses());
            watch.Stop();

            Console.WriteLine($"Execution Time: {watch.ElapsedMilliseconds} ms");
            Console.WriteLine("Done");
        }
    }
}
