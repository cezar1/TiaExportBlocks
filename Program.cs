using System;

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
            DateTime startedUtc = DateTime.UtcNow;
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

            exporter.Run(SiemensTia.GetProcesses());

            Benchmark benchmark = exporter.Benchmark;
            Console.WriteLine($"Execution Time: {(long)benchmark.Total.TotalMilliseconds} ms ({Benchmark.FormatDuration(benchmark.Total)})");
            if (options.Benchmark)
            {
                Console.Write(benchmark.FormatReport());
                string csv = options.BenchmarkCsv ?? Benchmark.DefaultCsvPath;
                try
                {
                    Benchmark.AppendCsv(csv, benchmark.ToCsvLine(startedUtc, options.ExportLocation, options.Incremental, options.BatchSize, exporter.Summaries));
                    Console.WriteLine("Benchmark appended to " + csv);
                }
                catch (Exception exc)
                {
                    Console.WriteLine("WARNING: could not write benchmark to " + csv + ": " + exc.Message);
                }
            }
            Console.WriteLine("Done");
            return 0;
        }
    }
}
