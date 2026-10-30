using Xunit;

namespace TiaExportBlocks.Tests
{
    public class CommandLineTests
    {
        private static CommandLineOptions Parse(params string[] args)
        {
            CommandLineOptions options;
            string error;
            Assert.True(CommandLineOptions.TryParse(args, out options, out error), error);
            return options;
        }

        private static string Fails(params string[] args)
        {
            CommandLineOptions options;
            string error;
            Assert.False(CommandLineOptions.TryParse(args, out options, out error));
            Assert.Null(options);
            Assert.False(string.IsNullOrEmpty(error));
            return error;
        }

        [Fact]
        public void PathOnly_IsFullExportOneBlockPerCall()
        {
            var o = Parse(@"C:\Export");

            Assert.Equal(@"C:\Export", o.ExportLocation);
            Assert.False(o.Incremental);
            Assert.Equal(1, o.BatchSize);
        }

        [Fact]
        public void AllOptions_InAnyOrder()
        {
            var o = Parse("--batch", "50", "--incremental", @"C:\Export\");

            Assert.Equal(@"C:\Export", o.ExportLocation);
            Assert.True(o.Incremental);
            Assert.Equal(50, o.BatchSize);
        }

        [Fact]
        public void Options_AreCaseInsensitive()
        {
            var o = Parse(@"C:\Export", "--INCREMENTAL", "--Batch", "7");

            Assert.True(o.Incremental);
            Assert.Equal(7, o.BatchSize);
        }

        [Theory]
        [InlineData("0")]
        [InlineData("-5")]
        [InlineData("abc")]
        [InlineData("1001")]
        public void Batch_InvalidNumber_IsAnError(string n)
        {
            Assert.Contains("--batch", Fails(@"C:\Export", "--batch", n));
        }

        [Fact]
        public void Batch_MissingNumber_IsAnError()
        {
            Assert.Contains("--batch", Fails(@"C:\Export", "--batch"));
        }

        [Fact]
        public void UnknownOption_IsAnError()
        {
            Assert.Contains("--full", Fails(@"C:\Export", "--full"));
        }

        [Fact]
        public void NoPath_IsAnError()
        {
            Fails("--incremental");
            Fails();
            Fails(null);
        }

        [Fact]
        public void TwoPaths_IsAnError()
        {
            Fails(@"C:\Export", @"C:\Other");
        }

        [Theory]
        [InlineData(@"C:\")]
        [InlineData(@"\\server\share")]
        public void Root_IsRefused(string root)
        {
            Assert.Contains("root", Fails(root, "--incremental"));
        }
    }
}
