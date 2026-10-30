using System.Linq;
using Xunit;

namespace TiaExportBlocks.Tests
{
    public class SourceSplitterTests
    {
        private const string Fb = "FUNCTION_BLOCK \"FB_Main\"\r\n{ S7_Optimized_Access := 'TRUE' }\r\nVERSION : 0.1\r\nBEGIN\r\n\t#x := 1;\r\nEND_FUNCTION_BLOCK\r\n\r\n";
        private const string Fc = "FUNCTION \"FC_Calc\" : Int\r\nVERSION : 0.1\r\nBEGIN\r\n\t#FC_Calc := 2;\r\nEND_FUNCTION\r\n\r\n";
        private const string Ob = "ORGANIZATION_BLOCK \"Main\"\r\nVERSION : 0.1\r\nBEGIN\r\nEND_ORGANIZATION_BLOCK\r\n\r\n";
        private const string Db = "DATA_BLOCK \"DB_Recipe\"\r\nVERSION : 0.1\r\n   VAR\r\n      x : Int;\r\n   END_VAR\r\nBEGIN\r\nEND_DATA_BLOCK\r\n\r\n";
        private const string Udt = "TYPE \"UDT_Motor\"\r\nVERSION : 0.1\r\n   STRUCT\r\n      Speed : Real;\r\n   END_STRUCT;\r\n\r\nEND_TYPE\r\n\r\n";

        [Fact]
        public void Split_AllBlockKinds_ReturnsEachExactly()
        {
            var parts = SourceSplitter.Split(Fb + Fc + Ob + Db + Udt);

            Assert.Equal(new[] { "FB_Main", "FC_Calc", "Main", "DB_Recipe", "UDT_Motor" }, parts.Select(p => p.Name));
            Assert.Equal(new[] { "FUNCTION_BLOCK", "FUNCTION", "ORGANIZATION_BLOCK", "DATA_BLOCK", "TYPE" }, parts.Select(p => p.Keyword));
            Assert.Equal(new[] { Fb, Fc, Ob, Db, Udt }, parts.Select(p => p.Text));
        }

        [Fact]
        public void Split_PartsConcatenateBackToTheInput()
        {
            string text = Db + Fb + Udt;

            Assert.Equal(text, string.Concat(SourceSplitter.Split(text).Select(p => p.Text)));
        }

        [Fact]
        public void Split_SingleBlock_ReturnsItUnchanged()
        {
            Assert.Equal(Fb, SourceSplitter.Split(Fb).Single().Text);
        }

        [Fact]
        public void Split_KeepsUnixLineEndingsAsTheyAre()
        {
            string fb = Fb.Replace("\r\n", "\n");
            string fc = Fc.Replace("\r\n", "\n");

            var parts = SourceSplitter.Split(fb + fc);

            Assert.Equal(new[] { fb, fc }, parts.Select(p => p.Text));
        }

        [Fact]
        public void Split_LastBlockWithoutTrailingNewline_IsKept()
        {
            string fc = "FUNCTION \"FC_1\" : Void\r\nBEGIN\r\nEND_FUNCTION";

            Assert.Equal(new[] { Fb, fc }, SourceSplitter.Split(Fb + fc).Select(p => p.Text));
        }

        [Fact]
        public void Split_EndKeywordWithTrailingSpaces_IsRecognized()
        {
            string fc = "FUNCTION \"FC_1\" : Void\r\nBEGIN\r\nEND_FUNCTION   \r\n\r\n";

            Assert.Equal(2, SourceSplitter.Split(fc + Fb).Count);
        }

        [Fact]
        public void Split_NamesWithSpacesDotsAndUnicode_AreRead()
        {
            string text = "FUNCTION \"Calc v2.1 – Ș\" : Void\r\nBEGIN\r\nEND_FUNCTION\r\n\r\n";

            Assert.Equal("Calc v2.1 – Ș", SourceSplitter.Split(text).Single().Name);
        }

        [Fact]
        public void Split_EndKeywordOfAnotherKindInsideBody_DoesNotEndTheBlock()
        {
            string fb = "FUNCTION_BLOCK \"FB_1\"\r\nBEGIN\r\nEND_FUNCTION\r\nEND_FUNCTION_BLOCK\r\n\r\n";

            Assert.Equal(new[] { fb, Fc }, SourceSplitter.Split(fb + Fc).Select(p => p.Text));
        }

        [Fact]
        public void Split_BlockHeaderLikeLineInsideBody_IsNotANewBlock()
        {
            string fb = "FUNCTION_BLOCK \"FB_1\"\r\nBEGIN\r\nFUNCTION \"x\"\r\nEND_FUNCTION_BLOCK\r\n\r\n";

            Assert.Equal("FB_1", SourceSplitter.Split(fb).Single().Name);
        }

        [Fact]
        public void Split_LeadingBlankLines_AreAllowed()
        {
            Assert.Single(SourceSplitter.Split("\r\n\r\n" + Fb));
        }

        [Theory]
        [InlineData("// header comment\r\n")]
        [InlineData("garbage")]
        public void Split_TextBeforeFirstBlock_ReturnsNull(string prefix)
        {
            Assert.Null(SourceSplitter.Split(prefix + Fb));
        }

        [Fact]
        public void Split_TextBetweenBlocks_ReturnsNull()
        {
            Assert.Null(SourceSplitter.Split(Fb + "unexpected\r\n" + Fc));
        }

        [Fact]
        public void Split_UnterminatedBlock_ReturnsNull()
        {
            Assert.Null(SourceSplitter.Split(Fb + "FUNCTION \"FC_1\" : Void\r\nBEGIN\r\n"));
        }

        [Fact]
        public void Split_DuplicateNames_ReturnsNull()
        {
            Assert.Null(SourceSplitter.Split(Fb + Fb));
        }

        [Fact]
        public void Split_EmptyText_ReturnsNoParts()
        {
            Assert.Empty(SourceSplitter.Split(""));
            Assert.Null(SourceSplitter.Split(null));
        }
    }
}
