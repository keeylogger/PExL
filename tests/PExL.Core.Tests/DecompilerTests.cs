using PExL.Core;
using Xunit;

namespace PExL.Core.Tests
{
    public class DecompilerTests
    {
        // Excel formula -> PExL -> Excel should land back on an equivalent formula.
        [Theory]
        [InlineData("=TRIM(B2)")]
        [InlineData("=UPPER(TRIM(B2))")]
        [InlineData("=PROPER(TRIM(CLEAN(SUBSTITUTE(A2,\"_\",\" \"))))")]
        [InlineData("=LEN(B2)")]
        [InlineData("=SUBSTITUTE(B2,\"-\",\"_\")")]
        [InlineData("=SUBSTITUTE(B2,\"-\",\"_\",1)")]
        [InlineData("=TEXTBEFORE(B2,\"-\")")]
        [InlineData("=TEXTAFTER(B2,\"-\")")]
        [InlineData("=TEXTBEFORE(B2,\"-\",-1)")]
        [InlineData("=TEXTAFTER(B2,\"-\",-1)")]
        [InlineData("=TEXTSPLIT(B2,\"-\")")]
        [InlineData("=INDEX(TEXTSPLIT(B2,\"-\"),2)")]
        [InlineData("=TEXTJOIN(\"-\",TRUE,A1,B1,C1)")]
        [InlineData("=ISNUMBER(SEARCH(\"x\",B2))")]
        [InlineData("=LEFT(B2,3)=\"USD\"")]
        [InlineData("=RIGHT(B2,4)=\".csv\"")]
        [InlineData("=XLOOKUP(D2,A1:A100,B1:B100)")]
        [InlineData("=XLOOKUP(D2,A1:A100,B1:B100,\"N/A\")")]
        [InlineData("=MATCH(D2,A1:A100,0)")]
        [InlineData("=IF(B2>10,\"High\",\"Low\")")]
        [InlineData("=IFS(B2>90,\"A\",B2>80,\"B\",TRUE,\"C\")")]
        [InlineData("=IFERROR(XLOOKUP(D2,A1:A100,B1:B100),\"not found\")")]
        [InlineData("=SUM(A1:A10)")]
        [InlineData("=COUNTA(A1:A10)")]
        [InlineData("=COUNT(A1:A10)")]
        [InlineData("=SUMIFS(A1:A10,B1:B10,\">100\")")]
        [InlineData("=SUMIFS(A1:A10,B1:B10,\">100\",C1:C10,\"West\")")]
        [InlineData("=COUNTIFS(B1:B10,\"West\")")]
        [InlineData("=AGGREGATE(9,6,A1:A10)")]
        [InlineData("=EDATE(A2,3)")]
        [InlineData("=YEAR(A2)")]
        [InlineData("=ROUND(A2,2)")]
        [InlineData("=ROUNDUP(A2,2)")]
        [InlineData("=DATE(2024,1,1)")]
        [InlineData("=FILTER(A1:C100,B1:B100>100)")]
        [InlineData("=SORT(A1:C100,2,-1)")]
        [InlineData("=UNIQUE(A1:A100)")]
        [InlineData("=TAKE(A1:A100,5)")]
        [InlineData("=VLOOKUP(D2,A1:B100,2,FALSE)")]
        [InlineData("=SUMPRODUCT(A1:A10,B1:B10)")]
        [InlineData("=A1+B1*2")]
        [InlineData("=(A1+B1)*2")]
        [InlineData("=A1>0")]
        public void RoundTrips(string formula)
        {
            string pexl = Decompiler.ToPExL(formula);
            string back = Transpiler.ToFormula(pexl);
            Assert.Equal(formula, back);
        }

        [Theory]
        [InlineData("=TRIM(B2)", "trim(B2)")]
        [InlineData("=XLOOKUP(D2,A1:A100,B1:B100,\"N/A\")", "find D2 within A1:A100 thenReturn B1:B100 ifMissing \"N/A\"")]
        [InlineData("=ISNUMBER(SEARCH(\"x\",B2))", "contains(B2, \"x\")")]
        [InlineData("=A1&B1&C1", "combine(A1, B1, C1)")]
        public void ProducesReadablePExL(string formula, string expected)
        {
            Assert.Equal(expected, Decompiler.ToPExL(formula));
        }

        [Fact]
        public void AppendsTargetWhenGiven()
        {
            Assert.Equal("trim(B2) -> C2", Decompiler.ToPExL("=TRIM(B2)", "C2"));
        }

        [Fact]
        public void NestedIfBecomesCheckBlock()
        {
            string pexl = Decompiler.ToPExL("=IF(B2>=90,\"A\",IF(B2>=80,\"B\",IF(B2>=70,\"C\",IF(B2>=60,\"D\",\"F\"))))");
            string expected =
                "check\n" +
                "  B2 >= 90 then \"A\"\n" +
                "  B2 >= 80 then \"B\"\n" +
                "  B2 >= 70 then \"C\"\n" +
                "  B2 >= 60 then \"D\"\n" +
                "  else \"F\"";
            Assert.Equal(expected, pexl);
        }

        [Fact]
        public void NestedIfCheckBlockRecompilesToEquivalentIfs()
        {
            string formula = "=IF(B2>=90,\"A\",IF(B2>=80,\"B\",IF(B2>=70,\"C\",IF(B2>=60,\"D\",\"F\"))))";
            string pexl = Decompiler.ToPExL(formula);
            string back = Transpiler.ToFormula(pexl);
            Assert.Equal("=IFS(B2>=90,\"A\",B2>=80,\"B\",B2>=70,\"C\",B2>=60,\"D\",TRUE,\"F\")", back);
        }

        [Fact]
        public void CheckBlockTargetGoesOnNewLine()
        {
            string pexl = Decompiler.ToPExL("=IF(B2>=90,\"A\",IF(B2>=80,\"B\",\"C\"))", "C2");
            Assert.EndsWith("\n-> C2", pexl);
            // and it still recompiles to the active cell target
            Assert.Equal("=IFS(B2>=90,\"A\",B2>=80,\"B\",TRUE,\"C\")", Transpiler.ToFormula(pexl));
        }

        [Fact]
        public void EmptyFormulaIsCommented()
        {
            Assert.StartsWith("//", Decompiler.ToPExL(""));
        }
    }
}
