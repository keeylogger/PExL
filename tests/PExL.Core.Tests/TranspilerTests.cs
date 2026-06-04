using System.Linq;
using PExL.Core;
using Xunit;

namespace PExL.Core.Tests
{
    public class TranspilerTests
    {
        [Theory]
        // pipes & text
        [InlineData("B2 |> trim |> upper", "=UPPER(TRIM(B2))")]
        [InlineData("trim(B2)", "=TRIM(B2)")]
        [InlineData("B2 |> split.First(\"-\") |> fromLeft", "=TEXTBEFORE(B2,\"-\")")]
        [InlineData("B2 |> split.First(\"-\") |> fromRight", "=TEXTAFTER(B2,\"-\")")]
        [InlineData("B2 |> split.Last(\"-\") |> fromLeft", "=TEXTBEFORE(B2,\"-\",-1)")]
        [InlineData("B2 |> split.Last(\"-\") |> fromRight", "=TEXTAFTER(B2,\"-\",-1)")]
        [InlineData("B2 |> split(\"-\") |> fromLeft", "=TEXTBEFORE(B2,\"-\")")]
        [InlineData("B2 |> split(\"-\") |> fromRight", "=TEXTAFTER(B2,\"-\",-1)")]
        [InlineData("B2 |> split(\"-\") |> at(2)", "=INDEX(TEXTSPLIT(B2,\"-\"),2)")]
        [InlineData("B2 |> split(\"-\") |> spill", "=TEXTSPLIT(B2,\"-\")")]
        [InlineData("combine(A1, B1, C1) with(\"-\")", "=TEXTJOIN(\"-\",TRUE,A1,B1,C1)")]
        [InlineData("replace(B2,\"-\",\"_\")", "=SUBSTITUTE(B2,\"-\",\"_\")")]
        [InlineData("replace.first(B2,\"-\",\"_\")", "=SUBSTITUTE(B2,\"-\",\"_\",1)")]
        [InlineData("upper(B2)", "=UPPER(B2)")]
        [InlineData("proper(B2)", "=PROPER(B2)")]
        [InlineData("length(B2)", "=LEN(B2)")]
        [InlineData("startsWith(B2,\"USD\")", "=LEFT(B2,3)=\"USD\"")]
        [InlineData("endsWith(B2,\".csv\")", "=RIGHT(B2,4)=\".csv\"")]
        // contains
        [InlineData("contains(B2,\"x\")", "=ISNUMBER(SEARCH(\"x\",B2))")]
        [InlineData("B2 contains \"x\"", "=ISNUMBER(SEARCH(\"x\",B2))")]
        // lookup
        [InlineData("find D2 within A1:A100 thenReturn B1:B100", "=XLOOKUP(D2,A1:A100,B1:B100)")]
        [InlineData("find D2 within A1:A100 thenReturn B1:B100 ifMissing \"N/A\"", "=XLOOKUP(D2,A1:A100,B1:B100,\"N/A\")")]
        [InlineData("D2 |> find(A1:A100, B1:B100)", "=XLOOKUP(D2,A1:A100,B1:B100)")]
        [InlineData("position D2 within A1:A100", "=MATCH(D2,A1:A100,0)")]
        // logic
        [InlineData("if B2 > 10 then \"High\" else \"Low\"", "=IF(B2>10,\"High\",\"Low\")")]
        [InlineData("A1 > 0 and B1 > 0", "=AND(A1>0,B1>0)")]
        [InlineData("not contains(B2,\"x\")", "=NOT(ISNUMBER(SEARCH(\"x\",B2)))")]
        // aggregation
        [InlineData("sum(A1:A10)", "=SUM(A1:A10)")]
        [InlineData("count(A1:A10)", "=COUNTA(A1:A10)")]
        [InlineData("countNum(A1:A10)", "=COUNT(A1:A10)")]
        [InlineData("sumWhere(A1:A10, B1:B10 > 100)", "=SUMIFS(A1:A10,B1:B10,\">100\")")]
        [InlineData("sumWhere(A1:A10, B1:B10 > 100 and C1:C10 = \"West\")", "=SUMIFS(A1:A10,B1:B10,\">100\",C1:C10,\"West\")")]
        [InlineData("countWhere(B1:B10 = \"West\")", "=COUNTIFS(B1:B10,\"West\")")]
        [InlineData("sum.ignoreErrors(A1:A10)", "=AGGREGATE(9,6,A1:A10)")]
        // dates / math
        [InlineData("addMonths(A2, 3)", "=EDATE(A2,3)")]
        [InlineData("addYears(A2, 1)", "=EDATE(A2,12)")]
        [InlineData("yearOf(A2)", "=YEAR(A2)")]
        [InlineData("round(A2, 2)", "=ROUND(A2,2)")]
        [InlineData("round.up(A2, 2)", "=ROUNDUP(A2,2)")]
        [InlineData("#2024-01-01#", "=DATE(2024,1,1)")]
        [InlineData("Date(\"2024-01-01\")", "=DATE(2024,1,1)")]
        // filter / shape
        [InlineData("filter(A1:C100) where(B1:B100 > 100)", "=FILTER(A1:C100,B1:B100>100)")]
        [InlineData("sort(A1:C100) by(2) descending", "=SORT(A1:C100,2,-1)")]
        [InlineData("unique(A1:A100)", "=UNIQUE(A1:A100)")]
        [InlineData("take(A1:A100, 5)", "=TAKE(A1:A100,5)")]
        // references
        [InlineData("col(\"A\")", "=A:A")]
        [InlineData("cell(\"C\", 2)", "=C2")]
        [InlineData("fixed(A1)", "=$A$1")]
        // escape hatches
        [InlineData("raw(\"SUMPRODUCT\", A1:A10, B1:B10)", "=SUMPRODUCT(A1:A10,B1:B10)")]
        [InlineData("legacy.vlookup(D2, A1:B100, 2)", "=VLOOKUP(D2,A1:B100,2,FALSE)")]
        public void Transpiles(string pexl, string expected)
        {
            Assert.Equal(expected, Transpiler.ToFormula(pexl));
        }

        [Fact]
        public void Split_Bind_FansOutToTwoCells()
        {
            var src = "B2 |> split.First(\"-\") :: parts\nparts |> fromLeft -> C2\nparts |> fromRight -> D2";
            var result = Transpiler.Transpile(src);

            Assert.Equal(2, result.Cells.Count);
            Assert.Equal("C2", result.Cells[0].Target);
            Assert.Equal("=TEXTBEFORE(B2,\"-\")", result.Cells[0].Formula);
            Assert.Equal("D2", result.Cells[1].Target);
            Assert.Equal("=TEXTAFTER(B2,\"-\")", result.Cells[1].Formula);
        }

        [Fact]
        public void Check_Block_EmitsIfs()
        {
            var src = "check B2:\n  if > 90 then \"A\"\n  if > 80 then \"B\"\n  else \"C\"";
            Assert.Equal("=IFS(B2>90,\"A\",B2>80,\"B\",TRUE,\"C\")", Transpiler.ToFormula(src));
        }

        [Fact]
        public void Check_Block_PlainForm_NoIfNoColon_WithTarget()
        {
            var src = "check\n  B2 >= 90 then \"A\"\n  B2 >= 80 then \"B\"\n  else \"F\"\n-> C2";
            Assert.Equal("=IFS(B2>=90,\"A\",B2>=80,\"B\",TRUE,\"F\")", Transpiler.ToFormula(src));
        }

        [Fact]
        public void Check_Block_SubjectNoColon_ImplicitOperator()
        {
            var src = "check B2\n  >= 90 then \"A\"\n  >= 80 then \"B\"\n  else \"F\"";
            Assert.Equal("=IFS(B2>=90,\"A\",B2>=80,\"B\",TRUE,\"F\")", Transpiler.ToFormula(src));
        }


        [Fact]
        public void IfError_WrapsPipedExpression()
        {
            var src = "find D2 within A1:A100 thenReturn B1:B100 |> ifError(\"not found\")";
            Assert.Equal("=IFERROR(XLOOKUP(D2,A1:A100,B1:B100),\"not found\")", Transpiler.ToFormula(src));
        }

        [Fact]
        public void Output_Target_IsCaptured()
        {
            var result = Transpiler.Transpile("B2 |> upper -> C2");
            Assert.Single(result.Cells);
            Assert.Equal("C2", result.Cells[0].Target);
            Assert.Equal("=UPPER(B2)", result.Cells[0].Formula);
        }

        [Fact]
        public void Comments_AreIgnored()
        {
            var src = "// clean it up\nB2 |> trim -> C2";
            var result = Transpiler.Transpile(src);
            Assert.Single(result.Cells);
            Assert.Equal("=TRIM(B2)", result.Cells[0].Formula);
        }

        // ---- globals (MakeGlobal) ----

        [Fact]
        public void MakeGlobal_Constant_ProducesNamedDefinition()
        {
            var result = Transpiler.Transpile("MakeGlobal(0.2) :: TaxRate");
            Assert.Empty(result.Cells);
            Assert.Single(result.Globals);
            Assert.Equal("TaxRate", result.Globals[0].Name);
            Assert.Equal("=0.2", result.Globals[0].Formula);
        }

        [Fact]
        public void MakeGlobal_Range_ProducesNamedDefinition()
        {
            var result = Transpiler.Transpile("MakeGlobal(A2:A100) :: SalesQ1");
            Assert.Single(result.Globals);
            Assert.Equal("SalesQ1", result.Globals[0].Name);
            Assert.Equal("=$A$2:$A$100", result.Globals[0].Formula);
        }

        [Fact]
        public void Global_IsEmittedVerbatim_WhenReferenced()
        {
            var src = "MakeGlobal(0.2) :: TaxRate\nA1 * TaxRate -> B1";
            var result = Transpiler.Transpile(src);
            Assert.Single(result.Globals);
            Assert.Single(result.Cells);
            Assert.Equal("B1", result.Cells[0].Target);
            Assert.Equal("=A1*TaxRate", result.Cells[0].Formula);
        }

        [Fact]
        public void Global_CanReferenceEarlierGlobals()
        {
            var src = "MakeGlobal(B1) :: Revenue\nMakeGlobal(C1) :: Cost\nMakeGlobal(Revenue - Cost) :: Margin";
            var result = Transpiler.Transpile(src);
            Assert.Equal(3, result.Globals.Count);
            Assert.Equal("Margin", result.Globals[2].Name);
            Assert.Equal("=Revenue-Cost", result.Globals[2].Formula);
        }

        [Fact]
        public void MakeGlobal_WithoutName_Throws()
        {
            Assert.Throws<PExL.Core.Diagnostics.PExLException>(() => Transpiler.Transpile("MakeGlobal(0.2)"));
        }

        [Fact]
        public void GlobalSynonym_Works()
        {
            var result = Transpiler.Transpile("global(0.2) :: TaxRate");
            Assert.Single(result.Globals);
            Assert.Equal("TaxRate", result.Globals[0].Name);
        }

        // ---- console commands (ShowGlobals) ----

        [Fact]
        public void ShowGlobals_ProducesCommand_NotFormula()
        {
            var result = Transpiler.Transpile("ShowGlobals()");
            Assert.Empty(result.Cells);
            Assert.Single(result.Commands);
            Assert.Equal("showGlobals", result.Commands[0].Name);
        }
    }
}
