using System;
using DocumentFormat.OpenXml;
using OfficeCli.Core;
using Xunit;

namespace OfficeCli.Tests.Core;

public class FormulaParserBugTests
{
    [Fact]
    public void Parse_NullInput_ThrowsNullReferenceException_Bug()
    {
        Assert.Throws<NullReferenceException>(() => FormulaParser.Parse(null));
    }

    [Fact]
    public void ToLatex_NullInput_ThrowsNullReferenceException_Bug()
    {
        Assert.Throws<NullReferenceException>(() => FormulaParser.ToLatex(null));
    }
}
