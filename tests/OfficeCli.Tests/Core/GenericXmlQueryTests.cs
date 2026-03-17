using System;
using OfficeCli.Core;
using Xunit;

namespace OfficeCli.Tests.Core;

public class GenericXmlQueryTests
{
    [Fact]
    public void ParsePathSegments_NullPath_ThrowsNullReferenceException_Bug()
    {
        // This confirms the bug where missing null check on 'path' 
        // results in NullReferenceException due to .Trim() call.
        Assert.Throws<NullReferenceException>(() => GenericXmlQuery.ParsePathSegments(null));
    }
}
