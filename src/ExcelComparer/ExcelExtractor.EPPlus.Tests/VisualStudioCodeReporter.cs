using ApprovalTests.Reporters;
using DiffEngine;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelExtractor.EPPlus.Tests;

public class VisualStudioCodeReporter : DiffToolReporter
{
    public static readonly VisualStudioCodeReporter INSTANCE = new();

    public VisualStudioCodeReporter() : base(DiffTool.VisualStudioCode)
    {
    }
}
