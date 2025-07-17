using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Core;

namespace Test2
{
    public partial class ThisWorkbook
    {
        protected override IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbon1(); // Match your class name
        }
    }
}

