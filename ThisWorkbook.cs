namespace Test2
{
    public partial class ThisWorkbook
    {
        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbon1();
        }
    }
}

