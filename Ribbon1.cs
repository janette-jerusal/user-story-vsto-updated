using Microsoft.Office.Core;
using System;
using System.Reflection;
using System.Runtime.InteropServices;

namespace Test2
{
    [ComVisible(true)]
    public class Ribbon1 : IRibbonExtensibility
    {
        private IRibbonUI ribbon;

        public string GetCustomUI(string ribbonID)
        {
            var assembly = Assembly.GetExecutingAssembly();
            using (var stream = assembly.GetManifestResourceStream("Test2.Ribbon1.xml"))
            using (var reader = new System.IO.StreamReader(stream))
            {
                return reader.ReadToEnd();
            }
        }

        public void Ribbon_Load(IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        public void OnCompareClick(IRibbonControl control)
        {
            System.Windows.Forms.MessageBox.Show("Compare button clicked!");
        }
    }
}
