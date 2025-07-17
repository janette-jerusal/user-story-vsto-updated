using System;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using Microsoft.Office.Core;

namespace Test2
{
    public class Ribbon1 : IRibbonExtensibility
    {
        private IRibbonUI ribbon;

        public string GetCustomUI(string ribbonID)
        {
            var asm = Assembly.GetExecutingAssembly();
            using (var stream = asm.GetManifestResourceStream("Test2.Ribbon1.xml"))
            using (var reader = new StreamReader(stream))
                return reader.ReadToEnd();
        }

        public void Ribbon_Load(IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        public void OnCompareClick(IRibbonControl control)
        {
            MessageBox.Show("🎯 Compare button clicked!");
            // TODO: Call your comparison logic here
        }
    }
}

