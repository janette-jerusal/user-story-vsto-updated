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
            using (Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("Test2.Ribbon1.xml"))
            using (StreamReader reader = new StreamReader(stream))
            {
                return reader.ReadToEnd();
            }
        }

        public void Ribbon_Load(IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        public void OnCompareClicked(IRibbonControl control)
        {
            MessageBox.Show("Compare button clicked!", "User Story Comparator");
        }
    }
}

