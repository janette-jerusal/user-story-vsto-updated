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
            var resourceName = "Test2.Ribbon1.xml"; // matches your project + file name

            using (Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName))
            using (StreamReader reader = new StreamReader(stream))
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
            MessageBox.Show("Compare logic goes here!");
        }
    }
}

