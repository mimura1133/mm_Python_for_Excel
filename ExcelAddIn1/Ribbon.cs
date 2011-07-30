using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace ExcelAddIn1
{
    public partial class Ribbon
    {
        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            new Form1().Show();
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            ThisAddIn.Run();
        }
    }
}
