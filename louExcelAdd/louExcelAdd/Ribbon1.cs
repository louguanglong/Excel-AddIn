using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace louExcelAdd
{
    public partial class Ribbon1
    {
        Excel.Application ExcelApp;
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            ExcelApp = Globals.ThisAddIn.Application;
            ExcelApp.ActiveCell.Value = DateTime.Now;
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            ExcelApp= Globals.ThisAddIn.Application;
            ExcelApp.ActiveCell.Value ="111";
        }
    }
}
