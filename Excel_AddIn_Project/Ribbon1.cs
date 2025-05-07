using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn2
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void exportCsv_Click(object sender, RibbonControlEventArgs e)
        {
            // getting worksheet
            var sheet = Globals.ThisAddIn.Application.ActiveSheet as Excel.Worksheet;

            var range = sheet.UsedRange;

            //initialize sb
            var sb = new StringBuilder();

            for (int r = 1; r <= range.Rows.Count; r++)
            {
                for (int c = 1; c <= range.Columns.Count; c++)
                {
                    // get cell value (it kept crashing when null)
                    var value = range.Cells[r, c].Value2;
                    sb.Append(value != null ? value.ToString() : "");

                    // add commas to separate columns
                    if (c != range.Columns.Count)
                        sb.Append(",");
                }

                // newline char in data file
                sb.AppendLine();
            }

            System.IO.File.WriteAllText(@"C:\Temp\export.csv", sb.ToString());

            System.Windows.Forms.MessageBox.Show("Success! Exported to C:\\Temp\\export.csv");
        }
        private void AutoFit_Click(object sender, RibbonControlEventArgs e)
        {
            var sheet = Globals.ThisAddIn.Application.ActiveSheet;
            // autofit is built into excel
            // used range is any cell to any cell that has data
            sheet.UsedRange.Columns.Autofit();
            System.Windows.Forms.MessageBox.Show("The colmuns have been auto fitted.");
        }

        private void Clear_Click(object sender, RibbonControlEventArgs e)
        {
            var sheet = Globals.ThisAddIn.Application.ActiveSheet as Excel.Worksheet;
            sheet.Cells.Clear();

            System.Windows.Forms.MessageBox.Show("The data has been cleared.");
        }
    }
}
