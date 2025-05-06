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

            // Getting the active worksheet from Excel
            var sheet = Globals.ThisAddIn.Application.ActiveSheet as Excel.Worksheet;

            
            var range = sheet.UsedRange;

            // Initialize string builder for CSV file contents
            var sb = new StringBuilder();

            // first loop thru row then column
            for (int r =1; r<=range.Rows.Count; r++)
            {
                for (int c = 1; c<= range.Columns.Count; c++)
                {
                    // add cell value to string builder
                    sb.Append(range.Cells[r, c].Value2);

                    // if not the last row, add comma
                    if (c != range.Columns.Count) sb.Append(",");
                }

                // add \n character
                sb.AppendLine();
            }

            //export 
            System.IO.File.WriteAllText(@"C:\Temp\export.csv", sb.ToString());

            //export successful message
            System.Windows.Forms.MessageBox.Show("Success! Exported to C:\\Temp\\export.csv");


            //end of button1 bracket
        }

    }
}
