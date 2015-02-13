using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop;
using Microsoft.Office.Interop.Excel;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            Microsoft.Office.Interop.Excel.Application excelApp;
            Workbook workbook;
            Worksheet worksheet;
           // object missing = System.Reflection.Missing.Value;

            excelApp = new Microsoft.Office.Interop.Excel.Application();
            workbook = excelApp.Workbooks.Open("C:\\ElevenFiftyPractice\\WindowsFormsApplication1\\Book1.xlsx", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets.get_Item(1);


           // var range = (Microsoft.Office.Interop.Excel.Range)worksheet.get_Range(worksheet.Cells[1, 1],
            //                worksheet.Cells[3, 3]).Value2.ToString();


            var range2 = worksheet.UsedRange;
            var result = "";

            for (var r = 1; r < range2.Rows.Count+1; r++)
            {
                for (var c = 1; c < range2.Columns.Count+1; c++)
                {
                    var cellAsString = "";
                    if((range2.Cells[r, c] as Microsoft.Office.Interop.Excel.Range).Value2 != null)
                    {
                      cellAsString = (string) (range2.Cells[r, c] as Microsoft.Office.Interop.Excel.Range).Value2.ToString();
                    }
                    result += cellAsString;
                    result += " ";
                }
            }

            MessageBox.Show(result);

            workbook.Close();
            excelApp.Quit();

            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                worksheet = null;
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                workbook = null;
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                excelApp = null;
            }
            catch (Exception ex)
            {
                worksheet = null;
                MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }


        }
    }
}
