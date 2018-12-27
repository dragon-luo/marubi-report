using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace Marubi
{
    public class ExcelUtils
    {
        public static void Export(System.Data.DataTable dt, string rptFilePath, string workSheetName)
        {             
            Microsoft.Office.Interop.Excel.Application xlsApp = new Microsoft.Office.Interop.Excel.Application();
            if (xlsApp == null) return;            
            xlsApp.Visible = false;
            

            Microsoft.Office.Interop.Excel.Workbooks workBooks = xlsApp.Workbooks;
            Microsoft.Office.Interop.Excel.Workbook workBook = workBooks.Add(Microsoft.Office.Interop.Excel.XlWBATemplate.xlWBATWorksheet);
            Microsoft.Office.Interop.Excel.Worksheet workSheet = (Microsoft.Office.Interop.Excel.Worksheet)workBook.Worksheets[1]; ;
            workSheet.Name = workSheetName;
            Microsoft.Office.Interop.Excel.Range range;            
            xlsApp.DisplayAlerts = false;
            xlsApp.Columns["B:B"].ColumnWidth = 24;
            xlsApp.Columns["E:E"].ColumnWidth = 10;
            xlsApp.Columns["F:F"].ColumnWidth = 40;
            xlsApp.Columns["G:G"].ColumnWidth = 9;

            for (int i = 0; i < dt.Columns.Count; i++)
            {
                workSheet.Cells[1, i + 1] = dt.Columns[i].Caption;
                range = (Microsoft.Office.Interop.Excel.Range)workSheet.Cells[1, i + 1];
                //range = xlsApp.Range[xlsApp.Cells[1, 1], xlsApp.Cells[1, dt.Columns.Count]];
                range.Font.FontStyle = "Times New Roman";
                range.Font.Bold = true;
                range.Font.Size = 10;
                range.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                range.VerticalAlignment = XlVAlign.xlVAlignCenter;
            }
            
            string tempId = "";
            int row = 1;

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                string productId = "";

                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    string fieldValue = dt.Rows[i][j].ToString();
                    if (dt.Columns[j].ColumnName == "包材库存" && string.IsNullOrEmpty(fieldValue))                                            
                        workSheet.Cells[i + 2, j + 1] = "#N/A";                        
                    else
                        workSheet.Cells[i + 2, j + 1] = fieldValue;

                    range = (Microsoft.Office.Interop.Excel.Range)workSheet.Cells[i + 2, j + 1];
                    //range = xlsApp.Range[xlsApp.Cells[2, 1], xlsApp.Cells[dt.Rows.Count + 1, dt.Columns.Count]];
                    range.Font.FontStyle = "Times New Roman";
                    range.Font.Bold = false;
                    range.Font.Size = 9;

                    productId = dt.Rows[i][0].ToString();

                    if (i == 0) 
                    {
                        tempId = productId;
                        row = i + 2;              
                    }

                    if (productId != tempId || i+ 1 == dt.Rows.Count)
                    {
                        range = workSheet.Range[workSheet.Cells[row, 1], workSheet.Cells[i + 1, 1]];
                        range.MergeCells = true;
                        range = workSheet.Range[workSheet.Cells[row, 2], workSheet.Cells[i + 1, 2]];
                        range.MergeCells = true;
                        range = workSheet.Range[workSheet.Cells[row, 3], workSheet.Cells[i + 1, 3]];
                        range.MergeCells = true;
                        range = workSheet.Range[workSheet.Cells[row, 4], workSheet.Cells[i + 1, 4]];
                        range.MergeCells = true;

                        tempId = productId;
                        row = i + 2;                        
                    }
                }
            }

            System.Reflection.Missing miss = System.Reflection.Missing.Value;

            workSheet.SaveAs(rptFilePath, miss, miss, miss, miss, miss, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, miss, miss, miss);
            workBook.Close(false, miss, miss);
            workBooks.Close();
            xlsApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workSheet);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workBook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workBooks);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlsApp);
      
            GC.Collect();
        }
   
        public static string CheckOfficeVersion()
        {

            string version = "";

            Microsoft.Win32.RegistryKey registerKey = Microsoft.Win32.Registry.LocalMachine;
            Microsoft.Win32.RegistryKey xls2003Key = registerKey.OpenSubKey(@"SOFTWARE\\Microsoft\\Office\\11.0\\Word\\InstallRoot\\");
            Microsoft.Win32.RegistryKey xls2007Key = registerKey.OpenSubKey(@"SOFTWARE\\Microsoft\\Office\\12.0\\Word\\InstallRoot\\");
            Microsoft.Win32.RegistryKey xls2010Key = registerKey.OpenSubKey(@"SOFTWARE\\Microsoft\\Office\\14.0\\Word\\InstallRoot\\");
            
            //检查本机是否安装 Office2003
            if (xls2003Key != null)
            {
                string xls2003File = xls2003Key.GetValue("Path").ToString();
                if (System.IO.File.Exists(xls2003File + "Excel.exe"))
                {
                    version = "2003";
                }
            }

            //检查本机是否安装 Office2007
            if (xls2003Key != null)
            {
                string xls2007File = xls2007Key.GetValue("Path").ToString();
                if (System.IO.File.Exists(xls2007File + "Excel.exe"))
                {
                    version = "2007";
                }
            }

            //检查本机是否安装 Office2010
            if (xls2003Key != null)
            {
                string xls2010File = xls2010Key.GetValue("Path").ToString();
                if (System.IO.File.Exists(xls2010File + "Excel.exe"))
                {
                    version = "2010";
                }
            }

            return version;
        }
    }
}
