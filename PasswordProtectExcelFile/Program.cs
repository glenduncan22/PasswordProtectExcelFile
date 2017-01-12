using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Runtime.InteropServices;

namespace PasswordProtectExcelFile
{
    class PasswordProtectExcelFile
    {
        static void Main(string[] args)
        {
            int i_debug_mode = Int32.Parse(args[0]);
            string s_sourcefile = args[1];
            string s_password = args[2];
            Microsoft.Office.Interop.Excel.Application excelApp = null;
            Microsoft.Office.Interop.Excel.Workbook excelWorkbook = null;
            try
            {
                if (i_debug_mode > 0)
                {
                    Console.WriteLine("The filename is: " + s_sourcefile);
                    Console.WriteLine("Step1 Open the Excel file: " + s_sourcefile);
                }
                //open the Excel application
                excelApp = new Microsoft.Office.Interop.Excel.ApplicationClass();
                excelWorkbook = excelApp.Workbooks.Open(s_sourcefile,
                0,
                false,
                Type.Missing,
                Type.Missing,
                Type.Missing,
                true,
                Microsoft.Office.Interop.Excel.XlPlatform.xlWindows,
                "",
                false,
                false,
                0,
                false,
                false,
                false);
                if (i_debug_mode > 0)
                {
                    Console.WriteLine("Step1 Open the Excel file Completed: " + s_sourcefile);
                    Console.WriteLine("__________________________________");
                    Console.WriteLine("Step2 Save the Excel file: " + s_sourcefile);
                }
                //Save the file with the "s_sourcefile" name
                excelApp.DisplayAlerts = false;
                excelWorkbook.SaveAs(s_sourcefile,
                Type.Missing,
                s_password, //password is used in the saveas function
                Type.Missing,
                Type.Missing,
                Type.Missing,
                Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive,
                Microsoft.Office.Interop.Excel.XlSaveConflictResolution.xlLocalSessionChanges,
                true,
                Type.Missing,
                Type.Missing,
                Type.Missing);
                if (i_debug_mode > 0)
                {
                    Console.WriteLine("Step2 save the Excel file Completed with password: " + s_sourcefile);
                    Console.WriteLine("__________________________________");
                    Console.WriteLine("Completed processing Excel File" + s_sourcefile);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("The error message is : " + e.Message);
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
                //Close the workbook
                excelWorkbook.Close(true, s_sourcefile, false);
                Marshal.FinalReleaseComObject(excelWorkbook);
                excelWorkbook = null;
                //Quit the application
                excelApp.Quit();
                Marshal.FinalReleaseComObject(excelApp);
                excelApp = null;
                //Console.WriteLine("All connections are closed");
            }
        }
    }
}
