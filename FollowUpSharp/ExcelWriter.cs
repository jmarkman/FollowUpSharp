using System;
using System.Collections.Generic;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace FollowUpSharp
{
    public class ExcelWriter
    {
        private ExcelPackage qfuExcel;
        private ExcelWorksheet ws;
        private string filepath = $@"{Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)}\FollowUpSharp\";

        public ExcelWriter()
        {
            qfuExcel = new ExcelPackage();
            qfuExcel.Workbook.Worksheets.Add("Quote Follow Ups");
            ws = qfuExcel.Workbook.Worksheets[1];
            ws.Name = "Records";

            // Value assignment to cells is direct
            ws.Cells["A1"].Value = "Control Number";
            ws.Cells["B1"].Value = "Due Date";
            ws.Cells["C1"].Value = "Effective Date";
            ws.Cells["D1"].Value = "Created Date";
            ws.Cells["E1"].Value = "Broker Name";
            ws.Cells["F1"].Value = "Parent Broker Company";
            ws.Cells["G1"].Value = "Broker First Name";
            ws.Cells["H1"].Value = "Broker Last Name";
            ws.Cells["I1"].Value = "Broker Email";
            ws.Cells["J1"].Value = "Underwriter First Name";
            ws.Cells["K1"].Value = "Underwriter Last Name";
            ws.Cells["L1"].Value = "Underwriter Email";
            ws.Cells["M1"].Value = "Named Insured";

            // Changing font style is a boolean value, while color changing is more direct
            ws.Cells["A1:M1"].Style.Font.Bold = true;
            ws.Cells["A1:M1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["A1:M1"].Style.Fill.BackgroundColor.SetColor(Color.DarkBlue);
            ws.Cells["A1:M1"].Style.Font.Color.SetColor(Color.White);
        }

        /// <summary>
        /// This method, using a list provided as an argument, adds the contents of that list to
        /// the sheet. Is a call to the method chain Cells.LoadFromCollection() provided in EPPlus. 
        /// </summary>
        /// <param name="_results">Accepts a list of IMSEntry objects as an argument.</param>
        public void AddResults(List<IMSEntry> _results)
        {
            try
            {
                ws.Cells["A2"].LoadFromCollection(_results);
                ws.Cells["A1:M1"].AutoFitColumns();
            }
            catch (Exception ex)
            {
                DateTime theDate = DateTime.Now;
                MessageBox.Show("The list passed to the method does not contain any insureds!");
                using (StreamWriter excelExceptionWriter = new StreamWriter(filepath + "ErrorLog.txt", true))
                {
                    excelExceptionWriter.WriteLine(theDate.ToString("MM-dd-yyyy"));
                    excelExceptionWriter.WriteLine(ex.GetType().Name);
                    excelExceptionWriter.WriteLine(ex.Message);
                }
            }

        }

        /// <summary>
        /// Saves the worksheet to a specified direcetory since EPPlus works with ExcelWriter in memory.
        /// </summary>
        public void SaveWS()
        {
            DateTime theDate = DateTime.Now; // Instantiate a datetime object to get today's date
            string today = theDate.ToString("MM-dd-yyyy"); // Get today's date as a string in the format month-day-year

            string excelWorkbookName = $"Follow Ups for {today}.xlsx";
            /*
             * Capture the ExcelWriter file currently being worked with in memory and store it in a byte array
             * so it can be saved to the user's computer
             */
            Byte[] sheetAsBinary = qfuExcel.GetAsByteArray();
            try
            {
                using (qfuExcel)
                {
                    File.WriteAllBytes(Path.Combine(filepath + @"\Quote Follow Ups Archive", excelWorkbookName), sheetAsBinary);
                }
            }
            catch (DirectoryNotFoundException)
            {
                using (qfuExcel)
                {
                    Directory.CreateDirectory($@"{Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)}\FollowUpSharp\Quote Follow Ups Archive\");
                    File.WriteAllBytes(Path.Combine(filepath + @"\Quote Follow Ups Archive", excelWorkbookName), sheetAsBinary);
                }
            }
            catch (IOException io)
            {
                MessageBox.Show("The list passed to the method does not contain any insureds!");
                using (StreamWriter excelExceptionWriter = new StreamWriter(filepath + "ErrorLog.txt", true))
                {
                    excelExceptionWriter.WriteLine(theDate.ToString("MM-dd-yyyy"));
                    excelExceptionWriter.WriteLine(io.GetType().Name);
                    excelExceptionWriter.WriteLine(io.Message + "\n");
                }
            }
            qfuExcel = null;
            ws = null;
        }
    }
}