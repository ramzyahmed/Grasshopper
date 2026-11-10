using Autodesk.Revit.DB;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;


namespace BIM_Integration
{

   public class ExcelUtils
    {

        public static List<Frame> getAllFrames(ref List<Beam> beams, ref List<Column> columns, Microsoft.Office.Interop.Excel.Worksheet ws)
        {
            List<Frame> allFrames = new List<Frame>();
            #region get_the_cells
            Excel.Range ID = ws.Range["A7"];
            Excel.Range Beta = ws.Range["T7"];
            Excel.Range count = ws.Range["F5"];

            #endregion

            int framesCount = Convert.ToInt32(count.Value);
            int numberColumns = 20;

            object[,] dataFromExcel = new object[framesCount, numberColumns]; ;

            Microsoft.Office.Interop.Excel.Range startCell = (Range)ws.Cells[7, 1];
            Microsoft.Office.Interop.Excel.Range endCell = (Range)ws.Cells[framesCount + 7 - 1, numberColumns];
            Microsoft.Office.Interop.Excel.Range range = (Range)ws.Range[startCell, endCell];

            dataFromExcel = (object[,])range.Value2;


            #region all Frames' data

            for (int i = 1; i <= framesCount; i++)
            {
                Frame F = new Frame();

                try
                {

                    F.ID = (Convert.ToString(dataFromExcel[i, 1]));
                    F.StartPoint.Id = (Convert.ToString(dataFromExcel[i, 2]));
                    F.EndPoint.Id = (Convert.ToString(dataFromExcel[i, 3]));
                    F.StartPoint.X = (Convert.ToDouble(dataFromExcel[i, 4]));
                    F.StartPoint.Y = (Convert.ToDouble(dataFromExcel[i, 5]));
                    F.StartPoint.Z = (Convert.ToDouble(dataFromExcel[i, 6]));

                    F.EndPoint.X = (Convert.ToDouble(dataFromExcel[i, 7]));
                    F.EndPoint.Y = (Convert.ToDouble(dataFromExcel[i, 8]));
                    F.EndPoint.Z = (Convert.ToDouble(dataFromExcel[i, 9]));


                    F.MidPoint.X = (F.StartPoint.X + F.EndPoint.X) / 2;
                    F.MidPoint.Y = (F.StartPoint.Y + F.EndPoint.Y) / 2;
                    F.MidPoint.Z = (F.StartPoint.Z + F.EndPoint.Z) / 2;


                    F.Sec = (Convert.ToString(dataFromExcel[i, 10]));

                    F.Lib = (Convert.ToString(dataFromExcel[i, 11]));
                    F.Shape = (Convert.ToString(dataFromExcel[i, 12]));

                    F.IsBUS = (Convert.ToBoolean(dataFromExcel[i, 13]));
                    F.IsCable = (Convert.ToBoolean(dataFromExcel[i, 14]));
                    F.IsLink = (Convert.ToBoolean(dataFromExcel[i, 15]));
                    F.IsConc = (Convert.ToBoolean(dataFromExcel[i, 16]));
                    F.Group = (Convert.ToString(dataFromExcel[i, 17]));

                    F.Material = (Convert.ToString(dataFromExcel[i, 18]));
                    F.Length = Point.getDistance(F.StartPoint, F.EndPoint);


                    F.rowNo = ID.Offset[i, 0].Row;


                    Point.getMemberType(F);
                    allFrames.Add(F);


                    if (F.MemberType == MemberType.Column || F.MemberType == MemberType.SlanttedColumn)
                    {
                        Column CL = new Column(F.ID, F.StartPoint, F.EndPoint, F.MidPoint, F.Beta, F.Shape, F.SecType, F.Sec, F.IsBUS, F.IsCable, F.IsLink, F.IsConc, F.RvtName, F.Lib, F.Material, F.Group, F.Length, F.MemberType, F.rowNo);
                        columns.Add(CL);
                    }

                    else
                    {
                        Beam BM = new Beam(F.ID, F.StartPoint, F.EndPoint, F.MidPoint, F.Beta, F.Shape, F.SecType, F.Sec, F.IsBUS, F.IsCable, F.IsLink, F.IsConc, F.RvtName, F.Lib, F.Material, F.Group, F.Length, F.MemberType, F.rowNo);
                        beams.Add(BM);
                    }




                    if (F.MemberType == MemberType.SlanttedColumn || F.MemberType == MemberType.Beam || F.MemberType == MemberType.VLBracing)
                    {
                        F.Beta = -(Convert.ToDouble(dataFromExcel[i, 20]));
                    }


                    else
                    {

                        F.Beta = (Convert.ToDouble(dataFromExcel[i, 20]) + 90.00);
                    }

                }

                catch (Exception ex2)
                {
                    System.Diagnostics.StackTrace trace = new System.Diagnostics.StackTrace(ex2, true);
                    RevitUtils.Log.Add("Error in getAllFrames Method in the ExcelUtils ... Frame no " + F.ID + ex2.Message + trace.GetFrame(trace.FrameCount - 1).GetFileLineNumber());
                    continue;
                }
            }



            #endregion

            return allFrames;


        }

        public static int CountCharsUsingForeach(string source, char toFind)
        {
            int count = 0;
            foreach (var ch in source)
            {
                if (ch == toFind)
                    count++;
            }
            return count;
        }


        public static void getExcel(string Type, ref Microsoft.Office.Interop.Excel.Application excel, ref Microsoft.Office.Interop.Excel._Workbook wb, ref Microsoft.Office.Interop.Excel._Worksheet ws)

        {
            string username = Environment.UserName;
            if (Type == "Staad")
            {
                string main_path = "C:\\Users\\" + username + "\\Desktop\\H Y D R A _temp\\ExcDataWithRvt\\Std To RVT.xlsx";

                excel = new Excel.Application();
                excel.DisplayAlerts = false;
                wb = (Excel.Workbook)excel.Workbooks.Open(main_path, System.Type.Missing, false);
                if (wb.ReadOnly)
                {
                    wb.Close(false);
                    wb = excel.Workbooks.Open(main_path, System.Type.Missing, false);
                    wb.SaveAs(main_path);
                }
                ws = (Excel.Worksheet)wb.Worksheets["Staad"];

            }

            else if (Type == "Sap")
            {
                string main_path = "C:\\Users\\" + username + "\\Desktop\\H Y D R A _temp\\ExcDataWithRvt\\Sap To RVT.xlsx";

                excel = new Excel.Application();
                excel.DisplayAlerts = false;
                wb = (Excel.Workbook)excel.Workbooks.Open(main_path, ReadOnly: false);
                ws = (Excel.Worksheet)wb.Worksheets["Sap"];

            }

            else if (Type == "Etabs")
            {
                string main_path = "C:\\Users\\" + username + "\\Desktop\\H Y D R A _temp\\ExcDataWithRvt\\Etabs To RVT.xlsx";

                excel = new Excel.Application();
                excel.DisplayAlerts = false;
                wb = (Excel.Workbook)excel.Workbooks.Open(main_path, ReadOnly: false);
                ws = (Excel.Worksheet)wb.Worksheets["Etabs"];
            }


            excel.Visible = false;
            return;


        }


        public static void CreateExcel(ref Microsoft.Office.Interop.Excel.Application app, ref Microsoft.Office.Interop.Excel._Workbook wb, ref Microsoft.Office.Interop.Excel._Worksheet ws)

        {

            string fileName = "Std To RVT.xlsx";
            string fileName2 = "Sap To RVT.xlsx";

            // Get all processes with the specified file name
            //Process[] processes = Process.GetProcessesByName("EXCEL").Where(p => p.MainWindowTitle.Contains(fileName)).ToArray();
            Process[] processes = Process.GetProcessesByName("EXCEL").ToArray();

            // Close all processes
            foreach (Process process in processes)
            {
                try
                {
                    string filePath = process.MainModule.FileName;
                    string processFileName = Path.GetFileName(filePath);

                    if (processFileName.Equals(fileName, StringComparison.InvariantCultureIgnoreCase) || processFileName.Equals(fileName2, StringComparison.InvariantCultureIgnoreCase))
                    {
                        process.Kill();
                    }
                }
                catch (Exception ex)
                {
                    // Handle any exceptions that might occur while trying to access the process information

                }
            }


            #region get_Excel_object   

            // creating Excel Application  
            app = new Microsoft.Office.Interop.Excel.Application();
            // creating new WorkBook within Excel application  
            wb = app.Workbooks.Add(System.Type.Missing);
            // creating new Excelsheet in workbook  
            ws = null;
            //worksheet.Name = "ExcDataWithRvt";
            // see the excel sheet behind the program  
            app.Visible = false;
            // get the reference of first sheet. By default its name is Sheet1.  
            // store its reference to worksheet  
            ws = (_Worksheet)wb.Sheets["Sheet1"];
            ws.Name = "Staad";

            int a = 6;
            ws.Cells[a - 1, "A"] = "UserUnits";
            //worksheet1.Cells[a - 1, "B"] = currentunits.ToString();
            ws.Cells[a - 1, "C"] = "ExportUnits";
            ws.Cells[a - 1, "D"] = "kN_mm_C";
            ws.Cells[a - 1, "E"] = "CountNumber";

            ws.Cells[a, "A"] = "ID";
            ws.Cells[a, "B"] = "StartNode";
            ws.Cells[a, "C"] = "EndNode";
            ws.Cells[a, "D"] = "startX";
            ws.Cells[a, "E"] = "startY";
            ws.Cells[a, "F"] = "startZ";
            ws.Cells[a, "G"] = "endX";
            ws.Cells[a, "H"] = "endY";
            ws.Cells[a, "I"] = "endZ";
            ws.Cells[a, "J"] = "Sec";
            ws.Cells[a, "K"] = "Lib";
            ws.Cells[a, "L"] = "Shape";
            ws.Cells[a, "M"] = "IsBUS";
            ws.Cells[a, "N"] = "IsCable";
            ws.Cells[a, "O"] = "IsLink";
            ws.Cells[a, "P"] = "IsConc";
            ws.Cells[a, "Q"] = "Group";
            ws.Cells[a, "R"] = "Material";
            ws.Cells[a, "S"] = "Length";
            ws.Cells[a, "T"] = "Beta";
            ws.Cells[a, "Y"] = "RvtName";
            ws.Cells[a, "Z"] = "RvtFamily";
            ws.Cells[a, "AA"] = "RvtID";
            ws.Cells[a, "AB"] = "RvtstartX";
            ws.Cells[a, "AC"] = "RvtstartY";
            ws.Cells[a, "AD"] = "RvtstartZ";
            ws.Cells[a, "AE"] = "RvtendX";
            ws.Cells[a, "AF"] = "RvtendY";
            ws.Cells[a, "AG"] = "RvtendZ";
            ws.Cells[a, "AH"] = "DeltastartX";
            ws.Cells[a, "AI"] = "DeltastartY";
            ws.Cells[a, "AJ"] = "DeltastartZ";
            ws.Cells[a, "AK"] = "DeltaendX";
            ws.Cells[a, "AL"] = "DeltaendY";
            ws.Cells[a, "AM"] = "DeltaendZ";


            string username = Environment.UserName;
            string main_path = "C:\\Users\\" + username + "\\Desktop\\H Y D R A _temp\\ExcDataWithRvt";
            Directory.CreateDirectory(main_path);
            app.DisplayAlerts = false;
            wb.SaveAs("C:\\Users\\" + username + "\\Desktop\\H Y D R A _temp\\ExcDataWithRvt\\Std To RVT.xlsx");

            #endregion

        }

        public static void CreateExcelForSAP(ref Microsoft.Office.Interop.Excel.Application app, ref Microsoft.Office.Interop.Excel._Workbook wb, ref Microsoft.Office.Interop.Excel._Worksheet ws,String analysisMOdel)

        {
            string fileName = "Sap To RVT.xlsx";
            if (analysisMOdel == "Etabs")
            {
                fileName = "Etabs To RVT.xlsx";
            }

            string fileName2 = "EXCEL.EXE";

            // Get all processes with the specified file name
            //Process[] processes = Process.GetProcessesByName("EXCEL").Where(p => p.MainWindowTitle.Contains(fileName)).ToArray();
            Process[] processes = Process.GetProcessesByName("EXCEL").ToArray();

            // Close all processes
            foreach (Process process in processes)
            {
                try
                {
                    string filePath = process.MainModule.FileName;
                    string processFileName = Path.GetFileName(filePath);

                    if (processFileName.Equals(fileName, StringComparison.InvariantCultureIgnoreCase) || processFileName.Equals(fileName2, StringComparison.InvariantCultureIgnoreCase))
                    {
                        process.Kill();
                    }
                }
                catch (Exception ex)
                {
                    // Handle any exceptions that might occur while trying to access the process information

                }
            }


            #region get_Excel_object   

            // creating Excel Application  
            app = new Microsoft.Office.Interop.Excel.Application();
            // creating new WorkBook within Excel application  
            wb = app.Workbooks.Add(System.Type.Missing);
            // creating new Excelsheet in workbook  
            ws = null;
            //worksheet.Name = "ExcDataWithRvt";
            // see the excel sheet behind the program  
            app.Visible = false;
            // get the reference of first sheet. By default its name is Sheet1.  
            // store its reference to worksheet  
            ws = (_Worksheet)wb.Sheets["Sheet1"];
            ws.Name = "Sap";

            int a = 6;
            ws.Cells[a - 1, "A"] = "UserUnits";
            //worksheet1.Cells[a - 1, "B"] = currentunits.ToString();
            ws.Cells[a - 1, "C"] = "ExportUnits";
            ws.Cells[a - 1, "D"] = "kN_mm_C";
            ws.Cells[a - 1, "E"] = "CountNumber";

            ws.Cells[a, "A"] = "ID";
            ws.Cells[a, "B"] = "StartNode";
            ws.Cells[a, "C"] = "EndNode";
            ws.Cells[a, "D"] = "startX";
            ws.Cells[a, "E"] = "startY";
            ws.Cells[a, "F"] = "startZ";
            ws.Cells[a, "G"] = "endX";
            ws.Cells[a, "H"] = "endY";
            ws.Cells[a, "I"] = "endZ";
            ws.Cells[a, "J"] = "Sec";
            ws.Cells[a, "K"] = "Lib";
            ws.Cells[a, "L"] = "Shape";
            ws.Cells[a, "M"] = "IsBUS";
            ws.Cells[a, "N"] = "IsCable";
            ws.Cells[a, "O"] = "IsLink";
            ws.Cells[a, "P"] = "IsConc";
            ws.Cells[a, "Q"] = "Group";
            ws.Cells[a, "R"] = "Material";
            ws.Cells[a, "S"] = "Length";
            ws.Cells[a, "T"] = "Beta";
            ws.Cells[a, "Y"] = "RvtName";
            ws.Cells[a, "Z"] = "RvtFamily";
            ws.Cells[a, "AA"] = "RvtID";
            ws.Cells[a, "AB"] = "RvtstartX";
            ws.Cells[a, "AC"] = "RvtstartY";
            ws.Cells[a, "AD"] = "RvtstartZ";
            ws.Cells[a, "AE"] = "RvtendX";
            ws.Cells[a, "AF"] = "RvtendY";
            ws.Cells[a, "AG"] = "RvtendZ";
            ws.Cells[a, "AH"] = "DeltastartX";
            ws.Cells[a, "AI"] = "DeltastartY";
            ws.Cells[a, "AJ"] = "DeltastartZ";
            ws.Cells[a, "AK"] = "DeltaendX";
            ws.Cells[a, "AL"] = "DeltaendY";
            ws.Cells[a, "AM"] = "DeltaendZ";


            string username = Environment.UserName;
            string main_path = "C:\\Users\\" + username + "\\Desktop\\H Y D R A _temp\\ExcDataWithRvt";
            Directory.CreateDirectory(main_path);
            app.DisplayAlerts = false;
            wb.SaveAs("C:\\Users\\" + username + "\\Desktop\\H Y D R A _temp\\ExcDataWithRvt\\"+ fileName);

            #endregion

        }

        public static void AddDataToWorksheet5(Workbook workbook, List<object> records, string sheetName)
        {

            #region definitions
            bool conutHighlight = false;
            bool LengthHighlight = false;
            bool TonnageHighlight = false;


            #endregion

            // Create a new worksheet
            Worksheet worksheet = (Worksheet)workbook.Sheets.Add();
            worksheet.Name = sheetName;

            // Add header row with property names
            int rowIndex = 1;
            int colIndex = 1;
            foreach (PropertyInfo property in records[0].GetType().GetProperties())
            {
                worksheet.Cells[rowIndex, colIndex] = property.Name;
                colIndex++;
            }

            // Add data rows
            rowIndex++;
            foreach (var record in records)
            {

                colIndex = 1;
                foreach (PropertyInfo property in record.GetType().GetProperties())
                {
                    Range cell = (Range)worksheet.Cells[rowIndex, colIndex];
                    cell.Value = property.GetValue(record);
                    colIndex++;
                }

                conutHighlight = false;
                LengthHighlight = false;
                TonnageHighlight = false;


                rowIndex++;
            }


            // Add a final row for column sums


           // int lastRow = rowIndex;
            //// Calculate column sums
            //for (int i = 1; i <= colIndex - 1; i++) // Loop through each column
            //{
            //    double sum = 0;
            //    for (int j = 2; j <= lastRow; j++) // Start from the second row to exclude header
            //    {
            //        double cellValue;
            //        if (double.TryParse((worksheet.Cells[j, i] as Range).Value?.ToString(), out cellValue)) // Try to parse cell value to double
            //        {
            //            sum += cellValue;
            //        }
            //    }

            //    if (sum != 0)
            //    {
            //        (worksheet.Cells[lastRow + 1, i] as Range).Value = sum; // Insert sum in the last row

            //    }
            //}





            // Autofit columns
            //Range usedRange = worksheet.UsedRange;
            Excel.Range usedRange = worksheet.Range["A1", worksheet.Cells[worksheet.Rows.Count, worksheet.Columns.Count - 2]];
            usedRange.Columns.AutoFit();

        }

        public static void AddDataToWorksheet(Workbook workbook, List<object> records, string sheetName)
        {
            #region definitions
            bool conutHighlight = false;
            bool LengthHighlight = false;
            bool TonnageHighlight = false;
            #endregion

            // Create a new worksheet
            Worksheet worksheet = (Worksheet)workbook.Sheets.Add();
            worksheet.Name = sheetName;

            // Define header and data array
            int rows = records.Count + 1; // +1 for header
            int cols = records[0].GetType().GetProperties().Length;
            object[,] data = new object[rows, cols];

            // Add header row with property names
            int colIndex = 0;
            foreach (PropertyInfo property in records[0].GetType().GetProperties())
            {
                data[0, colIndex] = property.Name;
                colIndex++;
            }

            // Add data rows
            for (int rowIndex = 0; rowIndex < records.Count; rowIndex++)
            {
                colIndex = 0;
                var record = records[rowIndex];
                foreach (PropertyInfo property in record.GetType().GetProperties())
                {
                    data[rowIndex + 1, colIndex] = property.GetValue(record); // Data starts at row 2 in Excel
                    colIndex++;
                }
            }

            // Define the range to fill in bulk
            Excel.Range startCell = (Range)worksheet.Cells[1, 1];
            Excel.Range endCell = (Range)worksheet.Cells[rows, cols];
            Excel.Range writeRange = worksheet.Range[startCell, endCell];

            // Write the data to the worksheet in bulk
            writeRange.Value = data;

            // Autofit columns
            writeRange.Columns.AutoFit();
        }

        public static void AddDataToWorksheet2(Workbook workbook, List<object> records, string sheetName)
        {

            #region definitions
            bool conutHighlight = false;
            bool LengthHighlight = false;
            bool TonnageHighlight = false;


            #endregion

            // Create a new worksheet
            Worksheet worksheet = (Worksheet)workbook.Sheets.Add();
            worksheet.Name = sheetName;

            // Add header row with property names
            int rowIndex = 1;
            int colIndex = 1;
            foreach (PropertyInfo property in records[0].GetType().GetProperties())
            {
                worksheet.Cells[rowIndex, colIndex] = property.Name;
                colIndex++;
            }

            // Add data rows
            rowIndex++;
            foreach (var record in records)
            {

                try
                {
                    #region connt diffrence

                    if ((string)record.GetType().GetProperty("RevitDesignGroup").GetValue(record) != (string)record.GetType().GetProperty("SAFESpring").GetValue(record))
                        conutHighlight = true;
                    #endregion

                }

                catch (Exception)
                {


                }



                colIndex = 1;
                foreach (PropertyInfo property in record.GetType().GetProperties())
                {
                    Range cell = (Range)worksheet.Cells[rowIndex, colIndex];
                    cell.Value = property.GetValue(record);


                    try
                    {
                        // Check if highlighting is needed and apply formatting
                        if (conutHighlight && (property.Name == "RevitDesignGroup" || property.Name == "SAFESpring"))
                        {
                            cell.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);

                        }

                        //if (LengthHighlight && (property.Name == HardCodedData.revitLength || property.Name == HardCodedData.sapLength))
                        //{
                        //    cell.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);

                        //}

                        //if (TonnageHighlight && (property.Name == HardCodedData.revitTonnage || property.Name == HardCodedData.sapTonnage))
                        //{
                        //    cell.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);

                        //}

                    }

                    catch (Exception)
                    {

                    }


                    colIndex++;
                }

                conutHighlight = false;
                LengthHighlight = false;
                TonnageHighlight = false;


                rowIndex++;
            }


            // Add a final row for column sums
            int lastRow = rowIndex;

            // Calculate column sums
            for (int i = 1; i <= colIndex - 1; i++) // Loop through each column
            {
                double sum = 0;
                for (int j = 2; j <= lastRow; j++) // Start from the second row to exclude header
                {
                    double cellValue;
                    if (double.TryParse((worksheet.Cells[j, i] as Range).Value?.ToString(), out cellValue)) // Try to parse cell value to double
                    {
                        sum += cellValue;
                    }
                }

                if (sum != 0)
                {
                    (worksheet.Cells[lastRow + 1, i] as Range).Value = sum; // Insert sum in the last row

                }
            }





            // Autofit columns
            //Range usedRange = worksheet.UsedRange;
            Excel.Range usedRange = worksheet.Range["A1", worksheet.Cells[worksheet.Rows.Count, worksheet.Columns.Count - 2]];
            usedRange.Columns.AutoFit();

        }

    }


}