using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using SAP2000v1;
using System;
using System.Collections.Generic;
using System.Linq;

using BIM_Integration.STAAD;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Threading;

namespace BIM_Integration
{
    [Transaction(TransactionMode.Manual)]
    class M_RvtToStaadUpdateSections : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {


                commandData.Application.Idling -= BtnAutomation.idleUpdate;
                commandData.Application.Idling -= BtnUpdateGeometry.idleUpdate;
                commandData.Application.Idling -= BtnUpdateSections.idleUpdate;

                #region get Revit Data
                UIDocument uidoc = commandData.Application.ActiveUIDocument;
                Document doc = uidoc.Document;
                // for revit 2021 or above use (ForgeTypeId dis = UnitTypeId.Millimeters;) however for revit 2020 or lower use ( ForgeTypeId dis = DisplayUnitType.DUT_MILLIMETERS;)
                ForgeTypeId dis = UnitTypeId.Millimeters;
                #endregion

                #region get Revit frames 
                List<Frame> RevitBeams = new List<Frame>();
                List<Frame> RevitColumns = new List<Frame>();
                List<Frame> RevitFrames = RevitUtils.getRevitFrames(uidoc, doc, dis, ref RevitBeams, ref RevitColumns, "Staad");
                #endregion

                #region Create_Excel_object   

                // creating Excel Application  
                Microsoft.Office.Interop.Excel.Application app = null;
                // creating new WorkBook within Excel application  
                Microsoft.Office.Interop.Excel._Workbook workbook = null;
                // creating new Excelsheet in workbook  
                Microsoft.Office.Interop.Excel._Worksheet worksheet1 = null;
                ExcelUtils.CreateExcel(ref app, ref workbook, ref worksheet1);

                #endregion


                RevitUtils.setRevitDataInExcelBulk(RevitFrames, (Microsoft.Office.Interop.Excel.Worksheet)worksheet1);

                #region saveExcel
                workbook.Save();
                workbook.Close();
                app.DisplayAlerts = true;
                app.Quit();
                #endregion

                #region write RvtToStd_1
                string username = Environment.UserName;
                string folder = @"C:\\Users\\" + username + "\\Desktop\\H Y D R A _temp\\ExcDataWithRvt\\";
                string fileName = "HYDRA Integration.txt";
                // Fullpath. You can direct hardcode it if you like.  
                string fullPath = folder + fileName;
                // An array of strings  
                string[] write = { "RvtToStd_Sections" };
                // If the file does not exists, it will create a new file.  
                // This method automatically opens the file, writes to it, and closes file  
                File.WriteAllLines(fullPath, write);



                #endregion


                //fromHydra
                DateTime lastModified = File.GetLastWriteTime(fullPath);
                //M_RvtToStaad.M_RvtToStaadUpdateSectionsExecute();



                #region reflect the staadDataFromExcelToRevitPramaeters
                while (true)
                {
                    // Check if the file has been modified
                    const int maxRetries = 10;
                    const int retryDelayMilliseconds = 500;

                    if (File.GetLastWriteTime(fullPath) > lastModified)
                    {
                        string firstLine = null;
                        int retryCount = 0;
                        bool fileAccessed = false;

                        FileStream fs = null;
                        StreamReader sr = null;
                        try
                        {
                            while (retryCount < maxRetries)
                            {
                                try
                                {
                                    using (fs = new FileStream(fullPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                                    using (sr = new StreamReader(fs))
                                    {
                                        firstLine = sr.ReadLine();
                                    }

                                    fileAccessed = true;
                                    break;
                                }
                                catch (IOException)
                                {
                                    retryCount++;
                                    Thread.Sleep(retryDelayMilliseconds);
                                }
                            }

                            if (fileAccessed)
                            {
                                if (firstLine == "StdToRvt_1")
                                {
                                    #region get_Excel_object
                                    Microsoft.Office.Interop.Excel.Application excel = null;
                                    Microsoft.Office.Interop.Excel._Workbook wb = null;
                                    Microsoft.Office.Interop.Excel._Worksheet ws = null;
                                    ExcelUtils.getExcel("Staad", ref excel, ref wb, ref ws);
                                    #endregion

                                    RevitUtils.assignStaadDataFromExcelToRevitParametersinBulk(doc, (Microsoft.Office.Interop.Excel.Worksheet)ws);
                                    wb.Close();
                                    excel.Quit();
                                    break;
                                }
                                else if (firstLine == "Quit")
                                {
                                    //File.WriteAllText(fullPath, string.Empty);
                                    break;
                                }
                            }
                            else
                            {
                                // File was still accessed by another process after retries
                                // Handle the error or retry logic here
                            }
                        }
                        finally

                        {

                            if (sr != null)
                                sr.Dispose();

                            if (fs != null)
                                fs.Dispose();

                        }


                    }


                }


                #endregion

                return Result.Succeeded;
            }

            catch (Exception ex2)
            {

                System.Diagnostics.StackTrace trace = new System.Diagnostics.StackTrace(ex2, true);
                TaskDialog.Show("Error", "error in beam no" + ex2.Message + trace.GetFrame(trace.FrameCount - 1).GetFileLineNumber());

                return Result.Failed;
            }



        }


    }
}


