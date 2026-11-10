using System;
using System.Collections.Generic;
using System.Linq;
using BIM_Integration.STAAD;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;


namespace BIM_Integration_Hydra
{

    class M_RvtToStaad
    {
        static public void M_RvtToStaadUpdateAllExecute()
        {
            try
            {
                #region get_Excel_object   

                Microsoft.Office.Interop.Excel.Application excel = null;
                Microsoft.Office.Interop.Excel._Workbook wb = null;
                Microsoft.Office.Interop.Excel._Worksheet ws = null;
                BIM_Integration.ExcelUtils.getExcel("Staad", ref excel, ref wb, ref ws);

                #endregion

                #region get Revit frames 
                List<Frame> RevitBeams = new List<Frame>();
                List<Frame> RevitColumns = new List<Frame>();
                List<Frame> RevitFrames = StaadUtils_HYDRA.getRevitFramesFromExcel(ref RevitBeams, ref RevitColumns, (Microsoft.Office.Interop.Excel.Worksheet)ws);

                #endregion

                #region get Staad frames
                List<Beam> staadBeams = new List<Beam>();
                List<Column> staadColumns = new List<Column>();
                List<Frame> staadFrames = StaadUtils_HYDRA.getStaadFrames(ref staadBeams, ref staadColumns);
                #endregion

                #region Mapping
                string groupName = null;
                List<SapRevitFrame> AllElements = new List<SapRevitFrame>();

                List<Frame> unMappedRevitFrames = new List<Frame>();
                List<Frame> RevitUnMappedBeams = new List<Frame>();
                List<Frame> RevitUnMappedColumns = new List<Frame>();


                List<Frame> unMappedStaadFrames = new List<Frame>();
                List<Frame> staadUnMappedBeams = new List<Frame>();
                List<Frame> staadUnMappedColumns = new List<Frame>();

                List<SapRevitFrame> mappedElements = new List<SapRevitFrame>();
                List<SapRevitFrame> staadMappedBeams = new List<SapRevitFrame>();
                List<SapRevitFrame> staadMappedColumns = new List<SapRevitFrame>();

                List<SapRevitFrame> mappedElementsById = new List<SapRevitFrame>();
                List<SapRevitFrame> mappedElementsByMid = new List<SapRevitFrame>();


                if (staadFrames.Count != 0)
                {
                    AllElements = SapRevitFrame.getMappedSapRevitByIdOnly(staadFrames, RevitFrames);
                    mappedElements = AllElements.Where(x => x.revitFrame != null).ToList();

                    mappedElementsById = mappedElements.Where(x => x.mappedBy == "ID").ToList();
                    mappedElementsByMid = mappedElements.Where(x => x.mappedBy == "MID").ToList();
                    SapRevitFrame.getUnMappedEelements(AllElements, staadFrames, RevitFrames, ref unMappedStaadFrames, ref unMappedRevitFrames);
                }



                else
                {
                    unMappedStaadFrames = staadFrames;
                    unMappedRevitFrames = RevitFrames;
                }



                #endregion

                #region forMappedFrames
                staadMappedBeams = mappedElements.Where(x => x.sapFrame.MemberType == MemberType.Beam || x.sapFrame.MemberType == MemberType.VLBracing).ToList();
                Beam.updateStaadGeomerty(staadMappedBeams, 0.05);
                Beam.updateStaadSections(staadMappedBeams, "BM");

                staadMappedColumns = mappedElements.Where(x => x.sapFrame.MemberType == MemberType.Column || x.sapFrame.MemberType == MemberType.SlanttedColumn).ToList();
                Beam.updateStaadGeomerty(staadMappedColumns, 0.05);
                Beam.updateStaadSections(staadMappedColumns, "COL");


                #endregion

                #region ForUnMappedFrames
                RevitUnMappedBeams = unMappedRevitFrames.Where(x => x.MemberType == MemberType.Beam || x.MemberType == MemberType.VLBracing).ToList();
                RevitUnMappedColumns = unMappedRevitFrames.Where(x => x.MemberType == MemberType.Column || x.MemberType == MemberType.SlanttedColumn).ToList();


                List<string> newSapCreatedElementsIds = new List<string>();
                string id = null;

                Dictionary<string, int> revitStaadSectionDicForBeams = StaadUtils_HYDRA.createStaadSectionForBeams(RevitUnMappedBeams);

                Dictionary<string, int> revitStaadSectionDicForCol = StaadUtils_HYDRA.createStaadSectionForColumns(RevitUnMappedColumns);

                foreach (var revitElement in RevitUnMappedBeams)
                {
                    id = StaadUtils_HYDRA.createStaadBeam(revitElement, revitStaadSectionDicForBeams);
                    newSapCreatedElementsIds.Add(id);
                    revitElement.ID = id;
                }

                foreach (var revitElement in RevitUnMappedColumns)
                {
                    id = StaadUtils_HYDRA.createStaadBeam(revitElement, revitStaadSectionDicForCol);
                    newSapCreatedElementsIds.Add(id);
                    revitElement.ID = id;
                }


                #endregion

                #region SetStaadDataInExcel
                List<Frame> allMappedSTaadFrames = staadMappedBeams.Select(x => x.sapFrame).Union(staadMappedColumns.Select(x => x.sapFrame)).ToList();
                List<Frame> allUnMappedStaadFrames = RevitUnMappedBeams.Union(RevitUnMappedColumns).ToList();
                StaadUtils_HYDRA.AssignStaadFramesData(ref allUnMappedStaadFrames);
                StaadUtils_HYDRA.setStaadDataInExcel(allMappedSTaadFrames.Union(allUnMappedStaadFrames).ToList(), (Microsoft.Office.Interop.Excel.Worksheet)ws);
                #endregion

                #region saveExcel
                wb.Save();
                wb.Close();
                excel.DisplayAlerts = true;
                excel.Quit();
                #endregion

                #region write to revit
                string username = Environment.UserName;
                string folder = @"C:\\Users\\" + username + "\\Desktop\\H Y D R A _temp\\ExcDataWithRvt\\";
                string fileName = "HYDRA Integration.txt";
                // Fullpath. You can direct hardcode it if you like.  
                string fullPath = folder + fileName;
                // An array of strings  
                string[] write = { "StdToRvt_1" };
                File.Create(fullPath).Close();

                using (FileStream fileStream = new FileStream(fullPath, FileMode.Append, FileAccess.Write, FileShare.Read))
                {
                    // Create a StreamWriter using the FileStream
                    using (StreamWriter streamWriter = new StreamWriter(fileStream))
                    {
                        // Write to the file
                        streamWriter.Write("StdToRvt_1");

                        // Flush the StreamWriter to ensure that all data is written to the file
                        streamWriter.Flush();
                    }

                    // Release the FileStream so that other processes can access the file
                    fileStream.Dispose();
                }


                #endregion

                #region create groups in staad
                // handle membersInSapButnotInRevit (group)
                groupName = "00_Members in staad but not in Revit";
                StaadUtils_HYDRA.createGroup(unMappedStaadFrames.Select(x => x.ID).ToList(), groupName.ToUpper());


                groupName = "00_Members in Revit but not in staad (Newly Created)";
                StaadUtils_HYDRA.createGroup(newSapCreatedElementsIds, groupName.ToUpper());

                groupName = "00_Matching Members(Sections and Geomerty Updated)";
                StaadUtils_HYDRA.createGroup(mappedElements.Select(x => x.sapFrame.ID).ToList(), groupName.ToUpper());


                //// delete the groups
                string grouptoDelete = "00_Matching Members(Only Geomerty Updated)";
                StaadUtils_HYDRA.deleteGroup(grouptoDelete.ToUpper());

                grouptoDelete = "00_Matching Members(Only Sections Updated)";
                StaadUtils_HYDRA.deleteGroup(grouptoDelete.ToUpper());
                #endregion


            }

            catch (Exception ex2)
            {

                System.Diagnostics.StackTrace trace = new System.Diagnostics.StackTrace(ex2, true);


            }



        }

        static public void M_RvtToStaadUpdateGeometryExecute()
        {
            try
            {


                #region get_Excel_object   

                Microsoft.Office.Interop.Excel.Application excel = null;
                Microsoft.Office.Interop.Excel._Workbook wb = null;
                Microsoft.Office.Interop.Excel._Worksheet ws = null;
                BIM_Integration.ExcelUtils.getExcel("Staad", ref excel, ref wb, ref ws);

                #endregion

                #region get Revit frames 
                List<Frame> RevitBeams = new List<Frame>();
                List<Frame> RevitColumns = new List<Frame>();
                List<Frame> RevitFrames = StaadUtils_HYDRA.getRevitFramesFromExcel(ref RevitBeams, ref RevitColumns, (Microsoft.Office.Interop.Excel.Worksheet)ws);

                #endregion

                #region get Staad frames
                List<Beam> staadBeams = new List<Beam>();
                List<Column> staadColumns = new List<Column>();
                List<Frame> staadFrames = StaadUtils_HYDRA.getStaadFrames(ref staadBeams, ref staadColumns);
                #endregion

                #region Mapping
                string groupName = null;
                List<SapRevitFrame> AllElements = new List<SapRevitFrame>();

                List<Frame> unMappedRevitFrames = new List<Frame>();
                List<Frame> RevitUnMappedBeams = new List<Frame>();
                List<Frame> RevitUnMappedColumns = new List<Frame>();


                List<Frame> unMappedStaadFrames = new List<Frame>();
                List<Frame> staadUnMappedBeams = new List<Frame>();
                List<Frame> staadUnMappedColumns = new List<Frame>();

                List<SapRevitFrame> mappedElements = new List<SapRevitFrame>();
                List<SapRevitFrame> staadMappedBeams = new List<SapRevitFrame>();
                List<SapRevitFrame> staadMappedColumns = new List<SapRevitFrame>();

                List<SapRevitFrame> mappedElementsById = new List<SapRevitFrame>();
                List<SapRevitFrame> mappedElementsByMid = new List<SapRevitFrame>();


                if (staadFrames.Count != 0)
                {
                    AllElements = SapRevitFrame.getMappedSapRevitByIdOnly(staadFrames, RevitFrames);
                    mappedElements = AllElements.Where(x => x.revitFrame != null).ToList();

                    mappedElementsById = mappedElements.Where(x => x.mappedBy == "ID").ToList();
                    mappedElementsByMid = mappedElements.Where(x => x.mappedBy == "MID").ToList();
                    SapRevitFrame.getUnMappedEelements(AllElements, staadFrames, RevitFrames, ref unMappedStaadFrames, ref unMappedRevitFrames);
                }



                else
                {
                    unMappedStaadFrames = staadFrames;
                    unMappedRevitFrames = RevitFrames;
                }



                #endregion

                #region forMappedFrames
                staadMappedBeams = mappedElements.Where(x => x.sapFrame.MemberType == MemberType.Beam || x.sapFrame.MemberType == MemberType.VLBracing).ToList();
                Beam.updateStaadGeomerty(staadMappedBeams, 0.05);


                staadMappedColumns = mappedElements.Where(x => x.sapFrame.MemberType == MemberType.Column || x.sapFrame.MemberType == MemberType.SlanttedColumn).ToList();
                Beam.updateStaadGeomerty(staadMappedColumns, 0.05);



                #endregion

                #region ForUnMappedFrames
                RevitUnMappedBeams = unMappedRevitFrames.Where(x => x.MemberType == MemberType.Beam || x.MemberType == MemberType.VLBracing).ToList();
                RevitUnMappedColumns = unMappedRevitFrames.Where(x => x.MemberType == MemberType.Column || x.MemberType == MemberType.SlanttedColumn).ToList();


                List<string> newSapCreatedElementsIds = new List<string>();
                string id = null;

                Dictionary<string, int> revitStaadSectionDicForBeams = StaadUtils_HYDRA.createStaadSectionForBeams(RevitUnMappedBeams);

                Dictionary<string, int> revitStaadSectionDicForCol = StaadUtils_HYDRA.createStaadSectionForColumns(RevitUnMappedColumns);

                foreach (var revitElement in RevitUnMappedBeams)
                {
                    id = StaadUtils_HYDRA.createStaadBeam(revitElement, revitStaadSectionDicForBeams);
                    newSapCreatedElementsIds.Add(id);
                    revitElement.ID = id;
                }

                foreach (var revitElement in RevitUnMappedColumns)
                {
                    id = StaadUtils_HYDRA.createStaadBeam(revitElement, revitStaadSectionDicForCol);
                    newSapCreatedElementsIds.Add(id);
                    revitElement.ID = id;
                }


                #endregion

                #region SetStaadDataInExcel
                List<Frame> allMappedSTaadFrames = staadMappedBeams.Select(x => x.sapFrame).Union(staadMappedColumns.Select(x => x.sapFrame)).ToList();
                List<Frame> allUnMappedStaadFrames = RevitUnMappedBeams.Union(RevitUnMappedColumns).ToList();
                StaadUtils_HYDRA.AssignStaadFramesData(ref allUnMappedStaadFrames);
                StaadUtils_HYDRA.setStaadDataInExcel(allMappedSTaadFrames.Union(allUnMappedStaadFrames).ToList(), (Microsoft.Office.Interop.Excel.Worksheet)ws);


                #endregion

                #region saveExcel
                wb.Save();
                wb.Close();
                excel.DisplayAlerts = true;
                excel.Quit();
                #endregion

                #region write to revit
                string username = Environment.UserName;
                string folder = @"C:\\Users\\" + username + "\\Desktop\\H Y D R A _temp\\ExcDataWithRvt\\";
                string fileName = "HYDRA Integration.txt";
                // Fullpath. You can direct hardcode it if you like.  
                string fullPath = folder + fileName;
                // An array of strings  
                string[] write = { "StdToRvt_1" };
                File.Create(fullPath).Close();

                using (FileStream fileStream = new FileStream(fullPath, FileMode.Append, FileAccess.Write, FileShare.Read))
                {
                    // Create a StreamWriter using the FileStream
                    using (StreamWriter streamWriter = new StreamWriter(fileStream))
                    {
                        // Write to the file
                        streamWriter.Write("StdToRvt_1");

                        // Flush the StreamWriter to ensure that all data is written to the file
                        streamWriter.Flush();
                    }

                    // Release the FileStream so that other processes can access the file
                    fileStream.Dispose();
                }


                #endregion

                #region create groups in staad
                // handle membersInSapButnotInRevit (group)
                groupName = "00_Members in staad but not in Revit";
                StaadUtils_HYDRA.createGroup(unMappedStaadFrames.Select(x => x.ID).ToList(), groupName.ToUpper());


                groupName = "00_Members in Revit but not in staad (Newly Created)";
                StaadUtils_HYDRA.createGroup(newSapCreatedElementsIds, groupName.ToUpper());

                groupName = "00_Matching Members(Sections and Geomerty Updated)";
                StaadUtils_HYDRA.createGroup(mappedElements.Select(x => x.sapFrame.ID).ToList(), groupName.ToUpper());


                //// delete the groups
                string grouptoDelete = "00_Matching Members(Only Geomerty Updated)";
                StaadUtils_HYDRA.deleteGroup(grouptoDelete.ToUpper());

                grouptoDelete = "00_Matching Members(Only Sections Updated)";
                StaadUtils_HYDRA.deleteGroup(grouptoDelete.ToUpper());
                #endregion


            }

            catch (Exception ex2)
            {

                System.Diagnostics.StackTrace trace = new System.Diagnostics.StackTrace(ex2, true);


            }



        }

        static public void M_RvtToStaadUpdateSectionsExecute()
        {

            try
            {
                #region get_Excel_object   

                Microsoft.Office.Interop.Excel.Application excel = null;
                Microsoft.Office.Interop.Excel._Workbook wb = null;
                Microsoft.Office.Interop.Excel._Worksheet ws = null;
                BIM_Integration.ExcelUtils.getExcel("Staad", ref excel, ref wb, ref ws);

                #endregion

                #region get Revit frames 
                List<Frame> RevitBeams = new List<Frame>();
                List<Frame> RevitColumns = new List<Frame>();
                List<Frame> RevitFrames = StaadUtils_HYDRA.getRevitFramesFromExcel(ref RevitBeams, ref RevitColumns, (Microsoft.Office.Interop.Excel.Worksheet)ws);

                #endregion

                #region get Staad frames
                List<Beam> staadBeams = new List<Beam>();
                List<Column> staadColumns = new List<Column>();
                List<Frame> staadFrames = StaadUtils_HYDRA.getStaadFrames(ref staadBeams, ref staadColumns);
                #endregion

                #region Mapping
                string groupName = null;
                List<SapRevitFrame> AllElements = new List<SapRevitFrame>();

                List<Frame> unMappedRevitFrames = new List<Frame>();
                List<Frame> RevitUnMappedBeams = new List<Frame>();
                List<Frame> RevitUnMappedColumns = new List<Frame>();


                List<Frame> unMappedStaadFrames = new List<Frame>();
                List<Frame> staadUnMappedBeams = new List<Frame>();
                List<Frame> staadUnMappedColumns = new List<Frame>();

                List<SapRevitFrame> mappedElements = new List<SapRevitFrame>();
                List<SapRevitFrame> staadMappedBeams = new List<SapRevitFrame>();
                List<SapRevitFrame> staadMappedColumns = new List<SapRevitFrame>();

                List<SapRevitFrame> mappedElementsById = new List<SapRevitFrame>();
                List<SapRevitFrame> mappedElementsByMid = new List<SapRevitFrame>();


                if (staadFrames.Count != 0)
                {
                    AllElements = SapRevitFrame.getMappedSapRevitByIdOnly(staadFrames, RevitFrames);
                    mappedElements = AllElements.Where(x => x.revitFrame != null).ToList();

                    mappedElementsById = mappedElements.Where(x => x.mappedBy == "ID").ToList();
                    mappedElementsByMid = mappedElements.Where(x => x.mappedBy == "MID").ToList();
                    SapRevitFrame.getUnMappedEelements(AllElements, staadFrames, RevitFrames, ref unMappedStaadFrames, ref unMappedRevitFrames);
                }



                else
                {
                    unMappedStaadFrames = staadFrames;
                    unMappedRevitFrames = RevitFrames;
                }



                #endregion

                #region forMappedFrames
                staadMappedBeams = mappedElements.Where(x => x.sapFrame.MemberType == MemberType.Beam || x.sapFrame.MemberType == MemberType.VLBracing).ToList();
                Beam.updateStaadSections(staadMappedBeams, "BM");

                staadMappedColumns = mappedElements.Where(x => x.sapFrame.MemberType == MemberType.Column || x.sapFrame.MemberType == MemberType.SlanttedColumn).ToList();
                Beam.updateStaadSections(staadMappedColumns, "COL");


                #endregion

                #region ForUnMappedFrames
                RevitUnMappedBeams = unMappedRevitFrames.Where(x => x.MemberType == MemberType.Beam || x.MemberType == MemberType.VLBracing).ToList();
                RevitUnMappedColumns = unMappedRevitFrames.Where(x => x.MemberType == MemberType.Column || x.MemberType == MemberType.SlanttedColumn).ToList();


                List<string> newSapCreatedElementsIds = new List<string>();
                string id = null;

                Dictionary<string, int> revitStaadSectionDicForBeams = StaadUtils_HYDRA.createStaadSectionForBeams(RevitUnMappedBeams);

                Dictionary<string, int> revitStaadSectionDicForCol = StaadUtils_HYDRA.createStaadSectionForColumns(RevitUnMappedColumns);

                foreach (var revitElement in RevitUnMappedBeams)
                {
                    id = StaadUtils_HYDRA.createStaadBeam(revitElement, revitStaadSectionDicForBeams);
                    newSapCreatedElementsIds.Add(id);
                    revitElement.ID = id;
                }

                foreach (var revitElement in RevitUnMappedColumns)
                {
                    id = StaadUtils_HYDRA.createStaadBeam(revitElement, revitStaadSectionDicForCol);
                    newSapCreatedElementsIds.Add(id);
                    revitElement.ID = id;
                }


                #endregion

                #region SetStaadDataInExcel
                List<Frame> allMappedSTaadFrames = staadMappedBeams.Select(x => x.sapFrame).Union(staadMappedColumns.Select(x => x.sapFrame)).ToList();
                List<Frame> allUnMappedStaadFrames = RevitUnMappedBeams.Union(RevitUnMappedColumns).ToList();
                StaadUtils_HYDRA.AssignStaadFramesData(ref allUnMappedStaadFrames);
                StaadUtils_HYDRA.setStaadDataInExcel(allMappedSTaadFrames.Union(allUnMappedStaadFrames).ToList(), (Microsoft.Office.Interop.Excel.Worksheet)ws);


                #endregion

                #region saveExcel
                wb.Save();
                wb.Close();
                excel.DisplayAlerts = true;
                excel.Quit();
                #endregion

                #region write to revit
                string username = Environment.UserName;
                string folder = @"C:\\Users\\" + username + "\\Desktop\\H Y D R A _temp\\ExcDataWithRvt\\";
                string fileName = "HYDRA Integration.txt";
                // Fullpath. You can direct hardcode it if you like.  
                string fullPath = folder + fileName;
                // An array of strings  
                string[] write = { "StdToRvt_1" };
                File.Create(fullPath).Close();

                using (FileStream fileStream = new FileStream(fullPath, FileMode.Append, FileAccess.Write, FileShare.Read))
                {
                    // Create a StreamWriter using the FileStream
                    using (StreamWriter streamWriter = new StreamWriter(fileStream))
                    {
                        // Write to the file
                        streamWriter.Write("StdToRvt_1");

                        // Flush the StreamWriter to ensure that all data is written to the file
                        streamWriter.Flush();
                    }

                    // Release the FileStream so that other processes can access the file
                    fileStream.Dispose();
                }


                #endregion

                #region create groups in staad
                // handle membersInSapButnotInRevit (group)
                groupName = "00_Members in staad but not in Revit";
                StaadUtils_HYDRA.createGroup(unMappedStaadFrames.Select(x => x.ID).ToList(), groupName.ToUpper());


                groupName = "00_Members in Revit but not in staad (Newly Created)";
                StaadUtils_HYDRA.createGroup(newSapCreatedElementsIds, groupName.ToUpper());

                groupName = "00_Matching Members(Sections and Geomerty Updated)";
                StaadUtils_HYDRA.createGroup(mappedElements.Select(x => x.sapFrame.ID).ToList(), groupName.ToUpper());


                //// delete the groups
                string grouptoDelete = "00_Matching Members(Only Geomerty Updated)";
                StaadUtils_HYDRA.deleteGroup(grouptoDelete.ToUpper());

                grouptoDelete = "00_Matching Members(Only Sections Updated)";
                StaadUtils_HYDRA.deleteGroup(grouptoDelete.ToUpper());
                #endregion


            }

            catch (Exception ex2)
            {

                System.Diagnostics.StackTrace trace = new System.Diagnostics.StackTrace(ex2, true);


            }


        }

    }
}


