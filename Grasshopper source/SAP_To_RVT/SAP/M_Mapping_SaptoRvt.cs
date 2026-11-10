using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Numerics;
using Frame = BIM_Integration.Frame;
using Point = BIM_Integration.Point;
using SapRevitFrame = BIM_Integration.SapRevitFrame;
using Excel = Microsoft.Office.Interop.Excel;
//using BIM_Integration.Forms;

namespace BIM_Integration
{
    [Transaction(TransactionMode.Manual)]
    internal class M_MappingSapToRvt : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {


            try
            {

                    #region get Revit Data
                    UIDocument uidoc = commandData.Application.ActiveUIDocument;
                Document doc = uidoc.Document;

                // for revit 2021 or above use (ForgeTypeId dis = UnitTypeId.Millimeters;) however for revit 2020 or lower use ( ForgeTypeId dis = DisplayUnitType.DUT_MILLIMETERS;)
                ForgeTypeId dis = UnitTypeId.Millimeters;

                List<Frame> RevitBeams = new List<Frame>();
                List<Frame> RevitColumns = new List<Frame>();
                List<Frame> RevitFrames = RevitUtils.getRevitFrames(uidoc, doc, dis, ref RevitBeams, ref RevitColumns, "Sap");

                #endregion

                    #region get all Sap Data from Excel 
                    List<Column> sapColumns = new List<Column>();
                    List<Beam> sapBeams = new List<Beam>();
                    List<Frame> sapFrames = new List<Frame>();

                    // get all beams, columns and all members
                    sapFrames = SapUtils.getSapFrames(ref sapBeams, ref sapColumns);

                    ChangeSapSectionNames.GetSectionsDimentions2(SapUtils.Get_sap_mode(), ref sapFrames);

                    #endregion

                    #region getRevitSymbols
                    List<Frame> sabBeams2 = sapFrames.Where(x => x.MemberType == MemberType.Beam || x.MemberType == MemberType.VLBracing).ToList();
                    sabBeams2 = sabBeams2.GroupBy(x => x.Sec).Select(y => y.First()).ToList();

                    Dictionary<string, FamilySymbol> beamsSymbolsDictionary = RevitUtils.createRvtFamilySectionForAllBeam(sabBeams2, doc, dis);


                    List<Frame> sabColumns2 = sapFrames.Where(x => x.MemberType == MemberType.Column || x.MemberType == MemberType.SlanttedColumn).ToList();
                    sabColumns2 = sabColumns2.GroupBy(x => x.Sec).Select(y => y.First()).ToList();

                    Dictionary<string, FamilySymbol> columnsSymbolsDictionary = RevitUtils.createRvtFamilySectionForAllColumn(sabColumns2, doc, dis);

                    #endregion

                    #region Mapping
                    string GroupName = null;
                    List<SapRevitFrame> AllElements = new List<SapRevitFrame>();

                    List<Frame> unMappedRevitFrames = new List<Frame>();
                    List<Frame> unMappedSapFrames = new List<Frame>();
                    List<Frame> SapUnMappedBeams = new List<Frame>();
                    List<Frame> SapUnMappedColumns = new List<Frame>();

                    List<SapRevitFrame> mappedElements = new List<SapRevitFrame>();
                    List<SapRevitFrame> SapMappedBeams = new List<SapRevitFrame>();
                    List<SapRevitFrame> sapMappedColumns = new List<SapRevitFrame>();

                    List<SapRevitFrame> mappedElementsById = new List<SapRevitFrame>();
                    List<SapRevitFrame> mappedElementsByMid = new List<SapRevitFrame>();


                    if (RevitFrames.Count != 0)
                    {
                        AllElements = SapRevitFrame.getMappedSapRevit(sapFrames, RevitFrames, 0.05);
                        mappedElements = AllElements.AsParallel().Where(x => x.revitFrame != null).ToList();


                        mappedElementsById = mappedElements.AsParallel().Where(x => x.mappedBy == "ID").ToList();
                        mappedElementsByMid = mappedElements.AsParallel().Where(x => x.mappedBy == "MID").ToList();


                        SapRevitFrame.getUnMappedEelements(AllElements, sapFrames, RevitFrames, ref unMappedSapFrames, ref unMappedRevitFrames);
                    }


                    else
                    {
                        unMappedSapFrames = sapFrames;
                        unMappedRevitFrames = RevitFrames;
                    }

                    #endregion

                    #region forMappedSapFrames
                    object[,] dataToExcel = new object[sapFrames.Count, 9];

                    SapMappedBeams = mappedElements.AsParallel().Where(x => x.sapFrame.MemberType == MemberType.Beam || x.sapFrame.MemberType == MemberType.VLBracing).ToList();
                    Beam.updateRevitSectionsForBeams2(doc, SapMappedBeams, beamsSymbolsDictionary, dis, null, "Sap", dataToExcel);
                    List<FamilyInstance> ChangedBeams = Beam.updateRevitGeomertyForBeams2(doc, SapMappedBeams, dis, null, 0.001, "Sap", dataToExcel);

                    sapMappedColumns = mappedElements.AsParallel().Where(x => x.sapFrame.MemberType == MemberType.Column || x.sapFrame.MemberType == MemberType.SlanttedColumn).ToList();
                    Column.updateRevitSectionsForColumns2(doc, sapMappedColumns, columnsSymbolsDictionary, dis, null, "Sap", dataToExcel);
                    List<FamilyInstance> ChangedColumns = Column.updateRevitGeomertyForColumns2(doc, sapMappedColumns, dis, null, 0.001, "Sap", dataToExcel);

                    #endregion

                    #region ForUnMappedFrames
                    //object[,] dataToExcel = new object[SapUnMappedBeams.Count + SapUnMappedColumns.Count, 9];
                    SapUnMappedBeams = unMappedSapFrames.AsParallel().Where(x => x.MemberType == MemberType.Beam || x.MemberType == MemberType.VLBracing).ToList();
                    SapUnMappedColumns = unMappedSapFrames.AsParallel().Where(x => x.MemberType == MemberType.Column || x.MemberType == MemberType.SlanttedColumn).ToList();

                    List<ElementId> newRevitCreatedElementsIds = new List<ElementId>();
                    newRevitCreatedElementsIds = Beam.CreateAllBeams3(doc, SapUnMappedBeams, dis, beamsSymbolsDictionary, null, "Sap", dataToExcel);
                    newRevitCreatedElementsIds = newRevitCreatedElementsIds.Union(Column.CreateAllColumns3(doc, SapUnMappedColumns, dis, columnsSymbolsDictionary, null, "Sap", ref dataToExcel)).ToList();



                    newRevitCreatedElementsIds.RemoveAll(x => x == null);


                    GroupName = "00_Members in REVIT but not in SAP";
                    View3D unmappedView = SapRevitFrame.createView(unMappedRevitFrames.Select(x => x.revitEleId).ToList(), commandData, GroupName);

                    GroupName = "00_Members in SAP but not in REVIT (Newly Created)";
                    View3D unmappedView2 = SapRevitFrame.createView(newRevitCreatedElementsIds, commandData, GroupName);

                    GroupName = "00_Matching Members(Sections and Geometry Updated)";
                    List<ElementId> updatedPlusChanged = mappedElements.Select(x => x.revitFrame.revitEleId).ToList().Union(ChangedBeams.Select(x => x.Id).ToList()).Union(ChangedColumns.Select(x => x.Id).ToList()).ToList();
                    View3D unmappedView3 = SapRevitFrame.createView(updatedPlusChanged, commandData, GroupName);


                    View3D view1 = new FilteredElementCollector(doc).OfClass(typeof(View3D)).Cast<View3D>().ToList().FirstOrDefault(x => x.Name == "00_Matching Members(Only Geometry Updated)");
                    View3D view2 = new FilteredElementCollector(doc).OfClass(typeof(View3D)).Cast<View3D>().ToList().FirstOrDefault(x => x.Name == "00_Matching Members(Only Sections Updated)");
                    View3D view3 = new FilteredElementCollector(doc).OfClass(typeof(View3D)).Cast<View3D>().ToList().FirstOrDefault(x => x.Name == "00_Members in REVIT but not in STAAD");
                    View3D view4 = new FilteredElementCollector(doc).OfClass(typeof(View3D)).Cast<View3D>().ToList().FirstOrDefault(x => x.Name == "00_Members in STAAD but not in REVIT (Newly Created)");

                    List<View3D> viewsToDeleted = new List<View3D>();
                    viewsToDeleted.Add(view1);
                    viewsToDeleted.Add(view2);
                    viewsToDeleted.Add(view3);
                    viewsToDeleted.Add(view4);

                    foreach (View3D V in viewsToDeleted)
                    {
                        if (V != null)
                        {
                            using (Transaction trns = new Transaction(doc, "Delete VIEW"))
                            {
                                trns.Start();
                                doc.Delete(V.Id);
                                trns.Commit();

                            }

                        }
                    }

                    #endregion


                return Result.Succeeded;


            }


            catch (Exception ex2)
            {

                System.Diagnostics.StackTrace trace = new System.Diagnostics.StackTrace(ex2, true);
                TaskDialog.Show("ErrorInMainMethod", "ErrorInMainMethod " + ex2.Message + trace.GetFrame(trace.FrameCount - 1).GetFileLineNumber());
                return Result.Failed;
            }







        }
    }

}