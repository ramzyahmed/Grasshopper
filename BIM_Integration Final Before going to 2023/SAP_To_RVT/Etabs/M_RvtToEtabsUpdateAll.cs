using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using ETABSv1;
using Automation = BIM_Integration.BtnAutomation;
//using CSiAPIv1;

namespace BIM_Integration
{
    [Transaction(TransactionMode.Manual)]
    class M_RvtToEtabsUpdateAll : IExternalCommand
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

                #region get Sap frames
                cSapModel mysapmodel = EtabsUtils.Get_sap_mode();
                List<Beam> sapBeams = new List<Beam>();
                List<Column> sapColumns = new List<Column>();
                List<Frame> sapFrames = EtabsUtils.getSapFrames(ref sapBeams, ref sapColumns);
                #endregion

                #region get Revit frames 
                List<Frame> RevitBeams = new List<Frame>();
                List<Frame> RevitColumns = new List<Frame>();
                List<Frame> RevitFrames = RevitUtils.getRevitFrames(uidoc, doc, dis, ref RevitBeams, ref RevitColumns, "Etabs");
                #endregion

                #region Mapping
                string groupName = null;
                List<SapRevitFrame> AllElements = new List<SapRevitFrame>();

                List<Frame> unMappedRevitFrames = new List<Frame>();
                List<Frame> RevitUnMappedBeams = new List<Frame>();
                List<Frame> RevitUnMappedColumns = new List<Frame>();


                List<Frame> unMappedSapFrames = new List<Frame>();
                List<Frame> SapUnMappedBeams = new List<Frame>();
                List<Frame> SapUnMappedColumns = new List<Frame>();

                List<SapRevitFrame> mappedElements = new List<SapRevitFrame>();
                List<SapRevitFrame> SapMappedBeams = new List<SapRevitFrame>();
                List<SapRevitFrame> sapMappedColumns = new List<SapRevitFrame>();

                List<SapRevitFrame> mappedElementsById = new List<SapRevitFrame>();
                List<SapRevitFrame> mappedElementsByMid = new List<SapRevitFrame>();


                if (sapFrames.Count != 0)
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



                // create sapSections
                Dictionary<string, string> sapSectionsDicForBeams = EtabsUtils.createSapSectionForBeams2(RevitBeams, mysapmodel);
                Dictionary<string, string> sapSectionsDicForColumns = EtabsUtils.createSapSectionForColumns2(RevitColumns, mysapmodel);


                #region forMappedFrames
                SapMappedBeams = mappedElements.AsParallel().Where(x => x.sapFrame.MemberType == MemberType.Beam || x.sapFrame.MemberType == MemberType.VLBracing).ToList();
                Beam.updateEtabsSections(SapMappedBeams, mysapmodel, doc, "BM", sapSectionsDicForBeams, sapSectionsDicForColumns);
                Beam.updateEtabsGeomerty(SapMappedBeams, mysapmodel, doc, 0.001);

                sapMappedColumns = mappedElements.AsParallel().Where(x => x.sapFrame.MemberType == MemberType.Column || x.sapFrame.MemberType == MemberType.SlanttedColumn).ToList();
                Beam.updateEtabsSections(sapMappedColumns, mysapmodel, doc, "COL", sapSectionsDicForBeams, sapSectionsDicForColumns);
                Beam.updateEtabsGeomerty(sapMappedColumns, mysapmodel, doc, 0.001);
                #endregion

                #region ForUnMappedFrames
                RevitUnMappedBeams = unMappedRevitFrames.AsParallel().Where(x => x.MemberType == MemberType.Beam || x.MemberType == MemberType.VLBracing).ToList();
                RevitUnMappedColumns = unMappedRevitFrames.AsParallel().Where(x => x.MemberType == MemberType.Column || x.MemberType == MemberType.SlanttedColumn).ToList();


                List<string> newSapCreatedElementsIds = new List<string>();
                string id = null;
                foreach (var revitElement in RevitUnMappedBeams)
                {
                    id = EtabsUtils.createSapBeam(revitElement, mysapmodel, doc, "BM", sapSectionsDicForBeams, sapSectionsDicForColumns);
                    newSapCreatedElementsIds.Add(id);
                }

                foreach (var revitElement in RevitUnMappedColumns)
                {

                    id = EtabsUtils.createSapColumn(revitElement, mysapmodel, doc, "COL", sapSectionsDicForBeams, sapSectionsDicForColumns);
                    newSapCreatedElementsIds.Add(id);
                }


                #endregion



                // handle membersInSapButnotInRevit (group)
                groupName = "00_Members in Sap but not in Revit";
                EtabsUtils.createGroupinSap(unMappedSapFrames.Select(x => x.ID).ToList(), mysapmodel, groupName);


                groupName = "00_Members in Revit but not in Sap (Newly Created)";
                EtabsUtils.createGroupinSap(newSapCreatedElementsIds, mysapmodel, groupName);


                groupName = "00_Matching Members(Sections and Geometry Updated)";
                EtabsUtils.createGroupinSap(mappedElements.Select(x => x.sapFrame.ID).ToList(), mysapmodel, groupName);

                // delete the group
                mysapmodel.GroupDef.Delete("00_Matching Members(Only Geomerty Updated)");

                // delete the group
                mysapmodel.GroupDef.Delete("00_Matching Members(Only Sections Updated)");


                mysapmodel = null;
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


