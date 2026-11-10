using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;



namespace BIM_Integration
{
    [Transaction(TransactionMode.Manual)]
    internal class M_BeamsToColumns : IExternalCommand
    {
        public string ModelType = "";
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {

            try
            {

                #region get Revit Data
                UIDocument uidoc = commandData.Application.ActiveUIDocument;
                Document doc = uidoc.Document;

                // for revit 2021 or above use (ForgeTypeId dis = UnitTypeId.Millimeters;) however for revit 2020 or lower use ( ForgeTypeId dis = DisplayUnitType.DUT_MILLIMETERS;)
                ForgeTypeId dis = UnitTypeId.Millimeters;



                #region get get Selected Beams
                List<ElementId> selectedElements = uidoc.Selection.GetElementIds().ToList();

                List<FamilyInstance> selectedBeams = new List<FamilyInstance>();

                foreach (ElementId elementId in selectedElements)
                {
                    Element element = doc.GetElement(elementId);

                    // Check if the element belongs to the Structural Columns or Structural Framing category
                    if ((BuiltInCategory)element.Category.Id.IntegerValue == BuiltInCategory.OST_StructuralFraming)
                    {
                        selectedBeams.Add(element as FamilyInstance);
                    }
                }


                #endregion
                #endregion

                List<Frame> newCreatedColumns = new List<Frame>();

                foreach (var instance in selectedBeams)
                {

                    AnalyticalToPhysicalAssociationManager analyManger = AnalyticalToPhysicalAssociationManager.GetAnalyticalToPhysicalAssociationManager(doc);
                    AnalyticalElement analytInstance = doc.GetElement(analyManger.GetAssociatedElementId(instance.Id)) as AnalyticalElement;

                    Frame frame = new Frame();

                    if (analytInstance.GetCurve() != null)
                    {
                        Point startPoint = new Point(analytInstance.GetCurve().GetEndPoint(0));
                        Point endPoint = new Point(analytInstance.GetCurve().GetEndPoint(1));

                        startPoint.X = UnitUtils.ConvertFromInternalUnits(startPoint.X, dis);
                        startPoint.Y = UnitUtils.ConvertFromInternalUnits(startPoint.Y, dis);
                        startPoint.Z = UnitUtils.ConvertFromInternalUnits(startPoint.Z, dis);

                        endPoint.X = UnitUtils.ConvertFromInternalUnits(endPoint.X, dis);
                        endPoint.Y = UnitUtils.ConvertFromInternalUnits(endPoint.Y, dis);
                        endPoint.Z = UnitUtils.ConvertFromInternalUnits(endPoint.Z, dis);


                        Point midPoint = Point.getMidPoint(startPoint, endPoint);
                        frame.RvtID = instance.Id.ToString();
                        frame.revitEleId = instance.Id;
                        frame.StartPoint = startPoint;
                        frame.EndPoint = endPoint;
                        frame.MidPoint = midPoint;
                        frame.Length = Point.getDistance(startPoint, endPoint);

                        #region GetAnalysisID

                        if (instance.LookupParameter("EtabsID").AsString() != null && instance.LookupParameter("EtabsID").AsString() != "")
                        {
                            frame.ID = instance.LookupParameter("EtabsID").AsString();
                            ModelType = "Etabs";

                        }
                        if (instance.LookupParameter("StaadID").AsString() != null && instance.LookupParameter("StaadID").AsString() != "")
                        {
                            frame.ID = instance.LookupParameter("StaadID").AsString();
                            ModelType = "Staad";

                        }
                        else
                        {
                            frame.ID = instance.LookupParameter("SapID").AsString();
                            ModelType = "Sap";
                        }

                        #endregion

                        frame.Material = instance.LookupParameter("Material").AsString();
                        frame.FamilyType = RevitUtils.getBeamSecType(instance.Symbol.FamilyName);


                        string shape = "";
                        string lib = "";
                        bool isbus = false;
                        RevitUtils.mappBetweenSections(frame.FamilyType, ref shape, ref lib, ref isbus);
                        frame.Shape = shape;
                        frame.Lib = lib;
                        frame.IsBUS = isbus;

                        frame.Sec = instance.Symbol.Name;

                        if (frame.Sec.Contains("UB"))
                            frame.Sec = frame.Sec.Replace("UB", "UKB");

                        if (frame.Sec.Contains("UC"))
                            frame.Sec = frame.Sec.Replace("UC", "UKC");


                        frame.RvtName = instance.Symbol.Name;
                        frame.MemberType = MemberType.SlanttedColumn;


                        #region getRotationAngle
                        double angle = 0.0;
                        frame.Beta = instance.LookupParameter("Cross-Section Rotation").AsDouble();
                        frame.Beta = Math.Round((frame.Beta) * 180 / Math.PI, 0) ;

                        if (frame.Beta >= 360)
                        {
                            frame.Beta = frame.Beta - 360;
                        }

                        else if (frame.Beta < 0)
                        {
                            frame.Beta = frame.Beta + 360;
                        }

                        #endregion

                    }

                    else
                    {
                        continue;
                    }

                    newCreatedColumns.Add(frame);
                }

                List<Frame> newCreatedColumns2 = newCreatedColumns.GroupBy(x => x.Sec).Select(y => y.First()).ToList();


                Dictionary<string, FamilySymbol> dic = RevitUtils.createRvtFamilySectionForAllColumn(newCreatedColumns2, doc, dis);


                #region Create Columns
                object[,] dataToExcel = new object[10, 9];

                if (ModelType == "Etabs")
                {
                    Column.CreateAllColumns3(doc, newCreatedColumns, dis, dic, null, "Staad", ref dataToExcel);
                }
                if (ModelType == "Staad")
                {
                    Column.CreateAllColumns3(doc, newCreatedColumns, dis, dic, null, "Staad", ref dataToExcel);
                }
                else
                {
                    Column.CreateAllColumns3(doc, newCreatedColumns, dis, dic, null, "Staad", ref dataToExcel);
                }
                #endregion


                // delete Beams
                using (Transaction trns = new Transaction(doc, "aa"))
                {
                    trns.Start();
                    selectedBeams.ForEach(x => doc.Delete(x.Id));
                    trns.Commit();
                }


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