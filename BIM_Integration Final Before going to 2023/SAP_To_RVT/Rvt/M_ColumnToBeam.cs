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
    internal class M_ColumnToBeam : IExternalCommand
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


                #region get Selected Columns
                List<ElementId> selectedElementids = uidoc.Selection.GetElementIds().ToList();

                List<FamilyInstance> selectedColumns = new List<FamilyInstance>();

                foreach (ElementId elementId in selectedElementids)
                {
                    Element element = doc.GetElement(elementId);

                    // Check if the element belongs to the Structural Columns or Structural Framing category
                    if ((BuiltInCategory)element.Category.Id.IntegerValue == BuiltInCategory.OST_StructuralColumns)
                    {
                        selectedColumns.Add(element as FamilyInstance);
                    }
                }


                #endregion

                #endregion

                List<Frame> newCreatedBeams = new List<Frame>();

                foreach (var instance in selectedColumns)
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
                        frame.Material = instance.LookupParameter("Material").AsString();
                        frame.FamilyType = RevitUtils.getColumnSecType(instance.Symbol.FamilyName);

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
                        frame.MemberType = MemberType.Column;

                        #region getRotationAngle
                        LocationPoint locationPoint = instance.Location as LocationPoint;
                        double angle = 0.0;

                        if (locationPoint == null)
                        {

                            //frame.Beta = instance.LookupParameter("Cross-Section Rotation").AsDouble() *180/Math.PI ;
                            frame.Beta = instance.LookupParameter("Cross-Section Rotation").AsDouble();
                            frame.Beta = Math.Round((frame.Beta) * 180 / Math.PI, 0);

                            if (frame.Beta >= 360)
                            {
                                frame.Beta = frame.Beta - 360;
                            }

                            else if (frame.Beta < 0)
                            {
                                frame.Beta = frame.Beta + 360;
                            }

                        }


                        else
                        {
                            angle = locationPoint.Rotation;


                            //frame.Beta = (angle *180/Math.PI) ;
                            frame.Beta = Math.Round((angle) * 180 / Math.PI, 0) + 90;


                            if (frame.Beta >= 360)
                            {
                                frame.Beta = frame.Beta - 360;
                            }

                            else if (frame.Beta < 0)
                            {
                                frame.Beta = frame.Beta + 360;
                            }


                        }

                        #endregion

                    }

                    else
                    {
                        continue;
                    }

                    newCreatedBeams.Add(frame);
                }

                List<Frame> newCreatedBeams2 = newCreatedBeams.GroupBy(x => x.Sec).Select(y => y.First()).ToList();


                Dictionary<string, FamilySymbol> dic = RevitUtils.createRvtFamilySectionForAllBeam(newCreatedBeams2, doc, dis);


                #region createBeams
                object[,] dataToExcel = new object[10, 9];

                if (ModelType == "Etabs")
                {
                    Beam.CreateAllBeams3(doc, newCreatedBeams, dis, dic, null, "Etabs", dataToExcel);

                }
                if (ModelType == "Staad")
                {
                    Beam.CreateAllBeams3(doc, newCreatedBeams, dis, dic, null, "Staad", dataToExcel);

                }

                else
                {
                    Beam.CreateAllBeams3(doc, newCreatedBeams, dis, dic, null, "Sap", dataToExcel);

                }
                #endregion




                // delete columns
                using (Transaction trns = new Transaction(doc, "aa"))
                {
                    trns.Start();
                    selectedColumns.ForEach(x => doc.Delete(x.Id));
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