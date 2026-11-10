using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using Microsoft.Office.Interop.Excel;
using Rhino.Geometry;

//using Rhino.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Numerics;
using System.Reflection;
using Color = Autodesk.Revit.DB.Color;
using Curve = Autodesk.Revit.DB.Curve;
using Excel = Microsoft.Office.Interop.Excel;
using Floor = Autodesk.Revit.DB.Floor;
using Line = Autodesk.Revit.DB.Line;
using Outline = Autodesk.Revit.DB.Outline;
using Parameter = Autodesk.Revit.DB.Parameter;

namespace BIM_Integration
{
    public static class RevitUtils
    {
        public static List<string> Log = new List<string>();
        private static Random random = new Random();

        public static List<Frame> getRevitFrames(UIDocument uidoc, Document doc, ForgeTypeId dis, ref List<Frame> Beams, ref List<Frame> Columns, string anaylsisModel)
        {



            AnalyticalToPhysicalAssociationManager analyManger = AnalyticalToPhysicalAssociationManager.GetAnalyticalToPhysicalAssociationManager(doc);

            using (Transaction allTrns = new Transaction(doc, "aaa"))
            {
                try
                {

                    allTrns.Start();

                    #region Revit Custom Error Handling
                    FailureHandlingOptions failOpt = allTrns.GetFailureHandlingOptions();
                    failOpt.SetFailuresPreprocessor(new RoomWarningSwallower());
                    allTrns.SetFailureHandlingOptions(failOpt);
                    #endregion

                    #region get Selected Elements 
                    ICollection<ElementId> selectedIds = uidoc.Selection.GetElementIds();

                    // Create a filter for Structural Framing and Structural Beams categories
                    List<BuiltInCategory> categories = new List<BuiltInCategory>();
                    categories.Add(BuiltInCategory.OST_StructuralFraming);
                    categories.Add(BuiltInCategory.OST_StructuralColumns);

                    // Create an ElementCategoryFilter for each category and combine them with a LogicalOrFilter
                    List<ElementFilter> categoryFilters = new List<ElementFilter>();
                    foreach (BuiltInCategory category in categories)
                    {
                        categoryFilters.Add(new ElementCategoryFilter(category));
                    }

                    LogicalOrFilter orFilter = new LogicalOrFilter(categoryFilters);

                    // Get the selected elements that match the filter
                    List<Element> selectedElements = new List<Element>();
                    foreach (ElementId id in selectedIds)
                    {
                        Element element = uidoc.Document.GetElement(id);
                        if (orFilter.PassesFilter(element))
                        {
                            selectedElements.Add(element);
                        }
                    }

                    List<FamilyInstance> selectedStrucralFraming = selectedElements.Where(x => x.Category.Id.IntegerValue == (int)BuiltInCategory.OST_StructuralFraming).Cast<FamilyInstance>().ToList();
                    List<FamilyInstance> selectedStrucralColumns = selectedElements.Where(x => x.Category.Id.IntegerValue == (int)BuiltInCategory.OST_StructuralColumns).Cast<FamilyInstance>().ToList();


                    #endregion

                    List<Frame> beamFarames = new List<Frame>();
                    List<Frame> columnFarames = new List<Frame>();
                    List<Frame> frames = new List<Frame>();

                    List<FamilyInstance> Allbeams = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralFraming).WhereElementIsNotElementType().Cast<FamilyInstance>().ToList();
                    List<FamilyInstance> Allcolumns = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralColumns).WhereElementIsNotElementType().Cast<FamilyInstance>().ToList();

                    if (selectedStrucralFraming.Count > 0 || selectedStrucralColumns.Count > 0)
                    {
                        Allbeams = selectedStrucralFraming;
                        Allcolumns = selectedStrucralColumns;
                    }

                    #region get data of beams

                    #region handling curvedBeams and null analytical Members (not finished yet)
                    List<FamilyInstance> strightBeams = new List<FamilyInstance>();
                    List<FamilyInstance> curvedBeams = new List<FamilyInstance>();
                    List<FamilyInstance> StrightElementsSplittedFromCurved = new List<FamilyInstance>();

                    foreach (var instance in Allbeams)
                    {


                        AnalyticalElement analytInstance = doc.GetElement(analyManger.GetAssociatedElementId(instance.Id)) as AnalyticalElement;
                        if (analytInstance != null)
                        {
                            Point startPoint = new Point(analytInstance.GetCurve().GetEndPoint(0));
                            Point endPoint = new Point(analytInstance.GetCurve().GetEndPoint(1));


                            Vector3 vec1 = new Vector3((float)(endPoint.X - startPoint.X), (float)(endPoint.Y - startPoint.Y), (float)(endPoint.Z - startPoint.Z));
                            vec1 = Vector3.Normalize(vec1);

                            LocationCurve locCurve = instance.Location as LocationCurve;
                            Curve curve = locCurve.Curve;
                            double totalLength = curve.Length;
                            Point midPoint = new Point(curve.Evaluate(totalLength / 2, false));


                            Vector3 vec2 = new Vector3((float)(midPoint.X - startPoint.X), (float)(midPoint.Y - startPoint.Y), (float)(midPoint.Z - startPoint.Z));
                            vec2 = Vector3.Normalize(vec2);


                            if (vec1.X == vec2.X && vec1.Y == vec2.Y && vec1.Z == vec2.Z)
                            {
                                strightBeams.Add(instance);

                            }

                            else
                            {
                                if (UnitUtils.ConvertFromInternalUnits((instance.Location as LocationCurve).Curve.Length, dis) > 6000)
                                {
                                    curvedBeams.Add(instance);
                                }
                                else
                                {
                                    strightBeams.Add(instance);
                                }

                            }

                        }
                        else
                        {


                            Curve c = (instance.Location as LocationCurve).Curve;
                            AnalyticalMember analyMember = AnalyticalMember.Create(doc, c);
                            analyManger.AddAssociation(analyMember.Id, instance.Id);
                            strightBeams.Add(instance);
                        }
                    }

                    // get stright and curved beams


                    // split curved beams to straight beams 
                    foreach (var instance in curvedBeams)
                    {
                        //SapRevitFrame.createView(curvedBeams.Select(x => x.Id).ToList(), doc, uidoc, "curvedBeams");

                        //return null;

                        StrightElementsSplittedFromCurved.Add(instance);




                    }



                    // union between  straightBeams and StraightBeamsFromSplittedCurvedBeams
                    strightBeams = strightBeams.Union(StrightElementsSplittedFromCurved).ToList();
                    #endregion

                    #region get Data Of Beams after handling curvedBeams

                    foreach (var instance in strightBeams.AsParallel())
                    {
                        try
                        {
                            Frame frame = new Frame();



                            AnalyticalElement analytInstance = doc.GetElement(analyManger.GetAssociatedElementId(instance.Id)) as AnalyticalElement;
                            if (analytInstance != null)
                            {

                                //Point startPoint = new Point(analytInstance.GetCurve().GetEndPoint(0));
                                //Point endPoint = new Point(analytInstance.GetCurve().GetEndPoint(1));

                                Point startPoint = new Point((instance.Location as LocationCurve).Curve.GetEndPoint(0));
                                Point endPoint = new Point((instance.Location as LocationCurve).Curve.GetEndPoint(1));



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




                                frame.Material = instance.LookupParameter(HardCodedData.MaterialParameter).AsString();

                                if (anaylsisModel == "Staad")
                                {
                                    frame.ID = instance.LookupParameter(HardCodedData.StaadIDParameter).AsString();
                                }
                                else if (anaylsisModel == "Sap")
                                {
                                    frame.ID = instance.LookupParameter(HardCodedData.SapIDParameter).AsString();

                                }
                                else if (anaylsisModel == "Etabs")
                                {
                                    frame.ID = instance.LookupParameter(HardCodedData.EtabsIDParameter).AsString();

                                }

                                frame.SubGroup = instance.LookupParameter(HardCodedData.RevitSubGroup).AsString();
                                frame.MainGroup = instance.LookupParameter(HardCodedData.RevitMainGroup).AsString();
                                frame.SubGroupUsage = instance.LookupParameter(HardCodedData.RevitSubGroupUsage).AsString();



                                frame.Beta = instance.LookupParameter("Cross-Section Rotation").AsDouble() * 180 / (Math.PI);
                                frame.RvtFamily = instance.Symbol.FamilyName;
                                frame.FamilyType = getBeamSecType(instance.Symbol.FamilyName);
                                frame.RvtName = (instance.Symbol.Name).Replace("x", "X");


                                Point.getMemberType(frame);

                                beamFarames.Add(frame);
                            }
                            // This handeled above in Allbeams loop
                            else
                            {
                                continue;
                            }
                        }


                        catch (Exception ex2)
                        {

                            System.Diagnostics.StackTrace trace = new System.Diagnostics.StackTrace(ex2, true);

                            RevitUtils.Log.Add("Error in getRevitFrames Method in the Beams part... BeamInstance no" + instance.Id + ex2.Message + trace.GetFrame(trace.FrameCount - 1).GetFileLineNumber());

                            // TaskDialog.Show("ErrorInGetRevitFramesMethod", "ErrorInRevitBeamID= "+ instance.Id + ex2.Message + trace.GetFrame(trace.FrameCount - 1).GetFileLineNumber());
                            continue;
                        }





                    }
                    #endregion

                    #endregion

                    #region get Data of columns
                    foreach (var instance in Allcolumns.AsParallel())
                    {
                        try
                        {
                            Frame frame = new Frame();


                            AnalyticalElement analytInstance = doc.GetElement(analyManger.GetAssociatedElementId(instance.Id)) as AnalyticalElement;
                            if (analytInstance == null)
                            {
                                // verticalColumn
                                LocationPoint locationPoint = instance.Location as LocationPoint;
                                if (locationPoint != null)
                                {
                                    XYZ columnPoint = locationPoint.Point;

                                    Level baseLevel = (doc.GetElement(instance.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_PARAM).AsElementId()) as Level);
                                    double baseLevelElev = baseLevel.ProjectElevation;
                                    double baseLevelOffset = instance.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_OFFSET_PARAM).AsDouble();
                                    double startptElev = baseLevelOffset + baseLevelElev;
                                    XYZ startPoint = new XYZ(columnPoint.X, columnPoint.Y, startptElev);

                                    Level TopLevel = (doc.GetElement(instance.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_PARAM).AsElementId()) as Level);
                                    double TopLevelElev = TopLevel.ProjectElevation;
                                    double topLevelOffset = instance.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_OFFSET_PARAM).AsDouble();
                                    double endPtElev = topLevelOffset + TopLevelElev;
                                    XYZ endPoint = new XYZ(columnPoint.X, columnPoint.Y, endPtElev);

                                    Line c = Line.CreateBound(startPoint, endPoint);

                                    AnalyticalMember analyMember = AnalyticalMember.Create(doc, c as Curve);

                                    analyManger.AddAssociation(analyMember.Id, instance.Id);

                                    analytInstance = analyMember;
                                }
                                else
                                {
                                    // slantted Column
                                    Curve locationCurve = (instance.Location as LocationCurve).Curve;

                                    AnalyticalMember analyMember = AnalyticalMember.Create(doc, locationCurve as Curve);

                                    analyManger.AddAssociation(analyMember.Id, instance.Id);

                                    analytInstance = analyMember;
                                }
                            }

                            if (analytInstance.GetCurve() != null)
                            {
                                //Point startPoint = new Point(analytInstance.GetCurve().GetEndPoint(0));
                                //Point endPoint = new Point(analytInstance.GetCurve().GetEndPoint(1));

                                Point startPoint = new Point();
                                Point endPoint = new Point();

                                GetTopAndBottomPoints(doc, instance.Id, ref startPoint, ref endPoint);

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



                                frame.Material = instance.LookupParameter(HardCodedData.MaterialParameter).AsString();

                                if (anaylsisModel == "Staad")
                                {
                                    frame.ID = instance.LookupParameter(HardCodedData.StaadIDParameter).AsString();
                                }
                                else if (anaylsisModel == "Sap")
                                {
                                    frame.ID = instance.LookupParameter(HardCodedData.SapIDParameter).AsString();

                                }
                                else if (anaylsisModel == "Etabs")
                                {
                                    frame.ID = instance.LookupParameter(HardCodedData.EtabsIDParameter).AsString();

                                }


                                frame.SubGroup = instance.LookupParameter(HardCodedData.RevitSubGroup).AsString();
                                frame.MainGroup = instance.LookupParameter(HardCodedData.RevitMainGroup).AsString();
                                frame.SubGroupUsage = instance.LookupParameter(HardCodedData.RevitSubGroupUsage).AsString();

                                frame.Group = instance.LookupParameter(HardCodedData.RevitSubGroup).AsString();


                                frame.FamilyType = getColumnSecType(instance.Symbol.FamilyName);
                                frame.RvtFamily = instance.Symbol.FamilyName;
                                frame.RvtName = (instance.Symbol.Name).Replace("x", "X"); ;
                                Point.getMemberType(frame);

                                #region getRotationAngle
                                LocationPoint locationPoint = instance.Location as LocationPoint;
                                double angle = 0.0;
                                if (locationPoint == null)
                                {
                                    frame.Beta = instance.LookupParameter("Cross-Section Rotation").AsDouble() * 180 / (Math.PI);
                                }
                                else
                                {
                                    angle = locationPoint.Rotation;
                                    frame.Beta = (angle * 180 / Math.PI) + 90;
                                }

                                #endregion

                                columnFarames.Add(frame);
                            }
                        }
                        catch (Exception ex)
                        {
                        }
                    }

                    #endregion

                    frames = beamFarames.Union(columnFarames).ToList();
                    Beams = beamFarames;
                    Columns = columnFarames;
                    //addRevitID(doc, frames);

                    allTrns.Commit();


#if AEC_Revit_2023

#else
                    // To return the Revit Elements in the Oringial Co-Ordinate System.
                    VectorHelper.GetTransformedFramesForRotationOnly(ref frames, HardCodedData.AngleValueInDegree, UseTheInverse: true);
#endif


                    return frames;
                }
                catch (Exception ex)
                {

                    throw;
                }
            }
        }

        public static void GetTopAndBottomPoints(Document doc, ElementId columnId, ref Point topPt, ref Point bottomPt)
        {
            // Get the structural column element by its Id
            Element columnElement = doc.GetElement(columnId);

            // Get the FamilyInstance of the structural column
            FamilyInstance columnInstance = columnElement as FamilyInstance;


            // Retrieve the geometry of the column
            Options geomOptions = new Options();

            GeometryElement geomElement = columnInstance.GetOriginalGeometry(geomOptions);



            // Iterate through the geometry objects and process them
            foreach (GeometryObject geomObj in geomElement)
            {
                if (geomObj is Solid solid)
                {
                    // Process the solid geometry to find top and bottom faces
                    foreach (Face face in solid.Faces)
                    {
                        PlanarFace planarFace = face as PlanarFace;
                        if (planarFace != null)
                        {
                            // Check if the face normal is pointing up (top face) or down (bottom face)
                            XYZ normal = planarFace.FaceNormal;
                            if (normal.IsAlmostEqualTo(XYZ.BasisZ))
                            {
                                // This is the top face
                                BoundingBoxUV bbox = planarFace.GetBoundingBox();
                                UV midPoint = (bbox.Min + bbox.Max) / 2;
                                XYZ topPoint = planarFace.Evaluate(midPoint);


                                if (topPoint.Z > topPt.Z)
                                {

                                    topPt.X = topPoint.X;
                                    topPt.Y = topPoint.Y;
                                    topPt.Z = topPoint.Z;

                                }
                            }

                            else if (normal.IsAlmostEqualTo(-XYZ.BasisZ))
                            {
                                // This is the bottom face
                                BoundingBoxUV bbox = planarFace.GetBoundingBox();
                                UV midPoint = (bbox.Min + bbox.Max) / 2;
                                XYZ bottomPoint = planarFace.Evaluate(midPoint);
                                if (bottomPoint.Z < bottomPt.Z)
                                {
                                    bottomPt.X = bottomPoint.X;
                                    bottomPt.Y = bottomPoint.Y;
                                    bottomPt.Z = bottomPoint.Z;
                                }
                            }
                        }
                    }
                }


                else if (geomObj is GeometryInstance geomInstance)
                {
                    // If it's a geometry instance, get the instance geometry and process similarly
                    GeometryElement instanceGeometry = geomInstance.GetInstanceGeometry();
                    foreach (GeometryObject instanceGeomObj in instanceGeometry)
                    {
                        if (instanceGeomObj is Solid instanceSolid)
                        {
                            foreach (Face face in instanceSolid.Faces)
                            {
                                PlanarFace planarFace = face as PlanarFace;
                                if (planarFace != null)
                                {
                                    XYZ normal = planarFace.FaceNormal;
                                    if (normal.IsAlmostEqualTo(XYZ.BasisZ))
                                    {
                                        BoundingBoxUV bbox = planarFace.GetBoundingBox();
                                        UV midPoint = (bbox.Min + bbox.Max) / 2;
                                        XYZ topPoint = planarFace.Evaluate(midPoint);
                                        if (topPoint.Z > topPt.Z)
                                        {

                                            topPt.X = topPoint.X;
                                            topPt.Y = topPoint.Y;
                                            topPt.Z = topPoint.Z;

                                        }
                                    }


                                    else if (normal.IsAlmostEqualTo(-XYZ.BasisZ))
                                    {
                                        BoundingBoxUV bbox = planarFace.GetBoundingBox();
                                        UV midPoint = (bbox.Min + bbox.Max) / 2;
                                        XYZ bottomPoint = planarFace.Evaluate(midPoint);
                                        if (bottomPoint.Z < bottomPt.Z)
                                        {
                                            bottomPt.X = bottomPoint.X;
                                            bottomPt.Y = bottomPoint.Y;
                                            bottomPt.Z = bottomPoint.Z;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }

        }

        public static List<Frame> getRevitFrames2(UIDocument uidoc, Document doc, ForgeTypeId dis, ref List<Frame> Beams, ref List<Frame> Columns, string anaylsisModel)
        {



            AnalyticalToPhysicalAssociationManager analyManger = AnalyticalToPhysicalAssociationManager.GetAnalyticalToPhysicalAssociationManager(doc);

            using (Transaction allTrns = new Transaction(doc, "aaa"))
            {
                try
                {

                    allTrns.Start();

                    #region Revit Custom Error Handling
                    FailureHandlingOptions failOpt = allTrns.GetFailureHandlingOptions();
                    failOpt.SetFailuresPreprocessor(new RoomWarningSwallower());
                    allTrns.SetFailureHandlingOptions(failOpt);
                    #endregion

                    #region get Selected Elements 
                    ICollection<ElementId> selectedIds = uidoc.Selection.GetElementIds();

                    // Create a filter for Structural Framing and Structural Beams categories
                    List<BuiltInCategory> categories = new List<BuiltInCategory>();
                    categories.Add(BuiltInCategory.OST_StructuralFraming);
                    categories.Add(BuiltInCategory.OST_StructuralColumns);

                    // Create an ElementCategoryFilter for each category and combine them with a LogicalOrFilter
                    List<ElementFilter> categoryFilters = new List<ElementFilter>();
                    foreach (BuiltInCategory category in categories)
                    {
                        categoryFilters.Add(new ElementCategoryFilter(category));
                    }

                    LogicalOrFilter orFilter = new LogicalOrFilter(categoryFilters);

                    // Get the selected elements that match the filter
                    List<Element> selectedElements = new List<Element>();
                    foreach (ElementId id in selectedIds)
                    {
                        Element element = uidoc.Document.GetElement(id);
                        if (orFilter.PassesFilter(element))
                        {
                            selectedElements.Add(element);
                        }
                    }

                    List<FamilyInstance> selectedStrucralFraming = selectedElements.Where(x => x.Category.Id.IntegerValue == (int)BuiltInCategory.OST_StructuralFraming).Cast<FamilyInstance>().ToList();
                    List<FamilyInstance> selectedStrucralColumns = selectedElements.Where(x => x.Category.Id.IntegerValue == (int)BuiltInCategory.OST_StructuralColumns).Cast<FamilyInstance>().ToList();


                    #endregion

                    List<Frame> beamFarames = new List<Frame>();
                    List<Frame> columnFarames = new List<Frame>();
                    List<Frame> frames = new List<Frame>();

                    List<FamilyInstance> Allbeams = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralFraming).WhereElementIsNotElementType().Cast<FamilyInstance>().ToList();
                    List<FamilyInstance> Allcolumns = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralColumns).WhereElementIsNotElementType().Cast<FamilyInstance>().ToList();

                    if (selectedStrucralFraming.Count > 0 || selectedStrucralColumns.Count > 0)
                    {
                        Allbeams = selectedStrucralFraming;
                        Allcolumns = selectedStrucralColumns;
                    }

                    #region get data of beams

                    #region handling curvedBeams and null analytical Members (not finished yet)
                    List<FamilyInstance> strightBeams = new List<FamilyInstance>();
                    List<FamilyInstance> curvedBeams = new List<FamilyInstance>();
                    List<FamilyInstance> StrightElementsSplittedFromCurved = new List<FamilyInstance>();

                    foreach (var instance in Allbeams)
                    {


                        AnalyticalElement analytInstance = doc.GetElement(analyManger.GetAssociatedElementId(instance.Id)) as AnalyticalElement;
                        if (analytInstance != null)
                        {
                            Point startPoint = new Point(analytInstance.GetCurve().GetEndPoint(0));
                            Point endPoint = new Point(analytInstance.GetCurve().GetEndPoint(1));


                            Vector3 vec1 = new Vector3((float)(endPoint.X - startPoint.X), (float)(endPoint.Y - startPoint.Y), (float)(endPoint.Z - startPoint.Z));
                            vec1 = Vector3.Normalize(vec1);

                            LocationCurve locCurve = instance.Location as LocationCurve;
                            Curve curve = locCurve.Curve;
                            double totalLength = curve.Length;
                            Point midPoint = new Point(curve.Evaluate(totalLength / 2, false));


                            Vector3 vec2 = new Vector3((float)(midPoint.X - startPoint.X), (float)(midPoint.Y - startPoint.Y), (float)(midPoint.Z - startPoint.Z));
                            vec2 = Vector3.Normalize(vec2);


                            if (vec1.X == vec2.X && vec1.Y == vec2.Y && vec1.Z == vec2.Z)
                            {
                                strightBeams.Add(instance);

                            }

                            else
                            {
                                if (UnitUtils.ConvertFromInternalUnits((instance.Location as LocationCurve).Curve.Length, dis) > 6000)
                                {
                                    curvedBeams.Add(instance);
                                }
                                else
                                {
                                    strightBeams.Add(instance);
                                }

                            }

                        }
                        else
                        {


                            Curve c = (instance.Location as LocationCurve).Curve;
                            AnalyticalMember analyMember = AnalyticalMember.Create(doc, c);
                            analyManger.AddAssociation(analyMember.Id, instance.Id);
                            strightBeams.Add(instance);
                        }
                    }

                    // get stright and curved beams


                    // split curved beams to straight beams 
                    foreach (var instance in curvedBeams)
                    {
                        //SapRevitFrame.createView(curvedBeams.Select(x => x.Id).ToList(), doc, uidoc, "curvedBeams");

                        //return null;

                        StrightElementsSplittedFromCurved.Add(instance);




                    }



                    // union between  straightBeams and StraightBeamsFromSplittedCurvedBeams
                    strightBeams = strightBeams.Union(StrightElementsSplittedFromCurved).ToList();
                    #endregion

                    #region get Data Of Beams after handling curvedBeams

                    foreach (var instance in strightBeams.AsParallel())
                    {
                        try
                        {
                            Frame frame = new Frame();

                            // frame.revitCurve = analytInstance.GetCurve();
                            frame.revitCurve = (instance.Location as LocationCurve).Curve;

                            frame.RhinoCurve = RhinoInside.Revit.Convert.Geometry.GeometryDecoder.ToCurve(frame.revitCurve);

                            Point startPoint = new Point(frame.revitCurve.GetEndPoint(0));
                            Point endPoint = new Point(frame.revitCurve.GetEndPoint(1));

                            startPoint.X = UnitUtils.ConvertFromInternalUnits(startPoint.X, dis);
                            startPoint.Y = UnitUtils.ConvertFromInternalUnits(startPoint.Y, dis);
                            startPoint.Z = UnitUtils.ConvertFromInternalUnits(startPoint.Z, dis);

                            endPoint.X = UnitUtils.ConvertFromInternalUnits(endPoint.X, dis);
                            endPoint.Y = UnitUtils.ConvertFromInternalUnits(endPoint.Y, dis);
                            endPoint.Z = UnitUtils.ConvertFromInternalUnits(endPoint.Z, dis);

                            Point midPoint = new Point(frame.revitCurve.Evaluate(.5, true));

                            midPoint.X = UnitUtils.ConvertFromInternalUnits(midPoint.X, dis);
                            midPoint.Y = UnitUtils.ConvertFromInternalUnits(midPoint.Y, dis);
                            midPoint.Z = UnitUtils.ConvertFromInternalUnits(midPoint.Z, dis);



                            frame.StartPoint = startPoint;
                            frame.EndPoint = endPoint;
                            frame.MidPoint = midPoint;


                            frame.RvtID = instance.Id.ToString();
                            frame.revitEleId = instance.Id;
                            frame.Length = Point.getDistance(startPoint, endPoint);
                            frame.Sec = instance.Symbol.Name;


                            frame.Material = instance.LookupParameter(HardCodedData.MaterialParameter).AsString();

                            if (anaylsisModel == "Staad")
                            {
                                frame.ID = instance.LookupParameter(HardCodedData.StaadIDParameter).AsString();
                            }
                            else if (anaylsisModel == "Sap")
                            {
                                frame.ID = instance.LookupParameter(HardCodedData.SapIDParameter).AsString();

                            }
                            else if (anaylsisModel == "Etabs")
                            {
                                frame.ID = instance.LookupParameter(HardCodedData.EtabsIDParameter).AsString();

                            }

                            frame.SubGroup = instance.LookupParameter(HardCodedData.RevitSubGroup).AsString();
                            frame.MainGroup = instance.LookupParameter(HardCodedData.RevitMainGroup).AsString();
                            frame.SubGroupUsage = instance.LookupParameter(HardCodedData.RevitSubGroupUsage).AsString();



                            frame.Beta = instance.LookupParameter("Cross-Section Rotation").AsDouble() * 180 / (Math.PI);
                            frame.RvtFamily = instance.Symbol.FamilyName;
                            frame.FamilyType = getBeamSecType(instance.Symbol.FamilyName);
                            frame.RvtName = (instance.Symbol.Name).Replace("x", "X");


                            Point.getMemberType(frame);

                            beamFarames.Add(frame);

                        }


                        catch (Exception ex2)
                        {

                            System.Diagnostics.StackTrace trace = new System.Diagnostics.StackTrace(ex2, true);

                            RevitUtils.Log.Add("Error in getRevitFrames Method in the Beams part... BeamInstance no" + instance.Id + ex2.Message + trace.GetFrame(trace.FrameCount - 1).GetFileLineNumber());

                            // TaskDialog.Show("ErrorInGetRevitFramesMethod", "ErrorInRevitBeamID= "+ instance.Id + ex2.Message + trace.GetFrame(trace.FrameCount - 1).GetFileLineNumber());
                            continue;
                        }





                    }
                    #endregion

                    #endregion

                    #region get Data of columns
                    foreach (var instance in Allcolumns.AsParallel())
                    {
                        try
                        {
                            Frame frame = new Frame();


                            AnalyticalElement analytInstance = doc.GetElement(analyManger.GetAssociatedElementId(instance.Id)) as AnalyticalElement;
                            if (analytInstance == null)
                            {
                                // verticalColumn
                                LocationPoint locationPoint = instance.Location as LocationPoint;
                                if (locationPoint != null)
                                {
                                    XYZ columnPoint = locationPoint.Point;

                                    Level baseLevel = (doc.GetElement(instance.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_PARAM).AsElementId()) as Level);
                                    double baseLevelElev = baseLevel.Elevation;
                                    double baseLevelOffset = instance.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_OFFSET_PARAM).AsDouble();
                                    double startptElev = baseLevelOffset + baseLevelElev;
                                    XYZ startPoint = new XYZ(columnPoint.X, columnPoint.Y, startptElev);

                                    Level TopLevel = (doc.GetElement(instance.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_PARAM).AsElementId()) as Level);
                                    double TopLevelElev = TopLevel.Elevation;
                                    double topLevelOffset = instance.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_OFFSET_PARAM).AsDouble();
                                    double endPtElev = topLevelOffset + TopLevelElev;
                                    XYZ endPoint = new XYZ(columnPoint.X, columnPoint.Y, endPtElev);

                                    Line c = Line.CreateBound(startPoint, endPoint);

                                    AnalyticalMember analyMember = AnalyticalMember.Create(doc, c as Curve);

                                    analyManger.AddAssociation(analyMember.Id, instance.Id);

                                    analytInstance = analyMember;
                                }
                                else
                                {
                                    // slantted Column
                                    Curve locationCurve = (instance.Location as LocationCurve).Curve;

                                    AnalyticalMember analyMember = AnalyticalMember.Create(doc, locationCurve as Curve);

                                    analyManger.AddAssociation(analyMember.Id, instance.Id);

                                    analytInstance = analyMember;
                                }
                            }


                            if (analytInstance.GetCurve() != null)
                            {

                                frame.revitCurve = analytInstance.GetCurve();

                                frame.RhinoCurve = RhinoInside.Revit.Convert.Geometry.GeometryDecoder.ToCurve(frame.revitCurve);

                                #region gePoints
                                Point3d stPt = RhinoInside.Revit.Convert.Geometry.GeometryDecoder.ToPoint3d(analytInstance.GetCurve().GetEndPoint(0));
                                Point3d endPt = RhinoInside.Revit.Convert.Geometry.GeometryDecoder.ToPoint3d(analytInstance.GetCurve().GetEndPoint(1));


                                Point stPoint = new Point(stPt);
                                Point endPoint = new Point(endPt);

                                Point mid = Point.getMidPoint(stPoint, endPoint);

                                frame.StartPoint = stPoint;
                                frame.EndPoint = endPoint;
                                frame.MidPoint = mid;

                                frame.Length = Point.getDistance(stPoint, endPoint);

                                #endregion
                                frame.RvtID = instance.Id.ToString();
                                frame.revitEleId = instance.Id;
                                frame.ID = instance.Id.ToString();


                                #region Parameters
                                frame.Sec = instance.Symbol.Name;
                                frame.RvtFamily = instance.Symbol.FamilyName;
                                frame.RvtName = instance.Symbol.Name;
                                frame.RvtName = (instance.Symbol.Name).Replace("x", "X"); ;

                                frame.MemberType = MemberType.Column;


                                frame.Material = instance.LookupParameter(HardCodedData.MaterialParameter).AsString();

                                if (anaylsisModel == "Staad")
                                {
                                    frame.ID = instance.LookupParameter(HardCodedData.StaadIDParameter).AsString();
                                }
                                else if (anaylsisModel == "Sap")
                                {
                                    frame.ID = instance.LookupParameter(HardCodedData.SapIDParameter).AsString();

                                }
                                else if (anaylsisModel == "Etabs")
                                {
                                    frame.ID = instance.LookupParameter(HardCodedData.EtabsIDParameter).AsString();

                                }


                                frame.SubGroup = instance.LookupParameter(HardCodedData.RevitSubGroup).AsString();
                                frame.MainGroup = instance.LookupParameter(HardCodedData.RevitMainGroup).AsString();
                                frame.SubGroupUsage = instance.LookupParameter(HardCodedData.RevitSubGroupUsage).AsString();
                                frame.Group = instance.LookupParameter(HardCodedData.RevitSubGroup).AsString();

                                #endregion

                                #region getRotationAngle
                                LocationPoint locationPoint = instance.Location as LocationPoint;
                                double angle = 0.0;
                                if (locationPoint == null)
                                {
                                    frame.Beta = instance.LookupParameter("Cross-Section Rotation").AsDouble() * 180 / (Math.PI);
                                }
                                else
                                {
                                    angle = locationPoint.Rotation;
                                    frame.Beta = (angle * 180 / Math.PI) + 90;
                                }

                                #endregion

                                columnFarames.Add(frame);
                            }
                        }
                        catch (Exception ex)
                        {
                            continue;
                        }
                    }

                    #endregion

                    frames = beamFarames.Union(columnFarames).ToList();
                    Beams = beamFarames;
                    Columns = columnFarames;
                    //addRevitID(doc, frames);

                    allTrns.Commit();


                    return frames;
                }


                catch (Exception ex)
                {
                    throw;
                }
            }
        }


        public static void getRevitAreas(UIDocument uidoc, Document doc, ForgeTypeId dis, ref List<Area> areas)
        {

            // Collect structural floors
            List<Floor> structuralFloors = new List<Floor>();
            List<Floor> archFloors = new List<Floor>();

            FilteredElementCollector collector = new FilteredElementCollector(doc);
            ICollection<Element> elementsOfType = collector.OfClass(typeof(Floor)).ToElements();
            foreach (Element elem in elementsOfType)
            {

                Floor floor = elem as Floor;
                Area a = new Area();

                AnalyticalToPhysicalAssociationManager analyManger = AnalyticalToPhysicalAssociationManager.GetAnalyticalToPhysicalAssociationManager(doc);
                // may be needed to be handled 
                AnalyticalElement analyModel = doc.GetElement(analyManger.GetAssociatedElementId(elem.Id)) as AnalyticalElement;


                if (analyModel != null)
                {
                    structuralFloors.Add(floor);

                    a.Sec = floor.FloorType.ToString();
                    a.revitEleId = floor.Id;

                    Sketch s = doc.GetElement(floor.SketchId) as Sketch;

                    CurveArrArray crvArray = s.Profile;
                    a.crvArray = crvArray;


                    List<Rhino.Geometry.Curve> rhinoCrvs = RhinoInside.Revit.Convert.Geometry.GeometryDecoder.ToCurveMany(crvArray).ToList();
                    a.crvList = rhinoCrvs;


                    areas.Add(a);
                }

                else
                {
                    archFloors.Add(floor);
                }
            }


        }


        public static SectType getBeamSecType(string rvtFamilyName)

        {
            SectType beamSecType = SectType.BM_BUS_I;

            #region getFamilyType

            if (rvtFamilyName == BM_2L_EQUALANGLE.FamilyName)
            {
                beamSecType = SectType.BM_2L_EQUALANGLE;
            }
            else if (rvtFamilyName == BM_BUS_I.FamilyName)
            {
                beamSecType = SectType.BM_BUS_I;
            }
            else if (rvtFamilyName == BM_BUS_I_Tapered.FamilyName)
            {
                beamSecType = SectType.BM_BUS_I_Tapered;
            }
            else if (rvtFamilyName == BM_BUS_PIPE_Tapered.FamilyName)
            {
                beamSecType = SectType.BM_BUS_PIPE_Tapered;
            }
            else if (rvtFamilyName == BM_BUS_TUBE_Tapered.FamilyName)
            {
                beamSecType = SectType.BM_BUS_TUBE_Tapered;
            }
            else if (rvtFamilyName == BM_HEA.FamilyName)
            {
                beamSecType = SectType.BM_HEA;
            }
            else if (rvtFamilyName == BM_HEAA.FamilyName)
            {
                beamSecType = SectType.BM_HEAA;
            }
            else if (rvtFamilyName == BM_HEB.FamilyName)
            {
                beamSecType = SectType.BM_HEB;
            }
            else if (rvtFamilyName == BM_HEM.FamilyName)
            {
                beamSecType = SectType.BM_HEM;
            }
            else if (rvtFamilyName == BM_IPE.FamilyName)
            {
                beamSecType = SectType.BM_IPE;
            }
            else if (rvtFamilyName == BM_IPEA.FamilyName)
            {
                beamSecType = SectType.BM_IPEA;
            }
            else if (rvtFamilyName == BM_IPEO.FamilyName)
            {
                beamSecType = SectType.BM_IPEO;
            }
            else if (rvtFamilyName == BM_L_EQUAL.FamilyName)
            {
                beamSecType = SectType.BM_L_EQUAL;
            }
            else if (rvtFamilyName == BM_L_UNEQUAL.FamilyName)
            {
                beamSecType = SectType.BM_L_UNEQUAL;
            }
            else if (rvtFamilyName == BM_PIPE.FamilyName)
            {
                beamSecType = SectType.BM_PIPE;
            }
            else if (rvtFamilyName == BM_TUBE.FamilyName)
            {
                beamSecType = SectType.BM_TUBE;
            }

            else if (rvtFamilyName == BM_UB.FamilyName)
            {
                beamSecType = SectType.BM_UB;
            }

            else if (rvtFamilyName == BM_UC.FamilyName)
            {
                beamSecType = SectType.BM_UC;
            }

            else if (rvtFamilyName == BM_UPN.FamilyName)
            {
                beamSecType = SectType.BM_UPN;
            }


            else if (rvtFamilyName == BM_W.FamilyName)
            {
                beamSecType = SectType.BM_W;
            }


            #endregion


            return beamSecType;
        }

        public static SectType getColumnSecType(string rvtFamilyName)

        {
            SectType columnSecType = SectType.BM_BUS_I;

            #region getFamilyType


            if (rvtFamilyName == COL_2L_EQUALANGLE.FamilyName)
            {
                columnSecType = SectType.COL_2L_EQUALANGLE;
            }
            else if (rvtFamilyName == COL_BUS_I.FamilyName)
            {
                columnSecType = SectType.COL_BUS_I;
            }
            else if (rvtFamilyName == COL_BUS_I_Tapered.FamilyName)
            {
                columnSecType = SectType.COL_BUS_I_Tapered;
            }
            else if (rvtFamilyName == COL_BUS_PIPE_Tapered.FamilyName)
            {
                columnSecType = SectType.COL_BUS_PIPE_Tapered;
            }
            else if (rvtFamilyName == COL_BUS_TUBE_Tapered.FamilyName)
            {
                columnSecType = SectType.COL_BUS_TUBE_Tapered;
            }
            else if (rvtFamilyName == COL_HEA.FamilyName)
            {
                columnSecType = SectType.COL_HEA;
            }
            else if (rvtFamilyName == COL_HEAA.FamilyName)
            {
                columnSecType = SectType.COL_HEAA;
            }
            else if (rvtFamilyName == COL_HEB.FamilyName)
            {
                columnSecType = SectType.COL_HEB;
            }
            else if (rvtFamilyName == COL_HEM.FamilyName)
            {
                columnSecType = SectType.COL_HEM;
            }
            else if (rvtFamilyName == COL_IPE.FamilyName)
            {
                columnSecType = SectType.COL_IPE;
            }
            else if (rvtFamilyName == COL_IPEA.FamilyName)
            {
                columnSecType = SectType.COL_IPEA;
            }
            else if (rvtFamilyName == COL_IPEO.FamilyName)
            {
                columnSecType = SectType.COL_IPEO;
            }
            else if (rvtFamilyName == COL_L_EQUAL.FamilyName)
            {
                columnSecType = SectType.COL_L_EQUAL;
            }
            else if (rvtFamilyName == COL_L_UNEQUAL.FamilyName)
            {
                columnSecType = SectType.COL_L_UNEQUAL;
            }
            else if (rvtFamilyName == COL_PIPE.FamilyName)
            {
                columnSecType = SectType.COL_PIPE;
            }
            else if (rvtFamilyName == COL_TUBE.FamilyName)
            {
                columnSecType = SectType.COL_TUBE;
            }

            else if (rvtFamilyName == COL_UB.FamilyName)
            {
                columnSecType = SectType.COL_UB;
            }

            else if (rvtFamilyName == COL_UC.FamilyName)
            {
                columnSecType = SectType.COL_UC;
            }

            else if (rvtFamilyName == COL_UPN.FamilyName)
            {
                columnSecType = SectType.COL_UPN;
            }


            else if (rvtFamilyName == COL_W.FamilyName)
            {
                columnSecType = SectType.COL_W;
            }


            #endregion
            return columnSecType;
        }


        public static void setRevitDataInExcelBulk(List<Frame> RevitFrames, Microsoft.Office.Interop.Excel.Worksheet ws)
        {


            int count = RevitFrames.Count;
            ws.Range["F5"].Value = Convert.ToString(count);
            Excel.Range StdID = ws.Range["A7"];
            Excel.Range RvtName = ws.Range["Y7"];
            Excel.Range RvtFamily = ws.Range["Z7"];
            Excel.Range RvtID = ws.Range["AA7"];
            Excel.Range RvtStartX = ws.Range["AB7"];
            Excel.Range RvtMemberType = ws.Range["AN7"];
            Excel.Range RvtBeta = ws.Range["AO7"];
            Excel.Range RvtLength = ws.Range["AP7"];

            object[,] dataToExcel = new object[RevitFrames.Count, 12];
            object[,] sapIds = new object[RevitFrames.Count, 1];

            object[,] memberTypeBetaLength = new object[RevitFrames.Count, 3];

            for (int i = 0; i < count; i++)
            {

                dataToExcel[i, 0] = Convert.ToString(RevitFrames[i].RvtName);
                dataToExcel[i, 1] = Convert.ToString(RevitFrames[i].RvtFamily);
                dataToExcel[i, 2] = Convert.ToString(RevitFrames[i].RvtID);
                dataToExcel[i, 3] = Convert.ToString(RevitFrames[i].StartPoint.X);
                dataToExcel[i, 4] = Convert.ToString(RevitFrames[i].StartPoint.Y);
                dataToExcel[i, 5] = Convert.ToString(RevitFrames[i].StartPoint.Z);
                dataToExcel[i, 6] = Convert.ToString(RevitFrames[i].EndPoint.X);
                dataToExcel[i, 7] = Convert.ToString(RevitFrames[i].EndPoint.Y);
                dataToExcel[i, 8] = Convert.ToString(RevitFrames[i].EndPoint.Z);


            }

            Microsoft.Office.Interop.Excel.Range startCell = (Range)ws.Cells[7, 25];
            Microsoft.Office.Interop.Excel.Range endCell = (Range)ws.Cells[dataToExcel.GetLength(0) + 7 - 1, 33];
            Microsoft.Office.Interop.Excel.Range range = (Range)ws.Range[startCell, endCell];
            range.Value2 = dataToExcel;




            // set ids of sap
            for (int i = 0; i < count; i++)
            {
                sapIds[i, 0] = Convert.ToString(RevitFrames[i].ID);
            }
            startCell = (Range)ws.Cells[7, 1];
            endCell = (Range)ws.Cells[RevitFrames.Count + 7 - 1, 1];
            range = (Range)ws.Range[startCell, endCell];
            range.Value2 = sapIds;



            // set MemberTybeBetaLength
            for (int i = 0; i < count; i++)
            {
                memberTypeBetaLength[i, 0] = Convert.ToString(RevitFrames[i].MemberType.ToString());
                memberTypeBetaLength[i, 1] = Convert.ToString(RevitFrames[i].Beta.ToString());
                memberTypeBetaLength[i, 2] = Convert.ToString(RevitFrames[i].Length.ToString());
            }


            startCell = (Range)ws.Cells[7, 40];
            endCell = (Range)ws.Cells[RevitFrames.Count + 7 - 1, 42];
            range = (Range)ws.Range[startCell, endCell];
            range.Value2 = memberTypeBetaLength;

        }

        public static void assignStaadDataFromExcelToRevitParametersinBulk(Document doc, Microsoft.Office.Interop.Excel.Worksheet ws)
        {
            using (Transaction trns = new Transaction(doc, "aa"))

            {
                trns.Start();
                FailureHandlingOptions failOpt = trns.GetFailureHandlingOptions();
                failOpt.SetFailuresPreprocessor(new RoomWarningSwallower());
                trns.SetFailureHandlingOptions(failOpt);


                int count = Convert.ToInt32(ws.Range["F5"].Value2);


                object[,] staadIds = new object[count, 1];
                object[,] revitIds = new object[count, 1];
                object[,] staadMaterial = new object[count, 1];

                Microsoft.Office.Interop.Excel.Range startCell = (Range)ws.Cells[7, 1];
                Microsoft.Office.Interop.Excel.Range endCell = (Range)ws.Cells[count + 7 - 1, 1];
                Microsoft.Office.Interop.Excel.Range range = (Range)ws.Range[startCell, endCell];
                staadIds = (object[,])range.Value2;


                startCell = (Range)ws.Cells[7, 18];
                endCell = (Range)ws.Cells[staadIds.Length + 7 - 1, 18];
                range = ws.Range[startCell, endCell];
                staadMaterial = (object[,])range.Value2;



                startCell = (Range)ws.Cells[7, 27];
                endCell = (Range)ws.Cells[staadIds.Length + 7 - 1, 27];
                range = ws.Range[startCell, endCell];
                revitIds = (object[,])range.Value2;



                for (int i = 1; i <= count; i++)
                {
                    try
                    {
                        Instance RevitElement = (Instance)doc.GetElement(new ElementId(Convert.ToInt32(revitIds[i, 1])));
                        RevitElement.LookupParameter("StaadID").Set((string)Convert.ToString(staadIds[i, 1]));
                        RevitElement.LookupParameter("RvtID").Set((string)Convert.ToString(revitIds[i, 1]));
                        RevitElement.LookupParameter("Material").Set((string)Convert.ToString(staadMaterial[i, 1]));
                    }


                    catch (Exception ex2)
                    {
                        TaskDialog.Show("error", ex2.Message + "  No of row" + i);
                        continue;
                    }

                }

                trns.Commit();
            }


        }


        public static List<Frame> TransformGeneralElementsCoordinates(UIDocument uiDoc, Document doc, string analysis, ref List<Frame> RevitFrames)
        {

            // for revit 2021 or above use (ForgeTypeId dis = UnitTypeId.Millimeters;) however for revit 2020 or lower use ( ForgeTypeId dis = DisplayUnitType.DUT_MILLIMETERS;)
            ForgeTypeId dis = UnitTypeId.Millimeters;


            #region PropmptToSelectAllTargetMembers
            List<Frame> sapFrames = new List<Frame>();
            List<Frame> RevitBeams = new List<Frame>();
            List<Frame> RevitColumns = new List<Frame>();

            TaskDialog.Show("TargetMembers", "Please Select AllElments in analysisModel and the corresponding in Revit Model");
            List<ElementId> revitSelectedElements = new RevitStructuralFramingSelector(uiDoc).SelectStructuralFramesAndColumns();
            RevitFrames = RevitUtils.getRevitFrames2(uiDoc, doc, dis, ref RevitBeams, ref RevitColumns, analysis);


            List<Frame> filteredRevitFrames = new List<Frame>();

            if (revitSelectedElements.Count > 0)
            {
                filteredRevitFrames = RevitFrames.Where(x => revitSelectedElements.Contains(x.revitEleId)).ToList();

            }
            else
            {
                filteredRevitFrames = RevitFrames;
            }

            if (analysis == "Sap")
            {
                sapFrames = SapUtils.getSapFrames();

            }

            else if (analysis == "Etabs")
            {
                sapFrames = EtabsUtils.getSapFrames();

            }

            #endregion

            #region PromptUserToSelectRefrenceMember
            TaskDialog.Show("MappingSelection", "Please Select Element in analysisModel and the corresponding one in Revit Model");
            ElementId revitElmentID = new RevitStructuralFramingSelector(uiDoc).SelectStructuralFraming();
            Frame SapRefElement = new Frame();

            if (analysis == "Sap")
            {
                SapRefElement = SapUtils.getSapFrames().FirstOrDefault();

            }

            else if (analysis == "Etabs")
            {
                SapRefElement = EtabsUtils.getSapFrames().FirstOrDefault();

            }

            Frame RevitRefElement = filteredRevitFrames.Where(x => x.revitEleId == revitElmentID).First();

            #endregion


            if (RevitRefElement != null)
            {
                #region CalculateMatrixes
                Matrix4x4 translationfMatrix = new Matrix4x4();
                Matrix4x4 rotationalMatrix = new Matrix4x4();


                double RotAngle = 0;
                rotationalMatrix = TransformationHelper.get_transformation_masterMatrix(SapRefElement, RevitRefElement, ref translationfMatrix, ref RotAngle);
                #endregion

                #region getNewTransformedFrames
                TransformationHelper.getTransfromedFrames(rotationalMatrix, ref sapFrames, translationfMatrix, RotAngle);

                if (analysis == "Sap")
                {
                    ChangeSapSectionNames.GetSectionsDimentions2(SapUtils.Get_sap_mode(), ref sapFrames, sapFrames.Select(x => x.ID).ToList());

                }

                else if (analysis == "Etabs")
                {
                    ChangeEtabsSectionNames.GetSectionsDimentions2(EtabsUtils.Get_sap_mode(), ref sapFrames, sapFrames.Select(x => x.ID).ToList());

                }

                #endregion
            }




            return sapFrames;


        }



        #region Get Levels

        public static Dictionary<Level, double> getLevelsDictionary()
        {
            #region GetRevitLevels

            UIDocument revitUiDoc = RhinoInside.Revit.Revit.ActiveUIDocument;
            Autodesk.Revit.DB.Document revitDoc = RhinoInside.Revit.Revit.ActiveDBDocument;
            // Create a dictionary to hold levels and their elevations
            Dictionary<Level, double> levelsDict = new Dictionary<Level, double>();

            //for revit 2021 or above use(ForgeTypeId dis = UnitTypeId.Millimeters;) however for revit 2020 or lower use(ForgeTypeId dis = DisplayUnitType.DUT_MILLIMETERS;)
            ForgeTypeId dis = UnitTypeId.Millimeters;

            // Filter to get all levels in the document
            FilteredElementCollector collector = new FilteredElementCollector(revitDoc);
            collector.OfClass(typeof(Level));

            // Iterate through all levels and add them to the dictionary
            foreach (Element element in collector)
            {
                Level level = element as Level;
                if (level != null)
                {
                    double elevation = level.Elevation;

                    levelsDict.Add(level, UnitUtils.ConvertFromInternalUnits(elevation, dis));
                }
            }
            #endregion

            return levelsDict;
        }


        // Function to get the highest level from a list of level names and a dictionary
        public static string GetHighestLevel(List<string> levelNames, Dictionary<Level, double> levelsDict)
        {
            Level highestLevel = null;
            double maxElevation = double.MinValue;

            foreach (var levelName in levelNames)
            {
                foreach (var kvp in levelsDict)
                {
                    if (kvp.Key.Name == levelName && kvp.Value > maxElevation)
                    {
                        highestLevel = kvp.Key;
                        maxElevation = kvp.Value;
                    }
                }
            }

            return highestLevel.Name;
        }

        // Function to get the lowest level from a list of level names and a dictionary
        public static string GetLowestLevel(List<string> levelNames, Dictionary<Level, double> levelsDict)
        {
            Level lowestLevel = null;
            double minElevation = double.MaxValue;

            foreach (var levelName in levelNames)
            {
                foreach (var kvp in levelsDict)
                {
                    if (kvp.Key.Name == levelName && kvp.Value < minElevation)
                    {
                        lowestLevel = kvp.Key;
                        minElevation = kvp.Value;
                    }
                }
            }

            return lowestLevel.Name;
        }


        public static void SetNearestbaseAndTopLevels(List<Frame> frames)
        {
            Dictionary<Level, double> revitLevels = getLevelsDictionary();


            foreach (var item in frames)
            {

                // Find the level with the smallest absolute difference in Z value
                Level nearestTopLevel = revitLevels
                    .OrderBy(x => Math.Abs(x.Value - item.EndPoint.Z)) // Order by absolute difference
                    .FirstOrDefault().Key; // 

                Level nearestBottomLevel = revitLevels
                    .OrderBy(x => Math.Abs(x.Value - item.StartPoint.Z)) // Order by absolute difference
                    .FirstOrDefault().Key; // 

                item.RevitTopLevel = nearestTopLevel.Name;
                item.RevitTopLevelOffset = item.EndPoint.Z - revitLevels[nearestTopLevel];


                item.RevitBotLevel = nearestBottomLevel.Name;
                item.RevitBotLevelOffset = item.StartPoint.Z - revitLevels[nearestBottomLevel];

            }


        }

        #endregion

        #region RelatedToFamilies 
        public static FamilySymbol createRvtFamilySectionForBeam(Frame bm, Document doc, ForgeTypeId dis)
        {
            FamilySymbol revitSection = null;

            try
            {
                #region CreatePipeSection

                if (bm.Shape == BM_PIPE.shape)
                {
                    string dimentionswithX;
                    string[] dimentions;
                    BM_PIPE.getPipeDimentionsFromAnalysisSectionNames(bm, out dimentionswithX, out dimentions);

                    revitSection = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralFraming).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(f => f.FamilyName == BM_PIPE.FamilyName).FirstOrDefault(X => X.Name.Contains(dimentionswithX));

                    // new
                    if (revitSection == null)
                    {
                        revitSection = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralFraming).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(f => f.FamilyName == BM_PIPE.FamilyName).FirstOrDefault(X => X.Name.Contains(dimentionswithX.ToLower()));
                        if (revitSection != null)
                        {
                            using (Transaction trns = new Transaction(doc, "aa"))
                            {
                                trns.Start();
                                revitSection.Name = revitSection.Name.ToUpper();

                                trns.Commit();
                            }
                        }

                    }


                    // Define new familyType 
                    if (revitSection == null)
                    {

                        #region difintions 

                        FamilySymbol baseSection = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralFraming).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(f => f.FamilyName == BM_PIPE.FamilyName).FirstOrDefault();


                        using (Transaction trns = new Transaction(doc, "aa"))
                        {
                            trns.Start();
                            FailureHandlingOptions failOpt = trns.GetFailureHandlingOptions();
                            failOpt.SetFailuresPreprocessor(new RoomWarningSwallower());
                            trns.SetFailureHandlingOptions(failOpt);
                            string name = "PIPE " + dimentionswithX;
                            revitSection = baseSection.Duplicate(name) as FamilySymbol;
                            trns.Commit();
                        }

                        string outerDiameterParam = BM_PIPE.outerDiameter;
                        string thickParam = BM_PIPE.thick;


                        Parameter pra_a_1 = revitSection.LookupParameter(outerDiameterParam);
                        Parameter pra_b_1 = revitSection.LookupParameter(thickParam);


                        #endregion

                        try
                        {
                            using (Transaction trns = new Transaction(doc, "aa"))
                            {
                                trns.Start();
                                FailureHandlingOptions failOpt = trns.GetFailureHandlingOptions();
                                failOpt.SetFailuresPreprocessor(new RoomWarningSwallower());
                                trns.SetFailureHandlingOptions(failOpt);
                                double pra_a_2 = UnitUtils.ConvertToInternalUnits(Convert.ToDouble(dimentions[0]), dis);
                                double pra_b_2 = UnitUtils.ConvertToInternalUnits(Convert.ToDouble(dimentions[1]), dis);

                                pra_a_1.Set(pra_a_2);
                                pra_b_1.Set(pra_b_2);


                                trns.Commit();

                            }
                        }

                        catch (Exception ex2)
                        {




                        }

                    }


                    return revitSection;
                }

                #endregion

                #region BoxSections

                if (bm.Shape == BM_TUBE.shape)
                {
                    string dimentionswithX;
                    string[] dimentions;
                    int noOfX;
                    BM_TUBE.getTubeDimnsionsFromAnalysisSectionNames(bm, out dimentionswithX, out dimentions, out noOfX);

                    // check familyType existance

                    revitSection = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralFraming).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(f => f.FamilyName == BM_TUBE.FamilyName).FirstOrDefault(X => X.Name.Contains(dimentionswithX) && X.Name.Count(b => b == 'X') == noOfX);

                    // new
                    if (revitSection == null)
                    {
                        revitSection = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralFraming).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(f => f.FamilyName == BM_TUBE.FamilyName).FirstOrDefault(X => X.Name.Contains(dimentionswithX.ToLower()) && X.Name.Count(b => b == 'x') == noOfX);
                        if (revitSection != null)
                        {
                            using (Transaction trns = new Transaction(doc, "aa"))
                            {
                                trns.Start();
                                revitSection.Name = revitSection.Name.ToUpper();

                                trns.Commit();
                            }
                        }

                    }

                    // Define new familyType 
                    if (revitSection == null)
                    {

                        #region difintions 

                        FamilySymbol baseSection = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralFraming).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(f => f.FamilyName == BM_TUBE.FamilyName).FirstOrDefault();

                        using (Transaction trns = new Transaction(doc, "aa"))
                        {
                            trns.Start();
                            string name = "TUBE " + dimentionswithX;
                            revitSection = baseSection.Duplicate(name) as FamilySymbol;
                            trns.Commit();
                        }

                        string webDepthPram = BM_TUBE.webDepth;
                        string flgWidthParm = BM_TUBE.flgWidth;
                        string webThickParm = BM_TUBE.webThick;
                        string flgThickParm = BM_TUBE.flgThick;
                        string outerRadius = BM_TUBE.outerRad;
                        string innerRadius = BM_TUBE.innerRad;


                        Parameter pra_a_1 = revitSection.LookupParameter(webDepthPram);
                        Parameter pra_b_1 = revitSection.LookupParameter(flgWidthParm);
                        Parameter pra_c_1 = revitSection.LookupParameter(webThickParm);
                        Parameter pra_d_1 = revitSection.LookupParameter(flgThickParm);
                        Parameter pra_e_1 = revitSection.LookupParameter(outerRadius);
                        Parameter pra_f_1 = revitSection.LookupParameter(innerRadius);

                        #endregion


                        try
                        {
                            using (Transaction trns = new Transaction(doc, "aa"))
                            {
                                trns.Start();
                                FailureHandlingOptions failOpt = trns.GetFailureHandlingOptions();
                                failOpt.SetFailuresPreprocessor(new RoomWarningSwallower());
                                trns.SetFailureHandlingOptions(failOpt);
                                double pra_a_2 = UnitUtils.ConvertToInternalUnits(Convert.ToDouble(dimentions[0]), dis);
                                double pra_b_2 = UnitUtils.ConvertToInternalUnits(Convert.ToDouble(dimentions[1]), dis);
                                double pra_c_2 = UnitUtils.ConvertToInternalUnits(Convert.ToDouble(dimentions[2]), dis);
                                double pra_d_2 = 0;
                                // flgThick
                                if (dimentions.Length == 4)
                                {
                                    // if tf != tw
                                    pra_d_2 = UnitUtils.ConvertToInternalUnits(Convert.ToDouble(dimentions[3]), dis);
                                }


                                else
                                {
                                    // if Tf == tw
                                    pra_d_2 = UnitUtils.ConvertToInternalUnits(Convert.ToDouble(dimentions[2]), dis);
                                }

                                double pra_e_2 = UnitUtils.ConvertToInternalUnits(Convert.ToDouble(4.0), dis);
                                double pra_f_2 = UnitUtils.ConvertToInternalUnits(Convert.ToDouble(4.0), dis);





                                pra_a_1.Set(pra_a_2);
                                pra_b_1.Set(pra_b_2);
                                pra_c_1.Set(pra_c_2);
                                pra_d_1.Set(pra_d_2);
                                pra_e_1.Set(pra_e_2);
                                pra_f_1.Set(pra_f_2);

                                trns.Commit();
                            }
                        }

                        catch (Exception ex2)
                        {



                        }


                    }

                    return revitSection;

                }

                #endregion

                #region I_B

                if (bm.IsBUS == true && bm.Shape == BM_BUS_I.shape)
                {

                    string webThick, flgThick, consWebDepth, consFlgDepth;
                    BM_BUS_I.getBUS_IDimensionsFromAnalysisSectionNames(bm, out webThick, out flgThick, out consWebDepth, out consFlgDepth);

                    revitSection = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralFraming).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(f => f.FamilyName == BM_BUS_I.FamilyName).FirstOrDefault(X => X.Name == bm.Sec);

                    // check familyType existance
                    if (revitSection == null)
                    {

                        // Define new familyType 
                        #region  Definitions
                        using (Transaction trns = new Transaction(doc, "aa"))
                        {
                            FamilySymbol baseSection = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralFraming).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(f => f.FamilyName == BM_BUS_I.FamilyName).FirstOrDefault();
                            trns.Start();
                            revitSection = baseSection.Duplicate(bm.Sec) as FamilySymbol;
                            trns.Commit();
                        }


                        string totalDepthParam = BM_BUS_I.totalBeamDepth;
                        string flgWidthParm = BM_BUS_I.flgWidth;
                        string webThickParm = BM_BUS_I.webThick;
                        string flgThickParm = BM_BUS_I.flgThick;

                        Parameter pra_a_1 = revitSection.LookupParameter(totalDepthParam);
                        Parameter pra_b_1 = revitSection.LookupParameter(flgWidthParm);
                        Parameter pra_c_1 = revitSection.LookupParameter(webThickParm);
                        Parameter pra_d_1 = revitSection.LookupParameter(flgThickParm);
                        #endregion

                        try
                        {
                            using (Transaction trns = new Transaction(doc, "aa"))
                            {
                                trns.Start();

                                double pra_a_2 = UnitUtils.ConvertToInternalUnits(Convert.ToDouble(consWebDepth), dis);
                                double pra_b_2 = UnitUtils.ConvertToInternalUnits(Convert.ToDouble(consFlgDepth), dis);
                                double pra_c_2 = UnitUtils.ConvertToInternalUnits(Convert.ToDouble(webThick), dis);
                                double pra_d_2 = UnitUtils.ConvertToInternalUnits(Convert.ToDouble(flgThick), dis);



                                pra_a_1.Set(pra_a_2);
                                pra_b_1.Set(pra_b_2);
                                pra_c_1.Set(pra_c_2);
                                pra_d_1.Set(pra_d_2);


                                trns.Commit();
                            }
                        }

                        catch (Exception ex2)
                        {



                        }


                    }

                    return revitSection;

                }







                #endregion

                #region I_H

                if (bm.IsBUS == false && bm.Shape == "I")
                {
                    #region filterBeteenTypes
                    string numeric = null;
                    string FamilyType = null;


                    if (bm.Sec.Contains("HE") && bm.Shape == "I")
                    {
                        if (bm.Sec.Contains("M"))
                        {
                            //string numeric = new String(bm.Sec.Where(Char.IsDigit).ToArray());
                            //revitSection = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralFraming).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(f => f.FamilyName == HEB_Beam).FirstOrDefault(X => new String(X.Name.Where(Char.IsDigit).ToArray()) == numeric);
                            FamilyType = "HEM";
                            goto IBeamFamilySelection;

                        }

                        else if (bm.Sec.Contains("B"))
                        {
                            FamilyType = "HEB";
                            goto IBeamFamilySelection;
                        }


                        else if (bm.Sec.Contains("A"))
                        {
                            int count = ExcelUtils.CountCharsUsingForeach(bm.Sec, 'A');
                            if (count == 2)
                            {
                                FamilyType = "HEAA";
                                goto IBeamFamilySelection;
                            }

                            else
                            {
                                FamilyType = "HEA";
                                goto IBeamFamilySelection;
                            }
                        }

                    }

                    else if (bm.Sec.Contains("IP") && bm.Shape == "I")
                    {
                        if (bm.Sec.Contains("O"))
                        {
                            FamilyType = "IPEO";
                            goto IBeamFamilySelection;

                        }

                        else if (bm.Sec.Contains("A"))
                        {
                            FamilyType = "IPEA";
                            goto IBeamFamilySelection;
                        }

                        else if (bm.Sec.Contains("R"))
                        {
                            FamilyType = "IPER";
                            goto IBeamFamilySelection;

                        }

                        else if (bm.Sec.Contains("V"))
                        {
                            FamilyType = "IPEV";
                            goto IBeamFamilySelection;
                        }

                        else if (bm.Sec.Contains("E"))
                        {
                            FamilyType = "IPE";
                            goto IBeamFamilySelection;

                        }

                    }

                    else if ((bm.Sec.Contains("UKC") || bm.Sec.Contains("UC")) && bm.Shape == "I" && bm.IsBUS == false && (bm.Lib.ToUpper().Contains("BSSHAPES") || bm.Lib.ToUpper().Contains("BRITISH") || bm.Lib.ToUpper().Contains("AUSTRAILIAN")))
                    {

                        FamilyType = "UC";
                        goto IBeamFamilySelection;
                    }

                    else if ((bm.Sec.Contains("UKB") || bm.Sec.Contains("UB")) && bm.Shape == "I" && bm.IsBUS == false && (bm.Lib.ToUpper().Contains("BSSHAPES") || bm.Lib.ToUpper().Contains("BRITISH") || bm.Lib.ToUpper().Contains("AUSTRAILIAN")))

                    {
                        FamilyType = "UB";
                        goto IBeamFamilySelection;
                    }

                    else if (bm.Sec.Contains("W") && bm.Shape == "I" && bm.IsBUS == false && (bm.Lib.ToUpper().Contains("AISC") || bm.Lib.Contains("American") || bm.Lib.Contains("Canadian")))
                    {
                        FamilyType = "W";
                        goto IBeamFamilySelection;
                    }



                #endregion

                #region SelectFamilyType

                IBeamFamilySelection:
                    switch (FamilyType)
                    {

                        case "HEAA":
                            numeric = new String(bm.Sec.Where(Char.IsDigit).ToArray());
                            revitSection = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralFraming).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(f => f.FamilyName == BM_HEAA.FamilyName).FirstOrDefault(X => new String(X.Name.Where(Char.IsDigit).ToArray()) == numeric);
                            if (revitSection == null)
                            {
                                return null;
                            }
                            return revitSection;


                        case "HEA":
                            numeric = new String(bm.Sec.Where(Char.IsDigit).ToArray());
                            revitSection = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralFraming).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(f => f.FamilyName == BM_HEA.FamilyName).FirstOrDefault(X => new String(X.Name.Where(Char.IsDigit).ToArray()) == numeric);
                            if (revitSection == null)
                            {
                                return null;
                            }
                            return revitSection;


                        case "HEB":
                            numeric = new String(bm.Sec.Where(Char.IsDigit).ToArray());
                            revitSection = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralFraming).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(f => f.FamilyName == BM_HEB.FamilyName).FirstOrDefault(X => new String(X.Name.Where(Char.IsDigit).ToArray()) == numeric);
                            if (revitSection == null)
                            {
                                return null;
                            }
                            return revitSection;

                        case "HEM":
                            numeric = new String(bm.Sec.Where(Char.IsDigit).ToArray());
                            revitSection = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralFraming).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(f => f.FamilyName == BM_HEM.FamilyName).FirstOrDefault(X => new String(X.Name.Where(Char.IsDigit).ToArray()) == numeric);
                            if (revitSection == null)
                            {
                                return null;
                            }
                            return revitSection;

                        case "IPE":
                            numeric = new String(bm.Sec.Where(Char.IsDigit).ToArray());
                            revitSection = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralFraming).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(f => f.FamilyName == BM_IPE.FamilyName).FirstOrDefault(X => new String(X.Name.Where(Char.IsDigit).ToArray()) == numeric);
                            if (revitSection == null)
                            {
                                return null;
                            }
                            return revitSection;

                        case "IPEA":
                            numeric = new String(bm.Sec.Where(Char.IsDigit).ToArray());
                            revitSection = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralFraming).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(f => f.FamilyName == BM_IPEA.FamilyName).FirstOrDefault(X => new String(X.Name.Where(Char.IsDigit).ToArray()) == numeric);
                            if (revitSection == null)
                            {
                                return null;
                            }
                            return revitSection;

                        case "IPEO":
                            numeric = new String(bm.Sec.Where(Char.IsDigit).ToArray());
                            revitSection = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralFraming).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(f => f.FamilyName == BM_IPEO.FamilyName).FirstOrDefault(X => new String(X.Name.Where(Char.IsDigit).ToArray()) == numeric);
                            if (revitSection == null)
                            {
                                return null;
                            }
                            return revitSection;


                        ////////////////////////
                        case "UB":

                            char cc = bm.Sec.SkipWhile(c => !char.IsDigit(c)).ToList()[0];
                            int lastSpacecharacterIndex = bm.Sec.IndexOf(cc);
                            string dimentionswithX = bm.Sec.Substring(lastSpacecharacterIndex);

                            revitSection = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralFraming).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(f => f.FamilyName == BM_UB.FamilyName).FirstOrDefault(X => X.Name.ToUpper().Contains(dimentionswithX.ToUpper()));
                            if (revitSection == null)
                            {
                                return null;
                            }
                            return revitSection;



                        case "UC":

                            cc = bm.Sec.SkipWhile(c => !char.IsDigit(c)).ToList()[0];
                            lastSpacecharacterIndex = bm.Sec.IndexOf(cc);
                            dimentionswithX = bm.Sec.Substring(lastSpacecharacterIndex);

                            revitSection = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralFraming).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(f => f.FamilyName == BM_UC.FamilyName).FirstOrDefault(X => X.Name.ToUpper().Contains(dimentionswithX.ToUpper()));

                            if (revitSection == null)
                            {
                                return null;
                            }
                            return revitSection;



                        case "W":
                            cc = bm.Sec.SkipWhile(c => !char.IsDigit(c)).ToList()[0];
                            lastSpacecharacterIndex = bm.Sec.IndexOf(cc);
                            dimentionswithX = bm.Sec.Substring(lastSpacecharacterIndex);
                            revitSection = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralFraming).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(f => f.FamilyName == BM_W.FamilyName).FirstOrDefault(X => X.Name.ToUpper().Contains(dimentionswithX.ToUpper()));

                            if (revitSection == null)
                            {
                                return null;
                            }
                            return revitSection;


                        default:
                            return null;
                    }

                    #endregion

                }


                #endregion

                #region C_H 

                if (bm.IsBUS != true && bm.Shape == BM_UPN.shape)
                {
                    string dimentionswithX = BM_UPN.getChannelDimnsionsFromAnalysisSectionNames(bm);

                    // check familyType existance
                    revitSection = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralFraming).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(f => f.FamilyName == BM_UPN.FamilyName).FirstOrDefault(X => X.Name.Contains(dimentionswithX));

                    // Define new familyType 
                    if (revitSection == null)
                    {
                        return null;
                    }

                    return revitSection;

                }

                #endregion

                #region I_T

                if (bm.Shape == "VARI")
                {
                    string webThick, flgThick, startWebDepth, endWebDepth, consFlgDepth;
                    RevitFamily.getTapered_IDimensionsFromAnalysisSectionNames(bm, out webThick, out flgThick, out startWebDepth, out endWebDepth, out consFlgDepth);

                    // check familyType existance
                    revitSection = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralFraming).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(f => f.FamilyName == BM_BUS_I_Tapered.FamilyName).FirstOrDefault(X => X.Name == bm.Sec);

                    // Define new familyType 
                    if (revitSection == null)
                    {

                        #region  Definitions
                        using (Transaction trns = new Transaction(doc, "aa"))
                        {
                            FamilySymbol baseSection = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralFraming).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(f => f.FamilyName == BM_BUS_I_Tapered.FamilyName).FirstOrDefault();
                            trns.Start();
                            revitSection = baseSection.Duplicate(bm.Sec) as FamilySymbol;
                            trns.Commit();
                        }


                        string startWebDepthParam = BM_BUS_I_Tapered.startWebDepth;
                        string endWebDepthParam = BM_BUS_I_Tapered.endWebDepth;
                        string flgWidthParm = BM_BUS_I_Tapered.flgWidth;
                        string webThickParm = BM_BUS_I_Tapered.webThick;
                        string flgThickParm = BM_BUS_I_Tapered.flgThick;

                        Parameter pra_a_1 = revitSection.LookupParameter(startWebDepthParam);
                        Parameter pra_b_1 = revitSection.LookupParameter(endWebDepthParam);
                        Parameter pra_c_1 = revitSection.LookupParameter(flgWidthParm);
                        Parameter pra_d_1 = revitSection.LookupParameter(webThickParm);
                        Parameter pra_e_1 = revitSection.LookupParameter(flgThickParm);
                        #endregion


                        try
                        {
                            using (Transaction trns = new Transaction(doc, "aa"))
                            {
                                trns.Start();

                                double pra_a_2 = UnitUtils.ConvertToInternalUnits(Convert.ToDouble(startWebDepth), dis);
                                double pra_b_2 = UnitUtils.ConvertToInternalUnits(Convert.ToDouble(endWebDepth), dis);
                                double pra_c_2 = UnitUtils.ConvertToInternalUnits(Convert.ToDouble(consFlgDepth), dis);
                                double pra_d_2 = UnitUtils.ConvertToInternalUnits(Convert.ToDouble(webThick), dis);
                                double pra_e_2 = UnitUtils.ConvertToInternalUnits(Convert.ToDouble(flgThick), dis);



                                pra_a_1.Set(pra_a_2);
                                pra_b_1.Set(pra_b_2);
                                pra_c_1.Set(pra_c_2);
                                pra_d_1.Set(pra_d_2);
                                pra_e_1.Set(pra_e_2);



                                trns.Commit();

                            }
                        }

                        catch (Exception ex2)
                        {



                        }


                    }

                    return revitSection;

                }




                #endregion

                #region Pipe_T

                if (bm.IsBUS == true && bm.Shape == "VARPipe")
                {
                    string startDiamaterValue, endDiamaterValue, PipeThicknessValue;
                    RevitFamily.getTaperedPipeDimensionsFromAnalysisSectionNames(bm, out startDiamaterValue, out endDiamaterValue, out PipeThicknessValue);

                    // check familyType existance
                    revitSection = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralFraming).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(f => f.FamilyName == BM_BUS_PIPE_Tapered.FamilyName).FirstOrDefault(X => X.Name == bm.Sec);


                    // Define new familyType 
                    if (revitSection == null)
                    {

                        #region difintions 

                        FamilySymbol baseSection = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralFraming).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(f => f.FamilyName == BM_BUS_PIPE_Tapered.FamilyName).FirstOrDefault();


                        using (Transaction trns = new Transaction(doc, "aa"))
                        {
                            trns.Start();
                            FailureHandlingOptions failOpt = trns.GetFailureHandlingOptions();
                            failOpt.SetFailuresPreprocessor(new RoomWarningSwallower());
                            trns.SetFailureHandlingOptions(failOpt);
                            revitSection = baseSection.Duplicate(bm.Sec) as FamilySymbol;
                            trns.Commit();
                        }

                        string startDiamParam = BM_BUS_PIPE_Tapered.startDiametr;
                        string endDiamParam = BM_BUS_PIPE_Tapered.endDiametr;
                        string startthicknessParam = BM_BUS_PIPE_Tapered.startThick;
                        string endthicknessParam = BM_BUS_PIPE_Tapered.endThick;


                        Parameter pra_a_1 = revitSection.LookupParameter(startDiamParam);
                        Parameter pra_b_1 = revitSection.LookupParameter(endDiamParam);
                        Parameter pra_c_1 = revitSection.LookupParameter(startthicknessParam);
                        Parameter pra_d_1 = revitSection.LookupParameter(endthicknessParam);


                        #endregion

                        try
                        {
                            using (Transaction trns = new Transaction(doc, "aa"))
                            {
                                trns.Start();
                                FailureHandlingOptions failOpt = trns.GetFailureHandlingOptions();
                                failOpt.SetFailuresPreprocessor(new RoomWarningSwallower());
                                trns.SetFailureHandlingOptions(failOpt);
                                double pra_a_2 = UnitUtils.ConvertToInternalUnits(Convert.ToDouble(startDiamaterValue), dis);
                                double pra_b_2 = UnitUtils.ConvertToInternalUnits(Convert.ToDouble(endDiamaterValue), dis);
                                double pra_c_2 = UnitUtils.ConvertToInternalUnits(Convert.ToDouble(PipeThicknessValue), dis);
                                double pra_d_2 = UnitUtils.ConvertToInternalUnits(Convert.ToDouble(PipeThicknessValue), dis);



                                pra_a_1.Set(pra_a_2);
                                pra_b_1.Set(pra_b_2);
                                pra_c_1.Set(pra_c_2);
                                pra_d_1.Set(pra_c_2);



                                trns.Commit();

                            }
                        }

                        catch (Exception ex2)
                        {


                        }

                    }

                    return revitSection;

                }

                #endregion

                #region Tube_T

                if (bm.Shape == BM_BUS_TUBE_Tapered.shape)
                {
                    string startWebDepth, endWebDepth, webThick, startFlgWidth, endFlgWidth, flgThick;
                    BM_BUS_TUBE_Tapered.getTaperedTubeDimensionsFromAnalysisSectionNames(bm, out startWebDepth, out endWebDepth, out webThick, out startFlgWidth, out endFlgWidth, out flgThick);

                    // check familyType existance
                    revitSection = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralFraming).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(f => f.FamilyName == BM_BUS_TUBE_Tapered.FamilyName).FirstOrDefault(X => X.Name == bm.Sec);

                    // Define new familyType 
                    if (revitSection == null)
                    {

                        #region  Definitions
                        using (Transaction trns = new Transaction(doc, "aa"))
                        {
                            FamilySymbol baseSection = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralFraming).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(f => f.FamilyName == BM_BUS_TUBE_Tapered.FamilyName).FirstOrDefault();
                            trns.Start();
                            FailureHandlingOptions failOpt = trns.GetFailureHandlingOptions();
                            failOpt.SetFailuresPreprocessor(new RoomWarningSwallower());
                            trns.SetFailureHandlingOptions(failOpt);
                            revitSection = baseSection.Duplicate(bm.Sec) as FamilySymbol;
                            trns.Commit();
                        }


                        string startWebDepthParm = BM_BUS_TUBE_Tapered.startWebDepth;
                        string endWebDepthParm = BM_BUS_TUBE_Tapered.endWebDepth;
                        string startFlgWidthParm = BM_BUS_TUBE_Tapered.startFlgWidth;
                        string endFlgWidthParm = BM_BUS_TUBE_Tapered.endFlgWidth;
                        string webThickParm = BM_BUS_TUBE_Tapered.overAllThick;
                        string flgThickParm = BM_BUS_TUBE_Tapered.overAllThick;


                        Parameter pra_a_1 = revitSection.LookupParameter(startWebDepthParm);
                        Parameter pra_b_1 = revitSection.LookupParameter(endWebDepthParm);
                        Parameter pra_c_1 = revitSection.LookupParameter(startFlgWidthParm);
                        Parameter pra_d_1 = revitSection.LookupParameter(endFlgWidthParm);
                        Parameter pra_e_1 = revitSection.LookupParameter(webThickParm);
                        Parameter pra_f_1 = revitSection.LookupParameter(flgThickParm);

                        #endregion


                        try
                        {
                            using (Transaction trns = new Transaction(doc, "aa"))
                            {
                                trns.Start();
                                FailureHandlingOptions failOpt = trns.GetFailureHandlingOptions();
                                failOpt.SetFailuresPreprocessor(new RoomWarningSwallower());
                                trns.SetFailureHandlingOptions(failOpt);
                                double pra_a_2 = UnitUtils.ConvertToInternalUnits(Convert.ToDouble(startWebDepth), dis);
                                double pra_b_2 = UnitUtils.ConvertToInternalUnits(Convert.ToDouble(endWebDepth), dis);
                                double pra_c_2 = UnitUtils.ConvertToInternalUnits(Convert.ToDouble(startFlgWidth), dis);
                                double pra_d_2 = UnitUtils.ConvertToInternalUnits(Convert.ToDouble(endFlgWidth), dis);
                                double pra_e_2 = UnitUtils.ConvertToInternalUnits(Convert.ToDouble(webThick), dis);
                                double pra_f_2 = UnitUtils.ConvertToInternalUnits(Convert.ToDouble(flgThick), dis);




                                pra_a_1.Set(pra_a_2);
                                pra_b_1.Set(pra_b_2);
                                pra_c_1.Set(pra_c_2);
                                pra_d_1.Set(pra_d_2);
                                pra_e_1.Set(pra_e_2);
                                pra_f_1.Set(pra_f_2);




                                trns.Commit();
                            }
                        }

                        catch (Exception ex2)
                        {


                        }


                    }


                    return revitSection;

                }



                #endregion

                #region Angles

                if (bm.Shape == BM_L_EQUAL.shape || bm.Shape == BM_L_UNEQUAL.shape || bm.Shape == BM_2L_EQUALANGLE.shape)
                {
                    bool isEQSingleAngle, isUNEQSingleAngle, isEQDblAngle;
                    string dimentionswithX, firstDimention;
                    RevitFamily.getAnglesDimensionsFromAnalysisSectionNames(bm, out isEQSingleAngle, out isUNEQSingleAngle, out isEQDblAngle, out dimentionswithX, out firstDimention);


                    #region  EQAngle

                    if (isEQSingleAngle)
                    {
                        string dimentionswithXDublicated = dimentionswithX;
                        int countOfX = dimentionswithX.ToUpper().Count(c => c == 'X');

                        if (countOfX < 2)
                        {
                            dimentionswithXDublicated = dimentionswithX.Insert(firstDimention.IndexOf("X") + 1, firstDimention + "X"); ;

                        }

                        // check familyType existance
                        // revitSection = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralFraming).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(f => f.FamilyName == BM_L_EQUAL.FamilyName).FirstOrDefault(X => X.Name.Contains(dimentionswithXDublicated) || X.Name.Contains(dimentionswithX));
                        revitSection = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralFraming).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(f => f.FamilyName == BM_L_EQUAL.FamilyName).FirstOrDefault(X => X.Name.Contains(dimentionswithXDublicated));



                        // Define new familyType 
                        if (revitSection == null)
                        {

                            return null;

                        }

                        return revitSection;

                    }

                    #endregion

                    #region  UQAngle

                    if (isUNEQSingleAngle)
                    {

                        // check familyType existance
                        revitSection = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralFraming).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(f => f.FamilyName == BM_L_UNEQUAL.FamilyName).FirstOrDefault(X => X.Name.Contains(dimentionswithX));

                        // Define new familyType 
                        if (revitSection == null)
                        {

                            return null;

                        }

                        return revitSection;

                    }

                    #endregion

                    #region  DblAngle

                    if (isEQDblAngle)
                    {
                        string dimentionswithXDublicated = dimentionswithX;

                        int countOfX = dimentionswithX.ToUpper().Count(c => c == 'X');
                        if (countOfX < 2)
                        {
                            dimentionswithXDublicated = dimentionswithX.Insert(firstDimention.IndexOf("X") + 1, firstDimention + "X"); ;

                        }

                        // check familyType existance
                        // revitSection = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralFraming).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(f => f.FamilyName == BM_2L_EQUALANGLE.FamilyName).FirstOrDefault(X => X.Name.Contains(dimentionswithXDublicated) || X.Name.Contains(dimentionswithX));

                        revitSection = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralFraming).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(f => f.FamilyName == BM_2L_EQUALANGLE.FamilyName).FirstOrDefault(X => X.Name.Contains(dimentionswithXDublicated));

                        // Define new familyType 
                        if (revitSection == null)
                        {

                            return null;

                        }

                        return revitSection;

                    }

                    #endregion


                }

                #endregion

                #region ConcRect
                if (bm.Shape == BM_ConcRect.shape)
                {
                    string dimentionswithX;
                    string[] dimentions;
                    BM_ConcRect.getConcRectDimensionsFromAnalysisSectionNames(bm, out dimentionswithX, out dimentions);

                    revitSection = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralFraming).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(f => f.FamilyName == BM_ConcRect.FamilyName).FirstOrDefault(X => X.Name.Contains(dimentionswithX));

                    // Define new familyType 
                    if (revitSection == null)
                    {

                        return null;
                        #region difintions 

                        FamilySymbol baseSection = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralFraming).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(f => f.FamilyName == BM_ConcRect.FamilyName).FirstOrDefault();


                        using (Transaction trns = new Transaction(doc, "aa"))
                        {
                            trns.Start();
                            FailureHandlingOptions failOpt = trns.GetFailureHandlingOptions();
                            failOpt.SetFailuresPreprocessor(new RoomWarningSwallower());
                            trns.SetFailureHandlingOptions(failOpt);
                            string name = "BM " + dimentionswithX;
                            revitSection = baseSection.Duplicate(name) as FamilySymbol;
                            trns.Commit();
                        }

                        string widthParam = BM_ConcRect.width;
                        string depthParam = BM_ConcRect.depth;

                        Parameter pra_a_1 = revitSection.LookupParameter(widthParam);
                        Parameter pra_b_1 = revitSection.LookupParameter(depthParam);

                        #endregion

                        try
                        {
                            using (Transaction trns = new Transaction(doc, "aa"))
                            {
                                trns.Start();
                                FailureHandlingOptions failOpt = trns.GetFailureHandlingOptions();
                                failOpt.SetFailuresPreprocessor(new RoomWarningSwallower());
                                trns.SetFailureHandlingOptions(failOpt);
                                double pra_a_2 = UnitUtils.ConvertToInternalUnits(Convert.ToDouble(dimentions[0]), dis);
                                double pra_b_2 = UnitUtils.ConvertToInternalUnits(Convert.ToDouble(dimentions[1]), dis);



                                pra_a_1.Set(pra_a_2);
                                pra_b_1.Set(pra_b_2);


                                trns.Commit();

                            }
                        }

                        catch (Exception ex2)
                        {

                            return null;


                        }

                    }


                    return revitSection;
                }
            }
            #endregion


            catch (Exception ex2)
            {
                System.Diagnostics.StackTrace trace = new System.Diagnostics.StackTrace(ex2, true);
                RevitUtils.Log.Add("Error in createRvtFamilySectionForBeam Method in RevitUtils (Section " + bm.Sec + " will be Dummy )... bm no " + bm.ID + ex2.Message + trace.GetFrame(trace.FrameCount - 1).GetFileLineNumber());
                return null;
            }

            return null;


        }

        public static Dictionary<string, FamilySymbol> createRvtFamilySectionForAllBeam(List<Frame> beams, Document doc, ForgeTypeId dis)
        {

            Dictionary<string, FamilySymbol> result = new Dictionary<string, FamilySymbol>();
            FamilySymbol notFountSection = new FilteredElementCollector(doc).WhereElementIsElementType().OfCategory(BuiltInCategory.OST_StructuralFraming).Cast<FamilySymbol>().FirstOrDefault(x => x.Name == "NotFoundSection");

            foreach (Frame bm in beams)
            {
                try
                {
                    FamilySymbol revitSection = null;

                    revitSection = createRvtFamilySectionForBeam(bm, doc, dis);
                    if (revitSection == null)
                    {
                        FamilySymbol notFoundSections2 = new FilteredElementCollector(doc).WhereElementIsElementType().OfCategory(BuiltInCategory.OST_StructuralFraming).Cast<FamilySymbol>().FirstOrDefault(x => x.Name == notFountSection.Name + "( " + bm.Sec + " )");

                        if (notFoundSections2 == null)
                        {
                            using (Transaction trns = new Transaction(doc, "aa"))
                            {
                                trns.Start();
                                notFoundSections2 = (FamilySymbol)notFountSection.Duplicate(notFountSection.Name + "( " + bm.Sec + " )");

                                trns.Commit();
                            }
                        }

                        revitSection = notFoundSections2;



                    }

                    result.Add(bm.Sec, revitSection);
                }


                catch (Exception)
                {
                    continue;
                }




            }


            return result;


        }

        public static FamilySymbol createRvtFamilySectionForColumn(Frame cl, Document doc, ForgeTypeId dis)
        {


            FamilySymbol revitSection = null;


            try
            {
                #region CreatePipeSection

                if (cl.Shape == COL_PIPE.shape)
                {
                    string dimentionswithX;
                    string[] dimentions;
                    COL_PIPE.getPipeDimentionsFromAnalysisSectionNames(cl, out dimentionswithX, out dimentions);

                    revitSection = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralColumns).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(f => f.FamilyName == COL_PIPE.FamilyName).FirstOrDefault(X => X.Name.Contains(dimentionswithX));

                    // new
                    if (revitSection == null)
                    {
                        revitSection = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralColumns).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(f => f.FamilyName == COL_PIPE.FamilyName).FirstOrDefault(X => X.Name.Contains(dimentionswithX.ToLower()));
                        if (revitSection != null)
                        {
                            using (Transaction trns = new Transaction(doc, "aa"))
                            {
                                trns.Start();
                                revitSection.Name = revitSection.Name.ToUpper();

                                trns.Commit();
                            }
                        }

                    }

                    // Define new familyType 
                    if (revitSection == null)
                    {

                        #region difintions 

                        FamilySymbol baseSection = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralColumns).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(f => f.FamilyName == COL_PIPE.FamilyName).FirstOrDefault();


                        using (Transaction trns = new Transaction(doc, "aa"))
                        {
                            trns.Start();
                            FailureHandlingOptions failOpt = trns.GetFailureHandlingOptions();
                            failOpt.SetFailuresPreprocessor(new RoomWarningSwallower());
                            trns.SetFailureHandlingOptions(failOpt);
                            string name = "PIPE " + dimentionswithX;
                            revitSection = baseSection.Duplicate(name) as FamilySymbol;
                            trns.Commit();
                        }

                        string outerDiameterParam = COL_PIPE.outerDiameter;
                        string thickParam = COL_PIPE.thick;


                        Parameter pra_a_1 = revitSection.LookupParameter(outerDiameterParam);
                        Parameter pra_b_1 = revitSection.LookupParameter(thickParam);


                        #endregion

                        try
                        {
                            using (Transaction trns = new Transaction(doc, "aa"))
                            {
                                trns.Start();
                                FailureHandlingOptions failOpt = trns.GetFailureHandlingOptions();
                                failOpt.SetFailuresPreprocessor(new RoomWarningSwallower());
                                trns.SetFailureHandlingOptions(failOpt);
                                double pra_a_2 = UnitUtils.ConvertToInternalUnits(Convert.ToDouble(dimentions[0]), dis);
                                double pra_b_2 = UnitUtils.ConvertToInternalUnits(Convert.ToDouble(dimentions[1]), dis);

                                pra_a_1.Set(pra_a_2);
                                pra_b_1.Set(pra_b_2);


                                trns.Commit();

                            }
                        }

                        catch (Exception ex2)
                        {




                        }

                    }


                    return revitSection;
                }

                #endregion

                #region BoxSections

                if (cl.Shape == COL_TUBE.shape)
                {

                    string dimentionswithX;
                    string[] dimentions;
                    int noOfX;
                    COL_TUBE.getTubeDimnsionsFromAnalysisSectionNames(cl, out dimentionswithX, out dimentions, out noOfX);

                    // check familyType existance

                    revitSection = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralColumns).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(f => f.FamilyName == COL_TUBE.FamilyName).FirstOrDefault(X => X.Name.Contains(dimentionswithX) && X.Name.Count(b => b == 'X') == noOfX);

                    // new
                    if (revitSection == null)
                    {
                        revitSection = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralColumns).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(f => f.FamilyName == COL_TUBE.FamilyName).FirstOrDefault(X => X.Name.Contains(dimentionswithX.ToLower()) && X.Name.Count(b => b == 'x') == noOfX);
                        if (revitSection != null)
                        {
                            using (Transaction trns = new Transaction(doc, "aa"))
                            {
                                trns.Start();
                                revitSection.Name = revitSection.Name.ToUpper();

                                trns.Commit();
                            }
                        }

                    }

                    // Define new familyType 
                    if (revitSection == null)
                    {

                        #region difintions 

                        FamilySymbol baseSection = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralColumns).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(f => f.FamilyName == COL_TUBE.FamilyName).FirstOrDefault();

                        using (Transaction trns = new Transaction(doc, "aa"))
                        {
                            trns.Start();
                            string name = "TUBE " + dimentionswithX;
                            revitSection = baseSection.Duplicate(name) as FamilySymbol;
                            trns.Commit();
                        }

                        string webDepthPram = COL_TUBE.webDepth;
                        string flgWidthParm = COL_TUBE.flgWidth;
                        string webThickParm = COL_TUBE.webThick;
                        string flgThickParm = COL_TUBE.flgThick;
                        string outerRadius = COL_TUBE.outerRad;
                        string innerRadius = COL_TUBE.innerRad;


                        Parameter pra_a_1 = revitSection.LookupParameter(webDepthPram);
                        Parameter pra_b_1 = revitSection.LookupParameter(flgWidthParm);
                        Parameter pra_c_1 = revitSection.LookupParameter(webThickParm);
                        Parameter pra_d_1 = revitSection.LookupParameter(flgThickParm);
                        Parameter pra_e_1 = revitSection.LookupParameter(outerRadius);
                        Parameter pra_f_1 = revitSection.LookupParameter(innerRadius);

                        #endregion


                        try
                        {
                            using (Transaction trns = new Transaction(doc, "aa"))
                            {
                                trns.Start();
                                FailureHandlingOptions failOpt = trns.GetFailureHandlingOptions();
                                failOpt.SetFailuresPreprocessor(new RoomWarningSwallower());
                                trns.SetFailureHandlingOptions(failOpt);
                                double pra_a_2 = UnitUtils.ConvertToInternalUnits(Convert.ToDouble(dimentions[0]), dis);
                                double pra_b_2 = UnitUtils.ConvertToInternalUnits(Convert.ToDouble(dimentions[1]), dis);
                                double pra_c_2 = UnitUtils.ConvertToInternalUnits(Convert.ToDouble(dimentions[2]), dis);
                                double pra_d_2 = 0;
                                // flgThick
                                if (dimentions.Length == 4)
                                {
                                    // if tf != tw
                                    pra_d_2 = UnitUtils.ConvertToInternalUnits(Convert.ToDouble(dimentions[3]), dis);
                                }


                                else
                                {
                                    // if Tf == tw
                                    pra_d_2 = UnitUtils.ConvertToInternalUnits(Convert.ToDouble(dimentions[2]), dis);
                                }

                                double pra_e_2 = UnitUtils.ConvertToInternalUnits(Convert.ToDouble(4.0), dis);
                                double pra_f_2 = UnitUtils.ConvertToInternalUnits(Convert.ToDouble(4.0), dis);





                                pra_a_1.Set(pra_a_2);
                                pra_b_1.Set(pra_b_2);
                                pra_c_1.Set(pra_c_2);
                                pra_d_1.Set(pra_d_2);
                                pra_e_1.Set(pra_e_2);
                                pra_f_1.Set(pra_f_2);

                                trns.Commit();
                            }
                        }

                        catch (Exception ex2)
                        {



                        }


                    }

                    return revitSection;

                }

                #endregion

                #region I_B

                if (cl.IsBUS == true && cl.Shape == COL_BUS_I.shape)
                {

                    string webThick, flgThick, consWebDepth, consFlgDepth;
                    COL_BUS_I.getBUS_IDimensionsFromAnalysisSectionNames(cl, out webThick, out flgThick, out consWebDepth, out consFlgDepth);


                    revitSection = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralColumns).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(f => f.FamilyName == COL_BUS_I.FamilyName).FirstOrDefault(X => X.Name == cl.Sec);

                    // check familyType existance
                    if (revitSection == null)
                    {

                        // Define new familyType 
                        #region  Definitions
                        using (Transaction trns = new Transaction(doc, "aa"))
                        {
                            FamilySymbol baseSection = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralColumns).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(f => f.FamilyName == COL_BUS_I.FamilyName).FirstOrDefault();
                            trns.Start();
                            revitSection = baseSection.Duplicate(cl.Sec) as FamilySymbol;
                            trns.Commit();
                        }


                        string totalDepthParam = COL_BUS_I.totalBeamDepth;
                        string flgWidthParm = COL_BUS_I.flgWidth;
                        string webThickParm = COL_BUS_I.webThick;
                        string flgThickParm = COL_BUS_I.flgThick;

                        Parameter pra_a_1 = revitSection.LookupParameter(totalDepthParam);
                        Parameter pra_b_1 = revitSection.LookupParameter(flgWidthParm);
                        Parameter pra_c_1 = revitSection.LookupParameter(webThickParm);
                        Parameter pra_d_1 = revitSection.LookupParameter(flgThickParm);
                        #endregion

                        try
                        {
                            using (Transaction trns = new Transaction(doc, "aa"))
                            {
                                trns.Start();

                                double pra_a_2 = UnitUtils.ConvertToInternalUnits(Convert.ToDouble(consWebDepth), dis);
                                double pra_b_2 = UnitUtils.ConvertToInternalUnits(Convert.ToDouble(consFlgDepth), dis);
                                double pra_c_2 = UnitUtils.ConvertToInternalUnits(Convert.ToDouble(webThick), dis);
                                double pra_d_2 = UnitUtils.ConvertToInternalUnits(Convert.ToDouble(flgThick), dis);



                                pra_a_1.Set(pra_a_2);
                                pra_b_1.Set(pra_b_2);
                                pra_c_1.Set(pra_c_2);
                                pra_d_1.Set(pra_d_2);


                                trns.Commit();
                            }
                        }

                        catch (Exception ex2)
                        {



                        }


                    }

                    return revitSection;

                }







                #endregion

                #region I_H

                if (cl.IsBUS == false && cl.Shape == "I")
                {
                    #region filterBeteenTypes
                    string numeric = null;
                    string FamilyType = null;


                    if (cl.Sec.Contains("HE") && cl.Shape == "I")
                    {
                        if (cl.Sec.Contains("M"))
                        {
                            //string numeric = new String(cl.Sec.Where(Char.IsDigit).ToArray());
                            //revitSection = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralColumns).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(f => f.FamilyName == HEB_Beam).FirstOrDefault(X => new String(X.Name.Where(Char.IsDigit).ToArray()) == numeric);
                            FamilyType = "HEM";
                            goto IBeamFamilySelection;

                        }

                        else if (cl.Sec.Contains("B"))
                        {
                            FamilyType = "HEB";
                            goto IBeamFamilySelection;
                        }


                        else if (cl.Sec.Contains("A"))
                        {
                            int count = ExcelUtils.CountCharsUsingForeach(cl.Sec, 'A');
                            if (count == 2)
                            {
                                FamilyType = "HEAA";
                                goto IBeamFamilySelection;
                            }

                            else
                            {
                                FamilyType = "HEA";
                                goto IBeamFamilySelection;
                            }
                        }

                    }

                    else if (cl.Sec.Contains("IP") && cl.Shape == "I")
                    {
                        if (cl.Sec.Contains("O"))
                        {
                            FamilyType = "IPEO";
                            goto IBeamFamilySelection;

                        }

                        else if (cl.Sec.Contains("A"))
                        {
                            FamilyType = "IPEA";
                            goto IBeamFamilySelection;
                        }

                        else if (cl.Sec.Contains("R"))
                        {
                            FamilyType = "IPER";
                            goto IBeamFamilySelection;

                        }

                        else if (cl.Sec.Contains("V"))
                        {
                            FamilyType = "IPEV";
                            goto IBeamFamilySelection;
                        }

                        else if (cl.Sec.Contains("E"))
                        {
                            FamilyType = "IPE";
                            goto IBeamFamilySelection;

                        }

                    }

                    else if ((cl.Sec.Contains("UKC") || cl.Sec.Contains("UC")) && cl.Shape == "I" && cl.IsBUS == false && (cl.Lib.ToUpper().Contains("BSSHAPES") || cl.Lib.ToUpper().Contains("BRITISH") || cl.Lib.ToUpper().Contains("AUSTRAILIAN")))
                    {

                        FamilyType = "UC";
                        goto IBeamFamilySelection;
                    }

                    else if ((cl.Sec.Contains("UKB") || cl.Sec.Contains("UB")) && cl.Shape == "I" && cl.IsBUS == false && (cl.Lib.ToUpper().Contains("BSSHAPES") || cl.Lib.ToUpper().Contains("BRITISH") || cl.Lib.ToUpper().Contains("AUSTRAILIAN")))

                    {
                        FamilyType = "UB";
                        goto IBeamFamilySelection;
                    }

                    else if (cl.Sec.Contains("W") && cl.Shape == "I" && cl.IsBUS == false && (cl.Lib.ToUpper().Contains("AISC") || cl.Lib.Contains("American") || cl.Lib.Contains("Canadian")))
                    {
                        FamilyType = "W";
                        goto IBeamFamilySelection;
                    }



                #endregion

                #region SelectFamilyType

                IBeamFamilySelection:
                    switch (FamilyType)
                    {

                        case "HEAA":
                            numeric = new String(cl.Sec.Where(Char.IsDigit).ToArray());
                            revitSection = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralColumns).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(f => f.FamilyName == COL_HEAA.FamilyName).FirstOrDefault(X => new String(X.Name.Where(Char.IsDigit).ToArray()) == numeric);
                            if (revitSection == null)
                            {
                                return null;
                            }
                            return revitSection;


                        case "HEA":
                            numeric = new String(cl.Sec.Where(Char.IsDigit).ToArray());
                            revitSection = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralColumns).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(f => f.FamilyName == COL_HEA.FamilyName).FirstOrDefault(X => new String(X.Name.Where(Char.IsDigit).ToArray()) == numeric);
                            if (revitSection == null)
                            {
                                return null;
                            }
                            return revitSection;


                        case "HEB":
                            numeric = new String(cl.Sec.Where(Char.IsDigit).ToArray());
                            revitSection = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralColumns).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(f => f.FamilyName == COL_HEB.FamilyName).FirstOrDefault(X => new String(X.Name.Where(Char.IsDigit).ToArray()) == numeric);
                            if (revitSection == null)
                            {
                                return null;
                            }
                            return revitSection;

                        case "HEM":
                            numeric = new String(cl.Sec.Where(Char.IsDigit).ToArray());
                            revitSection = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralColumns).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(f => f.FamilyName == COL_HEM.FamilyName).FirstOrDefault(X => new String(X.Name.Where(Char.IsDigit).ToArray()) == numeric);
                            if (revitSection == null)
                            {
                                return null;
                            }
                            return revitSection;

                        case "IPE":
                            numeric = new String(cl.Sec.Where(Char.IsDigit).ToArray());
                            revitSection = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralColumns).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(f => f.FamilyName == COL_IPE.FamilyName).FirstOrDefault(X => new String(X.Name.Where(Char.IsDigit).ToArray()) == numeric);
                            if (revitSection == null)
                            {
                                return null;
                            }
                            return revitSection;

                        case "IPEA":
                            numeric = new String(cl.Sec.Where(Char.IsDigit).ToArray());
                            revitSection = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralColumns).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(f => f.FamilyName == COL_IPEA.FamilyName).FirstOrDefault(X => new String(X.Name.Where(Char.IsDigit).ToArray()) == numeric);
                            if (revitSection == null)
                            {
                                return null;
                            }
                            return revitSection;

                        case "IPEO":
                            numeric = new String(cl.Sec.Where(Char.IsDigit).ToArray());
                            revitSection = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralColumns).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(f => f.FamilyName == COL_IPEO.FamilyName).FirstOrDefault(X => new String(X.Name.Where(Char.IsDigit).ToArray()) == numeric);
                            if (revitSection == null)
                            {
                                return null;
                            }
                            return revitSection;


                        ////////////////////////
                        case "UB":

                            char cc = cl.Sec.SkipWhile(c => !char.IsDigit(c)).ToList()[0];
                            int lastSpacecharacterIndex = cl.Sec.IndexOf(cc);
                            string dimentionswithX = cl.Sec.Substring(lastSpacecharacterIndex);

                            revitSection = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralColumns).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(f => f.FamilyName == COL_UB.FamilyName).FirstOrDefault(X => X.Name.ToUpper().Contains(dimentionswithX.ToUpper()));
                            if (revitSection == null)
                            {
                                return null;
                            }
                            return revitSection;



                        case "UC":

                            cc = cl.Sec.SkipWhile(c => !char.IsDigit(c)).ToList()[0];
                            lastSpacecharacterIndex = cl.Sec.IndexOf(cc);
                            dimentionswithX = cl.Sec.Substring(lastSpacecharacterIndex);

                            revitSection = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralColumns).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(f => f.FamilyName == COL_UC.FamilyName).FirstOrDefault(X => X.Name.ToUpper().Contains(dimentionswithX.ToUpper()));

                            if (revitSection == null)
                            {
                                return null;
                            }
                            return revitSection;



                        case "W":
                            cc = cl.Sec.SkipWhile(c => !char.IsDigit(c)).ToList()[0];
                            lastSpacecharacterIndex = cl.Sec.IndexOf(cc);
                            dimentionswithX = cl.Sec.Substring(lastSpacecharacterIndex);
                            revitSection = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralColumns).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(f => f.FamilyName == COL_W.FamilyName).FirstOrDefault(X => X.Name.ToUpper().Contains(dimentionswithX.ToUpper()));

                            if (revitSection == null)
                            {
                                return null;
                            }
                            return revitSection;


                        default:
                            return null;
                    }

                    #endregion
                }


                #endregion

                #region C_H 

                if (cl.IsBUS != true && cl.Shape == COL_UPN.shape)
                {
                    string dimentionswithX = COL_UPN.getChannelDimnsionsFromAnalysisSectionNames(cl);

                    // check familyType existance
                    revitSection = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralColumns).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(f => f.FamilyName == COL_UPN.FamilyName).FirstOrDefault(X => X.Name.Contains(dimentionswithX));

                    // Define new familyType 
                    if (revitSection == null)
                    {
                        return null;
                    }

                    return revitSection;

                }

                #endregion

                #region I_T

                if (cl.Shape == COL_BUS_I_Tapered.shape)
                {

                    string webThick, flgThick, startWebDepth, endWebDepth, consFlgDepth;
                    COL_BUS_I_Tapered.getTapered_IDimensionsFromAnalysisSectionNames(cl, out webThick, out flgThick, out startWebDepth, out endWebDepth, out consFlgDepth);


                    // check familyType existance
                    revitSection = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralColumns).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(f => f.FamilyName == COL_BUS_I_Tapered.FamilyName).FirstOrDefault(X => X.Name == cl.Sec);

                    // Define new familyType 
                    if (revitSection == null)
                    {

                        #region  Definitions
                        using (Transaction trns = new Transaction(doc, "aa"))
                        {
                            FamilySymbol baseSection = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralColumns).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(f => f.FamilyName == COL_BUS_I_Tapered.FamilyName).FirstOrDefault();
                            trns.Start();
                            revitSection = baseSection.Duplicate(cl.Sec) as FamilySymbol;
                            trns.Commit();
                        }


                        string startWebDepthParam = COL_BUS_I_Tapered.startWebDepth;
                        string endWebDepthParam = COL_BUS_I_Tapered.endWebDepth;
                        string startFlgWidthParm = COL_BUS_I_Tapered.startFlgWidth;
                        string endFlgWidthParm = COL_BUS_I_Tapered.endFlgWidth;
                        string webThickParm = COL_BUS_I_Tapered.webThick;
                        string flgThickParm = COL_BUS_I_Tapered.flgThick;

                        Parameter pra_a_1 = revitSection.LookupParameter(startWebDepthParam);
                        Parameter pra_b_1 = revitSection.LookupParameter(endWebDepthParam);
                        Parameter pra_c_1 = revitSection.LookupParameter(startFlgWidthParm);
                        Parameter pra_d_1 = revitSection.LookupParameter(endFlgWidthParm);
                        Parameter pra_e_1 = revitSection.LookupParameter(webThickParm);
                        Parameter pra_f_1 = revitSection.LookupParameter(flgThickParm);

                        #endregion


                        try
                        {
                            using (Transaction trns = new Transaction(doc, "aa"))
                            {
                                trns.Start();

                                double pra_a_2 = UnitUtils.ConvertToInternalUnits(Convert.ToDouble(startWebDepth), dis);
                                double pra_b_2 = UnitUtils.ConvertToInternalUnits(Convert.ToDouble(endWebDepth), dis);
                                double pra_c_2 = UnitUtils.ConvertToInternalUnits(Convert.ToDouble(consFlgDepth), dis);
                                double pra_d_2 = UnitUtils.ConvertToInternalUnits(Convert.ToDouble(consFlgDepth), dis);
                                double pra_e_2 = UnitUtils.ConvertToInternalUnits(Convert.ToDouble(webThick), dis);
                                double pra_f_2 = UnitUtils.ConvertToInternalUnits(Convert.ToDouble(flgThick), dis);



                                pra_a_1.Set(pra_a_2);
                                pra_b_1.Set(pra_b_2);
                                pra_c_1.Set(pra_c_2);
                                pra_d_1.Set(pra_d_2);
                                pra_e_1.Set(pra_e_2);
                                pra_f_1.Set(pra_f_2);




                                trns.Commit();

                            }
                        }

                        catch (Exception ex2)
                        {



                        }


                    }

                    return revitSection;

                }




                #endregion

                #region Pipe_T

                if (cl.IsBUS == true && cl.Shape == COL_BUS_PIPE_Tapered.shape)
                {

                    string startOuterDiamaterValue, endOuterDiamaterValue, PipeThicknessValue;
                    COL_BUS_PIPE_Tapered.getTaperedPipeDimensionsFromAnalysisSectionNames(cl, out startOuterDiamaterValue, out endOuterDiamaterValue, out PipeThicknessValue);


                    // check familyType existance
                    revitSection = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralColumns).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(f => f.FamilyName == COL_BUS_PIPE_Tapered.FamilyName).FirstOrDefault(X => X.Name == cl.Sec);

                    // Define new familyType 
                    if (revitSection == null)
                    {

                        #region difintions 

                        FamilySymbol baseSection = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralColumns).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(f => f.FamilyName == COL_BUS_PIPE_Tapered.FamilyName).FirstOrDefault();


                        using (Transaction trns = new Transaction(doc, "aa"))
                        {
                            trns.Start();
                            FailureHandlingOptions failOpt = trns.GetFailureHandlingOptions();
                            failOpt.SetFailuresPreprocessor(new RoomWarningSwallower());
                            trns.SetFailureHandlingOptions(failOpt);
                            revitSection = baseSection.Duplicate(cl.Sec) as FamilySymbol;
                            trns.Commit();
                        }

                        string startInnerDiamParam = COL_BUS_PIPE_Tapered.startInnerDiam;
                        string startOuterDiamParam = COL_BUS_PIPE_Tapered.startOuterDiam;
                        string endInnerDiamParam = COL_BUS_PIPE_Tapered.endInnerDiam;
                        string endOuterDiamParam = COL_BUS_PIPE_Tapered.endOuterDiam;


                        Parameter pra_a_1 = revitSection.LookupParameter(startInnerDiamParam);
                        Parameter pra_b_1 = revitSection.LookupParameter(startOuterDiamParam);
                        Parameter pra_c_1 = revitSection.LookupParameter(endInnerDiamParam);
                        Parameter pra_d_1 = revitSection.LookupParameter(endOuterDiamParam);


                        #endregion

                        try
                        {
                            using (Transaction trns = new Transaction(doc, "aa"))
                            {
                                trns.Start();
                                FailureHandlingOptions failOpt = trns.GetFailureHandlingOptions();
                                failOpt.SetFailuresPreprocessor(new RoomWarningSwallower());
                                trns.SetFailureHandlingOptions(failOpt);
                                double pra_a_2 = UnitUtils.ConvertToInternalUnits(Convert.ToDouble(startOuterDiamaterValue) - 2 * Convert.ToDouble(PipeThicknessValue), dis);
                                double pra_b_2 = UnitUtils.ConvertToInternalUnits(Convert.ToDouble(startOuterDiamaterValue), dis);
                                double pra_c_2 = UnitUtils.ConvertToInternalUnits(Convert.ToDouble(endOuterDiamaterValue) - 2 * Convert.ToDouble(PipeThicknessValue), dis);
                                double pra_d_2 = UnitUtils.ConvertToInternalUnits(Convert.ToDouble(endOuterDiamaterValue), dis);



                                pra_a_1.Set(pra_a_2);
                                pra_b_1.Set(pra_b_2);
                                pra_c_1.Set(pra_c_2);
                                pra_d_1.Set(pra_c_2);



                                trns.Commit();

                            }
                        }

                        catch (Exception ex2)
                        {


                        }

                    }

                    return revitSection;

                }

                #endregion

                #region Tube_T

                if (cl.Shape == COL_BUS_TUBE_Tapered.shape)
                {


                    string startWebDepth, endWebDepth, webThick, startFlgWidth, endFlgWidth, flgThick;
                    COL_BUS_TUBE_Tapered.getTaperedTubeDimensionsFromAnalysisSectionNames(cl, out startWebDepth, out endWebDepth, out webThick, out startFlgWidth, out endFlgWidth, out flgThick);


                    // check familyType existance
                    revitSection = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralColumns).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(f => f.FamilyName == COL_BUS_TUBE_Tapered.FamilyName).FirstOrDefault(X => X.Name == cl.Sec);

                    // Define new familyType 
                    if (revitSection == null)
                    {

                        #region  Definitions
                        using (Transaction trns = new Transaction(doc, "aa"))
                        {
                            FamilySymbol baseSection = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralColumns).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(f => f.FamilyName == COL_BUS_TUBE_Tapered.FamilyName).FirstOrDefault();
                            trns.Start();
                            FailureHandlingOptions failOpt = trns.GetFailureHandlingOptions();
                            failOpt.SetFailuresPreprocessor(new RoomWarningSwallower());
                            trns.SetFailureHandlingOptions(failOpt);
                            revitSection = baseSection.Duplicate(cl.Sec) as FamilySymbol;
                            trns.Commit();
                        }


                        string startWebDepthParm = COL_BUS_TUBE_Tapered.startWebDepth;
                        string endWebDepthParm = COL_BUS_TUBE_Tapered.endWebDepth;
                        string startFlgWidthParm = COL_BUS_TUBE_Tapered.startFlgWidth;
                        string endFlgWidthParm = COL_BUS_TUBE_Tapered.endFlgWidth;
                        string startWebThickParm = COL_BUS_TUBE_Tapered.startWebThick;
                        string endWebThickParm = COL_BUS_TUBE_Tapered.endFlgThick;
                        string StartFlgThickParm = COL_BUS_TUBE_Tapered.startFlgThick;
                        string endFlgThickParm = COL_BUS_TUBE_Tapered.endFlgThick;



                        Parameter pra_a_1 = revitSection.LookupParameter(startWebDepthParm);
                        Parameter pra_b_1 = revitSection.LookupParameter(endWebDepthParm);
                        Parameter pra_c_1 = revitSection.LookupParameter(startFlgWidthParm);
                        Parameter pra_d_1 = revitSection.LookupParameter(endFlgWidthParm);
                        Parameter pra_e_1 = revitSection.LookupParameter(startWebThickParm);
                        Parameter pra_f_1 = revitSection.LookupParameter(endWebThickParm);
                        Parameter pra_g_1 = revitSection.LookupParameter(StartFlgThickParm);
                        Parameter pra_h_1 = revitSection.LookupParameter(endFlgThickParm);

                        #endregion


                        try
                        {
                            using (Transaction trns = new Transaction(doc, "aa"))
                            {
                                trns.Start();
                                FailureHandlingOptions failOpt = trns.GetFailureHandlingOptions();
                                failOpt.SetFailuresPreprocessor(new RoomWarningSwallower());
                                trns.SetFailureHandlingOptions(failOpt);
                                double pra_a_2 = UnitUtils.ConvertToInternalUnits(Convert.ToDouble(startWebDepth), dis);
                                double pra_b_2 = UnitUtils.ConvertToInternalUnits(Convert.ToDouble(endWebDepth), dis);
                                double pra_c_2 = UnitUtils.ConvertToInternalUnits(Convert.ToDouble(startFlgWidth), dis);
                                double pra_d_2 = UnitUtils.ConvertToInternalUnits(Convert.ToDouble(endFlgWidth), dis);
                                double pra_e_2 = UnitUtils.ConvertToInternalUnits(Convert.ToDouble(webThick), dis);
                                double pra_f_2 = UnitUtils.ConvertToInternalUnits(Convert.ToDouble(webThick), dis);
                                double pra_g_2 = UnitUtils.ConvertToInternalUnits(Convert.ToDouble(flgThick), dis);
                                double pra_h_2 = UnitUtils.ConvertToInternalUnits(Convert.ToDouble(flgThick), dis);



                                pra_a_1.Set(pra_a_2);
                                pra_b_1.Set(pra_b_2);
                                pra_c_1.Set(pra_c_2);
                                pra_d_1.Set(pra_d_2);
                                pra_e_1.Set(pra_e_2);
                                pra_f_1.Set(pra_f_2);
                                pra_g_1.Set(pra_g_2);
                                pra_h_1.Set(pra_h_2);



                                trns.Commit();
                            }
                        }

                        catch (Exception ex2)
                        {


                        }


                    }


                    return revitSection;

                }



                #endregion

                #region Angles

                if (cl.Shape == COL_L_EQUAL.shape || cl.Shape == COL_L_UNEQUAL.shape || cl.Shape == COL_2L_EQUALANGLE.shape)
                {

                    bool isEQSingleAngle, isUNEQSingleAngle, isEQDblAngle;
                    string dimentionswithX, firstDimention;
                    RevitFamily.getAnglesDimensionsFromAnalysisSectionNames(cl, out isEQSingleAngle, out isUNEQSingleAngle, out isEQDblAngle, out dimentionswithX, out firstDimention);



                    #region  EQAngle

                    if (isEQSingleAngle)
                    {


                        string dimentionswithXDublicated = dimentionswithX;
                        int countOfX = dimentionswithX.ToUpper().Count(c => c == 'X');

                        if (countOfX < 2)
                        {
                            dimentionswithXDublicated = dimentionswithX.Insert(firstDimention.IndexOf("X") + 1, firstDimention + "X"); ;

                        }

                        // check familyType existance
                        //revitSection = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralColumns).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(f => f.FamilyName == COL_L_EQUAL.FamilyName).FirstOrDefault(X => X.Name.Contains(dimentionswithXDublicated) || X.Name.Contains(dimentionswithX));
                        revitSection = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralColumns).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(f => f.FamilyName == COL_L_EQUAL.FamilyName).FirstOrDefault(X => X.Name.Contains(dimentionswithXDublicated));                     // Define new familyType 
                        if (revitSection == null)
                        {

                            return null;

                        }

                        return revitSection;

                    }

                    #endregion

                    #region  UQAngle

                    if (isUNEQSingleAngle)
                    {

                        // check familyType existance
                        revitSection = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralColumns).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(f => f.FamilyName == COL_L_UNEQUAL.FamilyName).FirstOrDefault(X => X.Name.Contains(dimentionswithX));

                        // Define new familyType 
                        if (revitSection == null)
                        {

                            return null;

                        }

                        return revitSection;

                    }

                    #endregion

                    #region  DblAngle

                    if (isEQDblAngle)
                    {

                        string dimentionswithXDublicated = dimentionswithX;

                        int countOfX = dimentionswithX.ToUpper().Count(c => c == 'X');
                        if (countOfX < 2)
                        {
                            dimentionswithXDublicated = dimentionswithX.Insert(firstDimention.IndexOf("X") + 1, firstDimention + "X"); ;

                        }

                        // check familyType existance
                        //revitSection = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralColumns).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(f => f.FamilyName == COL_2L_EQUALANGLE.FamilyName).FirstOrDefault(X => X.Name.Contains(dimentionswithXDublicated) || X.Name.Contains(dimentionswithX));
                        revitSection = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralColumns).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(f => f.FamilyName == COL_2L_EQUALANGLE.FamilyName).FirstOrDefault(X => X.Name.Contains(dimentionswithXDublicated));


                        // Define new familyType 
                        if (revitSection == null)
                        {

                            return null;

                        }

                        return revitSection;

                    }

                    #endregion


                }

                #endregion

                #region ConcRect
                if (cl.Shape == COL_ConcRect.shape)
                {
                    string dimentionswithX;
                    string[] dimentions;
                    COL_ConcRect.getConcRectDimensionsFromAnalysisSectionNames(cl, out dimentionswithX, out dimentions);

                    revitSection = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralColumns).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(f => f.FamilyName == COL_ConcRect.FamilyName).FirstOrDefault(X => X.Name.Contains(dimentionswithX));

                    // Define new familyType 
                    if (revitSection == null)
                    {

                        #region difintions 

                        FamilySymbol baseSection = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralColumns).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(f => f.FamilyName == COL_ConcRect.FamilyName).FirstOrDefault();


                        using (Transaction trns = new Transaction(doc, "aa"))
                        {
                            trns.Start();
                            FailureHandlingOptions failOpt = trns.GetFailureHandlingOptions();
                            failOpt.SetFailuresPreprocessor(new RoomWarningSwallower());
                            trns.SetFailureHandlingOptions(failOpt);
                            string name = "COL " + dimentionswithX;
                            revitSection = baseSection.Duplicate(name) as FamilySymbol;
                            trns.Commit();
                        }

                        string widthParam = COL_ConcRect.width;
                        string depthParam = COL_ConcRect.depth;

                        Parameter pra_a_1 = revitSection.LookupParameter(widthParam);
                        Parameter pra_b_1 = revitSection.LookupParameter(depthParam);

                        #endregion

                        try
                        {
                            using (Transaction trns = new Transaction(doc, "aa"))
                            {
                                trns.Start();
                                FailureHandlingOptions failOpt = trns.GetFailureHandlingOptions();
                                failOpt.SetFailuresPreprocessor(new RoomWarningSwallower());
                                trns.SetFailureHandlingOptions(failOpt);
                                double pra_a_2 = UnitUtils.ConvertToInternalUnits(Convert.ToDouble(dimentions[0]), dis);
                                double pra_b_2 = UnitUtils.ConvertToInternalUnits(Convert.ToDouble(dimentions[1]), dis);



                                pra_a_1.Set(pra_a_2);
                                pra_b_1.Set(pra_b_2);


                                trns.Commit();

                            }
                        }

                        catch (Exception ex2)
                        {

                            return null;


                        }

                    }


                    return revitSection;
                }
            }
            #endregion


            catch (Exception ex2)
            {
                System.Diagnostics.StackTrace trace = new System.Diagnostics.StackTrace(ex2, true);
                RevitUtils.Log.Add("Error in createRvtFamilySectionForColumn Method in RevitUtils (Section " + cl.Sec + " will be Dummy )... col no " + cl.ID + ex2.Message + trace.GetFrame(trace.FrameCount - 1).GetFileLineNumber());
                return null;
            }


            return null;

        }

        public static Dictionary<string, FamilySymbol> createRvtFamilySectionForAllColumn(List<Frame> columns, Document doc, ForgeTypeId dis)
        {
            // parameters ( bm.isBUS, bm.Shape,doc,dis, bm.libarary (used in w,ukc and ukb))

            Dictionary<string, FamilySymbol> result = new Dictionary<string, FamilySymbol>();
            FamilySymbol notFountSection = new FilteredElementCollector(doc).WhereElementIsElementType().OfCategory(BuiltInCategory.OST_StructuralColumns).Cast<FamilySymbol>().FirstOrDefault(x => x.Name == "NotFoundSection");


            foreach (Frame cl in columns)
            {
                FamilySymbol revitSection = null;

                try
                {
                    revitSection = createRvtFamilySectionForColumn(cl, doc, dis);

                    if (revitSection == null)
                    {

                        FamilySymbol notFountSection2 = new FilteredElementCollector(doc).WhereElementIsElementType().OfCategory(BuiltInCategory.OST_StructuralColumns).Cast<FamilySymbol>().FirstOrDefault(x => x.Name == "NotFoundSection" + "( " + cl.Sec + " )");

                        if (notFountSection2 == null)
                        {
                            using (Transaction trns = new Transaction(doc, "aa"))
                            {
                                trns.Start();
                                FailureHandlingOptions failOpt = trns.GetFailureHandlingOptions();
                                failOpt.SetFailuresPreprocessor(new RoomWarningSwallower());
                                trns.SetFailureHandlingOptions(failOpt);

                                notFountSection2 = (FamilySymbol)notFountSection.Duplicate(notFountSection.Name + "( " + cl.Sec + " )");
                                //notFountSection2.LookupParameter(COL_TUBE.innerRad).Set(UnitUtils.ConvertToInternalUnits(4, dis));
                                //notFoundSections2.LookupParameter(COL_TUBE.outerRad).Set(UnitUtils.ConvertToInternalUnits(4, dis));
                                trns.Commit();


                            }
                        }
                        revitSection = notFountSection2;


                    }
                    result.Add(cl.Sec, revitSection);

                }


                catch (Exception ex2)
                {
                    System.Diagnostics.StackTrace trace = new System.Diagnostics.StackTrace(ex2, true);
                    TaskDialog.Show("Error", ex2.Message + "section name " + cl.Sec + "for columns" + " ,line of the error is " + trace.GetFrame(trace.FrameCount - 1).GetFileLineNumber());
                    continue;
                }


            }






            return result;

        }

        public static void mappBetweenSections(SectType familyType, ref string shape, ref string library, ref bool isBus)
        {
            isBus = false;

            switch (familyType)
            {
                case SectType.COL_BUS_I:
                    shape = SAP2000v1.eFramePropType.I.ToString();
                    isBus = true;
                    break;
                case SectType.COL_HEA:
                    shape = SAP2000v1.eFramePropType.I.ToString();
                    library = "euro";
                    break;
                case SectType.COL_HEAA:
                    shape = SAP2000v1.eFramePropType.I.ToString();
                    library = "euro";

                    break;
                case SectType.COL_HEB:
                    shape = SAP2000v1.eFramePropType.I.ToString();
                    library = "euro";

                    break;
                case SectType.COL_HEM:
                    shape = SAP2000v1.eFramePropType.I.ToString();
                    library = "euro";

                    break;
                case SectType.COL_IPE:
                    shape = SAP2000v1.eFramePropType.I.ToString();
                    library = "euro";

                    break;
                case SectType.COL_IPEA:
                    shape = SAP2000v1.eFramePropType.I.ToString();
                    library = "euro";

                    break;
                case SectType.COL_IPEO:
                    shape = SAP2000v1.eFramePropType.I.ToString();
                    library = "euro";

                    break;
                case SectType.COL_L_EQUAL:
                    shape = SAP2000v1.eFramePropType.Angle.ToString();
                    break;
                case SectType.COL_2L_EQUALANGLE:
                    shape = SAP2000v1.eFramePropType.DblAngle.ToString();

                    break;
                case SectType.COL_L_UNEQUAL:
                    shape = SAP2000v1.eFramePropType.Angle.ToString();

                    break;
                case SectType.COL_PIPE:
                    shape = SAP2000v1.eFramePropType.Pipe.ToString();
                    break;
                case SectType.COL_TUBE:
                    shape = SAP2000v1.eFramePropType.Box.ToString();

                    break;
                case SectType.COL_UPN:
                    shape = SAP2000v1.eFramePropType.Channel.ToString();

                    break;
                case SectType.COL_BUS_I_Tapered:
                    shape = "VARI";
                    isBus = true;
                    break;

                case SectType.COL_BUS_PIPE_Tapered:
                    shape = "VARPipe";
                    isBus = true;
                    break;

                case SectType.COL_BUS_TUBE_Tapered:
                    shape = "VARBox";
                    isBus = true;
                    break;

                case SectType.COL_W:
                    shape = SAP2000v1.eFramePropType.I.ToString();
                    library = "AISC";

                    break;
                case SectType.COL_UB:
                    shape = "I";
                    library = "BSShapes";

                    break;
                case SectType.COL_UC:
                    shape = "I";
                    library = "BSShapes";
                    break;



                case SectType.BM_BUS_I:
                    shape = SAP2000v1.eFramePropType.I.ToString();
                    isBus = true;
                    break;
                case SectType.BM_HEA:
                    shape = SAP2000v1.eFramePropType.I.ToString();
                    library = "euro";
                    break;
                case SectType.BM_HEAA:
                    shape = SAP2000v1.eFramePropType.I.ToString();
                    library = "euro";

                    break;
                case SectType.BM_HEB:
                    shape = SAP2000v1.eFramePropType.I.ToString();
                    library = "euro";

                    break;
                case SectType.BM_HEM:
                    shape = SAP2000v1.eFramePropType.I.ToString();
                    library = "euro";

                    break;
                case SectType.BM_IPE:
                    shape = SAP2000v1.eFramePropType.I.ToString();
                    library = "euro";

                    break;
                case SectType.BM_IPEA:
                    shape = SAP2000v1.eFramePropType.I.ToString();
                    library = "euro";

                    break;
                case SectType.BM_IPEO:
                    shape = SAP2000v1.eFramePropType.I.ToString();
                    library = "euro";

                    break;
                case SectType.BM_L_EQUAL:
                    shape = SAP2000v1.eFramePropType.Angle.ToString();
                    break;
                case SectType.BM_2L_EQUALANGLE:
                    shape = SAP2000v1.eFramePropType.DblAngle.ToString();

                    break;
                case SectType.BM_L_UNEQUAL:
                    shape = SAP2000v1.eFramePropType.Angle.ToString();

                    break;
                case SectType.BM_PIPE:
                    shape = SAP2000v1.eFramePropType.Pipe.ToString();
                    break;
                case SectType.BM_TUBE:
                    shape = SAP2000v1.eFramePropType.Box.ToString();

                    break;
                case SectType.BM_UPN:
                    shape = SAP2000v1.eFramePropType.Channel.ToString();

                    break;
                case SectType.BM_BUS_I_Tapered:
                    shape = "VARI";
                    isBus = true;
                    break;

                case SectType.BM_BUS_PIPE_Tapered:
                    shape = "VARPipe";
                    isBus = true;
                    break;

                case SectType.BM_BUS_TUBE_Tapered:
                    shape = "VARBox";
                    isBus = true;
                    break;

                case SectType.BM_W:
                    shape = SAP2000v1.eFramePropType.I.ToString();
                    library = "AISC";

                    break;
                case SectType.BM_UB:
                    shape = "I";
                    library = "BSShapes";

                    break;
                case SectType.BM_UC:
                    shape = "I";
                    library = "BSShapes";
                    break;

            }


        }

        #endregion

        #region AssignColors
        #region AssignColors
        public static void SetRevitColorsByIDs(Document doc, List<ElementId> ids, string Color)
        {
            using (Transaction trans = new Transaction(doc, "Assign Colors to Family Types"))
            {
                trans.Start();


                Color c;


                FillPatternElement solidPatternId = new FilteredElementCollector(doc)
                       .WherePasses(new ElementClassFilter(typeof(FillPatternElement)))
                       .ToElements().Cast<FillPatternElement>().ToList().First(x => x.GetFillPattern().IsSolidFill);


                OverrideGraphicSettings overrideSettings = new OverrideGraphicSettings();
                overrideSettings.SetSurfaceBackgroundPatternId(solidPatternId.Id);

                if (Color == "R")
                {
                    c = new Color(255, 0, 0);
                }
                else if (Color == "G")
                {
                    c = new Color(0, 255, 0);

                }

                else if (Color == "Y")
                {
                    c = new Color(255, 255, 0);

                }
                else
                {
                    c = new Color(0, 0, 0);
                }
                overrideSettings.SetSurfaceBackgroundPatternColor(c);


                foreach (var id in ids)
                {
                    doc.ActiveView.SetElementOverrides(id, overrideSettings);
                }
                trans.Commit();
            }
        }

        public static void changeColorBasedOnSections(Document doc)
        {

            try
            {


                // Define a dictionary to store family type ids and colors
                Dictionary<ElementId, Color> typeColorMap = new Dictionary<ElementId, Color>();

                // Collect all structural framing elements
                FilteredElementCollector framingCollector = new FilteredElementCollector(doc);
                ICollection<Element> framingElements = framingCollector
                    .OfCategory(BuiltInCategory.OST_StructuralFraming)
                    .OfClass(typeof(FamilyInstance))
                    .ToElements();

                // Collect all structural columns
                FilteredElementCollector columnCollector = new FilteredElementCollector(doc);
                ICollection<Element> columnElements = columnCollector
                    .OfCategory(BuiltInCategory.OST_StructuralColumns)
                    .OfClass(typeof(FamilyInstance))
                    .ToElements();

                // Combine the two sets of elements
                List<Element> allElements = new List<Element>();
                allElements.AddRange(framingElements);
                allElements.AddRange(columnElements);


                int colorIndex = 0; // Index to track color assignment
                foreach (FamilyInstance element in allElements)
                {
                    FamilySymbol familySymbol = element.Symbol;

                    // Check if the family type is already in the dictionary
                    if (!typeColorMap.ContainsKey(familySymbol.Id))
                    {
                        // Generate a color based on the color index
                        Color color = GetRandomColor();
                        typeColorMap.Add(familySymbol.Id, color);
                        colorIndex++;
                    }
                }

                // Apply colors to family types
                using (Transaction trans = new Transaction(doc, "Assign Colors to Family Types"))
                {
                    trans.Start();
                    FillPatternElement solidPatternId = new FilteredElementCollector(doc)
                   .WherePasses(new ElementClassFilter(typeof(FillPatternElement)))
                   .ToElements().Cast<FillPatternElement>().ToList().First(x => x.GetFillPattern().IsSolidFill);


                    foreach (KeyValuePair<ElementId, Color> kvp in typeColorMap)
                    {
                        FamilySymbol familySymbol = doc.GetElement(kvp.Key) as FamilySymbol;
                        if (familySymbol == null)
                            continue;

                        OverrideGraphicSettings overrideSettings = new OverrideGraphicSettings();
                        overrideSettings.SetSurfaceBackgroundPatternColor(kvp.Value);
                        overrideSettings.SetSurfaceBackgroundPatternId(solidPatternId.Id);

                        foreach (FamilyInstance element in allElements.OfType<FamilyInstance>().Where(x => x.Symbol.Id == kvp.Key))
                        {
                            doc.ActiveView.SetElementOverrides(element.Id, overrideSettings);
                        }
                    }

                    trans.Commit();
                }

            }

            catch (Exception ex2)

            {

                System.Diagnostics.StackTrace trace = new System.Diagnostics.StackTrace(ex2, true);
                TaskDialog.Show("Error", ex2.Message + " ,line of the error is " + trace.GetFrame(trace.FrameCount - 1).GetFileLineNumber());
            }




        }

        public static Color GetRandomColor()
        {

            byte red = (byte)random.Next(256);
            byte green = (byte)random.Next(256);
            byte blue = (byte)random.Next(256);
            return new Color(red, green, blue);
        }

        #endregion

        #endregion

        #region Key Schdule Manipulation


        public static void loopOnElementsInKeySchdule(UIDocument uidoc, Document doc, string scheduleName)
        {


            // Retrieve the key schedule by name
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            ViewSchedule keySchedule = collector.OfClass(typeof(ViewSchedule))
                                                .Cast<ViewSchedule>()
                                                .FirstOrDefault(vs => vs.Name == scheduleName);



            //// Collect all elements (rows) in the key schedule
            FilteredElementCollector scheduleCollector = new FilteredElementCollector(doc, keySchedule.Id);

            #region Assign new Element to schedule
            using (Transaction trans = new Transaction(doc, "Modify Key Schedule"))
            {

                trans.Start();
                TableData tableData = keySchedule.GetTableData();

                // Get the section data for the body
                TableSectionData sectionData = tableData.GetSectionData(SectionType.Body);
                // Insert a new row at the first row number


                for (int i = 0; i < 10; i++)
                {
                    sectionData.InsertRow(sectionData.LastRowNumber);
                    Element Newelement = scheduleCollector.ToElements().Last();
                    Newelement.LookupParameter("Key Name").Set(i.ToString());
                }
                trans.Commit();
            }

            #endregion


            #region GetAllElementsInSchdules
            //IList<Element> allElements = scheduleCollector.ToElements();

            //for (int i = 0; i < elements.Count; i++)
            //{
            //    Element element = elements[i];

            //    foreach (Parameter param in element.Parameters)
            //    {
            //        string paramName = param.Definition.Name;
            //        var value = param.AsValueString();
            //    }
            //}

            #endregion

            #region Get All fileds in a Schedule

            //IList<SchedulableField> schedulableFields = keySchedule.Definition.GetSchedulableFields();
            //for (int i = 0; i < schedulableFields.Count; i++)
            //{
            //    SchedulableField field = schedulableFields[i];
            //    string fieldName = field.GetName(doc);

            //}

            #endregion

        }


        #endregion

        #region  set Paramaters

        public static void setParametersSAP(FamilyInstance instance, Frame bm)
        {

            try
            {
                instance.LookupParameter("SapID").Set(bm.ID);
                instance.LookupParameter("Material").Set(bm.Material);
                instance.LookupParameter("RvtID").Set(instance.Id.ToString());
            }

            catch (Exception ex2)
            {
                System.Diagnostics.StackTrace trace = new System.Diagnostics.StackTrace(ex2, true);

                //TaskDialog.Show("Error", "Family " + instance.Symbol.FamilyName + " has no parameters (sapID or Matrial or RvtID)");


            }



        }

        public static void setParametersSTAAD(FamilyInstance instance, Frame bm)
        {

            try
            {
                instance.LookupParameter("StaadID").Set(bm.ID);
                instance.LookupParameter("Material").Set(bm.Material);
                instance.LookupParameter("RvtID").Set(instance.Id.ToString());
            }

            catch (Exception ex2)
            {
                System.Diagnostics.StackTrace trace = new System.Diagnostics.StackTrace(ex2, true);

                //TaskDialog.Show("Error", "Family " + instance.Symbol.FamilyName + " has no parameters (sapID or Matrial or RvtID)");


            }



        }

        public static void setParametersEtabs(FamilyInstance instance, Frame bm)
        {

            try
            {
                instance.LookupParameter("EtabsID").Set(bm.ID);
                instance.LookupParameter("Material").Set(bm.Material);
                instance.LookupParameter("RvtID").Set(instance.Id.ToString());
            }

            catch (Exception ex2)
            {
                System.Diagnostics.StackTrace trace = new System.Diagnostics.StackTrace(ex2, true);

                //TaskDialog.Show("Error", "Family " + instance.Symbol.FamilyName + " has no parameters (sapID or Matrial or RvtID)");


            }



        }

        #endregion

        #region  unUsed
        public static void convertFrameToMM(Frame frame, ForgeTypeId dis)
        {
            frame.StartPoint.X = UnitUtils.ConvertFromInternalUnits(frame.StartPoint.X, dis);
            frame.StartPoint.Y = UnitUtils.ConvertFromInternalUnits(frame.StartPoint.Y, dis);
            frame.StartPoint.Z = UnitUtils.ConvertFromInternalUnits(frame.StartPoint.Z, dis);

            frame.EndPoint.X = UnitUtils.ConvertFromInternalUnits(frame.EndPoint.X, dis);
            frame.EndPoint.Y = UnitUtils.ConvertFromInternalUnits(frame.EndPoint.Y, dis);
            frame.EndPoint.Z = UnitUtils.ConvertFromInternalUnits(frame.EndPoint.Z, dis);

            frame.MidPoint = Point.getMidPoint(frame.StartPoint, frame.EndPoint);
        }

        public static void addRevitID(Document doc, List<Frame> revitFrames)
        {

            try
            {
                using (Transaction trns = new Transaction(doc, "aa"))
                {
                    trns.Start();

                    foreach (var item in revitFrames)
                    {
                        ((Instance)doc.GetElement(item.revitEleId)).LookupParameter("RvtID").Set(item.revitEleId.ToString());
                    }

                    trns.Commit();

                }
            }


            catch (Exception ex2)
            {


            }

        }

        public static void getTransfromedRevitFrames(Matrix4x4 mat, ref List<Frame> revitFrames, Matrix4x4 trans)
        {

            foreach (Frame frame in revitFrames)
            {
                Vector3 startPt = Point.ToVector(frame.StartPoint);
                startPt = Vector3.Transform(startPt, mat);
                startPt = Vector3.Transform(startPt, trans);
                frame.StartPoint = new Point(startPt);

                Vector3 endPt = Point.ToVector(frame.EndPoint);
                endPt = Vector3.Transform(endPt, mat);
                endPt = Vector3.Transform(endPt, trans);
                frame.EndPoint = new Point(endPt);

                Vector3 midPt = Point.ToVector(frame.MidPoint);
                midPt = Vector3.Transform(midPt, mat);
                midPt = Vector3.Transform(midPt, trans);
                frame.MidPoint = new Point(midPt);

            }

        }

        public static Frame getRevitSelectedFrame(ExternalCommandData commandData, ForgeTypeId dis)
        {
            Frame frame = new Frame();
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            FamilyInstance ele = uidoc.Selection.GetElementIds().Select(e => doc.GetElement(e)).Cast<FamilyInstance>().ToList()[0];
            if (ele != null)
            {


                AnalyticalToPhysicalAssociationManager analyManger = AnalyticalToPhysicalAssociationManager.GetAnalyticalToPhysicalAssociationManager(doc);
                AnalyticalElement analytInstance = doc.GetElement(analyManger.GetAssociatedElementId(ele.Id)) as AnalyticalElement;


                frame.ID = ele.Id.ToString();
                frame.StartPoint = new Point(analytInstance.GetCurve().GetEndPoint(0));
                frame.EndPoint = new Point(analytInstance.GetCurve().GetEndPoint(1));
                frame.MidPoint = Point.getMidPoint(frame.StartPoint, frame.EndPoint);

                convertFrameToMM(frame, dis);
                frame.Length = Point.getDistance(frame.StartPoint, frame.EndPoint);

            }


            return frame;
        }


        public static Curve CreateCurveFromPoints(Frame bm, ForgeTypeId dis)
        {

            XYZ stPoint;
            XYZ endPoint;
            stPoint = new XYZ(UnitUtils.ConvertToInternalUnits(bm.StartPoint.X, dis), UnitUtils.ConvertToInternalUnits(bm.StartPoint.Y, dis), UnitUtils.ConvertToInternalUnits(bm.StartPoint.Z, dis));
            endPoint = new XYZ(UnitUtils.ConvertToInternalUnits(bm.EndPoint.X, dis), UnitUtils.ConvertToInternalUnits(bm.EndPoint.Y, dis), UnitUtils.ConvertToInternalUnits(bm.EndPoint.Z, dis));
            return Autodesk.Revit.DB.Line.CreateBound(stPoint, endPoint);


        }

        public static void setRevitDataInExcel(List<Frame> RevitFrames, Microsoft.Office.Interop.Excel.Worksheet ws)
        {


            int count = RevitFrames.Count;
            ws.Range["F5"].Value = Convert.ToString(count);
            Excel.Range StdID = ws.Range["A7"];
            Excel.Range RvtName = ws.Range["Y7"];
            Excel.Range RvtFamily = ws.Range["Z7"];
            Excel.Range RvtID = ws.Range["AA7"];
            Excel.Range RvtStartX = ws.Range["AB7"];
            Excel.Range RvtMemberType = ws.Range["AN7"];
            Excel.Range RvtBeta = ws.Range["AO7"];
            Excel.Range RvtLength = ws.Range["AP7"];



            for (int i = 0; i < count; i++)
            {
                StdID.Offset[i, 0].Value = Convert.ToString(RevitFrames[i].ID);
                RvtName.Offset[i, 0].Value = Convert.ToString(RevitFrames[i].RvtName);
                RvtFamily.Offset[i, 0].Value = Convert.ToString(RevitFrames[i].RvtFamily);
                RvtID.Offset[i, 0].Value = Convert.ToString(RevitFrames[i].RvtID);
                RvtStartX.Offset[i, 0].Value = Convert.ToString(RevitFrames[i].StartPoint.X);
                RvtStartX.Offset[i, 1].Value = Convert.ToString(RevitFrames[i].StartPoint.Y);
                RvtStartX.Offset[i, 2].Value = Convert.ToString(RevitFrames[i].StartPoint.Z);
                RvtStartX.Offset[i, 3].Value = Convert.ToString(RevitFrames[i].EndPoint.X);
                RvtStartX.Offset[i, 4].Value = Convert.ToString(RevitFrames[i].EndPoint.Y);
                RvtStartX.Offset[i, 5].Value = Convert.ToString(RevitFrames[i].EndPoint.Z);

                RvtMemberType.Offset[i, 0].Value = Convert.ToString(RevitFrames[i].MemberType.ToString());
                RvtBeta.Offset[i, 0].Value = Convert.ToString(RevitFrames[i].Beta.ToString());
                RvtLength.Offset[i, 0].Value = Convert.ToString(RevitFrames[i].Length.ToString());


            }


        }

        public static List<Frame> getRevitFramesFromExcel(ref List<Frame> Beams, ref List<Frame> Columns, Microsoft.Office.Interop.Excel.Worksheet ws)
        {
            List<Frame> frames = new List<Frame>();



            int count = Convert.ToInt32(ws.Range["F5"].Value);

            for (int i = 0; i < count; i++)
            {
                Frame f = new Frame();
                f.ID = Convert.ToString(ws.Range["A7"].Offset[i, 0].Value2);
                f.RvtName = (string)ws.Range["Y7"].Offset[i, 0].Value;
                f.RvtFamily = (string)ws.Range["Z7"].Offset[i, 0].Value;
                f.RvtID = Convert.ToString(ws.Range["AA7"].Offset[i, 0].Value);




                Point stPoint = new Point();
                stPoint.X = Convert.ToDouble(ws.Range["AB7"].Offset[i, 0].Value);
                stPoint.Y = Convert.ToDouble(ws.Range["AB7"].Offset[i, 1].Value);
                stPoint.Z = Convert.ToDouble(ws.Range["AB7"].Offset[i, 2].Value);

                Point endPoint = new Point();
                endPoint.X = Convert.ToDouble(ws.Range["AB7"].Offset[i, 3].Value);
                endPoint.Y = Convert.ToDouble(ws.Range["AB7"].Offset[i, 4].Value);
                endPoint.Z = Convert.ToDouble(ws.Range["AB7"].Offset[i, 5].Value);


                f.StartPoint = stPoint;
                f.EndPoint = endPoint;
                f.MidPoint = Point.getMidPoint(stPoint, endPoint);
                f.Length = Point.getDistance(f.StartPoint, f.EndPoint);
                f.Beta = Convert.ToDouble(ws.Range["AO7"].Offset[i, 0].Value);


                Point.getMemberType(f);

                if (f.MemberType == MemberType.Beam || f.MemberType == MemberType.VLBracing)
                {
                    f.FamilyType = getBeamSecType(f.RvtFamily);
                    Beams.Add(f);


                }

                else
                {
                    f.FamilyType = getColumnSecType(f.RvtFamily);
                    Columns.Add(f);
                }



                frames.Add(f);

            }



            return frames;
        }

        public static void assignStaadDataFromExcelToRevitParameters(Document doc, Microsoft.Office.Interop.Excel.Worksheet ws)
        {
            using (Transaction trns = new Transaction(doc, "aa"))

            {
                trns.Start();
                FailureHandlingOptions failOpt = trns.GetFailureHandlingOptions();
                failOpt.SetFailuresPreprocessor(new RoomWarningSwallower());
                trns.SetFailureHandlingOptions(failOpt);

                int count = Convert.ToInt32(ws.Range["F5"].Value2);

                for (int i = 0; i < count; i++)
                {

                    try
                    {
                        Instance RevitElement = (Instance)doc.GetElement(new ElementId(Convert.ToInt32(ws.Range["AA7"].Offset[i, 0].Value2)));
                        RevitElement.LookupParameter("StaadID").Set((string)Convert.ToString(ws.Range["A7"].Offset[i, 0].Value2));
                        RevitElement.LookupParameter("RvtID").Set((string)Convert.ToString(ws.Range["AA7"].Offset[i, 0].Value2));
                        RevitElement.LookupParameter("Material").Set((string)Convert.ToString(ws.Range["R7"].Offset[i, 0].Value2));
                    }
                    catch (Exception ex2)
                    {
                        TaskDialog.Show("error", ex2.Message + "  No of row" + i);
                        continue;
                    }


                }

                trns.Commit();




            }


        }
        #endregion

        #region  Extra

        #region  sheetsCreations

        public static FamilyInstance getTitleBlockInViewSHeet(Document doc, ViewSheet v)
        {


            var titleBlockTypeId = ElementId.InvalidElementId;
            var titleBlockFamInstList = new FilteredElementCollector(doc).WhereElementIsNotElementType().OfCategory(BuiltInCategory.OST_TitleBlocks).ToList();

            var titlBlock1 = titleBlockFamInstList.FirstOrDefault(x => x.OwnerViewId == v.Id);

            var titlBlock2 = titlBlock1 as FamilyInstance;



            if (titlBlock2 != null)
                return titlBlock2;

            return null;

        }

        public static Outline TitleBlockDrawingAreaBoundary(FamilySymbol FS)
        {


            if (FS.Category.Id.IntegerValue != Convert.ToInt32(BuiltInCategory.OST_TitleBlocks))
            {
                throw new ArgumentException("FS should be of a title block category.");
            }

            Document IntDoc = FS.Document;
            Document DocT = IntDoc.EditFamily(FS.Family);

            FilteredElementCollector FEC_Lines = new FilteredElementCollector(DocT);
            ElementCategoryFilter ECF_Lines = new ElementCategoryFilter(BuiltInCategory.OST_Lines);

            List<DetailLine> Lines = FEC_Lines.WherePasses(ECF_Lines).WhereElementIsNotElementType().OfType<DetailLine>().ToList();
            Outline OL = null;

            for (int ia = 0; ia <= Lines.Count - 1; ia++)
            {
                BoundingBoxXYZ BB = Lines[ia].get_BoundingBox(null);
                if (OL == null)
                {
                    OL = new Outline(BB.Min, BB.Max);
                }
                else
                {
                    OL.AddPoint(BB.Min);
                    OL.AddPoint(BB.Max);
                }
            }
            XYZ RoughCentre = (OL.MinimumPoint + OL.MaximumPoint) / 2;

            Line XLn = Line.CreateUnbound(RoughCentre, XYZ.BasisX);
            Line YLn = Line.CreateUnbound(RoughCentre, XYZ.BasisY);

            double YUp = double.PositiveInfinity;
            double YDown = double.NegativeInfinity;
            double XLeft = double.NegativeInfinity;
            double XRight = double.PositiveInfinity;

            int[] BoundsSet = new int[4] { 0, 0, 0, 0 };
            for (int ia = 0; ia <= Lines.Count - 1; ia++)
            {
                DetailLine LN = Lines[ia];
                Line C = LN.GeometryCurve as Line;

                IntersectionResultArray IRA = null;

                if (Math.Abs(C.Direction.DotProduct(YLn.Direction)) > 0.99)
                {
                    SetComparisonResult SCR = XLn.Intersect(C, out IRA);
                    if (SCR != SetComparisonResult.Overlap)
                        continue;

                    for (int ic = 0; ic <= IRA.Size - 1; ic++)
                    {
                        IntersectionResult IR = IRA.get_Item(ic);
                        XYZ Pt = IR.XYZPoint;

                        if (Pt.X > XLeft && Pt.X < RoughCentre.X)
                        {
                            XLeft = Pt.X;
                            BoundsSet[0] = 1;
                        }
                        else if (Pt.X < XRight && Pt.X > RoughCentre.X)
                        {
                            XRight = Pt.X;
                            BoundsSet[1] = 1;
                        }
                    }

                }
                else if (Math.Abs(C.Direction.DotProduct(YLn.Direction)) < 0.01)
                {
                    SetComparisonResult SCR = YLn.Intersect(C, out IRA);
                    if (SCR != SetComparisonResult.Overlap)
                        continue;

                    for (int ic = 0; ic <= IRA.Size - 1; ic++)
                    {
                        IntersectionResult IR = IRA.get_Item(ic);
                        XYZ Pt = IR.XYZPoint;

                        if (Pt.Y > YDown && Pt.Y < RoughCentre.Y)
                        {
                            YDown = Pt.Y;
                            BoundsSet[2] = 1;
                        }
                        else if (Pt.Y < YUp && Pt.Y > RoughCentre.Y)
                        {
                            YUp = Pt.Y;
                            BoundsSet[3] = 1;
                        }
                    }
                }
            }
            DocT.Close(false);

            if (!BoundsSet.Contains(0))
            {
                XYZ Min = new XYZ(XLeft, YDown, 0);
                XYZ Max = new XYZ(XRight, YUp, 0);

                return new Outline(Min, Max);
            }
            else
            {
                return null;
            }

        }

        #endregion




        #endregion

        #region Hydra Simulator

        public static void SetRevitFramesInExcelForHydra(ref List<Frame> revitElements, string AnalysisModelType, ref Microsoft.Office.Interop.Excel.Application app, ref Excel._Workbook wb, ref Excel._Worksheet ws)
        {
            try
            {
                List<Frame> frames = revitElements;


                object[,] arr = new object[frames.Count, 9];


                for (int i = 0; i < frames.Count; i++)
                {
                    arr[i, 0] = frames[i].RvtName;
                    arr[i, 1] = frames[i].RvtFamily;
                    arr[i, 2] = frames[i].RvtID;
                    arr[i, 3] = frames[i].StartPoint.X;
                    arr[i, 4] = frames[i].StartPoint.Y;
                    arr[i, 5] = frames[i].StartPoint.Z;
                    arr[i, 6] = frames[i].EndPoint.X;
                    arr[i, 7] = frames[i].EndPoint.Y;
                    arr[i, 8] = frames[i].EndPoint.Z;


                    frames[i].rowNo = i + 7;
                }


                ExcelUtils.CreateExcelForSAP(ref app, ref wb, ref ws, "Sap");


                Microsoft.Office.Interop.Excel.Range startCell = (Range)ws.Cells[7, 25];
                Microsoft.Office.Interop.Excel.Range endCell = (Range)ws.Cells[arr.GetLength(0) + 7 - 1, 33];
                Microsoft.Office.Interop.Excel.Range range = (Range)ws.Range[startCell, endCell];
                range.Value2 = arr;

                ws.Range["F5"].Value = frames.Count;
                wb.Save();

            }
            catch (Exception)
            {

                throw;
            }



        }

        #endregion

        #region Reflections methods 
        public static bool accessToSet_privateProperty(object instance, string proprtyName, object newValue)
        {
            bool result = false;

            Type tp = instance.GetType();


            FieldInfo filedInfo = tp.GetField(proprtyName);

            PropertyInfo[] arr = tp.GetProperties(BindingFlags.Public);



            if (filedInfo != null)
            {
                filedInfo.SetValue(instance, newValue);
                result = true;
            }


            return result;
        }


        #endregion
    }
}
