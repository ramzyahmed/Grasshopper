using Autodesk.Revit.DB;
using BIM_Integration;
using ETABSv1;
using Grasshopper.Kernel;
using Microsoft.Office.Interop.Excel;
using Rhino;
using Rhino.DocObjects;
using Rhino.Geometry;
using Rhino.Geometry.Collections;
using RhinoInside.Revit.Convert.Geometry;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions; // Ensure this namespace is included
using static BIM_Integration.EtabsUtils;
using Application = Microsoft.Office.Interop.Excel.Application;
using Area = BIM_Integration.Area;
using Curve = Rhino.Geometry.Curve;
using Frame = BIM_Integration.Frame;
using Plane = Rhino.Geometry.Plane;
using Point = BIM_Integration.Point;


namespace MyGrasshopperAssembly2
{
    public class RhinoEtabs
    {

        #region EtabsCheckFramesDrift
        public static void EtabsCheckFramesDrift(IGH_DataAccess DA)
        {
            bool run = false;

            double windFactor = 0;
            double seismicFactor = 0;
            List<String> ColumnsGroupsName = new List<string>();
            List<String> LoadCaseKeyNames = new List<string>();


            DA.GetData(0, ref run);
            DA.GetData(1, ref windFactor);
            DA.GetData(2, ref seismicFactor);
            DA.GetDataList(3, ColumnsGroupsName);
            DA.GetDataList(4, LoadCaseKeyNames);



            if (run)
            {

                List<DriftRecrd> recList = EtabsUtils.framesEndPointsDrifts(ColumnsGroupsName, windFactor, seismicFactor, LoadCaseKeyNames);

                List<DriftRecrd> targetList = recList
                    .GroupBy(x => x.pt.Id)
                    .Select(group => group.OrderBy(y => y.f).First())
                    .ToList();

                Rhino.RhinoDoc doc = Rhino.RhinoDoc.ActiveDoc;

                List<RhinoObject> rhinoPoints = doc.Objects.FindByObjectType(ObjectType.Point).ToList();



                foreach (var pt in rhinoPoints)
                {

                    ObjectAttributes attributes = pt.Attributes;
                    attributes.ObjectColor = System.Drawing.Color.Green;
                    attributes.ColorSource = Rhino.DocObjects.ObjectColorSource.ColorFromObject;
                    NameValueCollection userStrings = attributes.GetUserStrings();


                    #region colorize the poins
                    // sap Points
                    if (userStrings.AllKeys.FirstOrDefault(x => x == "ID") != null)
                    {
                        System.Drawing.Color clr = System.Drawing.Color.Green;

                        string id = (string)userStrings["ID"];


                        // unSafePoints
                        if (targetList.Select(x => x.pt.Id).Contains(id))
                        {
                            clr = System.Drawing.Color.Red;
                            attributes.SetUserString("Unsafe", "Yes");
                            DriftRecrd rec = targetList.First(x => x.pt.Id == id);
                            attributes.SetUserString("Info", "LoadComb: " + rec.loadCases + "/ Dir: " + rec.direction + "/ Fact: " + Math.Round(rec.factor));


                        }


                        // SafePoints
                        else
                        {
                            attributes.DeleteUserString("Unsafe");
                            attributes.DeleteUserString("Info");

                        }

                        //Point3d point3d = ((Point)pt.Geometry).Location;
                        //targetPoints.Add(point3d);

                        attributes.ObjectColor = clr;
                        pt.Attributes = attributes;

                        doc.Objects.ModifyAttributes(pt.Id, attributes, true);

                        RhinoDoc.ActiveDoc.Views.Redraw();

                    }
                    #endregion


                }






            }



        }

        #endregion

        #region EtabsGroupingCheck

        public static void EtabsGroupingCheck(IGH_DataAccess DA)
        {
            bool aa = false;
            DA.GetData(0, ref aa);

            if (aa)
            {
                #region createGroupingLists 
                // Input data
                string input = @"00.Building.InnerColumn
00.Building.OuterColumn
00.Building.CorridorBeam
00.Building.VertBracing
00.Building.HorBracing
00.Building.Ties
00.Building.PrimaryBeam
00.Building.EdgeBeam
00.Building.Cantilever
00.Building.FacadeBeam
00.Building.Facade.Ghost
00.Building.OverhangBeam
00.Megapod.Column
00.Megapod.Beam
00.Megapod.HorBracing
00.Megapod.VertBracing
00.Megapod.Ties
00.Skislope.Primary
00.Skislope.Cantilever
00.Skislope.Secondary
00.Skislope.Column
00.Skislope.Rigid
00.Skislope.EdgeBeam
00.Skislope.Bracing
00.F&B.Column
00.F&B.VertBracing
00.F&B.HorBracing
00.F&B.Ties
00.F&B.Beam
00.F&B.EdgeBeam
00.F&B.Cantilever
00.F&B.FacadeBeam
00.F&B.Facade.Ghost
00.F&B.Roof
00.MST.Truss.Diagonal
00.MST.Truss.Post
00.MST.Truss.Chord
00.MST.PrimaryBeam
00.MST.CorridorBeam
00.MST.Support
00.MST.Column
00.MST.HorBracing
00.MST.Ties
00.MST.VertBracing
00.MST.Cantilever
00.MST.EdgeBeam
00.MST.FacadeBeam
00.MST.Facade.Ghost
00.MST.OverhangBeam
00.Skileg.Primary
00.Skileg.Cantilever
00.Skileg.Secondary
00.Skileg.Column
00.Skileg.Ties
00.Skileg.Bracing
00.Skileg.EdgeBeam
00.Basement.Beam
00.Basement.Column
00.Plaza.Beam
00.Plaza.Column
00.Skileg.BottomChord
00.Skileg.Diagonal
00.Skileg.TopChord
00.Skileg.Diagonal
00.Skileg.Pylon";

                string inputArea = @"01.Hospitality.LP
01.Building.Terrace.LP
01.Building.NonAccTerrace.LP
01.Building.MEP.LP
01.Building.Hospitality.LP
01.Building.FOH.LP
01.Building.FireEscape.LP
01.Building.Facade.LP
01.Building.Circulation.LP
01.Building.Balcony.LP
01.Building.BOH.LP
01.Skislope.Lat3Dx.LP
01.Skislope.Lat3Sx.LP
01.Skislope.Lat2Dx.LP
01.Skislope.Lat2Sx.LP
01.Skislope.Lat1Dx.LP
01.Skislope.Lat1Sx.LP
01.Skislope.Side2.LP
01.Skislope.Side1.LP
01.Skislope.Center.LP
01.F&B.Roof.LP
01.Building.Floor.LP
01.Building.Pool.LP
01.Building.GeoFacade.LP
01.Building.Shaft and Void.LP
01.Building.Terrace W\O Land.LP
01.Building.NonAccTerrace W\O Land.LP
01.Building.Traffic.LP
01.Building.HeavyParking.LP
01.Building.PublicTerrace.LP
01.Skislope.Under.F&B_LP
01.Skislope.NotSkiable.Lat1Sx.LP
01.Skislope.NotSkiable.Lat1Dx.LP
01.Skislope.NotSkiable.Center.LP
01.Hospitality.LP
01.MEP.LP
01.Building.NonAccTerraceLand.LP
01.Skislope.LP
01.F&B.LP
01.Other
01.BOH.Stair
01.Private.Balcony.WithTiles
01.Puplic.Balcony.WithTiles
01.Public.Terrace.WithTiles
01.Building.NonAccTerrace.WithLandscape
01.Heavy.Parking.LP
01.Basement.PedestrianArea
01.Non Accessible.Terrace.WithLandscape
01.NonAccessible.Balcony.WithGravel
01.NonAccessible.Terrace.WithLandscape
01.Non Accessible.Balcony.WithGravel
01.Public.Balcony.WithTiles";



                #region create subListsForFrameGroups

                // Split the input into lines by newline characters
                string[] framesGroupsFromExcel = input.Split('\n');

                string[] framesGroupsFromExcel2 = framesGroupsFromExcel.Select(s => s.Length >= 2 ? s.Substring(0, s.Length - 1) : "").ToArray();


                // Create a list of lists to store the categorized strings
                List<List<string>> categorizedLists = new List<List<string>>();

                // Initialize the current category
                string currentCategory = "";

                foreach (string line in framesGroupsFromExcel2)
                {
                    // Extract the category from the line (e.g., "00.Building")
                    string category = GetCategory(line);

                    // Check if the category has changed
                    if (category != currentCategory)
                    {
                        // Create a new list for the current category
                        List<string> categoryList = new List<string>();
                        categorizedLists.Add(categoryList);
                        currentCategory = category;
                    }

                    // Add the line to the current category's list
                    categorizedLists[categorizedLists.Count - 1].Add(line);
                }

                #endregion


                // criteria of Areas not like frames
                #region createJustOneListForAreaGroups  

                // Split the input into lines by newline characters
                List<string> areaGroupsFromExcel = inputArea.Split('\n').ToList();
                areaGroupsFromExcel = areaGroupsFromExcel.Select(s => s.Length >= 2 ? s.Substring(0, s.Length - 1) : "").ToList();
                #endregion

                #endregion

                #region CheckFramesValidity

                List<Column> myColumns = new List<Column>();
                List<Beam> myBeams = new List<Beam>();

                List<Frame> sapFrames = null;
                sapFrames = EtabsUtils.getSapFrames(ref myBeams, ref myColumns);


                List<Frame> correctElemtns = new List<Frame>();
                List<Frame> notCorrectElemtns = new List<Frame>();



                foreach (var frame in sapFrames)
                {

                    int a = categorizedLists.Count(GroupList => GroupList.Intersect(frame.Groups).ToList().Count >= 2);


                    if (a > 0)
                    {
                        notCorrectElemtns.Add(frame);

                    }

                    else
                    {
                        correctElemtns.Add(frame);
                    }
                }

                #endregion

                #region CheckAreaGroupAssignmentsValidity
                List<string> correctAreasIDS = new List<string>();
                List<string> notCorrectAreasIDS = new List<string>();


                List<Area> sapAreas = EtabsUtils.getEtabsAreas();
                foreach (var sapArea in sapAreas)
                {
                    if (sapArea.ID == "261")
                    {
                        int aaa = 12;
                    }

                    if (sapArea.Groups.Intersect(areaGroupsFromExcel).Count() > 1)
                    {
                        notCorrectAreasIDS.Add(sapArea.ID);
                    }
                    else
                    {
                        correctAreasIDS.Add(sapArea.ID);

                    }
                }

                EtabsUtils.createGroupOfAreasinSap(correctAreasIDS, "A00AreaElementsWithCorrectGroups");
                EtabsUtils.createGroupOfAreasinSap(notCorrectAreasIDS, "A00AreaElementsWithNotCorrectGroups");


                #endregion
            }
        }

        static string GetCategory(string input)
        {
            int firstDotIndex = input.IndexOf('.');
            if (firstDotIndex != -1)
            {
                int secondDotIndex = input.IndexOf('.', firstDotIndex + 1);
                if (secondDotIndex != -1)
                {
                    return input.Substring(0, secondDotIndex);
                }
            }
            return input;
        }

        #endregion

        #region EtabsToRhino

        // Main 
        public static void EtabsToRhino(IGH_DataAccess DA)
        {
            bool aa = false;
            DA.GetData(0, ref aa);

            if (aa)
            {
                Rhino.RhinoDoc doc = Rhino.RhinoDoc.ActiveDoc;
                {

                    #region draw frames

                    List<Frame> sapFrames = null;
                    sapFrames = EtabsUtils.getSapFrames();



                    List<string> SapGroups = EtabsUtils.getGroupsList();


                    Dictionary<string, string> groupNameAttributeNameDIc = new Dictionary<string, string>();

                    for (int a = 0; a < SapGroups.Count; a++)
                    {
                        groupNameAttributeNameDIc.Add(SapGroups[a], "Group" + a);
                    }


                    List<Curve> curves = new List<Curve>();

                    foreach (var item in sapFrames)
                    {
                        Point3d stPoint = new Point3d(item.StartPoint.X, item.StartPoint.Y, item.StartPoint.Z);
                        Point3d endPoint = new Point3d(item.EndPoint.X, item.EndPoint.Y, item.EndPoint.Z);
                        LineCurve crv = new LineCurve(stPoint, endPoint);
                        Rhino.DocObjects.Tables.ObjectTable ot = Rhino.RhinoDoc.ActiveDoc.Objects;
                        ObjectAttributes attr = new ObjectAttributes();


                        attr.SetUserString("EtabsSection", item.Sec);
                        attr.SetUserString("Story", item.Story);
                        attr.SetUserString("Shape", item.Shape);
                        attr.SetUserString("StartPointId", item.StartPoint.Id);
                        attr.SetUserString("EndPointId", item.EndPoint.Id);
                        attr.SetUserString("Angle", item.Beta.ToString());
                        attr.SetUserString("Length", item.Length.ToString());
                        attr.SetUserString("ID", item.ID);
                       // attr.SetUserString("MainGroup", item.MainGroup);
                       // attr.SetUserString("SubGroup", item.SubGroup);

                        List<string> assignedGroups = new List<string>();
                        foreach (var group in SapGroups)
                        {
                            if (item.Groups.Contains(group))
                            {
                                attr.SetUserString(groupNameAttributeNameDIc[group], group);

                                assignedGroups.Add(groupNameAttributeNameDIc[group]);
                            }


                        }

                        attr.SetUserString("Groups", string.Join(",", assignedGroups));


                        if (item.MemberType == MemberType.Column)
                        {
                            attr.SetUserString("MemberType", "Column");
                        }

                        else
                        {
                            attr.SetUserString("MemberType", "Beam");

                        }

                        ot.AddCurve(crv, attr);
                        curves.Add(crv);
                    }

                    DA.SetDataList(0, curves);

                    #endregion

                    #region drawAreas
                    List<Area> sapAreas = EtabsUtils.getEtabsAreas();
                    List<Brep> rhinoAreas = new List<Brep>();
                    foreach (var sapArea in sapAreas)
                    {
                        ObjectAttributes attr = new ObjectAttributes();
                        Rhino.DocObjects.Tables.ObjectTable ot = doc.Objects;

                        if (sapArea.vertices.Count == 3)
                        {
                            Point3d pt1 = new Point3d(sapArea.vertices[0].X, sapArea.vertices[0].Y, sapArea.vertices[0].Z);
                            Point3d pt2 = new Point3d(sapArea.vertices[1].X, sapArea.vertices[1].Y, sapArea.vertices[1].Z);
                            Point3d pt3 = new Point3d(sapArea.vertices[2].X, sapArea.vertices[2].Y, sapArea.vertices[2].Z);


                            Brep srf = Brep.CreateFromCornerPoints(pt1, pt2, pt3, 0.0001);
                            attr.SetUserString("EtabsSection", sapArea.Sec);
                            attr.SetUserString("ID", sapArea.ID);
                            attr.SetUserString("Angle", sapArea.Beta.ToString());
                            attr.SetUserString("Pts", string.Join(",", sapArea.vertices.Select(x => x.Id)));
                            attr.SetUserString("AreaType", sapArea.type);
                            attr.SetUserString("MainGroup", sapArea.MainGroup);
                            attr.SetUserString("SubGroup", sapArea.SubGroup);

                            ot.AddBrep(srf, attr);
                            rhinoAreas.Add(srf);
                        }
                        if (sapArea.vertices.Count == 4)
                        {
                            Point3d pt1 = new Point3d(sapArea.vertices[0].X, sapArea.vertices[0].Y, sapArea.vertices[0].Z);
                            Point3d pt2 = new Point3d(sapArea.vertices[1].X, sapArea.vertices[1].Y, sapArea.vertices[1].Z);
                            Point3d pt3 = new Point3d(sapArea.vertices[2].X, sapArea.vertices[2].Y, sapArea.vertices[2].Z);
                            Point3d pt4 = new Point3d(sapArea.vertices[3].X, sapArea.vertices[3].Y, sapArea.vertices[3].Z);


                            Brep srf = Brep.CreateFromCornerPoints(pt1, pt2, pt3, pt4, 0.0001);
                            attr.SetUserString("EtabsSection", sapArea.Sec);
                            attr.SetUserString("ID", sapArea.ID);
                            attr.SetUserString("Angle", sapArea.Beta.ToString());
                            attr.SetUserString("Pts", string.Join(",", sapArea.vertices.Select(x => x.Id)));
                            attr.SetUserString("AreaType", sapArea.type);
                            attr.SetUserString("MainGroup", sapArea.MainGroup);
                            attr.SetUserString("SubGroup", sapArea.SubGroup);


                            ot.AddBrep(srf, attr);
                            rhinoAreas.Add(srf);

                        }
                        if (sapArea.vertices.Count > 4)
                        {

                            List<Point3d> rhinoPts = new List<Point3d>();
                            foreach (var pt in sapArea.vertices)
                            {
                                Point3d rhinoPt = new Point3d(pt.X, pt.Y, pt.Z);
                                rhinoPts.Add(rhinoPt);
                            }
                            rhinoPts.Add(rhinoPts[0]);

                            Rhino.Geometry.Polyline plyLine = new Rhino.Geometry.Polyline(rhinoPts);
                            // ot.AddPolyline(plyLine);
                            Curve crv = plyLine.ToNurbsCurve();

                            try
                            {
                                Brep srf = Brep.CreatePlanarBreps(crv, doc.ModelAbsoluteTolerance)[0];


                                attr.SetUserString("EtabsSection", sapArea.Sec);
                                attr.SetUserString("ID", sapArea.ID);
                                attr.SetUserString("Angle", sapArea.Beta.ToString());
                                attr.SetUserString("Pts", string.Join(",", sapArea.vertices.Select(x => x.Id)));
                                attr.SetUserString("AreaType", sapArea.type);
                                attr.SetUserString("MainGroup", sapArea.MainGroup);
                                attr.SetUserString("SubGroup", sapArea.SubGroup);

                                ot.AddBrep(srf, attr);
                                rhinoAreas.Add(srf);

                            }
                            catch (Exception)
                            {
                                continue;

                                throw;
                            }




                        }


                    }

                    DA.SetDataList(1, rhinoAreas);

                    #endregion

                    #region draw Points

                    List<Point> pts = new List<Point>();

                    pts.AddRange(sapFrames.Select(x => x.StartPoint).Union(sapFrames.Select(x => x.EndPoint)));

                    pts.AddRange(sapAreas.SelectMany(x => x.vertices));

                    pts = pts
                         .GroupBy(p => p.Id)
                         .Select(g => g.First())
                         .ToList();

                    List<Point3d> rhinoPTS = new List<Point3d>();
                    foreach (var pt in pts)
                    {
                        Point3d rhinoPt = new Point3d(pt.X, pt.Y, pt.Z);
                        ObjectAttributes att = new ObjectAttributes();
                        att.SetUserString("ID", pt.Id);

                        doc.Objects.AddPoint(rhinoPt, att);
                        rhinoPTS.Add(rhinoPt);
                    }


                    //DA.SetData(2, rhinoPTS);


                    #endregion

                    #region Layers

                    CreateFramesLayers(doc);

                    CreateAreaLayers(doc);

                    CreatePointsLayer(doc);

                    #endregion

                }

            }



        }

        // ForColumnMarks
        public static void EtabsToRhino2(IGH_DataAccess DA)
        {
            bool aa = false;
            DA.GetData(0, ref aa);

            if (aa)
            {
                Rhino.RhinoDoc doc = Rhino.RhinoDoc.ActiveDoc;
                {

                    #region draw frames

                    List<Frame> sapFrames = null;
                    sapFrames = EtabsUtils.getSapFrames();
                    EtabsUtils.AssignColumnsAndBeamsDimentions(ref sapFrames);
                    EtabsUtils.AssignColumnsMarks7(ref sapFrames);
                    EtabsUtils.addColumnRigidZone(sapFrames);

                    List<Curve> curves = new List<Curve>();

                    sapFrames = sapFrames.Where(x => x.Mark != null).OrderBy(x => ExtractNumber(x.Mark)).ToList();
                    foreach (var item in sapFrames)
                    {
                        Point3d stPoint = new Point3d(item.StartPoint.X, item.StartPoint.Y, item.StartPoint.Z);
                        Point3d endPoint = new Point3d(item.EndPoint.X, item.EndPoint.Y, item.EndPoint.Z);
                        LineCurve crv = new LineCurve(stPoint, endPoint);
                        Rhino.DocObjects.Tables.ObjectTable ot = Rhino.RhinoDoc.ActiveDoc.Objects;
                        ObjectAttributes attr = new ObjectAttributes();

                        attr.SetUserString("Section", item.Sec);
                        attr.SetUserString("Story", item.Story);
                        attr.SetUserString("TopStory", item.EtabsTopStory);
                        attr.SetUserString("BotStory", item.EtabsBotStory);
                        attr.SetUserString("Shape", item.Shape);
                        attr.SetUserString("StartPointId", item.StartPoint.Id);
                        attr.SetUserString("EndPointId", item.EndPoint.Id);
                        attr.SetUserString("Angle", item.Beta.ToString());
                        attr.SetUserString("Length", item.Length.ToString());
                        attr.SetUserString("ID", item.ID);
                        attr.SetUserString("MainGroup", item.MainGroup);
                        attr.SetUserString("SubGroup", item.SubGroup);
                        attr.SetUserString("ColumnMark", item.Mark);
                        attr.SetUserString("Deignation", ColumnRebarDesignData.getDesignationMark(item));
                        attr.SetUserString("ColumnType", item.ColumnsRebarDesign.Pattern);

                        attr.SetUserString("WidthBars", item.ColumnsRebarDesign.NumberOfWidthBars.ToString());
                        attr.SetUserString("DepthBars", item.ColumnsRebarDesign.NumberOfDepthBars.ToString());

                        //attr.SetUserString("DepthBars", item.ColumnsRebarDesign.NumberOfDepthBars.ToString());
                        //attr.SetUserString("WidthBars", item.ColumnsRebarDesign.NumberOfWidthBars.ToString());
                        //attr.SetUserString("BarDiameter", item.ColumnsRebarDesign.RebarSize);





                        if (item.MemberType == MemberType.Column)
                        {
                            attr.SetUserString("MemberType", "Column");
                        }

                        else
                        {
                            attr.SetUserString("MemberType", "Beam");

                        }

                        ot.AddCurve(crv, attr);
                        curves.Add(crv);
                    }

                    DA.SetDataList(0, curves);

                    #endregion
                    #region Layers

                    CreateFramesLayersColumnMark(doc);

                    // CreateAreaLayers(doc);

                    // CreatePointsLayer(doc);

                    #endregion

                }

            }



        }


        #region Exit50 
        // to be removed later

        public static void EtabsToRhinoFramesOnly(IGH_DataAccess DA)
        {
            bool aa = false;
            DA.GetData(0, ref aa);

            if (aa)
            {


                Rhino.RhinoDoc doc = Rhino.RhinoDoc.ActiveDoc;
                {

                    #region draw frames



                    List<List<Frame>> sapFramesGroups = new List<List<Frame>>();
                    sapFramesGroups = EtabsUtils.assignEtabsFramesGroups();

                    foreach (var groupFrames in sapFramesGroups)
                    {
                        foreach (var item in groupFrames)
                        {
                            Point3d stPoint = new Point3d(item.StartPoint.X, item.StartPoint.Y, item.StartPoint.Z);
                            Point3d endPoint = new Point3d(item.EndPoint.X, item.EndPoint.Y, item.EndPoint.Z);
                            LineCurve crv = new LineCurve(stPoint, endPoint);
                            Rhino.DocObjects.Tables.ObjectTable ot = Rhino.RhinoDoc.ActiveDoc.Objects;
                            ObjectAttributes attr = new ObjectAttributes();


                            //attr.SetUserString("Group", item.Group);

                            attr.SetUserString("Group", String.Join(",", item.Groups2));

                            if (item.Groups2.Count > 1)
                            {
                                attr.SetUserString("IsAssignedTo2Group", "True");

                            }
                            attr.SetUserString("Section", item.Sec);
                            attr.SetUserString("Shape", item.Shape);
                            attr.SetUserString("StartPointId", item.StartPoint.Id);
                            attr.SetUserString("EndPointId", item.EndPoint.Id);
                            attr.SetUserString("Angle", item.Beta.ToString());
                            attr.SetUserString("Length", item.Length.ToString());
                            attr.SetUserString("ID", item.ID);



                            #region getLayerKey

                            //string s = "Neglect";
                            //s = item.Groups.FirstOrDefault(x => x.ToLower().Contains("section"));
                            //attr.SetUserString("MainGroup", s);

                            string s = "Neglect";
                            List<string> targetGroups = new List<string>();
                            targetGroups = item.Groups.Where(x => x.ToLower().Contains("truss") && !x.ToLower().Contains("sec") && !x.ToLower().Contains("main")).ToList();

                            if (targetGroups.Count > 1)
                            {
                                string s2 = targetGroups.FirstOrDefault(x => x.Count(c => c == '-') == 2);
                                s = s2;
                            }

                            else if (targetGroups.Count == 1)
                            {
                                s = targetGroups[0];
                            }

                            else
                            {
                                s = "Neglect";
                            }

                            attr.SetUserString("MainGroup", s);


                            #endregion




                            ot.AddCurve(crv, attr);



                        }
                    }




                    #endregion


                    #region Create frames Layers
                    RhinoObject[] rhinoCurves = doc.Objects.FindByObjectType(ObjectType.Curve).Where(x => x.Attributes.GetUserString("Group") != null).ToArray();


                    #region get  list of Sections
                    List<string> groups = new List<string>();
                    foreach (var item in rhinoCurves)
                    {

                        ObjectAttributes att = item.Attributes;

                        string group = att.GetUserString("MainGroup");

                        groups.Add(group);

                    }
                    groups = groups.Distinct().ToList();
                    #endregion

                    int FrameLayerIndex;
                    Layer FrameLayer;
                    FrameLayer = doc.Layers.FindName("EtabsFrames");

                    if (FrameLayer == null)
                    {
                        FrameLayer = new Layer();
                        FrameLayer.Name = "Frames";
                        FrameLayerIndex = doc.Layers.Add(FrameLayer);
                    }

                    else
                    {
                        FrameLayerIndex = FrameLayer.Index;
                    }



                    FrameLayer = doc.Layers.FindName("EtabsFrames");

                    int i = 0;

                    foreach (var group in groups.Where(x => x != null))
                    {
                        int red = i * 20 % 256;
                        int green = i * 40 % 256;
                        int blue = i * 60 % 256;

                        System.Drawing.Color color = System.Drawing.Color.FromArgb(255, red, green, blue);

                        int subLayerIndex;
                        Layer subLayer = doc.Layers.FindName(group);


                        if (subLayer == null)
                        {
                            subLayer = new Layer();
                            subLayer.Name = group;
                            subLayer.Color = color;
                            subLayer.ParentLayerId = FrameLayer.Id;
                            subLayerIndex = doc.Layers.Add(subLayer);
                        }

                        else
                        {
                            subLayer.Name = group;
                            subLayer.Color = color;
                            subLayer.ParentLayerId = FrameLayer.Id;
                            subLayerIndex = subLayer.Index;
                        }




                        RhinoObject[] objects = doc.Objects.FindByUserString("MainGroup", group, true);
                        foreach (var item in objects)
                        {
                            item.Attributes.LayerIndex = subLayerIndex;
                            item.CommitChanges();
                        }

                        i++;
                    }

                    #endregion



                }
            }
        }
        #endregion


        public static void EtabsToRhinoFramesOnlyForRevit(IGH_DataAccess DA, List<Frame> EtabsFrames)
        {
            bool aa = false;
            DA.GetData(0, ref aa);

            if (aa)
            {

                Rhino.RhinoDoc doc = Rhino.RhinoDoc.ActiveDoc;
                {

                    #region draw frames

                    List<List<Frame>> sapFramesGroups = new List<List<Frame>>();
                    sapFramesGroups = EtabsUtils.assignEtabsFramesGroups();


                    foreach (var item in EtabsFrames)
                    {
                        Point3d stPoint = new Point3d(item.StartPoint.X, item.StartPoint.Y, item.StartPoint.Z);
                        Point3d endPoint = new Point3d(item.EndPoint.X, item.EndPoint.Y, item.EndPoint.Z);
                        LineCurve crv = new LineCurve(stPoint, endPoint);
                        Rhino.DocObjects.Tables.ObjectTable ot = Rhino.RhinoDoc.ActiveDoc.Objects;
                        ObjectAttributes attr = new ObjectAttributes();


                        attr.SetUserString("Sec", item.Sec);
                        attr.SetUserString("Shape", item.Shape);
                        attr.SetUserString("StartPointId", item.StartPoint.Id);
                        attr.SetUserString("EndPointId", item.EndPoint.Id);
                        attr.SetUserString("Angle", item.Beta.ToString());
                        attr.SetUserString("Length", item.Length.ToString());
                        attr.SetUserString("ID", item.ID);


                        ot.AddCurve(crv, attr);



                    }


                    #endregion
                    #region Create frames Layers
                    RhinoObject[] rhinoCurves = doc.Objects.FindByObjectType(ObjectType.Curve).Where(x => x.Attributes.GetUserString("Sec") != null).ToArray();


                    #region get  list of Sections
                    List<string> Sections = new List<string>();
                    foreach (var item in rhinoCurves)
                    {

                        ObjectAttributes att = item.Attributes;

                        string group = att.GetUserString("Sec");

                        Sections.Add(group);

                    }
                    Sections = Sections.Distinct().ToList();
                    #endregion

                    int FrameLayerIndex;
                    Layer FrameLayer;
                    FrameLayer = doc.Layers.FindName("EtabsFrames");

                    if (FrameLayer == null)
                    {
                        FrameLayer = new Layer();
                        FrameLayer.Name = "EtabsFrames";
                        FrameLayerIndex = doc.Layers.Add(FrameLayer);
                    }

                    else
                    {
                        FrameLayerIndex = FrameLayer.Index;
                    }



                    FrameLayer = doc.Layers.FindName("EtabsFrames");

                    int i = 0;

                    foreach (var sec in Sections.Where(x => x != null))
                    {
                        int red = i * 20 % 256;
                        int green = i * 40 % 256;
                        int blue = i * 60 % 256;

                        System.Drawing.Color color = System.Drawing.Color.FromArgb(255, red, green, blue);

                        int subLayerIndex;
                        Layer subLayer = doc.Layers.FindName(sec);


                        // if (subLayer == null || parentLayer != null)
                        // {
                        subLayer = new Layer();
                        subLayer.Name = sec;
                        subLayer.Color = color;
                        subLayer.ParentLayerId = FrameLayer.Id;
                        subLayerIndex = doc.Layers.Add(subLayer);
                        // }

                        //else
                        //{
                        //    subLayer.Name = sec;
                        //    subLayer.Color = color;
                        //    subLayer.ParentLayerId = FrameLayer.Id;
                        //    subLayerIndex = subLayer.Index;
                        //}




                        RhinoObject[] objects = doc.Objects.FindByUserString("Sec", sec, true);
                        foreach (var item in objects)
                        {
                            item.Attributes.LayerIndex = subLayerIndex;
                            item.CommitChanges();
                        }

                        i++;
                    }

                    #endregion
                }
            }
        }

        // Main
        public static void AddTextToDocument2(IGH_DataAccess DA)
        {

            Rhino.RhinoDoc doc = Rhino.RhinoDoc.ActiveDoc;

            List<RhinoObject> rhinoCurves = doc.Objects.FindByObjectType(ObjectType.Curve).ToList();

            foreach (var crv in rhinoCurves)
            {
                CurveObject curveObject = crv as Rhino.DocObjects.CurveObject;
                Curve curvGeom = curveObject.CurveGeometry;


                // Calculate midpoint of the curve
                Point3d midpoint = curvGeom.PointAtNormalizedLength(0.5);

                // Calculate tangent vector at the midpoint
                Vector3d tangent = curvGeom.TangentAt(0.5);

                // Create a plane at the midpoint with the tangent vector as the normal
                //Plane textPlane = new Plane(midpoint, tangent);

                //Plane textPlane = new Plane(midpoint, Vector3d.ZAxis);


                // Create a vector pointing in the Z direction
                Vector3d zVector = new Vector3d(0, 0, 1);

                // Calculate the normal vector by taking the cross product of the tangent vector and the zVector
                Vector3d normal = Vector3d.CrossProduct(tangent, zVector);






                // for vr Frames
                if ((normal.X == 0 && normal.Y == 0 && normal.Z == 0) || (tangent.X < 0.01 && tangent.Y < 0.1 && tangent.Z > 0.89))
                {

                    var activeView = doc.Views.ActiveView;
                    var constructionPlane = activeView.ActiveViewport.ConstructionPlane();

                    #region Choose Here

                    // Create a plane X and Z
                    Plane textPlane = new Plane(midpoint, Vector3d.XAxis, Vector3d.ZAxis);



                    // Create a plane at Y and Z
                    // Plane textPlane = new Plane(midpoint, Vector3d.XAxis, Vector3d.ZAxis);

                    #endregion

                    #region further developement to get the view plane and try to modify based on this
                    //if (constructionPlane.XAxis> constructionPlane.YAxis && constructionPlane.ZAxis > constructionPlane.YAxis)
                    //{
                    //    textPlane = new Plane(midpoint, Vector3d.XAxis, Vector3d.ZAxis);

                    //   // textPlane = new Plane(midpoint, normal, Vector3d.ZAxis);

                    //}

                    //else if (constructionPlane.YAxis > constructionPlane.XAxis && constructionPlane.ZAxis > constructionPlane.XAxis)
                    //{
                    //    textPlane = new Plane(midpoint, Vector3d.YAxis, Vector3d.ZAxis);
                    //   // textPlane = new Plane(midpoint, normal, Vector3d.ZAxis);


                    //} 
                    #endregion


                    textPlane.Origin = midpoint;

                    string text = crv.Attributes.GetUserString("Group");


                    if (text == "All")
                    {
                        continue;
                    }
                    double height = 500; // Adjust as needed
                    string font = "Arial"; // Specify the font

                    // Add the text to the document
                    Guid id = doc.Objects.AddText(text, textPlane, height, font, false, false);
                    //Guid id2 = doc.Objects.AddText(text, textPlane2, height, font, false, false);


                }





                // for hz frames
                else
                {
                    // Create a plane at the midpoint with the tangent vector as the X axis and the normal as the Y axis

                    Plane textPlane = new Plane(midpoint, tangent, zVector);
                    Plane textPlane2 = new Plane(midpoint, tangent, Vector3d.YAxis);



                    textPlane.Origin = midpoint;

                    string text = crv.Attributes.GetUserString("Group");


                    if (text == "All")
                    {
                        continue;
                    }
                    double height = 450; // Adjust as needed
                    string font = "Arial"; // Specify the font



                    // show HZ text in verPlane
                    //  Guid id = doc.Objects.AddText(text, textPlane, height, font, false, false);



                    //show hz text on HZPlane
                    Guid id = doc.Objects.AddText(text, textPlane2, height, font, false, false);

                }




            }









        }
        public static void AddTextToDocument(IGH_DataAccess DA)
        {
            //List<Frame> RhinoFrames = getRhinoFrames();

            Rhino.RhinoDoc doc = Rhino.RhinoDoc.ActiveDoc;

            List<RhinoObject> rhinoCurves = doc.Objects.FindByUserString("MemberType", "Column", true).ToList();


            foreach (var crv in rhinoCurves)
            {
                CurveObject curveObject = crv as Rhino.DocObjects.CurveObject;
                Curve curvGeom = curveObject.CurveGeometry;


                // Calculate midpoint of the curve
                Point3d midpoint = curvGeom.PointAtNormalizedLength(0.5);
                Point3d startPoint = curvGeom.PointAtStart;


                // Calculate tangent vector at the midpoint
                Vector3d tangent = curvGeom.TangentAt(0.5);

                // Create a plane at the midpoint with the tangent vector as the normal
                //Plane textPlane = new Plane(midpoint, tangent);

                //Plane textPlane = new Plane(midpoint, Vector3d.ZAxis);


                // Create a vector pointing in the Z direction
                Vector3d zVector = new Vector3d(0, 0, 1);

                // Calculate the normal vector by taking the cross product of the tangent vector and the zVector
                Vector3d normal = Vector3d.CrossProduct(tangent, zVector);






                // for vr Frames
                if ((normal.X == 0 && normal.Y == 0 && normal.Z == 0) || (tangent.X < 0.01 && tangent.Y < 0.1 && tangent.Z > 0.89))
                {

                    var activeView = doc.Views.ActiveView;
                    var constructionPlane = activeView.ActiveViewport.ConstructionPlane();

                    #region Choose Here

                    // Create a plane X and Z
                    Plane textPlane = new Plane(startPoint, Vector3d.XAxis, Vector3d.YAxis);



                    // Create a plane at Y and Z
                    // Plane textPlane = new Plane(midpoint, Vector3d.XAxis, Vector3d.ZAxis);

                    #endregion

                    #region further developement to get the view plane and try to modify based on this
                    //if (constructionPlane.XAxis> constructionPlane.YAxis && constructionPlane.ZAxis > constructionPlane.YAxis)
                    //{
                    //    textPlane = new Plane(midpoint, Vector3d.XAxis, Vector3d.ZAxis);

                    //   // textPlane = new Plane(midpoint, normal, Vector3d.ZAxis);

                    //}

                    //else if (constructionPlane.YAxis > constructionPlane.XAxis && constructionPlane.ZAxis > constructionPlane.XAxis)
                    //{
                    //    textPlane = new Plane(midpoint, Vector3d.YAxis, Vector3d.ZAxis);
                    //   // textPlane = new Plane(midpoint, normal, Vector3d.ZAxis);


                    //} 
                    #endregion


                    textPlane.Origin = midpoint;

                    string text = crv.Attributes.GetUserString("ColumnMark");


                    if (text == "All")
                    {
                        continue;
                    }
                    double height = 350; // Adjust as needed
                    string font = "Arial"; // Specify the font

                    // Add the text to the document
                    Guid id = doc.Objects.AddText(text, textPlane, height, font, false, false);
                    //Guid id2 = doc.Objects.AddText(text, textPlane2, height, font, false, false);


                }





                // for hz frames
                else
                {
                    // Create a plane at the midpoint with the tangent vector as the X axis and the normal as the Y axis

                    Plane textPlane = new Plane(midpoint, tangent, zVector);
                    Plane textPlane2 = new Plane(midpoint, tangent, Vector3d.YAxis);



                    textPlane.Origin = midpoint;

                    string text = crv.Attributes.GetUserString("Group");


                    if (text == "All")
                    {
                        continue;
                    }
                    double height = 450; // Adjust as needed
                    string font = "Arial"; // Specify the font



                    // show HZ text in verPlane
                    //  Guid id = doc.Objects.AddText(text, textPlane, height, font, false, false);



                    //show hz text on HZPlane
                    Guid id = doc.Objects.AddText(text, textPlane2, height, font, false, false);

                }




            }









        }



        #endregion

        #region RhinoToEtabs

        // Original
        public static void RhinoToEtabs(IGH_DataAccess DA)
        {

            bool aa = false;
            DA.GetData(0, ref aa);
            if (aa)
            {


               // EtabsUtils.clearGroupAssingments();
                List<Point> assignedSapPoints = modifyPoints3();
                modifyFrames2(assignedSapPoints);
                //modifyAreas2(assignedSapPoints);

                modifyAreas_WithoutInteractiveDataBase(assignedSapPoints);
            }
        }


        //toTranseferGroupsToEtabs
        public static void RhinoToEtabsFramesOnly02(IGH_DataAccess DA)
        {
            List<Frame> rhinoFrames = new List<Frame>();
            Rhino.RhinoDoc doc = RhinoDoc.ActiveDoc;

            List<RhinoObject> rhinoCurves = doc.Objects.FindByObjectType(ObjectType.Curve).ToList();


            #region New
            foreach (var crv in rhinoCurves)
            {
                NameValueCollection n = crv.Attributes.GetUserStrings();

                Frame rhinoFrame = new Frame();
                if (crv.Attributes.GetUserString("EndPointId") == null)
                {
                    continue;
                }

                rhinoFrame.SubGroup = "**" + crv.Attributes.GetUserString("EtabsSubGroup");
                rhinoFrame.ID = crv.Attributes.GetUserString("ID");

                rhinoFrames.Add(rhinoFrame);

            }
            #endregion

            cSapModel myEtabs = EtabsUtils.Get_sap_mode();
            var grouppedFrames = rhinoFrames.GroupBy(x => x.SubGroup);

            foreach (var group in grouppedFrames)
            {
                List<string> etabsIDS = group.ToList().Select(x => x.ID).ToList();
                string groupName = group.Key;
                if (!groupName.StartsWith("**"))
                {
                    groupName = "*NoAssingedGroups";
                }
                EtabsUtils.createGroupinSap(etabsIDS, myEtabs, groupName);


            }


        }


        // forColumnsTool
        public static void createGroupOfShownElementsInSap(IGH_DataAccess DA)
        {
            List<Frame> revitFrames = new List<Frame>();
            Rhino.RhinoDoc doc = RhinoDoc.ActiveDoc;
            Dictionary<Level, double> levelDic = RevitUtils.getLevelsDictionary();

            List<RhinoObject> revitObjects = doc.Objects.Where(x => x.Attributes.GetUserString("Family") != null).ToList();


            #region getFramesWithAttributes
            foreach (var crv in revitObjects)
            {
                NameValueCollection n = crv.Attributes.GetUserStrings();

                Frame rhinoFrame = new Frame();
                string designation = crv.Attributes.GetUserString("Deignation");
                string ColumnMark = crv.Attributes.GetUserString("ColumnMark");
                string ColumnType = crv.Attributes.GetUserString("ColumnType");
                string EtabsTopStory = crv.Attributes.GetUserString("EtabsTopStory");
                string EtabsBotStory = crv.Attributes.GetUserString("EtabsBotStory");
                string etabsID = crv.Attributes.GetUserString("EtabsCropssID");



                string RevitTopLevel = crv.Attributes.GetUserString("TopLevel");
                string RevitBotLevel = crv.Attributes.GetUserString("BotLevel");


                double RevitTopLevelOffset = double.Parse(crv.Attributes.GetUserString("TopLevelOffset"));
                double RevitBotLevelOffset = double.Parse(crv.Attributes.GetUserString("BotLevelOffset"));

                string WidthBars = crv.Attributes.GetUserString("WidthBars");
                string DepthBars = crv.Attributes.GetUserString("DepthBars");


                rhinoFrame.ColumnsRebarDesign.NumberOfWidthBars = int.Parse(WidthBars);
                rhinoFrame.ColumnsRebarDesign.NumberOfDepthBars = int.Parse(WidthBars);


                rhinoFrame.Designation = designation;
                rhinoFrame.Mark = ColumnMark;
                rhinoFrame.ColumnsRebarDesign.Pattern = ColumnMark;

                rhinoFrame.EtabsTopStory = EtabsTopStory;
                rhinoFrame.EtabsBotStory = EtabsBotStory;


                rhinoFrame.RevitTopLevel = RevitTopLevel;
                rhinoFrame.RevitBotLevel = RevitBotLevel;

                rhinoFrame.RevitTopLevelOffset = RevitTopLevelOffset;
                rhinoFrame.RevitBotLevelOffset = RevitBotLevelOffset;


                ///////// very Important
                CurveObject curveObject = crv as Rhino.DocObjects.CurveObject;
                Curve curvGeom = curveObject.CurveGeometry;


                // you will use these coordinates to create new sapFrames only 
                // new sap frames that either startPoitId = null  or endPointId = null

                Point startPt = new Point(null, curvGeom.PointAtStart.X, curvGeom.PointAtStart.Y, curvGeom.PointAtStart.Z);
                Point endPt = new Point(null, curvGeom.PointAtEnd.X, curvGeom.PointAtEnd.Y, curvGeom.PointAtEnd.Z);


                rhinoFrame.StartPoint = startPt;
                rhinoFrame.EndPoint = endPt;

                rhinoFrame.ID = crv.Attributes.GetUserString("ID");

                revitFrames.Add(rhinoFrame);

            }
            #endregion

            var ColumnsMarksGroups = revitFrames.GroupBy(x => x.Mark);


            List<object> records = new List<object>();
            foreach (var columnMarkGroup in ColumnsMarksGroups)
            {
                List<Frame> columnMarksframes = columnMarkGroup.ToList();
                columnMarksframes = columnMarksframes.OrderBy(x => x.StartPoint.Z).ToList();



                #region Handling if it's only one designation set the bottom level to the lowest level and the top level to the highest level
                List<string> Alldesignations = columnMarkGroup.ToList().Select(x => x.Designation).Distinct().ToList();
                List<string> AllrevitBotLvls = columnMarkGroup.ToList().Select(x => x.RevitBotLevel).Distinct().ToList();
                List<string> AllrevitTopLvls = columnMarkGroup.ToList().Select(x => x.RevitTopLevel).Distinct().ToList();

                List<string> AlletabsTopStories = columnMarkGroup.ToList().Select(x => x.EtabsTopStory).Distinct().ToList();
                List<string> AlletabsBotStories = columnMarkGroup.ToList().Select(x => x.EtabsBotStory).Distinct().ToList();

                if (Alldesignations.Count == 1)
                {
                    string hightsLevel = RevitUtils.GetHighestLevel(AllrevitTopLvls.Union(AllrevitBotLvls).ToList(), levelDic);
                    string lowestLevel = RevitUtils.GetLowestLevel(AllrevitTopLvls.Union(AllrevitBotLvls).ToList(), levelDic);


                    // to avoid base Level and Top level offsets inConsistnty because it will be handled below in the levelsGroup
                    if (hightsLevel != lowestLevel)
                    {
                        foreach (var item in columnMarkGroup.ToList())
                        {

                            item.RevitTopLevel = hightsLevel;
                            item.RevitBotLevel = lowestLevel;

                        }
                    }


                }

                #endregion

                #region Handling Some elements in same Group is splitted and some other not In Revit 
                var BottomLevelGroups = columnMarksframes.GroupBy(x => x.RevitBotLevel);

                foreach (var bottomLvlGroup in BottomLevelGroups)
                {
                    List<Frame> bottomLvlFrames = bottomLvlGroup.ToList();

                    List<string> designation = bottomLvlGroup.ToList().Select(x => x.Designation).Distinct().ToList();
                    List<string> etabsTopStories = bottomLvlGroup.ToList().Select(x => x.EtabsTopStory).Distinct().ToList();
                    List<string> etabsBotStories = bottomLvlGroup.ToList().Select(x => x.EtabsBotStory).Distinct().ToList();

                    string revitBotLvl = bottomLvlGroup.Key;
                    List<string> revitTopLvls = bottomLvlGroup.ToList().Select(x => x.RevitTopLevel).Distinct().ToList();

                    List<double> revitTopLvlOffset = bottomLvlGroup.ToList().Select(x => x.RevitTopLevelOffset).Distinct().ToList();
                    List<double> revitBotLvlOffset = bottomLvlGroup.ToList().Select(x => x.RevitBotLevelOffset).Distinct().ToList();


                    if (designation.Count == 1 && revitTopLvls.Count == 2 && revitTopLvls[0] == revitBotLvl)
                    {
                        string TargtedTopLvl = RevitUtils.GetHighestLevel(revitTopLvls, levelDic);

                        if (TargtedTopLvl != revitBotLvl)
                        {
                            foreach (var item in columnMarkGroup.ToList().Intersect(bottomLvlFrames))
                            {
                                item.RevitTopLevel = TargtedTopLvl;
                            }
                        }

                    }


                }

                #endregion



                var ToplevelsGroups = columnMarksframes.GroupBy(x => x.RevitTopLevel);

                foreach (var topLvlGroup in ToplevelsGroups)
                {
                    List<string> designation = topLvlGroup.ToList().Select(x => x.Designation).Distinct().ToList();
                    List<string> etabsTopStories = topLvlGroup.ToList().Select(x => x.EtabsTopStory).Distinct().ToList();
                    List<string> etabsBotStories = topLvlGroup.ToList().Select(x => x.EtabsBotStory).Distinct().ToList();

                    string revitTopLvl = topLvlGroup.Key;
                    List<string> revitBotLvls = topLvlGroup.ToList().Select(x => x.RevitBotLevel).Distinct().ToList();

                    List<double> revitTopLvlOffset = topLvlGroup.ToList().Select(x => x.RevitTopLevelOffset).Distinct().ToList();
                    List<double> revitBotLvlOffset = topLvlGroup.ToList().Select(x => x.RevitBotLevelOffset).Distinct().ToList();

                    string ColumnType = topLvlGroup.ToList().Select(x => x.ColumnsRebarDesign.Pattern).First();
                    string NoOfBarsWidth = topLvlGroup.ToList().Select(x => x.ColumnsRebarDesign.NumberOfWidthBars).First().ToString();
                    string NoOfBarsDepth = topLvlGroup.ToList().Select(x => x.ColumnsRebarDesign.NumberOfDepthBars).First().ToString();



                    #region Handling if the bottom level of any instance is same as Top Level and there is only one designation


                    // In Revit
                    if (revitBotLvls.Count > 1 && designation.Count == 1)
                    {
                        if (revitBotLvls.Contains(topLvlGroup.Key))
                        {
                            revitBotLvls.Remove(topLvlGroup.Key);
                        }

                    }


                    // In Etabs
                    if (etabsBotStories.Count > 1 && designation.Count == 1)
                    {
                        if (etabsTopStories.Any(x => etabsBotStories.Contains(x)))
                        {
                            string desiredStory = etabsTopStories.First(x => etabsBotStories.Contains(x));
                            etabsBotStories.Remove(desiredStory);
                        }
                    }



                    #endregion


                    #region Handilng if the segment is assigned to only one level for top and bottom... we assign the offsets in this case

                    if ((revitBotLvls.Count == 1 && designation.Count == 1) && (revitBotLvls[0] == topLvlGroup.Key))
                    {

                        revitBotLvls[0] = revitBotLvls[0] + "+" + Math.Round(revitBotLvlOffset.Min());
                        revitTopLvl = revitTopLvl + "+" + Math.Round(revitTopLvlOffset.Max());

                    }

                    #endregion


                    #region Handling if the segement is assignet to only one designation and there are many bottom levels and nothing of them is the top level so we should assign bottom level to the lowest bottom level of them

                    if (revitBotLvls.Count > 1 && designation.Count == 1)
                    {
                        if (!revitBotLvls.Contains(topLvlGroup.Key))
                        {
                            revitBotLvls = new List<string> { RevitUtils.GetLowestLevel(revitBotLvls, levelDic) };
                        }

                    }

                    #endregion


                    #region Handling if the top two segements both are assigned to same top level with diffrent designation like case of parapet,for the bottom segment it's bottom level is different than it's top level and for the top segement it's bottom level equal to it's top level
                    // i have seprated between the two records and assigned each one indidually 

                    if (designation.Count == 2 && revitBotLvls.Count == 2 && revitBotLvls[1] == revitTopLvl)
                    {
                        #region BottomRecord
                        string revitBottomLevel = revitBotLvls[0];
                        List<Frame> bottomFrames = topLvlGroup.ToList().Where(x => x.RevitBotLevel == revitBottomLevel).ToList();



                        object bottomRectod = new
                        {
                            Mark = columnMarkGroup.Key,

                            RvtBotLvl = revitBottomLevel,
                            RvtTopLvl = revitTopLvl,

                            EtabsBotStory = string.Join(",", bottomFrames.Select(x => x.EtabsBotStory).Distinct().ToList()),
                            EtabsTopStory = string.Join(",", bottomFrames.Select(x => x.EtabsTopStory).Distinct().ToList()),


                            Designation = designation[0],
                            Count = bottomFrames.Count
                        };
                        #endregion

                        #region Top Record
                        List<Frame> topFrames = topLvlGroup.ToList().Where(x => x.RevitBotLevel != revitBottomLevel).ToList();
                        double revitTopOffset = topFrames.Select(x => x.RevitTopLevelOffset).Distinct().ToList().Max();
                        double revitBottomOffset = topFrames.Select(x => x.RevitBotLevelOffset).Distinct().ToList().Min();

                        string revitTopLevelWithBaseOffset = revitBotLvls[1] + "+" + revitBottomOffset;
                        string revitTopLevelWithTopOffset = revitBotLvls[1] + "+" + revitTopOffset;

                        object topRecord = new
                        {
                            Mark = columnMarkGroup.Key,

                            RvtBotLvl = revitTopLevelWithBaseOffset,
                            RvtTopLvl = revitTopLevelWithTopOffset,

                            EtabsBotStory = string.Join(",", topFrames.Select(x => x.EtabsBotStory).Distinct().ToList()),
                            EtabsTopStory = string.Join(",", topFrames.Select(x => x.EtabsTopStory).Distinct().ToList()),


                            Designation = designation[1],
                            Count = topFrames.Count
                        };
                        #endregion



                        records.Add(bottomRectod);
                        records.Add(topRecord);

                        continue;
                    }


                    #endregion
               
                    object rec = new
                    {
                        Mark = columnMarkGroup.Key,

                        RvtBotLvl = string.Join(",", revitBotLvls),
                        RvtTopLvl = revitTopLvl,

                        EtabsBotStory = string.Join(",", etabsBotStories),
                        EtabsTopStory = string.Join(",", etabsTopStories),

                        ColumnType = ColumnType,
                        NoOfBarsWidth = NoOfBarsWidth,
                        NoOfBarsDepth = NoOfBarsDepth,


                        Designation = string.Join(",", designation),
                        Count = topLvlGroup.ToList().Count


                    };

                    records.Add(rec);

                }



            }
            #region Skip Duplicated Records 

            var orderedRecords = records
        .OrderBy(r => ExtractNumber(((dynamic)r).Mark))
        .ToList();


            List<object> Newrecords = new List<object>();
            List<Record> NewRecords2 = new List<Record>();

            var MarkGroups = orderedRecords.GroupBy(x => ((dynamic)x).Mark);
            foreach (var markGroup in MarkGroups)
            {
                var designationGroups = GroupByContinuousDesignation(markGroup.ToList());

                // Processing or printing the grouped designations
                int i = 1;
                foreach (var designationGroup in designationGroups)
                {
                    if (designationGroup.Count > 1)
                    {
                        #region Record
                        Record newRec2 = new Record
                        {
                            Mark = ((dynamic)designationGroup[0]).Mark,
                            SubMark = ((dynamic)designationGroup[0]).Mark + "-" + i,
                            RvtBotLvl = ((dynamic)designationGroup[0]).RvtBotLvl,
                            RvtTopLvl = ((dynamic)designationGroup[designationGroup.Count - 1]).RvtTopLvl,
                            EtabsBotStory = ((dynamic)designationGroup[0]).EtabsBotStory,
                            EtabsTopStory = ((dynamic)designationGroup[designationGroup.Count - 1]).EtabsTopStory,
                            ColumnType = ((dynamic)designationGroup[0]).ColumnType,
                            Designation = ((dynamic)designationGroup[0]).Designation,
                            Count = ((dynamic)designationGroup[0]).Count,
                            Description = "FROM" + ((dynamic)designationGroup[0]).RvtBotLvl + "TO UNDER SIDE OF" + ((dynamic)designationGroup[designationGroup.Count - 1]).RvtTopLvl
                        };

                        if (newRec2.ColumnType == "R")
                        {
                            var obj = ExtractNumbers(newRec2.Designation);
                            newRec2.Width = obj.firstNumber.ToString();
                            newRec2.Depth = obj.secondNumber.ToString();
                            newRec2.Diameter = "";
                            newRec2.numberOfBarsInWidthDir = int.Parse(((dynamic)designationGroup[0]).NoOfBarsWidth);
                            newRec2.numberOfBarsInDepthDir = int.Parse(((dynamic)designationGroup[0]).NoOfBarsDepth);


                        }

                        else if (newRec2.ColumnType == "C")
                        {
                            newRec2.Diameter = ExtractDiameter(newRec2.Designation);

                            newRec2.Width = "";
                            newRec2.Depth = "";
                            newRec2.numberOfBarsInDepthDir = 0;
                            newRec2.numberOfBarsInWidthDir = 0;

                        }
                        #endregion
                        object newRec = new
                        {
                            Mark = ((dynamic)designationGroup[0]).Mark,

                            SubMark = ((dynamic)designationGroup[0]).Mark + "-" + i,
                            RvtBotLvl = ((dynamic)designationGroup[0]).RvtBotLvl,
                            RvtTopLvl = ((dynamic)designationGroup[designationGroup.Count - 1]).RvtTopLvl,

                            EtabsBotStory = ((dynamic)designationGroup[0]).EtabsBotStory,
                            EtabsTopStory = ((dynamic)designationGroup[designationGroup.Count - 1]).EtabsTopStory,

                            ColumnType = ((dynamic)designationGroup[0]).ColumnType,

                            Designation = ((dynamic)designationGroup[0]).Designation,
                            Count = ((dynamic)designationGroup[0]).Count,


                        };

                        Newrecords.Add(newRec);
                        NewRecords2.Add(newRec2);
                    }

                    else
                    {

                        #region Record
                        Record newRec2 = new Record
                        {
                            Mark = ((dynamic)designationGroup[0]).Mark,
                            SubMark = ((dynamic)designationGroup[0]).Mark + "-" + i,
                            RvtBotLvl = ((dynamic)designationGroup[0]).RvtBotLvl,
                            RvtTopLvl = ((dynamic)designationGroup[0]).RvtTopLvl,
                            EtabsBotStory = ((dynamic)designationGroup[0]).EtabsBotStory,
                            EtabsTopStory = ((dynamic)designationGroup[0]).EtabsTopStory,
                            ColumnType = ((dynamic)designationGroup[0]).ColumnType,
                            Designation = ((dynamic)designationGroup[0]).Designation,
                            Count = ((dynamic)designationGroup[0]).Count,
                            Description = "FROM" + ((dynamic)designationGroup[0]).RvtBotLvl + "TO UNDER SIDE OF" + ((dynamic)designationGroup[0]).RvtTopLvl

                        };
                        if (newRec2.ColumnType == "R")
                        {
                            var obj = ExtractNumbers(newRec2.Designation);
                            newRec2.Width = obj.firstNumber.ToString();
                            newRec2.Depth = obj.secondNumber.ToString();
                            newRec2.Diameter = "";

                        }

                        else if (newRec2.ColumnType == "C")
                        {
                            newRec2.Diameter = ExtractDiameter(newRec2.Designation);

                            newRec2.Width = "";
                            newRec2.Depth = "";
                        }

                        #endregion
                        object newRec = new
                        {
                            Mark = ((dynamic)designationGroup[0]).Mark,

                            SubMark = ((dynamic)designationGroup[0]).Mark + "-" + i,
                            RvtBotLvl = ((dynamic)designationGroup[0]).RvtBotLvl,
                            RvtTopLvl = ((dynamic)designationGroup[0]).RvtTopLvl,

                            EtabsBotStory = ((dynamic)designationGroup[0]).EtabsBotStory,
                            EtabsTopStory = ((dynamic)designationGroup[0]).EtabsTopStory,

                            ColumnType = ((dynamic)designationGroup[0]).ColumnType,
                            Designation = ((dynamic)designationGroup[0]).Designation,
                            Count = ((dynamic)designationGroup[0]).Count

                        };

                        Newrecords.Add(newRec);
                        NewRecords2.Add(newRec2);

                    }

                    i++;
                }

            }


            #endregion

            RhinoRevit.RevitSchdulesManupulation("STRUCTURAL CONCRETE COLUMNS SCHEDULES - DETAILED DESIGN", NewRecords2);

            #region add in excel

            // Create Excel application
            Application excelApp = new Application();
            excelApp.Visible = true;

            // Create a new workbook
            Workbook workbook = excelApp.Workbooks.Add();

            // Add data to the first worksheet

            ExcelUtils.AddDataToWorksheet2(workbook, Newrecords, "ColumnsMarks");

            #endregion
        }


        //createGroupOfShownElementsInSap
        public static void RhinoToEtabsFramesOnly(IGH_DataAccess DA)
        {

            bool aa = false;
            DA.GetData(0, ref aa);
            if (aa)
            {
                cSapModel sapModel = Get_sap_mode();

                Rhino.RhinoDoc doc = RhinoDoc.ActiveDoc;

                List<RhinoObject> allRhinoCurves = doc.Objects.FindByObjectType(ObjectType.Curve).ToList();

                List<string> etabsIDS = new List<string>();
                foreach (var crv in allRhinoCurves)
                {
                    NameValueCollection n = crv.Attributes.GetUserStrings();
                    string AnalysisId = crv.Attributes.GetUserString("ID");
                    etabsIDS.Add(AnalysisId);
                }

                createGroupinSap(etabsIDS, sapModel, "0SelectedElements");
            }
        }



        //RhinoToEtabsSetIdsBasedOnParameters
        public static void RhinoToEtabsSetIdsBasedOnParameters(IGH_DataAccess DA)
        {

            bool aa = false;
            DA.GetData(0, ref aa);
            if (aa)
            {
                cSapModel sapModel = Get_sap_mode();

                Rhino.RhinoDoc doc = RhinoDoc.ActiveDoc;

                List<RhinoObject> etabsRhinoFrames = doc.Objects.FindByObjectType(ObjectType.Curve).Where(x=> x.Attributes.GetUserString("BeamMark") != null).ToList();


                var etabsGroups = etabsRhinoFrames
                                            .GroupBy(x => x.Attributes.GetUserString("BeamMark"))
                                            .ToDictionary(g => g.Key, g => g.ToList());


                foreach (var group in etabsGroups)
                {

                    string beamMark = group.Key;


                    if (group.Value.Count ==1)
                    {
                        string etabsID = group.Value[0].Attributes.GetUserString("ID");

                        sapModel.FrameObj.ChangeName(etabsID, beamMark);
                    }

                    else if (group.Value.Count > 1)
                    {
                        int count = 1;
                        int ret = 1;
                        foreach (var item in group.Value)
                        {
                            string etabsID = item.Attributes.GetUserString("ID");

                            ret = sapModel.FrameObj.ChangeName(etabsID, beamMark + "**-" + count);

                            count++;
                        }
                    }
    


                }

               
            }
        }



        public static List<Frame> getRhinoFrames(ref List<List<Frame>> duplicatedMembers)
        {
            string sourcePath = "";
            List<Frame> rhinoFrames = new List<Frame>();
            Rhino.RhinoDoc doc = RhinoDoc.ActiveDoc;

            List<RhinoObject> rhinoCurves = doc.Objects.FindByObjectType(ObjectType.Curve).ToList();


            #region New
            foreach (var crv in rhinoCurves)
            {
                NameValueCollection n = crv.Attributes.GetUserStrings();

                List<string> keys = n.AllKeys.ToList().Where(x => (x.Contains("Group") && x != "Groups")).ToList();
                List<string> newgroups = new List<string>();
                keys.ForEach(x => newgroups.Add(n[x]));


                Frame rhinoFrame = new Frame();
                rhinoFrame.Sec = crv.Attributes.GetUserString("EtabsSection");
                rhinoFrame.Shape = crv.Attributes.GetUserString("Shape");
                rhinoFrame.ID = crv.Attributes.GetUserString("ID");
                string stPtId = crv.Attributes.GetUserString("StartPointId");
                string endPtId = crv.Attributes.GetUserString("EndPointId");
                rhinoFrame.Groups = new List<string>();

                double beta = 0.00;
                double.TryParse(crv.Attributes.GetUserString("Angle"), out beta);
                rhinoFrame.Beta = beta;

                double length = 0;
                double.TryParse(crv.Attributes.GetUserString("Length"), out length);
                rhinoFrame.Length = length;
                rhinoFrame.Groups = newgroups;


                #region NewAttributes
                rhinoFrame.Mark = crv.Attributes.GetUserString("ColumnMark");
                string membType = crv.Attributes.GetUserString("MemberType");
                if (membType == "Column")
                {
                    rhinoFrame.MemberType = MemberType.Column;
                }
                else
                {
                    rhinoFrame.MemberType = MemberType.Beam;

                }

                rhinoFrame.Designation = crv.Attributes.GetUserString("Deignation");


                #endregion

                ///////// very Important
                CurveObject curveObject = crv as Rhino.DocObjects.CurveObject;
                Curve curvGeom = curveObject.CurveGeometry;


                // you will use these coordinates to create new sapFrames only 
                // new sap frames that either startPoitId = null  or endPointId = null

                Point startPt = new Point(stPtId, curvGeom.PointAtStart.X, curvGeom.PointAtStart.Y, curvGeom.PointAtStart.Z);
                Point endPt = new Point(endPtId, curvGeom.PointAtEnd.X, curvGeom.PointAtEnd.Y, curvGeom.PointAtEnd.Z);


                rhinoFrame.StartPoint = startPt;
                rhinoFrame.EndPoint = endPt;

                rhinoFrames.Add(rhinoFrame);

            }
            #endregion

            #region Duplication
            var grouppedFrames = rhinoFrames
                                            .Where(x => !string.IsNullOrWhiteSpace(x.ID)) // Filter out items with null, empty, or whitespace IDs
                                            .GroupBy(x => x.ID)
                                            .Select(group => group.ToList()) // Convert each group into a list
                                            .ToList(); // Convert the result into a list of lists




            List<string> existIds = rhinoFrames.Select(x => x.ID).ToList();


            foreach (var group in grouppedFrames.Where(x => x.Count > 1))
            {
                rhinoFrames.RemoveAll(item => group.Skip(1).Contains(item));
            }

            foreach (var group in grouppedFrames.Where(x => x.Count > 1))
            {
                string id = group[0].ID;

                foreach (var item in group.Skip(1))
                {
                    int i = 1;

                    string subId = id + " -" + i;
                    while (existIds.Contains(subId))
                    {
                        i++;
                        subId = id + " -" + i;
                    }

                    item.ID = subId;
                    existIds.Add(subId);

                }

            }


            duplicatedMembers = grouppedFrames.Where(x => x.Count > 1).ToList();
            #endregion

            #region old
            //foreach (var crv in rhinoCurves)
            //{
            //    NameValueCollection n =  crv.Attributes.GetUserStrings();

            //    List<string> keys = n.AllKeys.ToList().Where(x=> (x.Contains("Group") && x!= "Groups")).ToList();
            //    List<string> newgroups = new List<string>();
            //    keys.ForEach(x => newgroups.Add(n[x]));


            //    Frame rhinoFrame = new Frame();
            //    rhinoFrame.Sec = crv.Attributes.GetUserString("Section");
            //    rhinoFrame.Shape = crv.Attributes.GetUserString("Shape");
            //    rhinoFrame.Group = crv.Attributes.GetUserString("Group");
            //    rhinoFrame.ID = crv.Attributes.GetUserString("ID");
            //    string stPtId = crv.Attributes.GetUserString("StartPointId");
            //    string endPtId = crv.Attributes.GetUserString("EndPointId");

            //    rhinoFrame.Groups = new List<string>();

            //    string groups = crv.Attributes.GetUserString("Groups");
            //    double beta = 0.00;
            //    double.TryParse(crv.Attributes.GetUserString("Angle"), out beta);
            //    rhinoFrame.Beta = beta;

            //    double length = 0;
            //    double.TryParse(crv.Attributes.GetUserString("Length"), out length);
            //    rhinoFrame.Length = length;

            //    // set groups 
            //    if (groups != null)
            //    {
            //        List<string> groupAttriputesNames= crv.Attributes.GetUserString("Groups").Split(',').ToList();
            //        List<string> groupNames = new List<string>();


            //        foreach (var attriputeName in groupAttriputesNames)
            //        {
            //            string groupName = crv.Attributes.GetUserString(attriputeName);
            //            groupNames.Add(groupName);
            //        }
            //        rhinoFrame.Groups = groupNames;
            //    }



            //    ///////// very Important
            //    CurveObject curveObject = crv as Rhino.DocObjects.CurveObject;
            //    Curve curvGeom = curveObject.CurveGeometry;


            //    // you will use these coordinates to create new sapFrames only 
            //    // new sap frames that either startPoitId = null  or endPointId = null

            //    Point startPt = new Point(stPtId, curvGeom.PointAtStart.X, curvGeom.PointAtStart.Y, curvGeom.PointAtStart.Z);
            //    Point endPt = new Point(endPtId, curvGeom.PointAtEnd.X, curvGeom.PointAtEnd.Y, curvGeom.PointAtEnd.Z);


            //    rhinoFrame.StartPoint = startPt;
            //    rhinoFrame.EndPoint = endPt;

            //    rhinoFrames.Add(rhinoFrame);

            //}
            #endregion

            return rhinoFrames;
        }

        public static List<Point> getRhinoPoints()
        {
            List<Point> rhinoPts = new List<Point>();
            Rhino.RhinoDoc doc = RhinoDoc.ActiveDoc;

            List<RhinoObject> rhinoPoints = doc.Objects.FindByObjectType(ObjectType.Point).ToList();

            foreach (var pt in rhinoPoints)
            {

                Point rhinoPoint = new Point();
                rhinoPoint.Id = pt.Attributes.GetUserString("ID");

                ////// very important 
                PointObject ptObj = pt as PointObject;
                Rhino.Geometry.Point3d ptt = ptObj.PointGeometry.Location;

                // you shall use these cooordinates in sap 
                // you should get the matching points by id then change the coodrdinates

                // if there is not matching points (pt.id == null)  you should think in this case 
                // user add this point manually in rhino and not assign any attribute to it 

                rhinoPoint.X = ptt.X;
                rhinoPoint.Y = ptt.Y;
                rhinoPoint.Z = ptt.Z;


                rhinoPts.Add(rhinoPoint);
            }


            return rhinoPts;
        }

        public static List<Area> getRhinoAreas(ref List<List<Area>> duplicatedMembers)
        {
            cSapModel myEtabs = Get_sap_mode();
            List<Area> rhinoAreas = new List<Area>();
            Rhino.RhinoDoc doc = RhinoDoc.ActiveDoc;

            List<RhinoObject> rhinoBreps = doc.Objects.FindByObjectType(ObjectType.Brep).ToList();

            #region New

            foreach (var brep in rhinoBreps)
            {

                NameValueCollection n = brep.Attributes.GetUserStrings();

                List<string> keys = n.AllKeys.ToList().Where(x => (x.Contains("Group") && x != "Groups")).ToList();
                List<string> newgroups = new List<string>();
                keys.ForEach(x => newgroups.Add(n[x]));



                Area rhinoArea = new Area();

                rhinoArea.Sec = brep.Attributes.GetUserString("EtabsSection");
                rhinoArea.ID = brep.Attributes.GetUserString("ID");
                rhinoArea.Groups = new List<string>();

                string groups = brep.Attributes.GetUserString("Groups");

                double angle = 0.00;
                double.TryParse(brep.Attributes.GetUserString("Angle"), out angle);
                rhinoArea.Beta = angle;

                int NoOfPts = 0;
                int.TryParse(brep.Attributes.GetUserString("no Of Pts"), out NoOfPts);
                rhinoArea.NoOfPts = NoOfPts;

                rhinoArea.Groups = newgroups;




                // you will use this coordinates in case where  ids of the vertices is null (new Area)
                List<Point> points = new List<Point>();
                BrepObject brPObject = brep as BrepObject;
                Brep bb = brPObject.BrepGeometry;


                


                #region OLD Using Edges
                //BrepEdgeList edgeList = bb.Edges;
                //foreach (var edge in edgeList)
                //{
                //    Point3d pt1 = edge.PointAtStart;
                //    Point3d pt2 = edge.PointAtEnd;
                //    Point stPt = new Point(null, pt1.X, pt1.Y, pt1.Z);
                //    Point endPt = new Point(null, pt2.X, pt2.Y, pt2.Z);

                //    if (!points.Any(x => x.X == stPt.X && x.Y == stPt.Y && x.Z == stPt.Z))
                //    {
                //        points.Add(stPt);

                //    }


                //    if (!points.Any(x => x.X == endPt.X && x.Y == endPt.Y && x.Z == endPt.Z))
                //    {
                //        points.Add(endPt);

                //    }

                //}

                #endregion


                // Loop through each vertex in the Brep
                foreach (var vertex in bb.Vertices)
                {
                    // Get the 3D point location of the vertex
                    Point3d pt = vertex.Location;

                    // Create a new Point object based on the vertex location
                    Point vertexPoint = new Point(null, pt.X, pt.Y, pt.Z);

                    // Check if this point already exists in the list to avoid duplicates
                    if (!points.Any(p => p.X == vertexPoint.X && p.Y == vertexPoint.Y && p.Z == vertexPoint.Z))
                    {
                        points.Add(vertexPoint);
                    }
                }


                // No Need for it 
                #region get the counterClockWiseProerty

                // Get the first face
                BrepFace face = bb.Faces[0];

                // Find the centroid (center point) of the face
                AreaMassProperties areaProps = AreaMassProperties.Compute(face);
                Point3d center = areaProps.Centroid;

                // Get the normal of the face
                Rhino.Geometry.Surface surface = face.UnderlyingSurface();
                Vector3d normal = surface.NormalAt(surface.Domain(0).Mid, surface.Domain(1).Mid);

                // Retrieve unique vertices of the brep
                List<Point3d> vertices = bb.Vertices.Select(v => v.Location).Distinct().ToList();

                // Check if the vertices are counterclockwise
                bool isCounterClockwise = Area.AreVerticesCounterClockwise(vertices, center, normal);

                rhinoArea.IsCounterClockWise = isCounterClockwise;

                #endregion

                rhinoArea.vertices = points;

                rhinoArea.guid = brep.Id;
                rhinoAreas.Add(rhinoArea);
            }

            #endregion

            #region Duplication
            var grouppedAreas = rhinoAreas
                                            .Where(x => !string.IsNullOrWhiteSpace(x.ID)) // Filter out items with null, empty, or whitespace IDs
                                            .GroupBy(x => x.ID)
                                            .Select(group => group.ToList()) // Convert each group into a list
                                            .ToList(); // Convert the result into a list of lists




            List<string> existIds = rhinoAreas.Select(x => x.ID).ToList();


            foreach (var group in grouppedAreas.Where(x => x.Count > 1))
            {
                // There is a a problem with the splitted areas that rhino bring it's vertices in a random order
                // I created a method that retrieve any brep in a counterClockWise order
                // you need to retireve the rhinoBrep here and then use this method.

                //the method is as below

                rhinoAreas.RemoveAll(item => group.Skip(1).Contains(item));
                
                List<Area> elements = group.ToList();

                foreach (var areaElement in elements)
                {
                    var brep = doc.Objects.FindByObjectType(ObjectType.Brep).ToList().FirstOrDefault(x => x.Id == areaElement.guid);
                    BrepObject brPObject = brep as BrepObject;
                    Brep bb = brPObject.BrepGeometry;

                    var RhinoPoints2=  bb.Vertices.Select(x => x.Location).ToList();
                    var Rhinopoints = Area.GetCounterClockwiseVertices(bb.Faces[0]);

                    List<Point> pts = new List<Point>();
                    // Loop through each vertex in the Brep
                    foreach (var pt in Rhinopoints)
                    {
                        // Create a new Point object based on the vertex location
                        Point vertexPoint = new Point(null, pt.X, pt.Y, pt.Z);

                        // Check if this point already exists in the list to avoid duplicates
                        if (!pts.Any(p => p.X == vertexPoint.X && p.Y == vertexPoint.Y && p.Z == vertexPoint.Z))
                        {
                            pts.Add(vertexPoint);
                        }
                    }

                    areaElement.vertices = pts;
                }




            }

            foreach (var group in grouppedAreas.Where(x => x.Count > 1))
            {
                string id = group[0].ID;

                foreach (var item in group.Skip(1))
                {
                    int i = 1;

                    string subId = id + " -" + i;
                    while (existIds.Contains(subId))
                    {
                        i++;
                        subId = id + " -" + i;
                    }

                    item.ID = subId;
                    existIds.Add(subId);

                }

            }


            duplicatedMembers = grouppedAreas.Where(x => x.Count > 1).ToList();
            #endregion
          


            return rhinoAreas;
        }

        public static cSapModel Get_sap_mode()
        {
            /// cOAPI mySapObject = (cOAPI)System.Runtime.InteropServices.Marshal.GetActiveObject("CSI.SAP2000.API.SapObject");
            // ETABSv1.cSapModel mysapmodel = (cSapModel) System.Runtime.InteropServices.Marshal.GetActiveObject("CSI.SAP2000.API.SapModel");

            ETABSv1.cOAPI mySapObject2 = Marshal.GetActiveObject("CSI.ETABS.API.ETABSObject") as ETABSv1.cOAPI;

            mySapObject2.SapModel.SetPresentUnits(eUnits.Ton_mm_C);



            #region getSteelMaterial


            int ret;
            string material = null;

            int numberOfNames = 0;
            string[] materialList = new string[0];
            ret = mySapObject2.SapModel.PropMaterial.GetNameList(ref numberOfNames, ref materialList);
            eMatType matType = eMatType.Steel;
            int a = 0;
            foreach (var mat in materialList)
            {

                mySapObject2.SapModel.PropMaterial.GetTypeOAPI(mat, ref matType, ref a);

                if (matType == eMatType.Steel)
                {
                    material = mat;
                    break;
                }


            }






            #endregion

            #region createDummySectionInSap

            int dummySectionId = mySapObject2.SapModel.PropFrame.SetCircle("dummySection", material, 10);


            #endregion  



            return mySapObject2.SapModel;


            //ETABSv1.cHelper cHelper = new ETABSv1.Helper();
            //cOAPI mySapObject2 = cHelper.GetObject("CSI.SAP2000.API.SapObject");
            //ETABSv1.cSapModel mysapmodel = mySapObject2.SapModel;


        }

        // use
        public static List<Point> modifyPoints3()
        {

            cSapModel mySapModel = Get_sap_mode();
            List<Point> RhinoPoints = getRhinoPoints();
            List<List<Area>> duplicatedAreas = null;
            List<Area> RhinoAreas = getRhinoAreas(ref duplicatedAreas);

            List<List<Frame>> duplicatedGroups = null;
            List<Frame> RhinoFrames = getRhinoFrames(ref duplicatedGroups);

            List<Point> newVerticesForNewAreas = new List<Point>();

            foreach (Area area in RhinoAreas)
            {
                foreach (var pt in area.vertices)
                {
                    if (!RhinoPoints.Any(x => x.X == pt.X && x.Y == pt.Y && x.Z == pt.Z))
                    {
                        newVerticesForNewAreas.Add(pt);
                    }
                }


            }

            // new
            foreach (var area in duplicatedAreas.SelectMany(x=>x))
            {
                foreach (var pt in area.vertices)
                {
                    if (!RhinoPoints.Any(x => x.X == pt.X && x.Y == pt.Y && x.Z == pt.Z))
                    {
                        newVerticesForNewAreas.Add(pt);
                    }
                }
            }


            foreach (Frame frame in RhinoFrames)
            {

                if (!RhinoPoints.Any(x => x.X == frame.StartPoint.X && x.Y == frame.StartPoint.Y && x.Z == frame.StartPoint.Z))
                {
                    newVerticesForNewAreas.Add(frame.StartPoint);
                }

                if (!RhinoPoints.Any(x => x.X == frame.EndPoint.X && x.Y == frame.EndPoint.Y && x.Z == frame.EndPoint.Z))
                {
                    newVerticesForNewAreas.Add(frame.EndPoint);
                }

            }

            RhinoPoints.AddRange(newVerticesForNewAreas);



            // List<Point> newRhinoPoints = RhinoPoints.Where(x => x.Id == null).ToList();
            //  List<Point> ModifiedRhinoPoints = RhinoPoints.Where(x => x.Id != null).ToList();


            List<string> existingPointsSapIDS = EtabsUtils.getSapPointIDS();

            // all points will be added as a new points and sap will handle this issue 
            List<string> ids = EtabsUtils.addNewPointsNew(ref RhinoPoints);



            List<string> SapPointIdsToBeRmoved = existingPointsSapIDS.Where(x => !RhinoPoints.Select(a => a.Id).ToList().Contains(x)).ToList();
            EtabsUtils.createGroupOfPointsinSap(SapPointIdsToBeRmoved, "A00PointsToBeRemoved");

            return RhinoPoints;

        }

        // used
        public static void modifyFrames2(List<Point> assignedSapPts)
        {
            List<List<Frame>> duplicatedGroups = null;
            List<Frame> RhinoFrames = getRhinoFrames(ref duplicatedGroups);
            List<Frame> newRhinoFrames = RhinoFrames.Where(x => x.ID == null).ToList();
            List<Frame> existingRhinoFrames = RhinoFrames.Where(x => x.ID != null).ToList();


            List<Column> columns = new List<Column>();
            List<Beam> beams = new List<Beam>();
            List<Frame> existingSapFrames = EtabsUtils.getSapFrames(ref beams, ref columns); ;


            // get sapIds to be removed
            List<Frame> SapFramesToBeRmoved = existingSapFrames.Where(x => !RhinoFrames.Select(a => a.ID).ToList().Contains(x.ID)).ToList();
            List<string> SapidsToRemoved = SapFramesToBeRmoved.Select(x => x.ID).ToList();

            EtabsUtils.deleteFrameObjects(SapidsToRemoved);

            // add new Frames
            List<string> newSapIDS = EtabsUtils.createSapFrames(newRhinoFrames);

            // adjust existingFramesWithTheNewPointsIDs
            foreach (var item in existingRhinoFrames)
            {
                item.StartPoint.Id = assignedSapPts.First(x => x.X == item.StartPoint.X && x.Y == item.StartPoint.Y && x.Z == item.StartPoint.Z).Id;
                item.EndPoint.Id = assignedSapPts.First(x => x.X == item.EndPoint.X && x.Y == item.EndPoint.Y && x.Z == item.EndPoint.Z).Id;

            }

            // update frames assignments 
            EtabsUtils.updataAssignmentsFromRhinoToSap(existingRhinoFrames);


            // EtabsUtils.createGroupinSap(SapidsToRemoved, "A00SapFramesToBeRemoved");
            EtabsUtils.createGroupinSap(newSapIDS, "A00NewSapFrames");

            #region Handling Duplicated Members 

            foreach (var dupGroup in duplicatedGroups)
            {
                copyFrameAssignments(dupGroup[0], dupGroup.Skip(1).ToList());

            }


            #endregion

        }

        // used
        public static void modifyAreas2(List<Point> assignedSapPts)
        {
            List<List<Area>> duplicatedAreas = null;
            List<Area> RhinoAreas = getRhinoAreas(ref duplicatedAreas);
            List<Area> newRhinoAreas = RhinoAreas.Where(x => x.ID == null).ToList();
            List<Area> existingRhinoAreas = RhinoAreas.Where(x => x.ID != null).ToList();
            List<Area> existingSapAreas = EtabsUtils.getEtabsAreas();
            List<Area> SapAreasToBeRmoved = existingSapAreas.Where(x => !RhinoAreas.Select(a => a.ID).ToList().Contains(x.ID)).ToList();


            // deleteAreas
             EtabsUtils.deleteAreaObjects(SapAreasToBeRmoved.Select(x => x.ID).ToList());


            // create new Areas
            List<string> newSapAreaIDS = EtabsUtils.createSapArea(newRhinoAreas);

            #region Handling Duplicated Members 

            foreach (var dupGroup in duplicatedAreas)
            {
                copyAreaAssignments(dupGroup[0], dupGroup.Skip(1).ToList());

                if ((dupGroup.Skip(1).ToList().Count>0))
                {
                    RhinoAreas.AddRange(dupGroup.Skip(1).ToList());

                }
            }
            #endregion


            // adjust existingAreaWithTheNewPointsIDs
            foreach (var areaElement in RhinoAreas)
            {
                foreach (var pt in areaElement.vertices)
                {
                    pt.Id = assignedSapPts.First(x => x.X == pt.X && x.Y == pt.Y && x.Z == pt.Z).Id;
                }
            }





            // update Areas assignments
            EtabsUtils.updataAssignmentsFromRhinoToSap(RhinoAreas);


            //EtabsUtils.createGroupOfAreasinSap(SapAreasToBeRmoved.Select(x => x.ID).ToList(), "A00SapAreasToBeRemoved");
            EtabsUtils.createGroupOfAreasinSap(newSapAreaIDS, "A00NewSapaAreas");



      



        }

        public static void modifyAreas_WithoutInteractiveDataBase(List<Point> assignedSapPts)
        {
            List<List<Area>> duplicatedAreas = null;
            List<Area> RhinoAreas = getRhinoAreas(ref duplicatedAreas);

            List<Area> newRhinoAreas = RhinoAreas.Where(x => x.ID == null).ToList();
            List<Area> existingRhinoAreas = RhinoAreas.Where(x => x.ID != null).ToList();


            List<Area> existingSapAreas = EtabsUtils.getEtabsAreas();
            List<Area> SapAreasToBeRmoved = existingSapAreas.Where(x => !RhinoAreas.Select(a => a.ID).ToList().Contains(x.ID)).ToList();


            // deleteAreas
           // EtabsUtils.deleteAreaObjects(SapAreasToBeRmoved.Select(x => x.ID).ToList());


            // create new Areas
            List<string> newSapAreaIDS = EtabsUtils.createSapArea(newRhinoAreas);






            // adjust existingAreaWithTheNewPointsIDs
            foreach (var areaElement in RhinoAreas)
            {
                foreach (var pt in areaElement.vertices)
                {
                    pt.Id = assignedSapPts.First(x => x.X == pt.X && x.Y == pt.Y && x.Z == pt.Z).Id;
                }
            }



            #region Handling Duplicated Members 

            foreach (var dupGroup in duplicatedAreas)
            {
                copyAreaAssignments(dupGroup[0], dupGroup.Skip(1).ToList());

                if ((dupGroup.Skip(1).ToList().Count > 0))
                {
                    RhinoAreas.AddRange(dupGroup.Skip(1).ToList());

                }
            }
            #endregion


            // update Areas assignments
            // EtabsUtils.updataAssignmentsFromRhinoToSap(RhinoAreas);




            //EtabsUtils.createGroupOfAreasinSap(SapAreasToBeRmoved.Select(x => x.ID).ToList(), "A00SapAreasToBeRemoved");
            EtabsUtils.createGroupOfAreasinSap(newSapAreaIDS, "A00NewSapaAreas");







        }


        public static void copyAreaAssignments(Area origElem, List<Area> copiedElements)
        {

            #region def
            int numberGroups = 0;
            string[] Groups = new string[0];
            int ret;
            int UniformToFrameLoadNo = 0;
            int UniformLoadNo = 0;
            int TempLoadNo = 0;
            string[] AreaName = null;
            string[] LoadPatUF = null;
            string[] LoadPatU = null;
            string[] LoadPatT = null;

            string[] CSysUF = null;
            string[] CSysU = null;

            int[] DirUF = null;
            int[] DirU = null;

            double[] ValueUF = null;
            double[] ValueU = null;
            double[] ValueT = null;

            int[] DistTypeUF = null;

            int[] MyType = null;
            string[] PatternName = default;

            double ang = 0.0;
            bool advanced = false;

            int[] mypDir = new int[1];
            string[] myplpt = new string[1];
            double[] myplvector = new double[2];


            string[] myPoints = new string[2];
            double[] MyPlVect = new double[3];
            int[] MyPlDir = new int[2];
            string st_id = "";
            string end_id = "";
            Boolean aa = true;
            bool Advanced = true;
            int plane = 0;
            int plVect = 0;
            string cys = null;
            string prop = null;

            int noOfPoints = 0;
            double[] Xcoor = null;
            double[] Ycoor = null;
            double[] Zcoor = null;

            String ID = null;
            #endregion

            cSapModel mysap = Get_sap_mode();


            #region Get Assignments

            mysap.AreaObj.GetGroupAssign(origElem.ID, ref numberGroups, ref Groups);
            ret = mysap.AreaObj.GetLoadUniformToFrame(origElem.ID, ref UniformToFrameLoadNo, ref AreaName, ref LoadPatUF, ref CSysUF, ref DirUF, ref ValueUF, ref DistTypeUF);
            ret = mysap.AreaObj.GetLoadUniform(origElem.ID, ref UniformLoadNo, ref AreaName, ref LoadPatU, ref CSysU, ref DirU, ref ValueU);
            ret = mysap.AreaObj.GetLoadTemperature(origElem.ID, ref TempLoadNo, ref AreaName, ref LoadPatT, ref MyType, ref ValueT, ref PatternName);
            ret = mysap.AreaObj.GetLocalAxes(origElem.ID, ref ang, ref advanced);
           // ret = mysap.AreaObj.GetLocalAxesAdvanced(origElem.ID, ref Advanced, ref plane, ref plVect, ref cys, ref MyPlDir, ref myPoints, ref MyPlVect);
            ret = mysap.AreaObj.GetProperty(origElem.ID, ref prop);

            #endregion


            #region SetAssignments

            //createSapArea(copiedElements);

            foreach (var item in copiedElements)
            {


                #region defintions
                List<string> PtIDs = new List<string>();

                noOfPoints = item.vertices.Count;
                Xcoor = item.vertices.Select(x => x.X).ToArray();
                Ycoor = item.vertices.Select(x => x.Y).ToArray();
                Zcoor = item.vertices.Select(x => x.Z).ToArray();

                foreach (var pt in item.vertices)
                {
                    string stringId = null;
                    mysap.PointObj.AddCartesian(pt.X, pt.Y, pt.Z, ref stringId);
                    PtIDs.Add(stringId);

                }
                string[] PtsIdsArr = PtIDs.ToArray();

                ID = item.ID;
                ret = mysap.AreaObj.AddByPoint(noOfPoints, ref PtsIdsArr, ref ID, origElem.Sec, ID);
                // ret = mysap.AreaObj.AddByCoord(noOfPoints, ref Xcoor, ref Ycoor, ref Zcoor, ref ID, origElem.Sec, ID);

                item.ID = ID;




                #endregion



                // group assignments
                foreach (var g in Groups)
                {
                    ret = mysap.AreaObj.SetGroupAssign(item.ID, g, false);
                }

                // UniformtoFrameLoads
                for (int i = 0; i < UniformToFrameLoadNo; i++)
                {
                    ret = mysap.AreaObj.SetLoadUniformToFrame(item.ID, LoadPatUF[i], ValueUF[i], DirUF[i], DistTypeUF[i], false, CSysUF[i]);
                }

                // UniformLoads
                for (int i = 0; i < UniformLoadNo; i++)
                {
                    ret = mysap.AreaObj.SetLoadUniform(item.ID, LoadPatU[i], ValueU[i], DirU[i], false, CSysU[i]);
                }


                // temp loads
                for (int i = 0; i < TempLoadNo; i++)
                {
                    ret = mysap.AreaObj.SetLoadTemperature(item.ID, LoadPatT[i], MyType[i], ValueT[i]);
                }

                // LocalAxes
                ret = mysap.AreaObj.SetLocalAxes(item.ID, ang);
               // ret = mysap.AreaObj.SetLocalAxesAdvanced(item.ID, Advanced, plane, plVect, cys, ref MyPlDir, ref myPoints, ref MyPlVect);


                // property
                mysap.AreaObj.SetProperty(item.ID, prop);
            }

            #endregion


        }

        public static void EditOfExistingAreas(Area origElem, Area item)
        {

            #region def
            int numberGroups = 0;
            string[] Groups = new string[0];
            int ret;
            int UniformToFrameLoadNo = 0;
            int UniformLoadNo = 0;
            int TempLoadNo = 0;
            string[] AreaName = null;
            string[] LoadPatUF = null;
            string[] LoadPatU = null;
            string[] LoadPatT = null;

            string[] CSysUF = null;
            string[] CSysU = null;

            int[] DirUF = null;
            int[] DirU = null;

            double[] ValueUF = null;
            double[] ValueU = null;
            double[] ValueT = null;

            int[] DistTypeUF = null;

            int[] MyType = null;
            string[] PatternName = default;

            double ang = 0.0;
            bool advanced = false;

            int[] mypDir = new int[1];
            string[] myplpt = new string[1];
            double[] myplvector = new double[2];


            string[] myPoints = new string[2];
            double[] MyPlVect = new double[3];
            int[] MyPlDir = new int[2];
            string st_id = "";
            string end_id = "";
            Boolean aa = true;
            bool Advanced = true;
            int plane = 0;
            int plVect = 0;
            string cys = null;
            string prop = null;

            int noOfPoints = 0;
            double[] Xcoor = null;
            double[] Ycoor = null;
            double[] Zcoor = null;

            String ID = null;
            #endregion

            cSapModel mysap = Get_sap_mode();


            #region Get Assignments

            mysap.AreaObj.GetGroupAssign(origElem.ID, ref numberGroups, ref Groups);
            ret = mysap.AreaObj.GetLoadUniformToFrame(origElem.ID, ref UniformToFrameLoadNo, ref AreaName, ref LoadPatUF, ref CSysUF, ref DirUF, ref ValueUF, ref DistTypeUF);
            ret = mysap.AreaObj.GetLoadUniform(origElem.ID, ref UniformLoadNo, ref AreaName, ref LoadPatU, ref CSysU, ref DirU, ref ValueU);
            ret = mysap.AreaObj.GetLoadTemperature(origElem.ID, ref TempLoadNo, ref AreaName, ref LoadPatT, ref MyType, ref ValueT, ref PatternName);
            ret = mysap.AreaObj.GetLocalAxes(origElem.ID, ref ang, ref advanced);
            // ret = mysap.AreaObj.GetLocalAxesAdvanced(origElem.ID, ref Advanced, ref plane, ref plVect, ref cys, ref MyPlDir, ref myPoints, ref MyPlVect);
            ret = mysap.AreaObj.GetProperty(origElem.ID, ref prop);

            #endregion


            #region SetAssignments

            //createSapArea(copiedElements);




                #region defintions
                List<string> PtIDs = new List<string>();

                noOfPoints = item.vertices.Count;
                Xcoor = item.vertices.Select(x => x.X).ToArray();
                Ycoor = item.vertices.Select(x => x.Y).ToArray();
                Zcoor = item.vertices.Select(x => x.Z).ToArray();

                foreach (var pt in item.vertices)
                {
                    string stringId = null;
                    mysap.PointObj.AddCartesian(pt.X, pt.Y, pt.Z, ref stringId);
                    PtIDs.Add(stringId);

                }
                string[] PtsIdsArr = PtIDs.ToArray();

                ID = item.ID;

                ret = mysap.AreaObj.AddByPoint(noOfPoints, ref PtsIdsArr, ref ID, origElem.Sec, ID);
                if (ret ==1)
                {
                    return;
                }
                // ret = mysap.AreaObj.AddByCoord(noOfPoints, ref Xcoor, ref Ycoor, ref Zcoor, ref ID, origElem.Sec, ID);

                item.ID = ID;




                #endregion



                // group assignments
                foreach (var g in Groups)
                {
                    ret = mysap.AreaObj.SetGroupAssign(item.ID, g, false);
                }

                // UniformtoFrameLoads
                for (int i = 0; i < UniformToFrameLoadNo; i++)
                {
                    ret = mysap.AreaObj.SetLoadUniformToFrame(item.ID, LoadPatUF[i], ValueUF[i], DirUF[i], DistTypeUF[i], false, CSysUF[i]);
                }

                // UniformLoads
                for (int i = 0; i < UniformLoadNo; i++)
                {
                    ret = mysap.AreaObj.SetLoadUniform(item.ID, LoadPatU[i], ValueU[i], DirU[i], false, CSysU[i]);
                }


                // temp loads
                for (int i = 0; i < TempLoadNo; i++)
                {
                    ret = mysap.AreaObj.SetLoadTemperature(item.ID, LoadPatT[i], MyType[i], ValueT[i]);
                }

                // LocalAxes
                ret = mysap.AreaObj.SetLocalAxes(item.ID, ang);
                // ret = mysap.AreaObj.SetLocalAxesAdvanced(item.ID, Advanced, plane, plVect, cys, ref MyPlDir, ref myPoints, ref MyPlVect);


                // property
                mysap.AreaObj.SetProperty(item.ID, prop);
            

            #endregion


        }


        public static void copyFrameAssignments(Frame origElem, List<Frame> copiedElements)
        {

            #region def
            int numberGroups = 0;
            string[] Groups = new string[0];
            int CardinalPoint = 0;
            bool Mirror2 = false;
            bool StiffTransform = false;
            double[] Offset1 = new double[0];
            double[] Offset2 = new double[0];
            string CSys = null;
            int ret = 1;
            int myType = 0;
            double MaxSize = 0;
            int MinSect = 0;
            bool NoOut = false;
            bool NoutAndDesign = false;
            int NumberItemsF = 0;
            int NumberItemsP = 0;
            int NumberItemsT = 0;

            string[] LoadPatF = new string[0];
            string[] LoadPatP = new string[0];
            string[] LoadPatT = new string[0];

            string[] CSys2F = null;
            string[] CSys2P = null;
            string[] CSys2T = null;

            int[] DirF = new int[0];
            int[] DirP = new int[0];
            int[] DirT = new int[0];

            double[] RD1F = new double[0];
            double[] RD1P = new double[0];
            double[] RD1T = new double[0];

            double[] RD2F = new double[0];
            double[] RD2P = new double[0];
            double[] RD2T = new double[0];

            double[] Dist1F = new double[0];
            double[] Dist1P = new double[0];
            double[] Dist1T = new double[0];

            double[] Dist2F = new double[0];
            double[] Dist2P = new double[0];
            double[] Dist2T = new double[0];

            double[] Val1F = new double[0];
            double[] Val1P = new double[0];
            double[] Val1T = new double[0];

            double[] Val2F = new double[0];
            double[] Val2P = new double[0];
            double[] Val2T = new double[0];


            string[] FrameNameF = null;
            string[] FrameNameP = null;
            string[] FrameNameT = null;



            int[] MyTypeF = new int[0];
            int[] MyTypeP = new int[0];
            int[] MyTypeT = new int[0];


            double[] Val2 = new double[0];
            string[] patternName = null;


            bool[] II = null;
            bool[] JJ = null;
            double[] StartValue = null;
            double[] EndValue = null;

            string[] PointName = new string[0];
            int[] LCStep = new int[0];
            string[] CSys3 = new string[0];
            double[] F1 = new double[0];
            double[] F2 = new double[0];
            double[] F3 = new double[0];
            double[] M1 = new double[0];
            double[] M2 = new double[0];
            double[] M3 = new double[0];
            double[] Value = new double[0];

            bool[] restValues = new bool[0];

            eFramePropType e = eFramePropType.Circle;
            string[] StartSec = null;
            string[] EndSec = null;
            double[] MYLength = null;
            int[] Ei33 = null;
            int[] Ei22 = null;
            int color = 0;
            string Notes = null;
            string guid = null;
            int[] Mytype2 = null;


            string propName = null;
            string SAuto = null;

            #endregion

            cSapModel mysap = Get_sap_mode();


            #region Get Assignments
            mysap.FrameObj.GetGroupAssign(origElem.ID, ref numberGroups, ref Groups);
            mysap.FrameObj.GetInsertionPoint(origElem.ID, ref CardinalPoint, ref Mirror2, ref StiffTransform, ref Offset1, ref Offset2, ref CSys);
            mysap.FrameObj.GetOutputStations(origElem.ID, ref myType, ref MaxSize, ref MinSect, ref NoOut, ref NoutAndDesign);
            mysap.FrameObj.GetLoadDistributed(origElem.ID, ref NumberItemsF, ref FrameNameF, ref LoadPatF, ref MyTypeF,
                                              ref CSys2F, ref DirF, ref RD1F, ref RD2F, ref Dist1F, ref Dist2F, ref Val1F, ref Val2F);
            mysap.FrameObj.GetLoadPoint(origElem.ID, ref NumberItemsP, ref FrameNameP, ref LoadPatP, ref MyTypeP, ref CSys2P, ref DirP, ref RD1P, ref Dist1P, ref Val1P);
            mysap.FrameObj.GetLoadTemperature(origElem.ID, ref NumberItemsT, ref FrameNameT, ref LoadPatT, ref MyTypeT, ref Val1T, ref patternName);
            mysap.FrameObj.GetReleases(origElem.ID, ref II, ref JJ, ref StartValue, ref EndValue);
            mysap.FrameObj.GetSection(origElem.ID, ref propName, ref SAuto);
            #endregion

            #region SetAssignments
            foreach (var item in copiedElements)
            {
                string id = item.ID;
                ret = mysap.FrameObj.AddByCoord(item.StartPoint.X, item.StartPoint.Y, item.StartPoint.Z, item.EndPoint.X, item.EndPoint.Y, item.EndPoint.Z, ref id, origElem.Sec, id);
                item.ID = id;


                // group assignments
                foreach (var g in Groups)
                {
                    ret = mysap.FrameObj.SetGroupAssign(item.ID, g, false);
                }

                // insertionpoints
                ret = mysap.FrameObj.SetInsertionPoint(item.ID, CardinalPoint, Mirror2, StiffTransform, ref Offset1, ref Offset2, CSys);

                //outputstations
                ret = mysap.FrameObj.SetOutputStations(item.ID, myType, MaxSize, MinSect, NoOut, NoutAndDesign);


                // dist loads
                for (int i = 0; i < NumberItemsF; i++)
                {
                    ret = mysap.FrameObj.SetLoadDistributed(item.ID, LoadPatF[i], MyTypeF[i], DirF[i], RD1F[i], RD2F[i], Val1F[i], Val2F[i], CSys2F[i], true, false);
                }


                // points loads
                for (int i = 0; i < NumberItemsP; i++)
                {
                    ret = mysap.FrameObj.SetLoadPoint(item.ID, LoadPatP[i], MyTypeP[i], DirP[i], RD1P[i], Val1P[i], CSys2P[i], true, false);
                }


                // temp loads
                for (int i = 0; i < NumberItemsT; i++)
                {
                    ret = mysap.FrameObj.SetLoadTemperature(item.ID, LoadPatT[i], MyTypeT[i], Val1T[i], patternName[i], false);
                }

                /// local axes
                ret = mysap.FrameObj.SetLocalAxes(item.ID, origElem.Beta);




                // releases
                ret = mysap.FrameObj.SetReleases(item.ID, ref II, ref JJ, ref StartValue, ref EndValue);


                // SetProfile
                ret = mysap.FrameObj.SetSection(item.ID, propName);




            }

            #endregion


        }



        #endregion

        #region LAyersCreation

        private static void CreateFramesLayers(RhinoDoc doc)
        {



            #region CreteParentLayers

            // ParentFramesLAyer
            int ParentFramesLayerIndex;
            Layer ParentFramesLayer;
            ParentFramesLayer = doc.Layers.FindName("Frames");

            if (ParentFramesLayer == null)
            {
                ParentFramesLayer = new Layer();
                ParentFramesLayer.Name = "Frames";
                ParentFramesLayerIndex = doc.Layers.Add(ParentFramesLayer);
            }

            else
            {
                ParentFramesLayerIndex = ParentFramesLayer.Index;
            }

            ParentFramesLayer = doc.Layers.FindName("Frames");


            #endregion

            #region Create frames Layers

            RhinoObject[] rhinoCurves = doc.Objects.FindByObjectType(ObjectType.Curve).Where(x => x.Attributes.GetUserString("EtabsSection") != null).ToArray();

            #region get  list of MainGroups
            List<string> FramesMaingroups = new List<string>();
            List<string> FramesMainSections = new List<string>();

            foreach (var item in rhinoCurves)
            {

                ObjectAttributes att = item.Attributes;

                string Mgroup = att.GetUserString("MainGroup");
                string sec = att.GetUserString("EtabsSection");
                FramesMaingroups.Add(Mgroup);
                FramesMainSections.Add(sec);
            }

            FramesMaingroups = FramesMaingroups.Distinct().Where(x => x != null).ToList();
            FramesMainSections = FramesMainSections.Distinct().Where(x => x != null).ToList();
            #endregion

            foreach (var MG in FramesMaingroups)
            {

                List<string> subGroups = new List<string>();

                rhinoCurves = doc.Objects.FindByObjectType(ObjectType.Curve).Where(x => x.Attributes.GetUserString("MainGroup") == MG).ToArray();
                foreach (var item in rhinoCurves)
                {
                    ObjectAttributes att = item.Attributes;
                    string subgroup = att.GetUserString("SubGroup");
                    subGroups.Add(subgroup);

                }

                subGroups = subGroups.Distinct().ToList().Where(x => x != null).ToList();
                subGroups.Sort();

                #region Creation of FramesMAinLayer

                int MGLayerIndex;
                Layer MainGroupLayer;
                MainGroupLayer = doc.Layers.FindName(MG);

                if (MainGroupLayer != null && MainGroupLayer.ParentLayerId == ParentFramesLayer.Id)
                {
                    MGLayerIndex = MainGroupLayer.Index;

                }

                else
                {
                    MainGroupLayer = new Layer();
                    MainGroupLayer.Name = MG;
                    MainGroupLayer.ParentLayerId = ParentFramesLayer.Id;
                    MGLayerIndex = doc.Layers.Add(MainGroupLayer);
                }


                MainGroupLayer = doc.Layers.FindName(MG);

                #endregion

                int i = 0;
                foreach (var subGroup in subGroups)
                {
                    int red = i * 20 % 256;
                    int green = i * 40 % 256;
                    int blue = i * 60 % 256;

                    System.Drawing.Color color = System.Drawing.Color.FromArgb(255, red, green, blue);

                    int subLayerIndex;
                    Layer subLayer = doc.Layers.FindName(subGroup);


                    if (subLayer != null && subLayer.ParentLayerId == MainGroupLayer.Id)
                    {
                        subLayer.Name = subGroup;
                        subLayer.Color = color;
                        subLayer.ParentLayerId = MainGroupLayer.Id;
                        subLayerIndex = subLayer.Index;

                    }

                    else
                    {
                        subLayer = new Layer();
                        subLayer.Name = subGroup;
                        subLayer.Color = color;
                        subLayer.ParentLayerId = MainGroupLayer.Id;
                        subLayerIndex = doc.Layers.Add(subLayer);
                    }

                    subLayer = doc.Layers.FindName(subGroup);

                    RhinoObject[] Curvesobjects = doc.Objects.FindByUserString("SubGroup", subGroup, true).Where(x => x.ObjectType == ObjectType.Curve).ToArray();

                    foreach (var item in Curvesobjects)
                    {

                        item.Attributes.LayerIndex = subLayerIndex;
                        item.CommitChanges();
                    }

                    i++;
                }

            }

            if (FramesMaingroups.Count == 0)
            {

                int i = 0;
                foreach (var sec in FramesMainSections)
                {
                    int red = i * 20 % 256;
                    int green = i * 40 % 256;
                    int blue = i * 60 % 256;

                    System.Drawing.Color color = System.Drawing.Color.FromArgb(255, red, green, blue);

                    int subLayerIndex;
                    Layer subLayer = doc.Layers.FindName(sec);


                    if (subLayer != null)
                    {
                        subLayer.Name = sec;
                        subLayer.Color = color;
                        subLayer.ParentLayerId = ParentFramesLayer.Id;
                        subLayerIndex = subLayer.Index;

                    }

                    else
                    {
                        subLayer = new Layer();
                        subLayer.Name = sec;
                        subLayer.Color = color;
                        subLayer.ParentLayerId = ParentFramesLayer.Id;
                        subLayerIndex = doc.Layers.Add(subLayer);
                    }

                    subLayer = doc.Layers.FindName(sec);

                    RhinoObject[] CurveObjects = doc.Objects.FindByUserString("EtabsSection", sec, true).Where(x => x.ObjectType == ObjectType.Curve).ToArray();

                    foreach (var item in CurveObjects)
                    {

                        item.Attributes.LayerIndex = subLayerIndex;
                        item.CommitChanges();
                    }

                    i++;
                }

            }


            #endregion



        }

        private static void CreateAreaLayers(RhinoDoc doc)
        {
            #region CreteParentLayers

            // ParentAreasLayer
            int ParentAreasLayerIndex;
            Layer ParentAreasLayer;
            ParentAreasLayer = doc.Layers.FindName("Areas");

            if (ParentAreasLayer == null)
            {
                ParentAreasLayer = new Layer();
                ParentAreasLayer.Name = "Areas";
                ParentAreasLayerIndex = doc.Layers.Add(ParentAreasLayer);
            }

            else
            {
                ParentAreasLayerIndex = ParentAreasLayer.Index;
            }

            ParentAreasLayer = doc.Layers.FindName("Areas");

            #endregion


            #region Create Aras Layers
            RhinoObject[] rhinoBreps = doc.Objects.FindByObjectType(ObjectType.Brep).Where(x => x.Attributes.GetUserString("EtabsSection") != null).ToArray();

            #region get  list of AreaMainGroups
            List<string> AreasMaingroups = new List<string>();
            List<string> AreasSections = new List<string>();

            foreach (var item in rhinoBreps)
            {
                ObjectAttributes att = item.Attributes;

                string Mgroup = att.GetUserString("MainGroup");
                string sec = att.GetUserString("EtabsSection");
                AreasMaingroups.Add(Mgroup);
                AreasSections.Add(sec);
            }

            AreasMaingroups = AreasMaingroups.Distinct().ToList().Where(x => x != null).ToList();
            AreasSections = AreasSections.Distinct().ToList().Where(x => x != null).ToList();
            #endregion


            foreach (var MG in AreasMaingroups)
            {
                List<string> subGroups = new List<string>();

                rhinoBreps = doc.Objects.FindByObjectType(ObjectType.Brep).Where(x => x.Attributes.GetUserString("MainGroup") == MG).ToArray();
                foreach (var item in rhinoBreps)
                {
                    ObjectAttributes att = item.Attributes;
                    string subgroup = att.GetUserString("SubGroup");
                    subGroups.Add(subgroup);

                }

                subGroups = subGroups.Distinct().ToList().Where(x => x != null).ToList();
                subGroups.Sort();

                #region Creation of AreaMainLayer
                int MGLayerIndex;
                Layer MainGroupLayer;
                MainGroupLayer = doc.Layers.FindName(MG);

                if (MainGroupLayer != null && MainGroupLayer.ParentLayerId == ParentAreasLayer.Id)
                {
                    MGLayerIndex = MainGroupLayer.Index;

                }

                else
                {
                    MainGroupLayer = new Layer();
                    MainGroupLayer.Name = MG;
                    MainGroupLayer.ParentLayerId = ParentAreasLayer.Id;
                    MGLayerIndex = doc.Layers.Add(MainGroupLayer);
                }


                MainGroupLayer = doc.Layers.FindName(MG);

                #endregion

                int i = 0;
                foreach (var subGroup in subGroups)
                {
                    int red = i * 20 % 256;
                    int green = i * 40 % 256;
                    int blue = i * 60 % 256;

                    System.Drawing.Color color = System.Drawing.Color.FromArgb(255, red, green, blue);

                    int subLayerIndex;
                    Layer subLayer = doc.Layers.FindName(subGroup);


                    if (subLayer != null && subLayer.ParentLayerId == MainGroupLayer.Id)
                    {
                        subLayer.Name = subGroup;
                        subLayer.Color = color;
                        subLayer.ParentLayerId = MainGroupLayer.Id;
                        subLayerIndex = subLayer.Index;

                    }

                    else
                    {
                        subLayer = new Layer();
                        subLayer.Name = subGroup;
                        subLayer.Color = color;
                        subLayer.ParentLayerId = MainGroupLayer.Id;
                        subLayerIndex = doc.Layers.Add(subLayer);
                    }

                    subLayer = doc.Layers.FindName(subGroup);

                    RhinoObject[] AreaObjects = doc.Objects.FindByUserString("SubGroup", subGroup, true).Where(x => x.ObjectType == ObjectType.Brep).ToArray();

                    foreach (var item in AreaObjects)
                    {

                        item.Attributes.LayerIndex = subLayerIndex;
                        item.CommitChanges();
                    }

                    i++;
                }

            }

            if (AreasMaingroups.Count == 0)
            {

                int i = 0;
                foreach (var sec in AreasSections)
                {
                    int red = i * 20 % 256;
                    int green = i * 40 % 256;
                    int blue = i * 60 % 256;

                    System.Drawing.Color color = System.Drawing.Color.FromArgb(255, red, green, blue);

                    int subLayerIndex;
                    Layer subLayer = doc.Layers.FindName(sec);


                    if (subLayer != null)
                    {
                        subLayer.Name = sec;
                        subLayer.Color = color;
                        subLayer.ParentLayerId = ParentAreasLayer.Id;
                        subLayerIndex = subLayer.Index;

                    }

                    else
                    {
                        subLayer = new Layer();
                        subLayer.Name = sec;
                        subLayer.Color = color;
                        subLayer.ParentLayerId = ParentAreasLayer.Id;
                        subLayerIndex = doc.Layers.Add(subLayer);
                    }

                    subLayer = doc.Layers.FindName(sec);

                    RhinoObject[] AreaObjects = doc.Objects.FindByUserString("EtabsSection", sec, true).Where(x => x.ObjectType == ObjectType.Brep).ToArray();

                    foreach (var item in AreaObjects)
                    {

                        item.Attributes.LayerIndex = subLayerIndex;
                        item.CommitChanges();
                    }

                    i++;
                }

            }
            #endregion
        }

        private static void CreatePointsLayer(RhinoDoc doc)
        {
            #region craete Points Layer
            int ptLayerIndex;
            RhinoObject[] rhinoPoints = doc.Objects.FindByObjectType(ObjectType.Point);
            Layer ptLayer;
            ptLayer = doc.Layers.FindName("Joints");
            if (ptLayer == null)
            {

                ptLayer = new Layer();
                ptLayer.Name = "Joints";
                ptLayerIndex = doc.Layers.Add(ptLayer);
            }

            else
            {
                ptLayerIndex = ptLayer.Index;

            }


            rhinoPoints.ToList().Where(x => x.Attributes.GetUserString("ID") != null).ToList().ForEach(x =>
            {
                x.Attributes.LayerIndex = ptLayerIndex;
                x.CommitChanges();
            });

            #endregion
        }

        private static void CreateFramesLayersColumnMark(RhinoDoc doc)
        {



            #region CreteParentLayers

            // ParentFramesLAyer
            int ParentFramesLayerIndex;
            Layer ParentFramesLayer;
            ParentFramesLayer = doc.Layers.FindName("Frames");

            if (ParentFramesLayer == null)
            {
                ParentFramesLayer = new Layer();
                ParentFramesLayer.Name = "Frames";
                ParentFramesLayerIndex = doc.Layers.Add(ParentFramesLayer);
            }

            else
            {
                ParentFramesLayerIndex = ParentFramesLayer.Index;
            }

            ParentFramesLayer = doc.Layers.FindName("Frames");


            #endregion

            #region Create frames Layers

            RhinoObject[] rhinoCurves = doc.Objects.Where(x => x.Attributes.GetUserString("ColumnMark") != null && x.Attributes.GetUserString("StartPointId") != null).ToArray();

            #region get  list of MainGroups
            List<string> FramesMaingroups = new List<string>();
            List<string> FramesMainSections = new List<string>();

            foreach (var item in rhinoCurves)
            {

                ObjectAttributes att = item.Attributes;

                string Mgroup = att.GetUserString("MainGroup");
                string sec = att.GetUserString("ColumnMark");
                FramesMaingroups.Add(Mgroup);
                FramesMainSections.Add(sec);
            }

            FramesMaingroups = FramesMaingroups.Distinct().Where(x => x != null).ToList();
            FramesMainSections = FramesMainSections.Distinct().Where(x => x != null).ToList();
            #endregion

            foreach (var MG in FramesMaingroups)
            {

                List<string> subGroups = new List<string>();

                rhinoCurves = rhinoCurves.Where(x => x.Attributes.GetUserString("MainGroup") == MG).ToArray();
                foreach (var item in rhinoCurves)
                {
                    ObjectAttributes att = item.Attributes;
                    string subgroup = att.GetUserString("SubGroup");
                    subGroups.Add(subgroup);

                }

                subGroups = subGroups.Distinct().ToList().Where(x => x != null).ToList();
                subGroups.Sort();

                #region Creation of FramesMAinLayer

                int MGLayerIndex;
                Layer MainGroupLayer;
                MainGroupLayer = doc.Layers.FindName(MG);

                if (MainGroupLayer != null && MainGroupLayer.ParentLayerId == ParentFramesLayer.Id)
                {
                    MGLayerIndex = MainGroupLayer.Index;

                }

                else
                {
                    MainGroupLayer = new Layer();
                    MainGroupLayer.Name = MG;
                    MainGroupLayer.ParentLayerId = ParentFramesLayer.Id;
                    MGLayerIndex = doc.Layers.Add(MainGroupLayer);
                }


                MainGroupLayer = doc.Layers.FindName(MG);

                #endregion

                int i = 0;
                foreach (var subGroup in subGroups)
                {
                    int red = i * 20 % 256;
                    int green = i * 40 % 256;
                    int blue = i * 60 % 256;

                    System.Drawing.Color color = System.Drawing.Color.FromArgb(255, red, green, blue);

                    int subLayerIndex;
                    Layer subLayer = doc.Layers.FindName(subGroup);


                    if (subLayer != null && subLayer.ParentLayerId == MainGroupLayer.Id)
                    {
                        subLayer.Name = subGroup;
                        subLayer.Color = color;
                        subLayer.ParentLayerId = MainGroupLayer.Id;
                        subLayerIndex = subLayer.Index;

                    }

                    else
                    {
                        subLayer = new Layer();
                        subLayer.Name = subGroup;
                        subLayer.Color = color;
                        subLayer.ParentLayerId = MainGroupLayer.Id;
                        subLayerIndex = doc.Layers.Add(subLayer);
                    }

                    subLayer = doc.Layers.FindName(subGroup);

                    RhinoObject[] Curvesobjects = doc.Objects.FindByUserString("SubGroup", subGroup, true).ToArray();

                    Curvesobjects = Curvesobjects.Where(x => x.Attributes.GetUserString("StartPointId") != null).ToArray();

                    foreach (var item in Curvesobjects)
                    {

                        item.Attributes.LayerIndex = subLayerIndex;
                        item.CommitChanges();
                    }

                    i++;
                }

            }

            if (FramesMaingroups.Count == 0)
            {

                int i = 0;
                foreach (var sec in FramesMainSections)
                {
                    int red = i * 20 % 256;
                    int green = i * 40 % 256;
                    int blue = i * 60 % 256;

                    System.Drawing.Color color = System.Drawing.Color.FromArgb(255, red, green, blue);

                    int subLayerIndex;
                    Layer subLayer = doc.Layers.FindName(sec);


                    if (subLayer != null && subLayer.ParentLayerId == ParentFramesLayer.Id)
                    {
                        subLayer.Name = sec;
                        subLayer.Color = color;
                        subLayer.ParentLayerId = ParentFramesLayer.Id;
                        subLayerIndex = subLayer.Index;

                    }

                    else
                    {
                        subLayer = new Layer();
                        subLayer.Name = sec;
                        subLayer.Color = color;
                        subLayer.ParentLayerId = ParentFramesLayer.Id;
                        subLayerIndex = doc.Layers.Add(subLayer);
                    }

                    subLayer = doc.Layers.FindName(sec);

                    RhinoObject[] CurveObjects = doc.Objects.FindByUserString("ColumnMark", sec, true).ToArray();
                    CurveObjects = CurveObjects.Where(x => x.Attributes.GetUserString("StartPointId") != null).ToArray();

                    foreach (var item in CurveObjects)
                    {

                        item.Attributes.LayerIndex = subLayerIndex;
                        item.CommitChanges();
                    }

                    i++;
                }

            }


            #endregion



        }

        #endregion

        #region Utils

        static int ExtractNumber(string s)
        {
            var match = Regex.Match(s, @"C\s*(\d+)");
            return match.Success ? int.Parse(match.Groups[1].Value) : int.MaxValue;
        }

        public static List<List<object>> GroupByContinuousDesignation(List<object> items)
        {
            var groupedDesignations = new List<List<object>>();

            if (items == null || items.Count == 0)
                return groupedDesignations;

            List<object> currentGroup = new List<object>();
            string currentDesignation = ((dynamic)items[0]).Designation;

            foreach (var item in items)
            {
                var designation = ((dynamic)item).Designation;
                if (designation == currentDesignation)
                {
                    currentGroup.Add(item);
                }
                else
                {
                    groupedDesignations.Add(currentGroup);
                    currentGroup = new List<object> { item };
                    currentDesignation = designation;
                }
            }

            // Add the last group
            if (currentGroup.Count > 0)
            {
                groupedDesignations.Add(currentGroup);
            }

            return groupedDesignations;
        }

        public static (int firstNumber, int secondNumber) ExtractNumbers(string input)
        {
            // Regular expression to match the pattern
            string pattern = @"(\d+)X(\d+)-";
            Match match = Regex.Match(input, pattern);

            if (match.Success)
            {
                int firstNumber = int.Parse(match.Groups[1].Value);
                int secondNumber = int.Parse(match.Groups[2].Value);

                return (firstNumber, secondNumber);
            }

            throw new ArgumentException("Input string does not match the expected pattern");
        }

        public static string ExtractDiameter(string input)
        {
            // Regular expression to match the pattern D{Diameter}-
            string pattern = @"D(\d+)-";
            Match match = Regex.Match(input, pattern);

            if (match.Success)
            {
                // Parse the matched group to get the diameter
                string diameter = match.Groups[1].Value;
                return diameter;
            }

            throw new ArgumentException("Input string does not match the expected pattern");
        }

        public static (int Diameter, int NumberCBars, int RebarSize) ExtractRebarDesign(string input)
        {
            // Regular expression to match the pattern D{Diameter}-{NumberCBars}T{RebarSize}
            string pattern = @"D(\d+)-(\d+)T(\d+)";
            Match match = Regex.Match(input, pattern);

            if (match.Success)
            {
                // Parse the matched groups to get Diameter, NumberCBars, and RebarSize
                int diameter = int.Parse(match.Groups[1].Value);
                int numberCBars = int.Parse(match.Groups[2].Value);
                int rebarSize = int.Parse(match.Groups[3].Value);

                return (diameter, numberCBars, rebarSize);
            }

            throw new ArgumentException("Input string does not match the expected pattern");
        }

        #endregion

    }

    #region Classes
    public static class LinqExtensions
    {
        public static IEnumerable<IEnumerable<T>> GroupByContinuous<T, TKey>(this IEnumerable<T> source, Func<T, TKey> keySelector)
        {
            var group = new List<T>();
            TKey lastKey = default(TKey);
            bool first = true;
            foreach (var item in source)
            {
                var key = keySelector(item);
                if (!first && !EqualityComparer<TKey>.Default.Equals(key, lastKey))
                {
                    yield return group;
                    group = new List<T>();
                }
                group.Add(item);
                lastKey = key;
                first = false;
            }
            if (group.Count > 0)
            {
                yield return group;
            }
        }
    }




    public class CSIFLoorConnectivityDB
    {
        public string UniqueName { get; set; } = string.Empty;
        public string UniquePt1 { get; set; } = string.Empty;
        public string UniquePt2 { get; set; } = string.Empty;
        public string UniquePt3 { get; set; } = string.Empty;
        public string UniquePt4 { get; set; } = string.Empty;
        public string GUID { get; set; } = string.Empty;
        public List<string> AllPointIds { get; set; } = new List<string>();
    }

    #endregion
}
