using Autodesk.Revit.DB;
using BIM_Integration;
using Grasshopper.Kernel;
using Microsoft.Office.Interop.Excel;
using Rhino;
using Rhino.DocObjects;
using Rhino.Geometry;
using Rhino.Geometry.Collections;
using SAP2000v1;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;
using System.Runtime.InteropServices;
using static Autodesk.Revit.DB.SpecTypeId;
using static BIM_Integration.SapUtils;
using Area = BIM_Integration.Area;
using Curve = Rhino.Geometry.Curve;
using Frame = BIM_Integration.Frame;
using Point = BIM_Integration.Point;
using String = System.String;

namespace MyGrasshopperAssembly2
{
    internal class RhinoSap
    {

        #region CheckFramesDrift

        public static void CheckFramesDrift(IGH_DataAccess DA)
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

                List<DriftRecrd> recList = SapUtils.framesEndPointsDrifts(ColumnsGroupsName, windFactor, seismicFactor, LoadCaseKeyNames);

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

        #region MainMethods

        public static void RhinoToSap(IGH_DataAccess DA)
        {

            bool aa = false;
            DA.GetData(0, ref aa);
            if (aa)
            {
                // SapUtils.clearGroupAssingments();
                List<Point> assignedSapPoints = modifyPoints3();
                modifyFrames2(assignedSapPoints);
                modifyAreas2(assignedSapPoints);
            }
        }
        public static void RhinoToSapFramesOnly(IGH_DataAccess DA)
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

                rhinoFrame.SubGroup = crv.Attributes.GetUserString("SapSubGroup");
                rhinoFrame.ID = crv.Attributes.GetUserString("ID");

                rhinoFrames.Add(rhinoFrame);

            }
            #endregion

            cSapModel mySapModel = SapUtils.Get_sap_mode();
            var grouppedFrames = rhinoFrames.GroupBy(x => x.SubGroup);

            foreach (var group in grouppedFrames)
            {
                List<string> etabsIDS = group.ToList().Select(x => x.ID).ToList();
                string groupName = group.Key;
                if (!groupName.StartsWith("2"))
                {
                    groupName = "*NoAssingedGroups";
                }

                SapUtils.createGroupinSap(etabsIDS, mySapModel, groupName);


            }


        }

        public static void RhinoToSapCreateGroup(IGH_DataAccess DA)
        {

            bool aa = false;
            DA.GetData(0, ref aa);
            if (aa)
            {
                 //createGroupOfShownElementsInSap();
                SetBeamMarkFromRevitInSapThroughRhino();
            }
        }

        public static void SapToRhino(IGH_DataAccess DA)
        {
            bool aa = false;
            DA.GetData(0, ref aa);

            if (aa)
            {
                Rhino.RhinoDoc doc = Rhino.RhinoDoc.ActiveDoc;
                {


                    #region getGroupAssignments

                    List<string> SapGroups = SapUtils.getGroupsList();

                    Dictionary<string, string> groupNameAttributeNameDIc = new Dictionary<string, string>();

                    for (int a = 0; a < SapGroups.Count; a++)
                    {
                        groupNameAttributeNameDIc.Add(SapGroups[a], "Group" + a);
                    }


                    #endregion

                    #region draw frames

                    List<Column> myColumns = new List<Column>();
                    List<Beam> myBeams = new List<Beam>();

                    List<Frame> sapFrames = null;
                    sapFrames = SapUtils.getSapFrames(ref myBeams, ref myColumns);

                    List<Curve> curves = new List<Curve>();

                    cSapModel mysapModel = SapUtils.Get_sap_mode();

                    foreach (var item in sapFrames)
                    {
                        Point3d stPoint = new Point3d(item.StartPoint.X, item.StartPoint.Y, item.StartPoint.Z);
                        Point3d endPoint = new Point3d(item.EndPoint.X, item.EndPoint.Y, item.EndPoint.Z);
                        LineCurve crv = new LineCurve(stPoint, endPoint);
                        Rhino.DocObjects.Tables.ObjectTable ot = Rhino.RhinoDoc.ActiveDoc.Objects;
                        ObjectAttributes attr = new ObjectAttributes();

                        attr.SetUserString("Section", item.Sec);
                        attr.SetUserString("Shape", item.Shape);
                        attr.SetUserString("StartPointId", item.StartPoint.Id);
                        attr.SetUserString("EndPointId", item.EndPoint.Id);
                        // attr.SetUserString("Group", item.Group);
                        attr.SetUserString("Angle", item.Beta.ToString());
                        attr.SetUserString("Length", item.Length.ToString());
                        attr.SetUserString("BasePlateID", item.BasePlateID);
                        attr.SetUserString("ID", item.ID);

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
                        attr.SetUserString("DistrLoadAssignments", item.distLoadAssignments);

                        ot.AddCurve(crv, attr);
                        curves.Add(crv);

                        //ot.AddPoint(stPoint);
                        //ot.AddPoint(endPoint);


                    }

                    DA.SetDataList(0, curves);

                    #endregion

                    #region drawAreas
                    List<Area> sapAreas = SapUtils.getSapAreas();
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
                            attr.SetUserString("Section", sapArea.Sec);
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
                            attr.SetUserString("Section", sapArea.Sec);
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


                                attr.SetUserString("Section", sapArea.Sec);
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

                        bool[] restArray = null;
                        mysapModel.PointObj.GetRestraint(pt.Id, ref restArray);
                        string boolString = string.Join(",", Array.ConvertAll(restArray, b => b.ToString()));

                        Point3d rhinoPt = new Point3d(pt.X, pt.Y, pt.Z);
                        ObjectAttributes att = new ObjectAttributes();
                        att.SetUserString("ID", pt.Id);
                        att.SetUserString("Restrains", boolString);

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


        public static void RhinoToSapTranseferBeamMarksGroup(IGH_DataAccess DA)
        {


            bool aa = false;
            DA.GetData(0, ref aa);
            if (aa)
            {
                SetBeamMarkFromRevitInSapThroughRhino();
            }

        }

        #endregion

        #region SapGroupingCheck

        public static void SapGroupingCheck(IGH_DataAccess DA)
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
                sapFrames = SapUtils.getSapFrames(ref myBeams, ref myColumns);


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

                SapUtils.createGroupinSap(correctElemtns.Select(x => x.ID).ToList(), SapUtils.Get_sap_mode(), "A00FramesElementsWithCorrectGroups");

                SapUtils.createGroupinSap(notCorrectElemtns.Select(x => x.ID).ToList(), SapUtils.Get_sap_mode(), "A00FramesElementsWithNotCorrectGroups");

                #endregion

                #region CheckAreaGroupAssignmentsValidity
                List<string> correctAreasIDS = new List<string>();
                List<string> notCorrectAreasIDS = new List<string>();


                List<Area> sapAreas = SapUtils.getSapAreas();
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

                SapUtils.createGroupOfAreasinSap(correctAreasIDS, "A00AreaElementsWithCorrectGroups");
                SapUtils.createGroupOfAreasinSap(notCorrectAreasIDS, "A00AreaElementsWithNotCorrectGroups");


                #endregion
            }
        }

        #region methods
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
        #endregion

      

        #region SetPointLocalAxesBasedOnConnectedColumn
        public static void SetPointLocalAxesBasedOnConnectedColumn(IGH_DataAccess DA)
        {
            bool aa = false;
            string columnsGroupName = "";

            DA.GetData(0, ref aa);
            DA.GetData(1, ref columnsGroupName);

            if (aa)
            {

                #region definitions

                int[] array = new int[1];
                string[] array2 = new string[1];
                double[] array3 = new double[2];
                string[] PlPt = new string[2];
                double[] PlVect = new double[3];
                int[] PlDir = new int[2];
                string text = "";
                string text2 = "";
                bool flag = true;
                bool Active = true;
                int Plane = 0;
                int PlVectOpt = 0;
                string PlCSys = null;
                double num = 0.0;



                #endregion

                #region setpointAxes

                List<Column> myColumns = new List<Column>();
                List<Beam> myBeams = new List<Beam>();

                List<Frame> sapFrames = null;
                sapFrames = SapUtils.getSapFrames(ref myBeams, ref myColumns);


                sapFrames = sapFrames.Where(x => x.Groups.Contains(columnsGroupName)).ToList();

                SAP2000v1.cSapModel mysapModel = BIM_Integration.SapUtils.Get_sap_mode();


                foreach (var frame in sapFrames)
                {
                    BIM_Integration.Point stPoint = frame.StartPoint;
                    BIM_Integration.Point endpoint = frame.EndPoint;

                    double angleAdusted = 0;

                    if (frame.MemberType == MemberType.Column)
                    {
                        angleAdusted = 90;
                    }



                    #region advancedLocalAxes

                    mysapModel.PointObj.SetLocalAxes(stPoint.Id, frame.Beta + angleAdusted, 0, 0);
                    mysapModel.PointObj.SetLocalAxes(endpoint.Id, frame.Beta + angleAdusted, 0, 0);


                    // mysapModel.FrameObj.GetLocalAxesAdvanced(frame.ID, ref Active, ref Plane, ref PlVectOpt, ref PlCSys, ref PlDir, ref PlPt, ref PlVect);
                    //text = PlPt[0];
                    //text2 = PlPt[1];

                    //mysapModel.PointObj.SetLocalAxesAdvanced(stPoint.Id, Active, Plane, PlCSys, PlDir, PlPt, PlVect, Plane,1,"",);

                    #endregion
                }

                #endregion

            }
        }


        #endregion

        #region Methods
        public static List<Frame> getRhinoFrames(ref List<List<Frame>> duplicatedMembers)
        {
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
                #region GetLayerName
                int layerIndex = crv.Attributes.LayerIndex;
                Layer layer = doc.Layers[layerIndex];

                // Child layer
                string childLayerName = layer.Name;

                // Parent layer
                string parentLayerName = "";
                if (layer.ParentLayerId != Guid.Empty)
                {
                    Layer parentLayer = doc.Layers.FindId(layer.ParentLayerId);
                    if (parentLayer != null)
                        parentLayerName = parentLayer.Name;
                }

                // Build the string: Parent_Child
                string layerString = string.IsNullOrEmpty(parentLayerName)
                    ? childLayerName
                    : parentLayerName + "_" + childLayerName;

                rhinoFrame.LayerName = layerString;
                #endregion

                rhinoFrame.Sec = crv.Attributes.GetUserString("Section");
                rhinoFrame.Shape = crv.Attributes.GetUserString("Shape");
                rhinoFrame.ID = crv.Attributes.GetUserString("ID");
                string stPtId = crv.Attributes.GetUserString("StartPointId");
                string endPtId = crv.Attributes.GetUserString("EndPointId");

                rhinoFrame.distLoadAssignments = crv.Attributes.GetUserString("DistrLoadAssignments");

                rhinoFrame.Groups = new List<string>();



                double beta = 0.00;
                double.TryParse(crv.Attributes.GetUserString("Angle"), out beta);
                rhinoFrame.Beta = beta;

                double length = 0;
                double.TryParse(crv.Attributes.GetUserString("Length"), out length);
                rhinoFrame.Length = length;


                rhinoFrame.Groups = newgroups;



                ///////// very Important
                CurveObject curveObject = crv as Rhino.DocObjects.CurveObject;
                Curve curvGeom = curveObject.CurveGeometry;


                rhinoFrame.SerialNumber = (int) curveObject.RuntimeSerialNumber; // add this property in Frame class


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
                                  .Where(x => !string.IsNullOrWhiteSpace(x.ID))
                                  .GroupBy(x => x.ID)
                                  .Select(group => group.OrderBy(f => f.SerialNumber).ToList()) // sort by creation order
                                  .ToList();



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

        public static List<Point> getRhinoPoints(ref List <List <Point>> DuplicatedPoints)
        {
            List<Point> rhinoPts = new List<Point>();
            Rhino.RhinoDoc doc = RhinoDoc.ActiveDoc;

            List<RhinoObject> rhinoPoints = doc.Objects.FindByObjectType(ObjectType.Point).ToList();

            foreach (var pt in rhinoPoints)
            {

                Point rhinoPoint = new Point();
                rhinoPoint.Id = pt.Attributes.GetUserString("ID");
                string ptRestrain = pt.Attributes.GetUserString("Restrains");

                if (ptRestrain != null)
                {
                    // Split the string by comma to get individual boolean values
                    string[] boolValues = ptRestrain.Split(',');

                    // Convert string array to boolean array
                    bool[] boolArray = Array.ConvertAll(boolValues, bool.Parse);



                    rhinoPoint.Restrains = boolArray;
                }



                ////// very important 
                PointObject ptObj = pt as PointObject;
                Rhino.Geometry.Point3d ptt = ptObj.PointGeometry.Location;

                rhinoPoint.X = ptt.X;
                rhinoPoint.Y = ptt.Y;
                rhinoPoint.Z = ptt.Z;


                rhinoPoint.SerialNumber = (int) pt.RuntimeSerialNumber;

                rhinoPts.Add(rhinoPoint);
            }


            #region HandlingDuplications
            var grouppedPoints = rhinoPts
                                .Where(x => !string.IsNullOrWhiteSpace(x.Id)) // Filter out items with null, empty, or whitespace IDs
                                .GroupBy(x => x.Id)
                                .Select(group => group.OrderBy(p => p.SerialNumber).ToList()) // sort by creation order
                                .Where(group => group.Count > 1) // keep only duplicates
                                .ToList(); // final result: List<List<Point>>


            DuplicatedPoints = grouppedPoints.ToList();
            #endregion

            return rhinoPts;
        }

        public static List<Area> getRhinoAreas(ref List<List<Area>> duplicatedMembers)
        {
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

                rhinoArea.Sec = brep.Attributes.GetUserString("Section");
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



                // you will use this coordinates in case where ids of the vertices is null(new Area)
                BrepObject brPObject = brep as BrepObject;
                Brep bb = brPObject.BrepGeometry;
                BrepEdgeList edgeList = bb.Edges;
                List<Point> points = new List<Point>();

                #region OLD using edges
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


                //points =Point.RemoveDuplicatePoints(points, 0.01);

                //rhinoArea.vertices = Point.OrderPointsCounterClockwise2(points);

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

                rhinoArea.vertices = points;
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
                rhinoAreas.RemoveAll(item => group.Skip(1).Contains(item));
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


            #region Old

            //foreach (var brep in rhinoBreps)
            //{

            //    Area rhinoArea = new Area();

            //    rhinoArea.Sec = brep.Attributes.GetUserString("Section");
            //    rhinoArea.ID = brep.Attributes.GetUserString("ID");
            //    rhinoArea.Groups = new List<string>();

            //    string groups = brep.Attributes.GetUserString("Groups");

            //    double angle = 0.00;
            //    double.TryParse(brep.Attributes.GetUserString("Angle"), out angle);
            //    rhinoArea.Beta = angle;

            //    int NoOfPts = 0;
            //    int.TryParse(brep.Attributes.GetUserString("no Of Pts"), out NoOfPts);
            //    rhinoArea.NoOfPts = NoOfPts;

            //    // set groups
            //    if (groups != null)
            //    {
            //        List<string> groupAttriputesNames = brep.Attributes.GetUserString("Groups").Split(',').ToList();
            //        List<string> groupNames = new List<string>();


            //        foreach (var attriputeName in groupAttriputesNames)
            //        {
            //            string groupName = brep.Attributes.GetUserString(attriputeName);
            //            groupNames.Add(groupName);
            //        }
            //        rhinoArea.Groups = groupNames;
            //    }


            //    Rhino.DocObjects.ObjectAttributes attributes = brep.Attributes;


            //    // to be deleted
            //    // Get list of user data types attached to object
            //    NameValueCollection userDataKeys = attributes.GetUserStrings();



            //    if (rhinoArea.NoOfPts == 3)
            //    {
            //        Point pt1 = new Point();
            //        pt1.Id = brep.Attributes.GetUserString("Point1_ID");

            //        Point pt2 = new Point();
            //        pt2.Id = brep.Attributes.GetUserString("Point2_ID");

            //        Point pt3 = new Point();
            //        pt3.Id = brep.Attributes.GetUserString("Point3_ID");

            //        rhinoArea.vertices = new List<Point> { pt1, pt2, pt3 };

            //    }

            //    if (rhinoArea.NoOfPts == 4)
            //    {
            //        Point pt1 = new Point();
            //        pt1.Id = brep.Attributes.GetUserString("Point1_ID");

            //        Point pt2 = new Point();
            //        pt2.Id = brep.Attributes.GetUserString("Point2_ID");

            //        Point pt3 = new Point();
            //        pt3.Id = brep.Attributes.GetUserString("Point3_ID");

            //        Point pt4 = new Point();
            //        pt3.Id = brep.Attributes.GetUserString("Point4_ID");

            //        rhinoArea.vertices = new List<Point> { pt1, pt2, pt3, pt4 };

            //    }


            //    // you will use this coordinates in case where  ids of the vertices is null (new Area)
            //    BrepObject brPObject = brep as BrepObject;
            //    Brep bb = brPObject.BrepGeometry;

            //    BrepEdgeList edgeList = bb.Edges;


            //    List<Point> points = new List<Point>();
            //    foreach (var edge in edgeList)
            //    {
            //        Point3d pt1 = edge.PointAtStart;
            //        Point3d pt2 = edge.PointAtEnd;
            //        Point stPt = new Point(null, pt1.X, pt1.Y, pt1.Z);
            //        Point endPt = new Point(null, pt2.X, pt2.Y, pt2.Z);

            //        if (!points.Any(x => x.X == stPt.X && x.Y == stPt.Y && x.Z == stPt.Z))
            //        {
            //            points.Add(stPt);

            //        }


            //        if (!points.Any(x => x.X == endPt.X && x.Y == endPt.Y && x.Z == endPt.Z))
            //        {
            //            points.Add(endPt);

            //        }

            //    }

            //    rhinoArea.vertices = points;
            //    rhinoAreas.Add(rhinoArea);
            //}

            #endregion





            return rhinoAreas;
        }




        public static cSapModel Get_sap_mode()
        {
            /// cOAPI mySapObject = (cOAPI)System.Runtime.InteropServices.Marshal.GetActiveObject("CSI.SAP2000.API.SapObject");
            // SAP2000v1.cSapModel mysapmodel = (cSapModel) System.Runtime.InteropServices.Marshal.GetActiveObject("CSI.SAP2000.API.SapModel");

            SAP2000v1.cOAPI mySapObject2 = ComActiveObject.Get("CSI.SAP2000.API.SapObject") as SAP2000v1.cOAPI;

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


            //SAP2000v1.cHelper cHelper = new SAP2000v1.Helper();
            //cOAPI mySapObject2 = cHelper.GetObject("CSI.SAP2000.API.SapObject");
            //SAP2000v1.cSapModel mysapmodel = mySapObject2.SapModel;


        }


        // used
        public static List<Point> modifyPoints3()
        {

            cSapModel mySapModel = Get_sap_mode();

            List<List<Frame>> duplicatedMembers = null;
            List<List<Area>> duplicatedAreas = null;

            List<List<Point>> duplicatedPoints = null;

            List<Point> RhinoPoints = getRhinoPoints(ref duplicatedPoints);
            List<Area> RhinoAreas = getRhinoAreas(ref duplicatedAreas);
            List<Frame> RhinoFrames = getRhinoFrames(ref duplicatedMembers);

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


            List<string> existingPointsSapIDS = SapUtils.getSapPointIDS();

            // all points will be added as a new points and sap will handle this issue 
            List<string> ids = SapUtils.addNewPointsNew(ref RhinoPoints);


            //  SapUtils.createGroupOfPointsinSap(ids, "A00NewPoints");



            #region Handling Duplicated points 

            foreach (var dupGroup in duplicatedPoints)
            {
                copyPointAssignments(dupGroup[0], dupGroup.Skip(1).ToList());

            }
            #endregion


            List<string> SapPointIdsToBeRmoved = existingPointsSapIDS.Where(x => !RhinoPoints.Select(a => a.Id).ToList().Contains(x)).ToList();
            SapUtils.createGroupOfPointsinSap(SapPointIdsToBeRmoved, "A00PointsToBeRemoved");

            return RhinoPoints;

        }


        // used
        public static void modifyFrames2(List<Point> assignedSapPts)
        {
            List<List<Frame>> duplicatedGroups = null;
            List<Frame> RhinoFrames = getRhinoFrames(ref duplicatedGroups);

            List<Frame> newRhinoFrames = RhinoFrames.Where(x => x.ID == null).ToList();
            List<Frame> existingRhinoFrames = RhinoFrames.Where(x => x.ID != null).ToList();

            List<string> existingFramesSapIDS = SapUtils.getSapFrameIDS();
            List<string> existingPoitsSapIDS = SapUtils.getSapPointIDS();



            List<Column> columns = new List<Column>();
            List<Beam> beams = new List<Beam>();
            List<Frame> existingSapFrames = SapUtils.getSapFrames(ref beams, ref columns); ;




            // add new Frames
            List<string> newSapIDS = SapUtils.createSapFrames(newRhinoFrames);


            List<Point> usedPoints = new List<Point>();

            // adjust existingFramesWithTheNewPointsIDs
            foreach (var item in existingRhinoFrames)
            {

                Point stPoint = assignedSapPts.First(x => x.X == item.StartPoint.X && x.Y == item.StartPoint.Y && x.Z == item.StartPoint.Z);
                Point endPoint = assignedSapPts.First(x => x.X == item.EndPoint.X && x.Y == item.EndPoint.Y && x.Z == item.EndPoint.Z);
                item.StartPoint.Id = stPoint.Id;
                item.EndPoint.Id = endPoint.Id;
                usedPoints.Add(stPoint);
                usedPoints.Add(endPoint);

            }


            /// 
            // get sapIds to be removed
            //List<Frame> SapFramesToBeRmoved = existingSapFrames.Where(x => !RhinoFrames.Select(a => a.ID).ToList().Contains(x.ID)).ToList();
            //List<string> SapidsToRemoved = SapFramesToBeRmoved.Select(x => x.ID).ToList();

            //SapUtils.createGroupinSap(SapidsToRemoved, SapUtils.Get_sap_mode(), "A00FramesToBeRemoved");
            //SapUtils.deleteFrameObjects(SapidsToRemoved);



            // update frames assignments 
            SapUtils.updataAssignmentsFromRhinoToSap(existingRhinoFrames);


            // SapUtils.createGroupinSap(SapidsToRemoved, "A00SapFramesToBeRemoved");
            SapUtils.createGroupinSap(newSapIDS, "A00NewSapFrames");

            cSapModel mysapmodel = SapUtils.Get_sap_mode();



            // get sapIds to be removed
            List<Frame> SapFramesToBeRmoved = existingSapFrames.Where(x => !RhinoFrames.Select(a => a.ID).ToList().Contains(x.ID)).ToList();
            List<string> SapidsToRemoved = SapFramesToBeRmoved.Select(x => x.ID).ToList();

            SapUtils.createGroupinSap(SapidsToRemoved, SapUtils.Get_sap_mode(), "A00FramesToBeRemoved");
            SapUtils.deleteFrameObjects(SapidsToRemoved);


            #region Handling Duplicated Members 

            foreach (var dupGroup in duplicatedGroups)
            {
                copyFrameAssignments(dupGroup[0], dupGroup.Skip(1).ToList(), ref usedPoints);

            }



            List<Point> unUsedPoints = assignedSapPts.Except(usedPoints).ToList();


            foreach (var item in unUsedPoints)
            {
                //mysapmodel.PointObj.DeleteRestraint(item.Id);
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
            List<Area> existingSapAreas = SapUtils.getSapAreas();
            List<Area> SapAreasToBeRmoved = existingSapAreas.Where(x => !RhinoAreas.Select(a => a.ID).ToList().Contains(x.ID)).ToList();


            // create new Areas
            List<string> newSapAreaIDS = SapUtils.createSapArea(newRhinoAreas);


            // adjust existingAreaWithTheNewPointsIDs
            foreach (var areaElement in existingRhinoAreas)
            {
                foreach (var pt in areaElement.vertices)
                {
                    pt.Id = assignedSapPts.First(x => x.X == pt.X && x.Y == pt.Y && x.Z == pt.Z).Id;
                }
            }

            // update Areas assignments
            SapUtils.updataAssignmentsFromRhinoToSap(existingRhinoAreas);





            //deleteAreas
            SapUtils.deleteAreaObjects(SapAreasToBeRmoved.Select(x => x.ID).ToList());


            // Create Groups
            SapUtils.createGroupOfAreasinSap(SapAreasToBeRmoved.Select(x => x.ID).ToList(), "A00SapAreasToBeRemoved");
            SapUtils.createGroupOfAreasinSap(newSapAreaIDS, "A00NewSapaAreas");


            #region Handling Duplicated Members 

            foreach (var dupGroup in duplicatedAreas)
            {
                copyAreaAssignments(dupGroup[0], dupGroup.Skip(1).ToList());

            }
            #endregion

        }



        public static void copyFrameAssignments(Frame origElem, List<Frame> copiedElements, ref List<Point> usedPoints)
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
            string StartPtId = null;
            string endtPtId = null;

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

            usedPoints.Add(origElem.StartPoint);
            usedPoints.Add(origElem.EndPoint);


            #region SetAssignments
            foreach (var item in copiedElements)
            {
                string id = item.ID;
                ret = mysap.FrameObj.AddByCoord(item.StartPoint.X, item.StartPoint.Y, item.StartPoint.Z, item.EndPoint.X, item.EndPoint.Y, item.EndPoint.Z, ref id, origElem.Sec, id);
                item.ID = id;


                mysap.FrameObj.GetPoints(item.ID, ref StartPtId , ref endtPtId);
                item.StartPoint.Id = StartPtId;
                item.EndPoint.Id = endtPtId;


                
                usedPoints.Add(item.StartPoint);
                usedPoints.Add(item.EndPoint);



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
            System.Boolean aa = true;
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
            ret = mysap.AreaObj.GetLocalAxesAdvanced(origElem.ID, ref Advanced, ref plane, ref plVect, ref cys, ref MyPlDir, ref myPoints, ref MyPlVect);
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
                ret = mysap.AreaObj.SetLocalAxesAdvanced(item.ID, Advanced, plane, plVect, cys, ref MyPlDir, ref myPoints, ref MyPlVect);


                // property
                mysap.AreaObj.SetProperty(item.ID, prop);
            }

            #endregion


        }


        public static void copyPointAssignments(Point origElem, List<Point> copiedElements)
        {

            #region def
            #region def

            string[] Groups = new string[0];
            double[] Offset1 = new double[0];
            double[] Offset2 = new double[0];

            int NumberItems = 0;
            string[] FrameName = new string[0];
            string[] LoadPat = new string[0];


            int[] MyType = new int[0];
            int[] Dir = new int[0];
            double[] RD1 = new double[0];
            double[] RD2 = new double[0];
            double[] Dist1 = new double[0];
            double[] Dist2 = new double[0];
            double[] Val1 = new double[0];
            double[] Val2 = new double[0];



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

            List<Point> PointsDiaphram = new List<Point>();
            string[] constrainNames = null;

            int ret = 0;
            double a = 0.0, b = 0.0, c = 0.0;
            bool advanced = false;

            eConstraintType e = eConstraintType.Line;
            #endregion


            #endregion

            cSapModel mysap = Get_sap_mode();


            #region Get Assignments
            #region Pointload
            mysap.PointObj.GetLoadForce(origElem.Id, ref NumberItems, ref PointName, ref LoadPat, ref LCStep, ref CSys3, ref F1, ref F2, ref F3, ref M1, ref M2, ref M3);


            mysap.PointObj.GetRestraint(origElem.Id, ref restValues);
            mysap.PointObj.GetLocalAxes(origElem.Id, ref a, ref b, ref c, ref advanced);


            #endregion
            #endregion

            #region SetAssignments
            foreach (var item in copiedElements)
            {
                string id = item.Id;

                for (int i = 0; i < NumberItems; i++)
                {
                    double[] value = new double[] { F1[0], F2[0], F3[0], M1[0], M2[0], M3[0] };

                    mysap.PointObj.SetLoadForce(item.Id, LoadPat[i], ref value, false, CSys3[0]);
                }

                mysap.PointObj.SetRestraint(item.Id, ref restValues);


                mysap.PointObj.SetLocalAxes(item.Id, a, b, c);





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
            RhinoObject[] rhinoCurves = doc.Objects.FindByObjectType(ObjectType.Curve).Where(x => x.Attributes.GetUserString("Section") != null && x.Attributes.GetUserString("StartPointId") != null).ToArray();

            #region get  list of MainGroups
            List<string> FramesMaingroups = new List<string>();
            List<string> FramesMainSections = new List<string>();

            foreach (var item in rhinoCurves)
            {

                ObjectAttributes att = item.Attributes;

                string Mgroup = att.GetUserString("MainGroup");
                string sec = att.GetUserString("Section");
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

                    RhinoObject[] CurveObjects = rhinoCurves.Where(x=> x.Attributes.GetUserString("Section") == sec).ToArray();

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
            RhinoObject[] rhinoBreps = doc.Objects.FindByObjectType(ObjectType.Brep).Where(x => x.Attributes.GetUserString("Section") != null).ToArray();

            #region get  list of AreaMainGroups
            List<string> AreasMaingroups = new List<string>();
            List<string> AreasSections = new List<string>();

            foreach (var item in rhinoBreps)
            {
                ObjectAttributes att = item.Attributes;

                string Mgroup = att.GetUserString("MainGroup");
                string sec = att.GetUserString("Section");
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

                    RhinoObject[] AreaObjects = doc.Objects.FindByUserString("Section", sec, true).Where(x => x.ObjectType == ObjectType.Brep).ToArray();

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


        #endregion

        #region SetGroupInSapAndSetBeamMarkFromRevitInSapThroughRhino
        public static void createGroupOfShownElementsInSap()
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


        public static void SetBeamMarkFromRevitInSapThroughRhino()
        {
            Document revitDoc  = RhinoInside.Revit.Revit.ActiveDBDocument;

            cSapModel sapModel = Get_sap_mode();
            List<Frame> sapFrames = new List<Frame>();
            Rhino.RhinoDoc doc = RhinoDoc.ActiveDoc;

            List<RhinoObject> sapRhinoCurves = doc.Objects.FindByObjectType(ObjectType.Curve).Where(x=> x.Attributes.GetUserString("EndPointId") != null).ToList();
            foreach (var crv in sapRhinoCurves)
            {
                Frame safFrame = new Frame();

                NameValueCollection n = crv.Attributes.GetUserStrings();
                string AnalysisId = crv.Attributes.GetUserString("ID");

                try
                {
                    int revitCropsspondingID = int.Parse(crv.Attributes.GetUserString("RevitCropssID"));
                    ElementId elementId = new ElementId(revitCropsspondingID);
                    FamilyInstance Revitelement = revitDoc.GetElement(elementId) as FamilyInstance;
                    string SubGroupFromRevit = Revitelement.LookupParameter(HardCodedData.RevitSubGroup).AsString();
                    safFrame.ID = AnalysisId;
                    safFrame.SubGroup = "0000" + SubGroupFromRevit;

                    sapFrames.Add(safFrame);
                }
                catch (Exception)
                {

                    continue;
                }

                 
            }



            var sapGroups = sapFrames.GroupBy(x => x.SubGroup);

            foreach (var item in sapGroups)
            {
                string groupKey = item.Key;
                List<Frame> sapElements = item.ToList();

                 createGroupinSap(sapElements.Select(x=> x.ID).ToList(), sapModel, groupKey);

            }


        }

        #endregion



    }
}
