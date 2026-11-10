using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using BIM_Integration;
using BIM_Integration_Hydra;
using Grasshopper.Kernel;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json.Linq;
using Rhino;
using Rhino.DocObjects;
using Rhino.Geometry;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using ARUI = Autodesk.Revit.UI;
using Curve = Rhino.Geometry.Curve;
using Floor = Autodesk.Revit.DB.Floor;
using Frame = BIM_Integration.Frame;
using Line = Autodesk.Revit.DB.Line;
using MemberType = BIM_Integration.MemberType;
using ObjectType = Rhino.DocObjects.ObjectType;
using Parameter = Autodesk.Revit.DB.Parameter;
using Point = BIM_Integration.Point;

namespace MyGrasshopperAssembly2
{
    public class RhinoRevit
    {


        #region CompiledMethods

        public static void RevitToRhino(IGH_DataAccess DA)
        {

            bool aa = false;
            DA.GetData(0, ref aa);

            if (aa)
            {
                //for revit 2021 or above use(ForgeTypeId dis = UnitTypeId.Millimeters;) however for revit 2020 or lower use(ForgeTypeId dis = DisplayUnitType.DUT_MILLIMETERS;)
                ForgeTypeId dis = UnitTypeId.Millimeters;

                UIDocument revitUiDoc = RhinoInside.Revit.Revit.ActiveUIDocument;
                Document revitDoc = RhinoInside.Revit.Revit.ActiveDBDocument;
                Rhino.RhinoDoc doc = RhinoDoc.ActiveDoc;


                // Main
                #region Original

                //CreateRhinoCurvesAsSolides(dis, revitUiDoc, revitDoc, doc); 

                CreateRhinoCurves(dis, revitUiDoc, revitDoc, doc);

                //RhinoRevit.createRhinoSurfaces(dis, revitUiDoc, revitDoc, doc);
                #endregion


                #region forSepcificReason

                //RhinoRevit.drawRhinoFramesForSepcificReason(dis, revitUiDoc, revitDoc, doc);
                // RhinoRevit.createRhinoSurfaces(dis, revitUiDoc, revitDoc, doc);


                #endregion


                #region safe and Revit integration
                 //drawSafePoints();
                //PilesFromRevitToRhino();
                //MapBetweenSafeAndRevit();
                //ExtractToExcel(); 
                #endregion

            }



        }

        public static void RhinoToRevit(IGH_DataAccess DA)
        {


            bool aa = false;
            DA.GetData(0, ref aa);


            if (aa)
            {

                ForgeTypeId dis = UnitTypeId.Millimeters;

                UIDocument revitUiDoc = RhinoInside.Revit.Revit.ActiveUIDocument;
                Document revitDoc = RhinoInside.Revit.Revit.ActiveDBDocument;


                #region UpdatesRevitElements

                //updateRevitFloorsAndWalls(revitUiDoc, revitDoc, dis);
                updateRevitFrames(revitUiDoc, revitDoc, dis);

                #endregion

                #region SetAnalyisIDSForUnMappedRevitElements
               // SetAnalysisIDForMappedInRevitElements(revitDoc);
                // createGroupOfShownElementsInEtabs();
                //createGroupOfShownElementsInREvit();
                
                #endregion

            }



        }

        public static void RhinoToRevitCreateGroupOfShownElemenets(IGH_DataAccess DA)
        {


            bool aa = false;
            DA.GetData(0, ref aa);


            if (aa)
            {

                ForgeTypeId dis = UnitTypeId.Millimeters;

                UIDocument revitUiDoc = RhinoInside.Revit.Revit.ActiveUIDocument;
                Document revitDoc = RhinoInside.Revit.Revit.ActiveDBDocument;




                #region createGroupOfShownElementsInREvit
               // createGroupOfShownElementsInREvit();


                CreateOrUpdateSelectionFilterForShownElements("0SelectedFilter_Tower01");



                //RhinoToRevitFlippTheConnectionsSymbols(DA);

                //RhinoToRevitAssignParamterValue(DA);


                #endregion

            }



        }


        public static void RhinoToRevitSetAnalysisIDs(IGH_DataAccess DA)
        {


            bool aa = false;
            DA.GetData(0, ref aa);


            if (aa)
            {

                ForgeTypeId dis = UnitTypeId.Millimeters;

                UIDocument revitUiDoc = RhinoInside.Revit.Revit.ActiveUIDocument;
                Document revitDoc = RhinoInside.Revit.Revit.ActiveDBDocument;




                #region SetAnalyisIDSForUnMappedRevitElements
                //SetAnalysisIDForMappedInRevitElements(revitDoc);

                SetParameterFromRevitElementToAnother(revitDoc);

                #endregion

            }



        }


        //RhinoToRevitSetParameterInRhino
        public static void RhinoToRevitSetParameterInRhino(IGH_DataAccess DA)
        {


            bool aa = false;
            DA.GetData(0, ref aa);


            if (aa)
            {
                Rhino.RhinoDoc doc = RhinoDoc.ActiveDoc;

                #region SetBeamMarkAttributes


                List<RhinoObject> allRhinoCurves = doc.Objects.FindByObjectType(ObjectType.Curve).ToList();
                List<RhinoObject> RevitRhinoFrames = allRhinoCurves.Where(x => x.Attributes.GetUserString("Family") != null && x.Attributes.GetUserString("IsMapped") == "True").ToList();
                List<RhinoObject> EtabsRhinoFrames = allRhinoCurves.Where(x => x.Attributes.GetUserString("EndPointId") != null && x.Attributes.GetUserString("IsMapped") == "True").ToList();

                foreach (var etabsElement in EtabsRhinoFrames)
                {
                    string RevitCrossID = etabsElement.Attributes.GetUserString("RevitCropssID").ToString();

                    RhinoObject CroesspRevitRhinoFrame =  RevitRhinoFrames.FirstOrDefault(x => x.Attributes.GetUserString("ID") == RevitCrossID);

                    string beamMark = CroesspRevitRhinoFrame.Attributes.GetUserString("SapID");


                    etabsElement.Attributes.SetUserString("BeamMark", beamMark);

                }

             

                #endregion

            }



        }


        public static void RhinoToRevitFlippTheConnectionsSymbols(IGH_DataAccess DA)
        {


            bool aa = false;
            DA.GetData(0, ref aa);


            if (aa)
            {

                ForgeTypeId dis = UnitTypeId.Millimeters;

                UIDocument revitUiDoc = RhinoInside.Revit.Revit.ActiveUIDocument;
                Document revitDoc = RhinoInside.Revit.Revit.ActiveDBDocument;

                List<Frame> rhinoFrames = new List<Frame>();
                Rhino.RhinoDoc doc = RhinoDoc.ActiveDoc;

                List<RhinoObject> allRhinoCurves = doc.Objects.FindByObjectType(ObjectType.Curve).ToList();
                List<RhinoObject> revitFlippedElements = allRhinoCurves.Where(x => x.Attributes.GetUserString("Family") != null && x.Attributes.GetUserString("0IsFlipped") == "Yes").ToList();

                #region New
                using (Transaction trns = new Transaction(revitDoc, "AssignElements"))
                {
                    trns.Start();
                    foreach (var crv in revitFlippedElements)
                    {

                        NameValueCollection n = crv.Attributes.GetUserStrings();
                        string RevitID = crv.Attributes.GetUserString("ID");

                        ElementId RevitEleID = new ElementId(int.Parse(RevitID));
                        Element revitElement = revitDoc.GetElement(RevitEleID);


                        if (revitElement != null)
                        {
                            string SM2 = revitElement.LookupParameter(HardCodedData.StartM2Release).AsString();
                            string SM3 = revitElement.LookupParameter(HardCodedData.StartM3Release).AsString();

                            revitElement.LookupParameter(HardCodedData.StartM2Release).Set(revitElement.LookupParameter(HardCodedData.EndM2Release).AsString());
                            revitElement.LookupParameter(HardCodedData.StartM3Release).Set(revitElement.LookupParameter(HardCodedData.EndM3Release).AsString());

                            revitElement.LookupParameter(HardCodedData.EndM2Release).Set(SM2);
                            revitElement.LookupParameter(HardCodedData.EndM3Release).Set(SM3);
                        }



                    }

                    trns.Commit();

                }

                #endregion

            }



        }


        public static void RhinoToRevitAssignParamterValue(IGH_DataAccess DA)
        {


            bool aa = false;
            DA.GetData(0, ref aa);


            if (aa)
            {

                string ParamterToBeSetInRevit = HardCodedData.BasePlateID;
                string ParamterToBeCompiedFromRhino = "BasePlateID";



                ForgeTypeId dis = UnitTypeId.Millimeters;

                UIDocument revitUiDoc = RhinoInside.Revit.Revit.ActiveUIDocument;
                Document revitDoc = RhinoInside.Revit.Revit.ActiveDBDocument;

                List<Frame> rhinoFrames = new List<Frame>();
                Rhino.RhinoDoc doc = RhinoDoc.ActiveDoc;

                List<RhinoObject> allRhinoCurves = doc.Objects.FindByObjectType(ObjectType.Curve).ToList();
                List<RhinoObject> revitCurves = allRhinoCurves.Where(x => x.Attributes.GetUserString("Family") != null).ToList();

                #region New
                using (Transaction trns = new Transaction(revitDoc, "AssignElements"))
                {
                    try
                    {
                        trns.Start();
                        foreach (var crv in revitCurves)
                        {

                            NameValueCollection n = crv.Attributes.GetUserStrings();
                            string RevitID = crv.Attributes.GetUserString("ID");
                            string ParamterValue = crv.Attributes.GetUserString(ParamterToBeCompiedFromRhino);


                            ElementId RevitEleID = new ElementId(int.Parse(RevitID));
                            Element revitElement = revitDoc.GetElement(RevitEleID);


                            if (revitElement != null)
                            {
                                revitElement.LookupParameter(HardCodedData.BasePlateID).Set(ParamterValue);
                            }
                        }


                        trns.Commit();

                    }
                    catch (Exception)
                    {

                        throw;
                    }
                  


                }

                #endregion


            }



        }



        public static void RhinoToRhino2(IGH_DataAccess DA)
        {

            bool aa = false;
            DA.GetData(0, ref aa);


            if (aa)
            {
                List<RevitArea> rehinoAres = RhinoRevit.getRhinoAreas();
                RhinoRevit.setRhinoCurve(rehinoAres);
            }



        }

        public static void RevitToRhinoFramesOnly(IGH_DataAccess DA)
        {

            bool aa = false;
            DA.GetData(0, ref aa);


            if (aa)
            {
                // for revit 2021 or above use (ForgeTypeId dis = UnitTypeId.Millimeters;) however for revit 2020 or lower use ( ForgeTypeId dis = DisplayUnitType.DUT_MILLIMETERS;)
                ForgeTypeId dis = UnitTypeId.Millimeters;

                UIDocument revitUiDoc = RhinoInside.Revit.Revit.ActiveUIDocument;
                Document revitDoc = RhinoInside.Revit.Revit.ActiveDBDocument;
                Rhino.RhinoDoc doc = RhinoDoc.ActiveDoc;

                RhinoRevit.drawRhinoFrames(dis, revitUiDoc, revitDoc, doc);


            }



        }

        public static void RevitAndEtabsToRhino(IGH_DataAccess DA)
        {

            bool aa = false;
            DA.GetData(0, ref aa);

            if (aa)
            {

                // for revit 2021 or above use (ForgeTypeId dis = UnitTypeId.Millimeters;) however for revit 2020 or lower use ( ForgeTypeId dis = DisplayUnitType.DUT_MILLIMETERS;)
                ForgeTypeId dis = UnitTypeId.Millimeters;

                UIDocument revitUiDoc = RhinoInside.Revit.Revit.ActiveUIDocument;
                Document revitDoc = RhinoInside.Revit.Revit.ActiveDBDocument;
                Rhino.RhinoDoc doc = RhinoDoc.ActiveDoc;


                //drawSafePoints();
                //PilesFromRevitToRhino();
                //MapBetweenSafeAndRevit();
                AssignSafeDataToRevitItem(revitDoc);




                //List<Frame> RevitBeams = new List<Frame>();
                //List<Frame> RevitColumns = new List<Frame>();
                // List<Frame> RevitFrames = getRevitFrames(revitUiDoc, revitDoc, dis, ref RevitBeams, ref RevitColumns);


                //List<Frame> EtabsFrames = new List<Frame>();
                //EtabsFrames = RevitUtils.TransformGeneralElementsCoordinates(revitUiDoc, revitDoc, "Etabs", ref RevitFrames);
                //ChangeEtabsSectionNames.GetSectionsDimentions2(EtabsUtils.Get_sap_mode(), ref EtabsFrames);



                //RhinoRevit.RevitToRhinoFramesOnlyForEtabs(RevitFrames);
                //RhinoEtabs.EtabsToRhinoFramesOnlyForRevit(DA, EtabsFrames);


            }



        }


        #region Piles 

        public static void PilesMethod(IGH_DataAccess DA)
        {

            bool aa = false;
            DA.GetData(0, ref aa);

            if (aa)
            {
                #region safe and Revit integration
                //drawSafePoints();
                PilesFromRevitToRhino();

               // MapBetweenSafeAndRevit();


                //ExtractToExcel(); 
                #endregion


            }



        }



        #endregion


        #endregion

        #region Areas

        #region Attributes
        public static void setRhinoAreaAttributes(ref ObjectAttributes attr, RevitArea item, string TypeOfArea)
        {

            attr.SetUserString("Section", item.Sec.ToString());
            attr.SetUserString("ID", item.revitEleId.ToString());
            attr.SetUserString("Type", TypeOfArea);
            attr.SetUserString("Level", item.Level.ToString());

        }

        public static void getRhinoAreaAttributes(ObjectAttributes attr, ref RevitArea rhinoArea)
        {

            rhinoArea.Sec = attr.GetUserString("Section");
            rhinoArea.ID = attr.GetUserString("ID");
            rhinoArea.type = attr.GetUserString("Type");
            rhinoArea.Level = attr.GetUserString("Level");


        }
        #endregion

        public static void getRevitFloors(UIDocument uidoc, Document doc, ForgeTypeId dis, ref List<RevitArea> areas)
        {

            // Collect structural floors
            List<Floor> structuralFloors = new List<Floor>();
            List<Floor> archFloors = new List<Floor>();

            #region getSelected
            Selection sel = uidoc.Selection;

            // Retrieve the selected elements
            ICollection<ElementId> selectedIds = sel.GetElementIds();
            List<Element> selectedStructuralFloors = new List<Element>();

            // Iterate through selected elements and filter for structural floors
            foreach (ElementId id in selectedIds)
            {
                Element element = doc.GetElement(id);

                // Check if the element is a floor and has the "Structural" parameter set to true
                if (element is Floor)
                {
                    Parameter structuralParam = element.get_Parameter(BuiltInParameter.FLOOR_PARAM_IS_STRUCTURAL);
                    if (structuralParam != null && structuralParam.AsInteger() == 1)
                    {
                        selectedStructuralFloors.Add(element);
                    }
                }
            }


            FilteredElementCollector collector = new FilteredElementCollector(doc);
            ICollection<Element> revitFloors = collector.OfClass(typeof(Floor)).ToElements();

            if (selectedStructuralFloors.Count > 0)
            {
                revitFloors = selectedStructuralFloors;
            }


            #endregion



            foreach (Element elem in revitFloors)
            {

                Floor floor = elem as Floor;
                RevitArea a = new RevitArea();
                a.rowNo = 0;



                AnalyticalToPhysicalAssociationManager analyManger = AnalyticalToPhysicalAssociationManager.GetAnalyticalToPhysicalAssociationManager(doc);
                // may be needed to be handled 
                AnalyticalElement analyModel = doc.GetElement(analyManger.GetAssociatedElementId(elem.Id)) as AnalyticalElement;



                if (analyModel != null || analyModel == null)
                {
                    structuralFloors.Add(floor);

                    a.Sec = floor.FloorType.Name.ToString();
                    a.revitEleId = floor.Id;
                    a.Level = (doc.GetElement(floor.LevelId) as Level).Name;

                    double offset = floor.LookupParameter("Height Offset From Level").AsDouble();
                    a.hightOffsetFromLevel = UnitUtils.ConvertFromInternalUnits(offset, dis);

                    Sketch s = doc.GetElement(floor.SketchId) as Sketch;




                    CurveArrArray crvArray = s.Profile;
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
        public static void getRevitWalls(UIDocument uidoc, Document doc, ForgeTypeId dis, ref List<RevitArea> walls)
        {


            // Create a filtered element collector
            FilteredElementCollector collector = new FilteredElementCollector(doc);

            // Define the category for structural walls
            BuiltInCategory category = BuiltInCategory.OST_Walls;

            // Filter elements by category
            ICollection<Wall> elements = collector.OfCategory(category).WhereElementIsNotElementType().ToElements().Cast<Wall>().ToList();


            foreach (Wall elem in elements)
            {

                RevitArea a = new RevitArea();

                AnalyticalToPhysicalAssociationManager analyManger = AnalyticalToPhysicalAssociationManager.GetAnalyticalToPhysicalAssociationManager(doc);
                // may be needed to be handled 
                AnalyticalElement analyModel = doc.GetElement(analyManger.GetAssociatedElementId(elem.Id)) as AnalyticalElement;



                // if (analyModel != null)

                if (true)

                {
                    a.Sec = elem.WallType.Name.ToString();
                    a.revitEleId = elem.Id;
                    a.Level = (doc.GetElement(elem.LevelId) as Level).Name;

                    Sketch s = doc.GetElement(elem.SketchId) as Sketch;
                    if (s == null)
                    {
                        using (Transaction trns = new Transaction(doc, "startCraeteProfile"))
                        {
                            trns.Start("startCraeteProfile");
                            s = elem.CreateProfileSketch();
                            trns.Commit();

                        }

                    }

                    CurveArrArray crvArray = s.Profile;
                    List<Rhino.Geometry.Curve> rhinoCrvs = RhinoInside.Revit.Convert.Geometry.GeometryDecoder.ToCurveMany(crvArray).ToList();
                    a.crvList = rhinoCrvs;

                    walls.Add(a);

                }


            }


        }


        public static void createRhinoSurfaces(ForgeTypeId dis, UIDocument revitUiDoc, Document revitDoc, RhinoDoc doc)
        {
            #region Floors
            List<RevitArea> areas = new List<RevitArea>();

            getRevitFloors(revitUiDoc, revitDoc, dis, ref areas);

            foreach (var item in areas)
            {

                Floor f = revitDoc.GetElement(item.revitEleId) as Floor;


                Brep surface = Brep.CreatePlanarBreps(item.crvList)[0];

                Rhino.Geometry.Transform Zoffset = Rhino.Geometry.Transform.Translation(0, 0, item.hightOffsetFromLevel);
                bool t = surface.Transform(Zoffset);

                Rhino.DocObjects.Tables.ObjectTable ot = Rhino.RhinoDoc.ActiveDoc.Objects;
                ObjectAttributes attr = new ObjectAttributes();
                setRhinoAreaAttributes(ref attr, item, "Area");
                ot.AddBrep(surface, attr);

            }
            #endregion


        }

        public static List<RevitArea> getRhinoAreas()
        {
            List<RevitArea> rhinoAreas = new List<RevitArea>();
            Rhino.RhinoDoc doc = RhinoDoc.ActiveDoc;

            List<RhinoObject> rhinoSurfaces = doc.Objects.FindByObjectType(ObjectType.Brep).ToList();



            #region New

            foreach (var surface in rhinoSurfaces)
            {

                List<Curve> boundaryCurves = new List<Curve>();
                List<Curve> innerCurve2 = new List<Curve>();
                List<Curve> OuterCurve2 = new List<Curve>();


                NameValueCollection n = surface.Attributes.GetUserStrings();
                ObjectAttributes attr = surface.Attributes;
                RevitArea rhinoArea = new RevitArea();

                getRhinoAreaAttributes(attr, ref rhinoArea);

                if (attr.GetUserString("ID") == null || attr.GetUserString("ID") == "")
                {
                    // continue;
                }

                rhinoArea.rhinoID = surface.Id.ToString();

                BrepObject brPObject = surface as BrepObject;

                Brep brep = brPObject.BrepGeometry;

                if (brep.Faces.Count > 1)
                {
                    continue;
                }

                rhinoArea.bottomCenterPoints = RevitArea.GetAverageOfTwoLowestVertices(brep);

                List<Curve> brepEdges = new List<Curve>();

                List<Rhino.Geometry.Curve[]> crvList = new List<Rhino.Geometry.Curve[]>();

                if (brep != null)
                {
                    foreach (BrepEdge edge in brep.Edges)
                    {
                        Curve curve = edge.ToNurbsCurve();
                        if (curve != null)
                        {

                            crvList.Add(curve.DuplicateSegments());

                        }
                    }

                    // Get the inner boundary curves (holes)
                    foreach (BrepFace face in brep.Faces)
                    {
                        foreach (BrepLoop loop in face.Loops)
                        {
                            if (loop.LoopType == BrepLoopType.Inner)
                            {
                                innerCurve2.Add(loop.To3dCurve());
                            }

                            else if (loop.LoopType == BrepLoopType.Outer)
                            {
                                OuterCurve2.Add(loop.To3dCurve());

                            }

                        }
                    }

                    rhinoArea.outerCurves = OuterCurve2;
                    rhinoArea.innerCurves = innerCurve2;

                    try
                    {
                        rhinoArea.targetZ = OuterCurve2[0].PointAt(0).Z;

                    }
                    catch (Exception)
                    {

                        continue;
                    }
                }

                rhinoAreas.Add(rhinoArea);
            }

            #endregion

            #region Multible IDs Handling
            var groupedAreas = rhinoAreas.GroupBy(x => x.ID);

            foreach (var group in groupedAreas)
            {
                // Get the first item in the group
                var firstItem = group.FirstOrDefault();

                if (firstItem != null)
                {
                    // Set the ID property to null for all other items in the group
                    foreach (var item in group.Skip(1))
                    {
                        item.ID = null;
                    }
                }
            }

            #endregion

            return rhinoAreas;
        }

        public static List<Rhino.Geometry.Extrusion> getRhinoExtrution()
        {
            List<RevitArea> rhinoAreas = new List<RevitArea>();
            Rhino.RhinoDoc doc = RhinoDoc.ActiveDoc;

            List<RhinoObject> rhinoExtrusion = doc.Objects.FindByObjectType(ObjectType.Extrusion).ToList();

            List<Rhino.Geometry.Extrusion> AllExtrusion = new List<Rhino.Geometry.Extrusion>();

            foreach (var extrusion in rhinoExtrusion)
            {


                ExtrusionObject brPObject = extrusion as ExtrusionObject;

                Rhino.Geometry.Extrusion ex = brPObject.ExtrusionGeometry;

                AllExtrusion.Add(ex);

            }
            return AllExtrusion;


        }


        public static List<Rhino.Geometry.Brep> getRhinoBrep()
        {
            List<RevitArea> rhinoAreas = new List<RevitArea>();
            Rhino.RhinoDoc doc = RhinoDoc.ActiveDoc;

            List<RhinoObject> rhinoBreps = doc.Objects.FindByObjectType(ObjectType.Brep).ToList();

            List<Rhino.Geometry.Brep> AllBreps = new List<Rhino.Geometry.Brep>();

            foreach (var brep in rhinoBreps)
            {

                BrepObject brPObject = brep as BrepObject;
                ObjectAttributes attr = brPObject.Attributes;

                Rhino.Geometry.Brep br = brPObject.BrepGeometry;

                if (attr.GetUserString("ID") != null)
                {
                    continue;
                }
                AllBreps.Add(br);

            }
            return AllBreps;


        }


        public static void updateRevitFloorsAndWalls(UIDocument uidoc, Document revitDoc, ForgeTypeId targetUnitConvertor)
        {

            #region Definitions

            #region getRevitLevels

            List<Level> revitLvels = new FilteredElementCollector(revitDoc)
                                .OfClass(typeof(Level)).Cast<Level>().ToList();

            #endregion

            #region getFloorTypes
            List<FloorType> floorTypes = new FilteredElementCollector(revitDoc).OfClass(typeof(FloorType)).Cast<FloorType>().ToList();


            #endregion

            #region getWallTypes

            List<WallType> wallTypes = new FilteredElementCollector(revitDoc).OfClass(typeof(WallType)).Cast<WallType>().ToList();


            #endregion


            Selection selection = uidoc.Selection;
            ICollection<ElementId> selectedIds = selection.GetElementIds();

            ICollection<Element> revitFloorElements = null;
            ICollection<Wall> revitWallsElements = null;

            if (selectedIds.Count > 0)
            {
                // Filter selected elements to only include Floors
                revitFloorElements = new FilteredElementCollector(revitDoc, selectedIds)
                    .OfClass(typeof(Floor))
                    .ToElements();

                // Filter selected elements to only include Walls
                revitWallsElements = new FilteredElementCollector(revitDoc, selectedIds)
                    .OfCategory(BuiltInCategory.OST_Walls)
                    .WhereElementIsNotElementType()
                    .Cast<Wall>()
                    .ToList();
            }

            // Check if any selected elements are Floors or Walls
            if (revitFloorElements?.Count > 0 || revitWallsElements?.Count > 0)
            {
                // Work with the selected Floors and Walls
            }
            else
            {
                // Collect all Floors
                revitFloorElements = new FilteredElementCollector(revitDoc)
                    .OfClass(typeof(Floor))
                    .ToElements();

                // Collect all Walls
                revitWallsElements = new FilteredElementCollector(revitDoc)
                    .OfCategory(BuiltInCategory.OST_Walls)
                    .WhereElementIsNotElementType()
                    .Cast<Wall>()
                    .ToList();
            }

            // Proceed with processing revitFloorElements and revitWallsElements
            List<RevitArea> rhinoAreas = RhinoRevit.getRhinoAreas();


            // List<RevitArea> rhinoFloors = rhinoAreas.Where(x => x.type == "Area").ToList();
            List<RevitArea> rhinoWalls = rhinoAreas.Where(x => x.type == "Wall").ToList();

            List<RevitArea> rhinoFloors = rhinoAreas.Except(rhinoWalls).ToList();

            List<Rhino.Geometry.Extrusion> rhinoExtrusions = RhinoRevit.getRhinoExtrution();

            List<Rhino.Geometry.Brep> rhinoBreps = RhinoRevit.getRhinoBrep();


            List<string> revitIdsParametersInRhino = rhinoAreas.Select(x => x.ID).ToList();




            #endregion

            #region Floors

            foreach (var rhinoArea in rhinoFloors)
            {
                Element FloorElment = revitFloorElements.FirstOrDefault(x => x.Id.ToString() == rhinoArea.ID);

                if (FloorElment != null)
                {

                    Floor fl = FloorElment as Floor;

                    double currentLevelElev = (revitDoc.GetElement(fl.get_Parameter(BuiltInParameter.LEVEL_PARAM).AsElementId()) as Level).ProjectElevation;

                    double targetZ = UnitUtils.ConvertToInternalUnits(rhinoArea.targetZ, targetUnitConvertor) - currentLevelElev;



                    using (SketchEditScope sketchEditScope2 = new SketchEditScope(revitDoc, "Edit Sketch"))
                    {
                        if (sketchEditScope2.IsSketchEditingSupported(fl.SketchId))
                        {
                            Element sketchElement = revitDoc.GetElement(fl.SketchId);
                            Sketch sketch = sketchElement as Sketch;

                            sketchEditScope2.Start(sketch.Id);


                            using (Transaction tr = new Transaction(revitDoc, "changeToArch"))
                            {
                                tr.Start("change");

                                fl.LookupParameter("Structural").Set(0);

                                tr.Commit();


                            }





                            using (TransactionGroup trnsGroup = new TransactionGroup(revitDoc, "StartTransGroup"))
                            {

                                trnsGroup.Start();

                                using (Transaction transaction = new Transaction(revitDoc, "Modify sketch"))
                                {


                                    transaction.Start();
                                    FailureHandlingOptions failOpt = transaction.GetFailureHandlingOptions();
                                    failOpt.SetFailuresPreprocessor(new RoomWarningSwallower());
                                    transaction.SetFailureHandlingOptions(failOpt);

                                    #region Remove old SketchProfile 

                                    // Get sketch elements
                                    var sketchEleIds = sketch.GetAllElements();

                                    revitDoc.Delete(sketchEleIds.ToArray());



                                    #endregion


                                    List<Rhino.Geometry.Curve[]> crvList = new List<Rhino.Geometry.Curve[]>();


                                    var revitPlane = sketch.SketchPlane.GetPlane();
                                    var rhinoPlane = RhinoInside.Revit.Convert.Geometry.GeometryDecoder.ToPlane(revitPlane);
                                    foreach (Curve i in rhinoArea.outerCurves)
                                    {

                                        var loop = Curve.ProjectToPlane(i, rhinoPlane);
                                        var segments = loop.DuplicateSegments();
                                        crvList.Add(segments);

                                    }

                                    foreach (Curve i in rhinoArea.innerCurves)
                                    {
                                        var loop = Curve.ProjectToPlane(i, rhinoPlane);
                                        var segments = loop.DuplicateSegments();
                                        crvList.Add(segments);

                                    }

                                    IEnumerable<Rhino.Geometry.Curve> flatCrvList = crvList.SelectMany(x => x);
                                    foreach (var i in flatCrvList)
                                    {
                                        Autodesk.Revit.DB.Curve c = RhinoInside.Revit.Convert.Geometry.GeometryEncoder.ToCurve(i);
                                        revitDoc.Create.NewModelCurve(c, sketch.SketchPlane);
                                    }

                                    using (var uiApplication = new Autodesk.Revit.UI.UIApplication(revitDoc.Application))
                                    {
                                        EventHandler<ARUI.Events.DialogBoxShowingEventArgs> DialogBoxShowing = null;

                                        try
                                        {
                                            uiApplication.DialogBoxShowing += DialogBoxShowing = (sender, args) =>
                                            {
                                                if (args.DialogId == "TaskDialog_Sketch_Edits_Discarded")
                                                    args.OverrideResult(1001 /*IDYES*/);
                                            };

                                            revitDoc.Regenerate();



                                            transaction.Commit();

                                        }

                                        catch (Autodesk.Revit.Exceptions.InvalidOperationException) { return; }
                                        finally { uiApplication.DialogBoxShowing -= DialogBoxShowing; }
                                    }




                                }


                                //using (Transaction trns2 = new Transaction(revitDoc, "modifySpanDirection"))
                                //{
                                //    trns2.Start();
                                //    fl.SpanDirectionAngle = 0;
                                //    trns2.Commit();

                                //}



                                trnsGroup.Commit();

                            }


                            sketchEditScope2.Commit(new RoomWarningSwallower());




                        }





                    }




                    using (Transaction tr = new Transaction(revitDoc, "changeToSTR"))
                    {
                        tr.Start("change");

                        fl.LookupParameter("Structural").Set(1);
                        fl.LookupParameter("Height Offset From Level").Set(targetZ);

                        tr.Commit();


                    }

                    using (Transaction trx = new Transaction(revitDoc, "SetFloorType"))
                    {
                        trx.Start("SetFloorType");

                        FloorType ft = floorTypes.FirstOrDefault(x => x.Name == rhinoArea.Sec);

                        if (ft == null)
                        {
                            ft = floorTypes.FirstOrDefault();
                        }

                        fl.FloorType = ft;



                    }
                }


                // create new floor
                else
                {
                    #region creteNewTheFloor

                    #region getCurveArray
                    List<Rhino.Geometry.Curve[]> crvList = new List<Rhino.Geometry.Curve[]>();
                    List<CurveLoop> crvLoops = new List<CurveLoop>();


                    foreach (Curve i in rhinoArea.outerCurves)
                    {


                        Autodesk.Revit.DB.CurveLoop c = RhinoInside.Revit.Convert.Geometry.GeometryEncoder.ToCurveLoop(i);
                        crvLoops.Add(c);

                        var segments = i.DuplicateSegments();
                        crvList.Add(segments);

                    }

                    foreach (Curve i in rhinoArea.innerCurves)
                    {
                        Autodesk.Revit.DB.CurveLoop c = RhinoInside.Revit.Convert.Geometry.GeometryEncoder.ToCurveLoop(i);
                        crvLoops.Add(c);

                        var segments = i.DuplicateSegments();
                        crvList.Add(segments);

                    }

                    IEnumerable<Rhino.Geometry.Curve> flatCrvList = crvList.SelectMany(x => x);

                    CurveArray crvAraay = new CurveArray();
                    foreach (var i in flatCrvList)
                    {
                        Autodesk.Revit.DB.Curve c = RhinoInside.Revit.Convert.Geometry.GeometryEncoder.ToCurve(i);
                        crvAraay.Append(c);
                    }
                    #endregion

                    #region get desired Level
                    Level lvl = revitLvels.FirstOrDefault(x => x.Name == rhinoArea.Level);

                    if (lvl == null)
                    {
                        lvl = revitLvels.FirstOrDefault();
                    }
                    #endregion

                    #region getDesiredFloorType

                    FloorType floorType = floorTypes.FirstOrDefault(x => x.Name == rhinoArea.Sec);

                    if (floorType == null)
                    {
                        floorType = floorTypes.FirstOrDefault();
                    }

                    #endregion

                    using (Transaction tx = new Transaction(revitDoc, "Create Floor"))
                    {

                        double currentLevelElev = lvl.LookupParameter("Elevation").AsDouble();

                        double targetZ = UnitUtils.ConvertToInternalUnits(rhinoArea.targetZ, targetUnitConvertor) - currentLevelElev;


                        tx.Start();

                        //revitDoc.Create.NewFloor(crvAraay, floorType, lvl, true);


                        Floor f = Floor.Create(revitDoc, crvLoops, floorType.Id, lvl.Id);
                        f.LookupParameter("Height Offset From Level").Set(targetZ);
                        f.LookupParameter("Structural").Set(1);
                        tx.Commit();

                        RhinoDoc.ActiveDoc.Objects.Find(new Guid(rhinoArea.rhinoID)).Attributes.SetUserString("ID", f.Id.ToString());

                    }



                    #endregion


                }
            }


            #endregion

            #region Walls

            // Create a filtered element collector
            FilteredElementCollector Wallscollector = new FilteredElementCollector(revitDoc);
            // Define the category for structural walls
            BuiltInCategory wallCategory = BuiltInCategory.OST_Walls;

            // Filter elements by category
            ICollection<Wall> revitWalls = Wallscollector.OfCategory(wallCategory).WhereElementIsNotElementType().ToElements().Cast<Wall>().ToList();

            List<Wall> JoinedWalls = new List<Wall>();

            foreach (var rhinoArea in rhinoWalls)
            {
                Wall wallInstance = revitWalls.FirstOrDefault(x => x.Id.ToString() == rhinoArea.ID);



                if (wallInstance != null)
                {
                    using (Transaction tx = new Transaction(revitDoc, "disJointsWalls"))
                    {
                        tx.Start("dissAllow");
                        if (WallUtils.IsWallJoinAllowedAtEnd(wallInstance, 0))
                        {
                            WallUtils.DisallowWallJoinAtEnd(wallInstance, 0);
                            JoinedWalls.Add(wallInstance);
                        }
                        if (WallUtils.IsWallJoinAllowedAtEnd(wallInstance, 1))
                        {
                            WallUtils.DisallowWallJoinAtEnd(wallInstance, 1);
                            JoinedWalls.Add(wallInstance);

                        }
                        tx.Commit();
                    }

                    using (SketchEditScope sketchEditScope = new SketchEditScope(revitDoc, "Edit Sketch"))
                    {


                        if (wallInstance.SketchId == null)
                        {
                            using (Transaction trns = new Transaction(revitDoc, "createsketch"))
                            {
                                trns.Start("CreateSketch");
                                wallInstance.CreateProfileSketch();
                                trns.Commit();
                            }
                        }


                        if (sketchEditScope.IsSketchEditingSupported(wallInstance.SketchId))
                        {
                            Element sketchElement = revitDoc.GetElement(wallInstance.SketchId);
                            Sketch sketch = sketchElement as Sketch;
                            sketchEditScope.Start(sketch.Id);


                            using (TransactionGroup trnsGroup = new TransactionGroup(revitDoc, "StartTransGroup"))
                            {

                                trnsGroup.Start();

                                using (Transaction transaction = new Transaction(revitDoc, "Modify sketch"))
                                {
                                    transaction.Start();

                                    #region Remove old SketchProfile 

                                    var sketchEleIds = sketch.GetAllElements();

                                    sketch.SketchPlane.GetPlane();
                                    List<ElementId> modelLines = new List<ElementId>();
                                    foreach (var item in sketchEleIds)
                                    {
                                        Element ele = revitDoc.GetElement(item);

                                        if (ele.Name == "Model Lines")
                                        {
                                            modelLines.Add(item);

                                        }
                                    }

                                    revitDoc.Delete(modelLines);

                                    #endregion


                                    List<Rhino.Geometry.Curve[]> crvList = new List<Rhino.Geometry.Curve[]>();

                                    Autodesk.Revit.DB.CurveArray crvArray = new CurveArray();

                                    var revitPlane = sketch.SketchPlane.GetPlane();
                                    var rhinoPlane = RhinoInside.Revit.Convert.Geometry.GeometryDecoder.ToPlane(revitPlane);



                                    foreach (Curve i in rhinoArea.outerCurves)
                                    {

                                        var loop = Curve.ProjectToPlane(i, rhinoPlane);
                                        var segments = loop.DuplicateSegments();
                                        crvList.Add(segments);

                                    }



                                    foreach (Curve i in rhinoArea.innerCurves)
                                    {
                                        var loop = Curve.ProjectToPlane(i, rhinoPlane);
                                        var segments = loop.DuplicateSegments();
                                        crvList.Add(segments);

                                    }

                                    IEnumerable<Rhino.Geometry.Curve> flatCrvList = crvList.SelectMany(x => x);



                                    //foreach (var i in flatCrvList)
                                    //{
                                    //    Autodesk.Revit.DB.Curve c = RhinoInside.Revit.Convert.Geometry.GeometryEncoder.ToCurve(i);
                                    //    crvArray.Append(c); 
                                    //}
                                    //revitDoc.Create.NewModelCurveArray(crvArray, sketch.SketchPlane);


                                    foreach (var i in flatCrvList)
                                    {


                                        Autodesk.Revit.DB.Curve c = null;

                                        c = RhinoInside.Revit.Convert.Geometry.GeometryEncoder.ToCurve(i);


                                        //if (i.IsArc())
                                        //{
                                        //    c = RhinoInside.Revit.Convert.Geometry.GeometryEncoder.ToArc(i);
                                        //}

                                        //else if (true)
                                        //{
                                        //    c = RhinoInside.Revit.Convert.Geometry.GeometryEncoder.ToCurve(i);


                                        //}



                                        revitDoc.Create.NewModelCurve(c, sketch.SketchPlane);

                                    }


                                    using (var uiApplication = new Autodesk.Revit.UI.UIApplication(revitDoc.Application))
                                    {
                                        EventHandler<ARUI.Events.DialogBoxShowingEventArgs> DialogBoxShowing = null;

                                        try
                                        {
                                            uiApplication.DialogBoxShowing += DialogBoxShowing = (sender, args) =>
                                            {
                                                if (args.DialogId == "TaskDialog_Sketch_Edits_Discarded")
                                                    args.OverrideResult(1001 /*IDYES*/);
                                            };

                                            revitDoc.Regenerate();



                                            transaction.Commit();
                                        }

                                        catch (Autodesk.Revit.Exceptions.InvalidOperationException) { return; }
                                        finally { uiApplication.DialogBoxShowing -= DialogBoxShowing; }
                                    }




                                }


                                trnsGroup.Commit();

                            }


                            sketchEditScope.Commit(new RoomWarningSwallower());




                        }





                    }

                    using (Transaction trx = new Transaction(revitDoc, "SetWallType"))
                    {
                        trx.Start("SetWallType");

                        WallType wt = wallTypes.FirstOrDefault(x => x.Name == rhinoArea.Sec);

                        if (wt == null)
                        {
                            wt = wallTypes.FirstOrDefault();
                        }

                        wallInstance.WallType = wt;



                        //LocationCurve a = (wallInstance.Location as LocationCurve);

                        //  (wallInstance.Location as LocationCurve).Move(RhinoInside.Revit.Convert.Geometry.GeometryEncoder.ToXYZ( rhinoArea.bottomCenterPoints)- a.Curve.Evaluate(0.5, true));


                        trx.Commit();


                    }




                }

                // create new wall
                else
                {
                    #region getCurveArray
                    List<Rhino.Geometry.Curve[]> crvList = new List<Rhino.Geometry.Curve[]>();

                    foreach (Curve i in rhinoArea.outerCurves)
                    {

                        var segments = i.DuplicateSegments();
                        crvList.Add(segments);

                    }

                    foreach (Curve i in rhinoArea.innerCurves)
                    {
                        var segments = i.DuplicateSegments();
                        crvList.Add(segments);

                    }

                    IEnumerable<Rhino.Geometry.Curve> flatCrvList = crvList.SelectMany(x => x);

                    List<Autodesk.Revit.DB.Curve> curveList = new List<Autodesk.Revit.DB.Curve>();

                    //CurveArray crvAraay = new CurveArray();
                    foreach (var i in flatCrvList)
                    {
                        Autodesk.Revit.DB.Curve c = RhinoInside.Revit.Convert.Geometry.GeometryEncoder.ToCurve(i);
                        curveList.Add(c);
                    }
                    #endregion

                    #region get desired Level
                    Level lvl = revitLvels.FirstOrDefault(x => x.Name == rhinoArea.Level);

                    if (lvl == null)
                    {
                        lvl = revitLvels.FirstOrDefault();
                    }
                    #endregion

                    #region getDesiredFloorType

                    WallType wt = wallTypes.FirstOrDefault(x => x.Name == rhinoArea.Sec);

                    if (wt == null)
                    {
                        wt = wallTypes.FirstOrDefault();
                    }


                    #endregion

                    using (Transaction tx = new Transaction(revitDoc, "Create Floor"))
                    {
                        tx.Start();

                        Wall w = Wall.Create(revitDoc, curveList, wt.Id, lvl.Id, true);

                        RhinoDoc.ActiveDoc.Objects.Find(new Guid(rhinoArea.rhinoID)).Attributes.SetUserString("ID", w.Id.ToString());

                        tx.Commit();
                    }
                }
            }

            foreach (var wall in JoinedWalls)
            {
                using (Transaction tx = new Transaction(revitDoc, "allowJoints"))
                {
                    tx.Start("allow");
                    WallUtils.AllowWallJoinAtEnd(wall, 0);
                    WallUtils.AllowWallJoinAtEnd(wall, 1);
                    tx.Commit();
                }

            }



            #endregion

            #region Extrusions

            foreach (var extr in rhinoExtrusions)
            {
                Solid s = RhinoInside.Revit.Convert.Geometry.GeometryEncoder.ToSolid(extr);



                using (Transaction t = new Transaction(revitDoc, "Create sphere direct shape"))
                {
                    t.Start();

                    FailureHandlingOptions failOpt = t.GetFailureHandlingOptions();
                    failOpt.SetFailuresPreprocessor(new RoomWarningSwallower());
                    t.SetFailureHandlingOptions(failOpt);

                    DirectShapeLibrary directShapeLibrary = DirectShapeLibrary.GetDirectShapeLibrary(revitDoc);
                    DirectShapeType directShapeType = DirectShapeType.Create(revitDoc, "Tested", new ElementId(BuiltInCategory.OST_GenericModel));
                    directShapeType.SetShape(new List<GeometryObject>() { s });
                    directShapeLibrary.AddDefinitionType("Tested", directShapeType.Id);
                    DirectShape ds = DirectShape.CreateElementInstance(revitDoc, directShapeType.Id, directShapeType.Category.Id, "Tested", Autodesk.Revit.DB.Transform.Identity);
                    ds.SetTypeId(directShapeType.Id);
                    ds.ApplicationId = "Application id";
                    ds.ApplicationDataId = "Geometry object id";
                    ds.SetShape(new GeometryObject[] { s });
                    t.Commit();
                }
            }


            #endregion

            #region Breps

            foreach (var brep in rhinoBreps)
            {
                Solid s = RhinoInside.Revit.Convert.Geometry.GeometryEncoder.ToSolid(brep);

                using (Transaction t = new Transaction(revitDoc, "Create sphere direct shape"))
                {
                    t.Start();
                    FailureHandlingOptions failOpt = t.GetFailureHandlingOptions();
                    failOpt.SetFailuresPreprocessor(new RoomWarningSwallower());
                    t.SetFailureHandlingOptions(failOpt);

                    DirectShapeLibrary directShapeLibrary = DirectShapeLibrary.GetDirectShapeLibrary(revitDoc);
                    DirectShapeType directShapeType = DirectShapeType.Create(revitDoc, "Tested2", new ElementId(BuiltInCategory.OST_GenericModel));
                    directShapeType.SetShape(new List<GeometryObject>() { s });
                    directShapeLibrary.AddDefinitionType("Tested2", directShapeType.Id);
                    DirectShape ds = DirectShape.CreateElementInstance(revitDoc, directShapeType.Id, directShapeType.Category.Id, "Tested2", Autodesk.Revit.DB.Transform.Identity);
                    ds.SetTypeId(directShapeType.Id);
                    ds.ApplicationId = "Application id2";
                    ds.ApplicationDataId = "Geometry object id2";
                    ds.SetShape(new GeometryObject[] { s });
                    t.Commit();
                }
            }


            #endregion

            #region removeRevitAres not Exist in Rhino
            //List<string> elementsToBeRemoved = revitFloorElements.Select(x => x.Id.ToString()).ToList().Except(revitIdsParametersInRhino).ToList();


            //foreach (var item in revitFloorElements.Where(x => elementsToBeRemoved.Contains(x.Id.ToString())))
            //{

            //    using (Transaction tr = new Transaction(revitDoc, "delete"))
            //    {
            //        tr.Start("delete");
            //        revitDoc.Delete(item.Id);
            //        tr.Commit();

            //    }
            //}
            #endregion
            #region removeRevitWalls not Exist in Rhino
            //List<string> WallsToBeRemoved = revitWallsElements.Select(x => x.Id.ToString()).ToList().Except(revitIdsParametersInRhino).ToList();


            //foreach (var item in revitWallsElements.Where(x => WallsToBeRemoved.Contains(x.Id.ToString())))
            //{

            //    using (Transaction tr = new Transaction(revitDoc, "delete"))
            //    {
            //        tr.Start("delete");
            //        revitDoc.Delete(item.Id);
            //        tr.Commit();

            //    }
            //}
            #endregion

        }


        #endregion

        #region Frames
        public static List<Frame> getRevitFrames(UIDocument uidoc, Document doc, ForgeTypeId dis, ref List<Frame> Beams, ref List<Frame> Columns)
        {

            using (Transaction allTrns = new Transaction(doc, "aaa"))
            {
                allTrns.Start();
                FailureHandlingOptions failOpt = allTrns.GetFailureHandlingOptions();
                failOpt.SetFailuresPreprocessor(new RoomWarningSwallower());
                allTrns.SetFailureHandlingOptions(failOpt);

                #region get Selected Elements 
    //            ICollection<ElementId> selectedIds = uidoc.Selection.GetElementIds();

    //            // Create a filter for Structural Framing and Structural Beams categories
    //            List<BuiltInCategory> categories = new List<BuiltInCategory>();
    //            categories.Add(BuiltInCategory.OST_StructuralFraming);
    //            categories.Add(BuiltInCategory.OST_StructuralColumns);


    //            // Create an ElementCategoryFilter for each category and combine them with a LogicalOrFilter
    //            List<ElementFilter> categoryFilters = new List<ElementFilter>();
    //            foreach (BuiltInCategory category in categories)
    //            {
    //                categoryFilters.Add(new ElementCategoryFilter(category));
    //            }
    //            //
    //            LogicalOrFilter orFilter = new LogicalOrFilter(categoryFilters);

    //            // Get the selected elements that match the filter
    //            List<Element> selectedElements = new List<Element>();
    //            foreach (ElementId id in selectedIds)
    //            {
    //                Element element = uidoc.Document.GetElement(id);


    //                FilteredElementCollector tagCollector = new FilteredElementCollector(doc)
    //.OfCategory(BuiltInCategory.OST_Tags);



    //                if (orFilter.PassesFilter(element))
    //                {
    //                    selectedElements.Add(element);
    //                }
    //            }


    //            List<FamilyInstance> selectedStrucralFraming = selectedElements.Where(x => x.Category.Id.IntegerValue == (int)BuiltInCategory.OST_StructuralFraming).Cast<FamilyInstance>().ToList();
    //            List<FamilyInstance> selectedStrucralColumns = selectedElements.Where(x => x.Category.Id.IntegerValue == (int)BuiltInCategory.OST_StructuralColumns).Cast<FamilyInstance>().ToList();


                #endregion

                #region get Selected Elements including the groups
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

                    if (element is Autodesk.Revit.DB.Group group) // Handle the case where the element is a group
                    {
                        // Iterate through the elements inside the group and apply the filter
                        foreach (ElementId groupMemberId in group.GetMemberIds())
                        {
                            Element groupMember = uidoc.Document.GetElement(groupMemberId);
                            if (orFilter.PassesFilter(groupMember))
                            {
                                selectedElements.Add(groupMember);
                            }
                        }
                    }
                    else if (orFilter.PassesFilter(element)) // Apply the filter normally for non-group elements
                    {
                        selectedElements.Add(element);
                    }
                }

                // Now you can filter the selected elements based on their categories
                List<FamilyInstance> selectedStrucralFraming = selectedElements
                    .Where(x => x.Category.Id.IntegerValue == (int)BuiltInCategory.OST_StructuralFraming)
                    .Cast<FamilyInstance>().ToList();

                List<FamilyInstance> selectedStrucralColumns = selectedElements
                    .Where(x => x.Category.Id.IntegerValue == (int)BuiltInCategory.OST_StructuralColumns)
                    .Cast<FamilyInstance>().ToList();

                #endregion


                List<Frame> beamFarames = new List<Frame>();
                List<Frame> columnFarames = new List<Frame>();
                List<Frame> frames = new List<Frame>();

                //List<FamilyInstance> Allbeams = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralFraming).WhereElementIsNotElementType().Cast<FamilyInstance>().ToList();
                //List<FamilyInstance> Allcolumns = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralColumns).WhereElementIsNotElementType().Cast<FamilyInstance>().ToList();

                List<FamilyInstance> Allbeams = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_StructuralFraming)
                    .WhereElementIsNotElementType()
                    .Where(x => !(x is DirectShape))  // Filter out DirectShape elements
                    .Cast<FamilyInstance>()
                    .ToList();

                // Get all columns, avoiding DirectShape elements
                List<FamilyInstance> Allcolumns = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_StructuralColumns)
                    .WhereElementIsNotElementType()
                    .Where(x => !(x is DirectShape))  // Filter out DirectShape elements
                    .Cast<FamilyInstance>()
                    .ToList();


                if (selectedStrucralFraming.Count > 0 || selectedStrucralColumns.Count > 0)
                {
                    Allbeams = selectedStrucralFraming;
                    Allcolumns = selectedStrucralColumns;
                }



                #region get data of beams


                #region handling curvedBeams (not implemented yet)
                List<FamilyInstance> strightBeams = new List<FamilyInstance>();
                List<FamilyInstance> curvedBeams = new List<FamilyInstance>();
                List<FamilyInstance> StrightElementsSplittedFromCurved = new List<FamilyInstance>();


                // union between  straightBeams and StraightBeamsFromSplittedCurvedBeams
                strightBeams = Allbeams;


                #endregion

                #region get Data Of Beams after handling curvedBeams

                foreach (var instance in strightBeams.AsParallel())
                {
                    try
                    {
                        Frame frame = new Frame();


                        //AnalyticalToPhysicalAssociationManager analyManger = AnalyticalToPhysicalAssociationManager.GetAnalyticalToPhysicalAssociationManager(doc);
                        //AnalyticalElement analytInstance = doc.GetElement(analyManger.GetAssociatedElementId(instance.Id)) as AnalyticalElement;

                        //if (analytInstance ==null)
                        //{
                        //    Autodesk.Revit.DB.Curve c = (instance.Location as LocationCurve).Curve;
                        //    XYZ fstPt = c.GetEndPoint(0);
                        //    XYZ sntPt = c.GetEndPoint(1);

                        //    //double zOffsetValue = instance.get_Parameter(BuiltInParameter.Z_OFFSET_VALUE).AsDouble();

                        //    double zOffsetValue = 0.0;


                        //    XYZ newFstPt = new XYZ(fstPt.X, fstPt.Y, fstPt.Z + zOffsetValue);
                        //    XYZ newsntPt = new XYZ(sntPt.X, sntPt.Y, sntPt.Z + zOffsetValue);
                        //    Line l = Line.CreateBound(newFstPt, newsntPt);
                        //    AnalyticalMember analyMember = AnalyticalMember.Create(doc, l);
                        //    analyManger.AddAssociation(analyMember.Id, instance.Id);
                        //}


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

                        frame.Length = UnitUtils.ConvertFromInternalUnits(frame.revitCurve.Length, dis);


                        frame.Sec = instance.Symbol.Name;


                        frame.RvtID = instance.Id.ToString();
                        frame.revitEleId = instance.Id;
                        frame.ID = instance.Id.ToString();

                        try
                        {
                           // frame.SapID = instance.LookupParameter("ACM_SE_SBM_SapID").AsString();

                            //frame.SapID = instance.LookupParameter("ACM_SE_Zone_Level_ID").AsString()+ "-"+ instance.LookupParameter("ACM_SE_Main Type ID").AsString();

                            frame.SapID = instance.LookupParameter("ACM_SE_SBM_SapID").AsString();

                        }
                        catch (Exception)
                        {

                            
                        }

                        frame.Beta = instance.LookupParameter("Cross-Section Rotation").AsDouble() * 180 / (Math.PI);

                        try
                        {
                            frame.MainGroup = instance.LookupParameter("ACM_SE_Zone_Level_ID").AsString();
                            frame.SubGroup = instance.LookupParameter("BEAM TYPE").AsString();
                        }
                        catch (Exception)
                        {

                            goto AA;
                        }



                    AA:
                        frame.RvtFamily = instance.Symbol.FamilyName;

                        frame.RvtName = instance.Symbol.Name;
                        frame.MemberType = MemberType.Beam;

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


                        // AnalyticalToPhysicalAssociationManager analyManger = AnalyticalToPhysicalAssociationManager.GetAnalyticalToPhysicalAssociationManager(doc);
                        //  AnalyticalElement analytInstance = doc.GetElement(analyManger.GetAssociatedElementId(instance.Id)) as AnalyticalElement;


                        #region OLD
                        //// verticalColumn
                        //LocationPoint locationPoint = instance.Location as LocationPoint;
                        //if (locationPoint != null)
                        //{
                        //    XYZ columnPoint = locationPoint.Point;

                        //    Level baseLevel = (doc.GetElement(instance.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_PARAM).AsElementId()) as Level);
                        //    double baseLevelElev = baseLevel.ProjectElevation;
                        //    double baseLevelOffset = instance.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_OFFSET_PARAM).AsDouble();
                        //    double startptElev = baseLevelOffset + baseLevelElev;
                        //    XYZ startPoint = new XYZ(columnPoint.X, columnPoint.Y, startptElev);

                        //    Level TopLevel = (doc.GetElement(instance.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_PARAM).AsElementId()) as Level);
                        //    double TopLevelElev = TopLevel.ProjectElevation;
                        //    double topLevelOffset = instance.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_OFFSET_PARAM).AsDouble();
                        //    double endPtElev = topLevelOffset + TopLevelElev;
                        //    XYZ endPoint = new XYZ(columnPoint.X, columnPoint.Y, endPtElev);

                        //    Line c = Line.CreateBound(startPoint, endPoint);

                        //    AnalyticalMember analyMember = AnalyticalMember.Create(doc, c as Autodesk.Revit.DB.Curve);

                        //    frame.revitCurve = analyMember.GetCurve();

                        //    //analyManger.AddAssociation(analyMember.Id, instance.Id);
                        //    analytInstance = analyMember;
                        //}


                        //// slantted Column
                        //else
                        //{
                        //    Autodesk.Revit.DB.Curve locationCurve = (instance.Location as LocationCurve).Curve;

                        //    AnalyticalMember analyMember = AnalyticalMember.Create(doc, locationCurve as Autodesk.Revit.DB.Curve);

                        //    frame.revitCurve = analyMember.GetCurve();

                        //    //analyManger.AddAssociation(analyMember.Id, instance.Id);

                        //    analytInstance = analyMember;
                        //}

                        #endregion

                        #region NEw

                        Parameter colStyleParam = instance.LookupParameter("Column Style");
                        colStyleParam.Set((int)SlantedOrVerticalColumnType.CT_EndPoint);

                        Autodesk.Revit.DB.Curve locationCurve = (instance.Location as LocationCurve).Curve;

                        // AnalyticalMember analyMember = AnalyticalMember.Create(doc, locationCurve as Autodesk.Revit.DB.Curve);

                        //frame.revitCurve = analyMember.GetCurve();

                        frame.revitCurve = locationCurve;

                        //analyManger.AddAssociation(analyMember.Id, instance.Id);

                        // analytInstance = analyMember;


                        #endregion



                        if (true)
                        {
                            #region gePoints

                            frame.RhinoCurve = RhinoInside.Revit.Convert.Geometry.GeometryDecoder.ToCurve(frame.revitCurve);

                            Point3d stPt = RhinoInside.Revit.Convert.Geometry.GeometryDecoder.ToPoint3d(frame.revitCurve.GetEndPoint(0));
                            Point3d endPt = RhinoInside.Revit.Convert.Geometry.GeometryDecoder.ToPoint3d(frame.revitCurve.GetEndPoint(1));




                            Point stPoint = new Point(stPt);
                            Point endPoint = new Point(endPt);

                            Point mid = Point.getMidPoint(stPoint, endPoint);

                            frame.StartPoint = stPoint;
                            frame.EndPoint = endPoint;
                            frame.MidPoint = mid;

                            frame.Length = Point.getDistance(stPoint, endPoint);



                            // another Approatch
                            // Point startPoint = new Point();
                            // Point endPoint = new Point();

                            // GetTopAndBottomPoints(doc, instance.Id, ref startPoint, ref endPoint);

                            // XYZ revStPt = new XYZ(startPoint.X, startPoint.Y, startPoint.Z);
                            // XYZ revEndPt = new XYZ(endPoint.X, endPoint.Y, endPoint.Z);

                            //Line ln=  Line.CreateBound(revStPt, revEndPt);
                            // frame.RhinoCurve = RhinoInside.Revit.Convert.Geometry.GeometryDecoder.ToCurve(ln);

                            // Point mid = Point.getMidPoint(startPoint, endPoint);

                            // frame.StartPoint = startPoint;
                            // frame.EndPoint = endPoint;
                            // frame.MidPoint = mid;

                            // frame.Length = Point.getDistance(startPoint, endPoint);


                            #endregion

                            frame.RvtID = instance.Id.ToString();
                            frame.revitEleId = instance.Id;
                            frame.ID = instance.Id.ToString();

                            frame.Sec = instance.Symbol.Name;
                            frame.RvtFamily = instance.Symbol.FamilyName;
                            frame.RvtName = instance.Symbol.Name;


                            frame.MemberType = MemberType.Column;


                            try
                            {
                                frame.MainGroup = instance.LookupParameter("ACM_SE_Zone_Level_ID").AsString();
                                frame.SubGroup = instance.LookupParameter("COLUMN TYPE").AsString();
                            }
                            catch (Exception ex)
                            {

                                goto BB;

                            }


                        BB:

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
                                // frame.Beta = (angle * 180 / Math.PI) + 90;
                                frame.Beta = (angle * 180 / Math.PI);

                            }

                            #endregion

                            //Autodesk.Revit.DB.Structure.StructuralFramingUtils.DisallowJoinAtEnd(instance, 0);
                            //Autodesk.Revit.DB.Structure.StructuralFramingUtils.DisallowJoinAtEnd(instance, 1);


                            columnFarames.Add(frame);
                        }


                        colStyleParam.Set((int)SlantedOrVerticalColumnType.CT_Vertical);

                    }

                    catch (Exception ex2)
                    {

                        System.Diagnostics.StackTrace trace = new System.Diagnostics.StackTrace(ex2, true);
                        RevitUtils.Log.Add("Error in getRevitFrames Method in the Columns part... Columninstance no" + instance.Id + ex2.Message + trace.GetFrame(trace.FrameCount - 1).GetFileLineNumber());

                        // TaskDialog.Show("ErrorInGetRevitFramesMethod", "ErrorInRevitBeamID= "+ instance.Id + ex2.Message + trace.GetFrame(trace.FrameCount - 1).GetFileLineNumber());
                        continue;
                    }

                }





                #endregion


                frames = beamFarames.Union(columnFarames).ToList();

                Beams = beamFarames;
                Columns = columnFarames;

                allTrns.Commit();

                return frames;
            }

        }


        public static List<Frame> getRevitFramesForRevit22(UIDocument uidoc, Document doc, ForgeTypeId dis, ref List<Frame> Beams, ref List<Frame> Columns)
        {

            using (Transaction allTrns = new Transaction(doc, "aaa"))
            {
                allTrns.Start();
                FailureHandlingOptions failOpt = allTrns.GetFailureHandlingOptions();
                failOpt.SetFailuresPreprocessor(new RoomWarningSwallower());
                allTrns.SetFailureHandlingOptions(failOpt);

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
                //
                LogicalOrFilter orFilter = new LogicalOrFilter(categoryFilters);

                // Get the selected elements that match the filter
                List<Element> selectedElements = new List<Element>();
                foreach (ElementId id in selectedIds)
                {
                    Element element = uidoc.Document.GetElement(id);


                    FilteredElementCollector tagCollector = new FilteredElementCollector(doc)
    .OfCategory(BuiltInCategory.OST_Tags);



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

                // Get all beams, avoiding DirectShape elements
                List<FamilyInstance> Allbeams = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_StructuralFraming)
                    .WhereElementIsNotElementType()
                    .Where(x => !(x is DirectShape))  // Filter out DirectShape elements
                    .Cast<FamilyInstance>()
                    .ToList();

                // Get all columns, avoiding DirectShape elements
                List<FamilyInstance> Allcolumns = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_StructuralColumns)
                    .WhereElementIsNotElementType()
                    .Where(x => !(x is DirectShape))  // Filter out DirectShape elements
                    .Cast<FamilyInstance>()
                    .ToList();

                if (selectedStrucralFraming.Count > 0 || selectedStrucralColumns.Count > 0)
                {
                    Allbeams = selectedStrucralFraming;
                    Allcolumns = selectedStrucralColumns;
                }



                #region get data of beams


                #region handling curvedBeams (not implemented yet)
                List<FamilyInstance> strightBeams = new List<FamilyInstance>();
                List<FamilyInstance> curvedBeams = new List<FamilyInstance>();
                List<FamilyInstance> StrightElementsSplittedFromCurved = new List<FamilyInstance>();


                // union between  straightBeams and StraightBeamsFromSplittedCurvedBeams
                strightBeams = Allbeams;


                #endregion

                #region get Data Of Beams after handling curvedBeams

                foreach (var instance in strightBeams.AsParallel())
                {
                    try
                    {
                        Frame frame = new Frame();


                        //AnalyticalToPhysicalAssociationManager analyManger = AnalyticalToPhysicalAssociationManager.GetAnalyticalToPhysicalAssociationManager(doc);
                        //AnalyticalElement analytInstance = doc.GetElement(analyManger.GetAssociatedElementId(instance.Id)) as AnalyticalElement;

                        //if (analytInstance ==null)
                        //{
                        //    Autodesk.Revit.DB.Curve c = (instance.Location as LocationCurve).Curve;
                        //    XYZ fstPt = c.GetEndPoint(0);
                        //    XYZ sntPt = c.GetEndPoint(1);

                        //    //double zOffsetValue = instance.get_Parameter(BuiltInParameter.Z_OFFSET_VALUE).AsDouble();

                        //    double zOffsetValue = 0.0;


                        //    XYZ newFstPt = new XYZ(fstPt.X, fstPt.Y, fstPt.Z + zOffsetValue);
                        //    XYZ newsntPt = new XYZ(sntPt.X, sntPt.Y, sntPt.Z + zOffsetValue);
                        //    Line l = Line.CreateBound(newFstPt, newsntPt);
                        //    AnalyticalMember analyMember = AnalyticalMember.Create(doc, l);
                        //    analyManger.AddAssociation(analyMember.Id, instance.Id);
                        //}


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

                        frame.Length = UnitUtils.ConvertFromInternalUnits(frame.revitCurve.Length, dis);


                        frame.Sec = instance.Symbol.Name;


                        frame.RvtID = instance.Id.ToString();
                        frame.revitEleId = instance.Id;
                        frame.ID = instance.Id.ToString();



                        frame.Beta = instance.LookupParameter("Cross-Section Rotation").AsDouble() * 180 / (Math.PI);

                        try
                        {
                            frame.MainGroup = instance.LookupParameter("ACM_SE_Zone_Level_ID").AsString();
                            frame.SubGroup = instance.LookupParameter("BEAM TYPE").AsString();
                        }
                        catch (Exception)
                        {

                            goto AA;
                        }



                    AA:
                        frame.RvtFamily = instance.Symbol.FamilyName;

                        frame.RvtName = instance.Symbol.Name;
                        frame.MemberType = MemberType.Beam;

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



                        #region OLD
                        //// verticalColumn
                        //LocationPoint locationPoint = instance.Location as LocationPoint;
                        //if (locationPoint != null)
                        //{
                        //    XYZ columnPoint = locationPoint.Point;

                        //    Level baseLevel = (doc.GetElement(instance.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_PARAM).AsElementId()) as Level);
                        //    double baseLevelElev = baseLevel.ProjectElevation;
                        //    double baseLevelOffset = instance.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_OFFSET_PARAM).AsDouble();
                        //    double startptElev = baseLevelOffset + baseLevelElev;
                        //    XYZ startPoint = new XYZ(columnPoint.X, columnPoint.Y, startptElev);

                        //    Level TopLevel = (doc.GetElement(instance.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_PARAM).AsElementId()) as Level);
                        //    double TopLevelElev = TopLevel.ProjectElevation;
                        //    double topLevelOffset = instance.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_OFFSET_PARAM).AsDouble();
                        //    double endPtElev = topLevelOffset + TopLevelElev;
                        //    XYZ endPoint = new XYZ(columnPoint.X, columnPoint.Y, endPtElev);

                        //    Line c = Line.CreateBound(startPoint, endPoint);

                        //    AnalyticalMember analyMember = AnalyticalMember.Create(doc, c as Autodesk.Revit.DB.Curve);

                        //    frame.revitCurve = analyMember.GetCurve();

                        //    //analyManger.AddAssociation(analyMember.Id, instance.Id);
                        //    analytInstance = analyMember;
                        //}


                        //// slantted Column
                        //else
                        //{
                        //    Autodesk.Revit.DB.Curve locationCurve = (instance.Location as LocationCurve).Curve;

                        //    AnalyticalMember analyMember = AnalyticalMember.Create(doc, locationCurve as Autodesk.Revit.DB.Curve);

                        //    frame.revitCurve = analyMember.GetCurve();

                        //    //analyManger.AddAssociation(analyMember.Id, instance.Id);

                        //    analytInstance = analyMember;
                        //}

                        #endregion

                        #region NEw

                        Parameter colStyleParam = instance.LookupParameter("Column Style");
                        colStyleParam.Set((int)SlantedOrVerticalColumnType.CT_EndPoint);

                        Autodesk.Revit.DB.Curve locationCurve = (instance.Location as LocationCurve).Curve;


                        frame.revitCurve = locationCurve;




                        #endregion



                        if (true)
                        {
                            #region gePoints

                            frame.RhinoCurve = RhinoInside.Revit.Convert.Geometry.GeometryDecoder.ToCurve(frame.revitCurve);

                            Point3d stPt = frame.RhinoCurve.PointAtStart;
                            Point3d endPt = frame.RhinoCurve.PointAtEnd;




                            Point stPoint = new Point(stPt);
                            Point endPoint = new Point(endPt);

                            Point mid = Point.getMidPoint(stPoint, endPoint);

                            frame.StartPoint = stPoint;
                            frame.EndPoint = endPoint;
                            frame.MidPoint = mid;

                            frame.Length = Point.getDistance(stPoint, endPoint);



                            // another Approatch
                            // Point startPoint = new Point();
                            // Point endPoint = new Point();

                            // GetTopAndBottomPoints(doc, instance.Id, ref startPoint, ref endPoint);

                            // XYZ revStPt = new XYZ(startPoint.X, startPoint.Y, startPoint.Z);
                            // XYZ revEndPt = new XYZ(endPoint.X, endPoint.Y, endPoint.Z);

                            //Line ln=  Line.CreateBound(revStPt, revEndPt);
                            // frame.RhinoCurve = RhinoInside.Revit.Convert.Geometry.GeometryDecoder.ToCurve(ln);

                            // Point mid = Point.getMidPoint(startPoint, endPoint);

                            // frame.StartPoint = startPoint;
                            // frame.EndPoint = endPoint;
                            // frame.MidPoint = mid;

                            // frame.Length = Point.getDistance(startPoint, endPoint);


                            #endregion

                            frame.RvtID = instance.Id.ToString();
                            frame.revitEleId = instance.Id;
                            frame.ID = instance.Id.ToString();

                            frame.Sec = instance.Symbol.Name;
                            frame.RvtFamily = instance.Symbol.FamilyName;
                            frame.RvtName = instance.Symbol.Name;


                            frame.MemberType = MemberType.Column;


                            try
                            {
                                frame.MainGroup = instance.LookupParameter("ACM_SE_Zone_Level_ID").AsString();
                                frame.SubGroup = instance.LookupParameter("COLUMN TYPE").AsString();
                            }
                            catch (Exception ex)
                            {

                                goto BB;

                            }


                        BB:

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
                                // frame.Beta = (angle * 180 / Math.PI) + 90;
                                frame.Beta = (angle * 180 / Math.PI);

                            }

                            #endregion

                            //Autodesk.Revit.DB.Structure.StructuralFramingUtils.DisallowJoinAtEnd(instance, 0);
                            //Autodesk.Revit.DB.Structure.StructuralFramingUtils.DisallowJoinAtEnd(instance, 1);


                            columnFarames.Add(frame);
                        }


                        colStyleParam.Set((int)SlantedOrVerticalColumnType.CT_Vertical);

                    }

                    catch (Exception ex2)
                    {

                        System.Diagnostics.StackTrace trace = new System.Diagnostics.StackTrace(ex2, true);
                        RevitUtils.Log.Add("Error in getRevitFrames Method in the Columns part... Columninstance no" + instance.Id + ex2.Message + trace.GetFrame(trace.FrameCount - 1).GetFileLineNumber());

                        // TaskDialog.Show("ErrorInGetRevitFramesMethod", "ErrorInRevitBeamID= "+ instance.Id + ex2.Message + trace.GetFrame(trace.FrameCount - 1).GetFileLineNumber());
                        continue;
                    }

                }





                #endregion


                frames = beamFarames.Union(columnFarames).ToList();

                Beams = beamFarames;
                Columns = columnFarames;

                allTrns.Commit();

                return frames;
            }

        }



        public static List<Frame> getRevitFramesFor22(UIDocument uidoc, Document doc, ForgeTypeId dis, ref List<Frame> Beams, ref List<Frame> Columns)
        {

            using (Transaction allTrns = new Transaction(doc, "aaa"))
            {
                allTrns.Start();
                FailureHandlingOptions failOpt = allTrns.GetFailureHandlingOptions();
                failOpt.SetFailuresPreprocessor(new RoomWarningSwallower());
                allTrns.SetFailureHandlingOptions(failOpt);
                #region getLevel

                FilteredElementCollector collector = new FilteredElementCollector(doc);
                Level revitLvl = collector.OfClass(typeof(Level)).First() as Level;
                double LevelProjectElevation = revitLvl.ProjectElevation;
                double levlElevation = revitLvl.Elevation;
                double Zdiffrence = UnitUtils.ConvertFromInternalUnits(levlElevation - LevelProjectElevation, dis);



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
                //
                LogicalOrFilter orFilter = new LogicalOrFilter(categoryFilters);

                // Get the selected elements that match the filter
                List<Element> selectedElements = new List<Element>();
                foreach (ElementId id in selectedIds)
                {
                    Element element = uidoc.Document.GetElement(id);


                    FilteredElementCollector tagCollector = new FilteredElementCollector(doc)
    .OfCategory(BuiltInCategory.OST_Tags);



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


                #region handling curvedBeams (not implemented yet)
                List<FamilyInstance> strightBeams = new List<FamilyInstance>();
                List<FamilyInstance> curvedBeams = new List<FamilyInstance>();
                List<FamilyInstance> StrightElementsSplittedFromCurved = new List<FamilyInstance>();


                // union between  straightBeams and StraightBeamsFromSplittedCurvedBeams
                strightBeams = Allbeams;


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
                        #region Cosider ProjectElevation

                        // Define the translation vector (0, 0, 500)
                        Vector3d translationVector = new Vector3d(0, 0, Zdiffrence);

                        // Create the translation transformation
                        Rhino.Geometry.Transform translation = Rhino.Geometry.Transform.Translation(translationVector);

                        // Apply the transformation to the curve
                        frame.RhinoCurve.Transform(translation);
                        #endregion

                        #region OldPoints
                        //Point startPoint = new Point(frame.revitCurve.GetEndPoint(0));
                        //Point endPoint = new Point(frame.revitCurve.GetEndPoint(1));

                        //startPoint.X = UnitUtils.ConvertFromInternalUnits(startPoint.X, dis);
                        //startPoint.Y = UnitUtils.ConvertFromInternalUnits(startPoint.Y, dis);
                        //startPoint.Z = UnitUtils.ConvertFromInternalUnits(startPoint.Z, dis) + Zdiffrence;

                        //endPoint.X = UnitUtils.ConvertFromInternalUnits(endPoint.X, dis);
                        //endPoint.Y = UnitUtils.ConvertFromInternalUnits(endPoint.Y, dis);
                        //endPoint.Z = UnitUtils.ConvertFromInternalUnits(endPoint.Z, dis) + Zdiffrence;

                        //Point midPoint = new Point(frame.revitCurve.Evaluate(.5, true));

                        //midPoint.X = UnitUtils.ConvertFromInternalUnits(midPoint.X, dis);
                        //midPoint.Y = UnitUtils.ConvertFromInternalUnits(midPoint.Y, dis);
                        //midPoint.Z = UnitUtils.ConvertFromInternalUnits(midPoint.Z, dis) + Zdiffrence; 
                        #endregion

                        #region NewPoints
                        Point3d stPt = frame.RhinoCurve.PointAtStart;
                        Point3d endPt = frame.RhinoCurve.PointAtEnd;



                        Point startPoint = new Point(stPt);
                        Point endPoint = new Point(endPt);

                        Point midPoint = Point.getMidPoint(startPoint, endPoint);

                        frame.StartPoint = startPoint;
                        frame.EndPoint = endPoint;
                        frame.MidPoint = midPoint;

                        #endregion


                        frame.StartPoint = startPoint;
                        frame.EndPoint = endPoint;
                        frame.MidPoint = midPoint;

                        frame.Length = UnitUtils.ConvertFromInternalUnits(frame.revitCurve.Length, dis);


                        frame.Sec = instance.Symbol.Name;


                        frame.RvtID = instance.Id.ToString();
                        frame.revitEleId = instance.Id;
                        frame.ID = instance.Id.ToString();



                        frame.Beta = instance.LookupParameter("Cross-Section Rotation").AsDouble() * 180 / (Math.PI);

                        try
                        {
                            // frame.MainGroup = "RevitEleemtnts";
                            //frame.MainGroup = instance.LookupParameter("ACM_SE_Zone_Level_ID").AsString();

                            frame.MainGroup = instance.LookupParameter("ACM_SE_Zone_Level_ID").AsString();
                            frame.SubGroup = instance.LookupParameter("ACM_SE_Main Type ID").AsValueString();

                        }
                        catch (Exception)
                        {

                            goto AA;
                        }



                    AA:
                        frame.RvtFamily = instance.Symbol.FamilyName;

                        frame.RvtName = instance.Symbol.Name;
                        frame.MemberType = MemberType.Beam;

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

                        if (true)
                        {

                            if (instance.Id.ToString() == "868817")
                            {
                                int a = 0;
                            }

                            Parameter colStyleParam = instance.LookupParameter("Column Style");
                            colStyleParam.Set((int)SlantedOrVerticalColumnType.CT_EndPoint);

                            //trns.Start();

                            Autodesk.Revit.DB.Curve locationCurve = (instance.Location as LocationCurve).Curve;




                            frame.revitCurve = locationCurve;
                            frame.RhinoCurve = RhinoInside.Revit.Convert.Geometry.GeometryDecoder.ToCurve(frame.revitCurve);

                            #region Cosider ProjectElevation

                            // Define the translation vector (0, 0, 500)
                            Vector3d translationVector = new Vector3d(0, 0, Zdiffrence);

                            // Create the translation transformation
                            Rhino.Geometry.Transform translation = Rhino.Geometry.Transform.Translation(translationVector);

                            // Apply the transformation to the curve
                            frame.RhinoCurve.Transform(translation);
                            #endregion



                            if (true)
                            {
                                #region gePoints


                                Point3d stPt = frame.RhinoCurve.PointAtStart;
                                Point3d endPt = frame.RhinoCurve.PointAtEnd;



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

                                frame.Sec = instance.Symbol.Name;
                                frame.RvtFamily = instance.Symbol.FamilyName;
                                frame.RvtName = instance.Symbol.Name;


                                frame.MemberType = MemberType.Column;


                                try
                                {
                                    frame.MainGroup = instance.LookupParameter("ACM_SE_Zone_Level_ID").AsString();
                                    frame.SubGroup = instance.LookupParameter("ACM_SE_Main Type ID").AsValueString();

                                    //frame.MainGroup = "RevitEleemtnts";
                                    //frame.SubGroup = instance.Symbol.LookupParameter("Type Mark").AsValueString();


                                }


                                catch (Exception ex)
                                {

                                    goto BB;

                                }


                            BB:

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
                                    // frame.Beta = (angle * 180 / Math.PI) + 90;
                                    frame.Beta = (angle * 180 / Math.PI);

                                }

                                #endregion


                                //Autodesk.Revit.DB.Structure.StructuralFramingUtils.DisallowJoinAtEnd(instance, 0);
                                //Autodesk.Revit.DB.Structure.StructuralFramingUtils.DisallowJoinAtEnd(instance, 1);


                                columnFarames.Add(frame);
                            }




                            colStyleParam.Set((int)SlantedOrVerticalColumnType.CT_Vertical);

                        }



                    }


                    catch (Exception ex2)
                    {

                        System.Diagnostics.StackTrace trace = new System.Diagnostics.StackTrace(ex2, true);
                        RevitUtils.Log.Add("Error in getRevitFrames Method in the Columns part... Columninstance no" + instance.Id + ex2.Message + trace.GetFrame(trace.FrameCount - 1).GetFileLineNumber());

                        // TaskDialog.Show("ErrorInGetRevitFramesMethod", "ErrorInRevitBeamID= "+ instance.Id + ex2.Message + trace.GetFrame(trace.FrameCount - 1).GetFileLineNumber());
                        continue;
                    }

                }





                #endregion


                frames = beamFarames.Union(columnFarames).ToList();

                Beams = beamFarames;
                Columns = columnFarames;

                allTrns.Commit();

                return frames;
            }

        }

        public static List<Frame> getRevitFramesWithoutZOffsetForRhino(UIDocument uidoc, Document doc, ForgeTypeId dis, ref List<Frame> Beams, ref List<Frame> Columns)
        {

            using (Transaction allTrns = new Transaction(doc, "aaa"))
            {
                allTrns.Start();
                FailureHandlingOptions failOpt = allTrns.GetFailureHandlingOptions();
                failOpt.SetFailuresPreprocessor(new RoomWarningSwallower());
                allTrns.SetFailureHandlingOptions(failOpt);
                #region getLevel

                FilteredElementCollector collector = new FilteredElementCollector(doc);
                Level revitLvl = collector.OfClass(typeof(Level)).First() as Level;
                double LevelProjectElevation = revitLvl.ProjectElevation;
                double levlElevation = revitLvl.Elevation;
                double Zdiffrence = UnitUtils.ConvertFromInternalUnits(levlElevation - LevelProjectElevation, dis);



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
                //
                LogicalOrFilter orFilter = new LogicalOrFilter(categoryFilters);

                // Get the selected elements that match the filter
                List<Element> selectedElements = new List<Element>();
                foreach (ElementId id in selectedIds)
                {
                    Element element = uidoc.Document.GetElement(id);


                    FilteredElementCollector tagCollector = new FilteredElementCollector(doc)
    .OfCategory(BuiltInCategory.OST_Tags);



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


                #region handling curvedBeams (not implemented yet)
                List<FamilyInstance> strightBeams = new List<FamilyInstance>();
                List<FamilyInstance> curvedBeams = new List<FamilyInstance>();
                List<FamilyInstance> StrightElementsSplittedFromCurved = new List<FamilyInstance>();


                // union between  straightBeams and StraightBeamsFromSplittedCurvedBeams
                strightBeams = Allbeams;


                #endregion

                #region get Data Of Beams after handling curvedBeams

                foreach (var instance in strightBeams.AsParallel())
                {
                    try
                    {
                        Frame frame = new Frame();

                        XYZ stPoint = (instance.Location as LocationCurve).Curve.GetEndPoint(0);
                        double z = UnitUtils.ConvertFromInternalUnits(stPoint.Z, dis);

                        // frame.revitCurve = analytInstance.GetCurve();
                        frame.revitCurve = (instance.Location as LocationCurve).Curve;

                        frame.RhinoCurve = RhinoInside.Revit.Convert.Geometry.GeometryDecoder.ToCurve(frame.revitCurve);
                        #region Cosider ProjectElevation

                        // Define the translation vector (0, 0, 500)
                        Vector3d translationVector = new Vector3d(0, 0, Zdiffrence);

                        // Create the translation transformation
                        Rhino.Geometry.Transform translation = Rhino.Geometry.Transform.Translation(translationVector);

                        // Apply the transformation to the curve /////////////////////////
                        // frame.RhinoCurve.Transform(translation);
                        #endregion

                        #region OldPoints
                        //Point startPoint = new Point(frame.revitCurve.GetEndPoint(0));
                        //Point endPoint = new Point(frame.revitCurve.GetEndPoint(1));

                        //startPoint.X = UnitUtils.ConvertFromInternalUnits(startPoint.X, dis);
                        //startPoint.Y = UnitUtils.ConvertFromInternalUnits(startPoint.Y, dis);
                        //startPoint.Z = UnitUtils.ConvertFromInternalUnits(startPoint.Z, dis) + Zdiffrence;

                        //endPoint.X = UnitUtils.ConvertFromInternalUnits(endPoint.X, dis);
                        //endPoint.Y = UnitUtils.ConvertFromInternalUnits(endPoint.Y, dis);
                        //endPoint.Z = UnitUtils.ConvertFromInternalUnits(endPoint.Z, dis) + Zdiffrence;

                        //Point midPoint = new Point(frame.revitCurve.Evaluate(.5, true));

                        //midPoint.X = UnitUtils.ConvertFromInternalUnits(midPoint.X, dis);
                        //midPoint.Y = UnitUtils.ConvertFromInternalUnits(midPoint.Y, dis);
                        //midPoint.Z = UnitUtils.ConvertFromInternalUnits(midPoint.Z, dis) + Zdiffrence; 
                        #endregion

                        #region NewPoints
                        Point3d stPt = frame.RhinoCurve.PointAtStart;
                        Point3d endPt = frame.RhinoCurve.PointAtEnd;



                        Point startPoint = new Point(stPt);
                        Point endPoint = new Point(endPt);

                        Point midPoint = Point.getMidPoint(startPoint, endPoint);

                        frame.StartPoint = startPoint;
                        frame.EndPoint = endPoint;
                        frame.MidPoint = midPoint;

                        #endregion


                        frame.StartPoint = startPoint;
                        frame.EndPoint = endPoint;
                        frame.MidPoint = midPoint;

                        frame.Length = UnitUtils.ConvertFromInternalUnits(frame.revitCurve.Length, dis);


                        frame.Sec = instance.Symbol.Name;


                        frame.RvtID = instance.Id.ToString();
                        frame.revitEleId = instance.Id;
                        frame.ID = instance.Id.ToString();



                        frame.Beta = instance.LookupParameter("Cross-Section Rotation").AsDouble() * 180 / (Math.PI);

                        try
                        {
                            // frame.MainGroup = "RevitEleemtnts";
                            //frame.MainGroup = instance.LookupParameter("ACM_SE_Zone_Level_ID").AsString();

                            frame.MainGroup = instance.LookupParameter("ACM_SE_Zone_Level_ID").AsString();
                            frame.SubGroup = instance.LookupParameter("ACM_SE_Main Type ID").AsValueString();

                        }
                        catch (Exception)
                        {

                            goto AA;
                        }



                    AA:
                        frame.RvtFamily = instance.Symbol.FamilyName;

                        frame.RvtName = instance.Symbol.Name;
                        frame.MemberType = MemberType.Beam;

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

                        if (true)
                        {

                            if (instance.Id.ToString() == "868817")
                            {
                                int a = 0;
                            }

                            Parameter colStyleParam = instance.LookupParameter("Column Style");
                            colStyleParam.Set((int)SlantedOrVerticalColumnType.CT_EndPoint);

                            //trns.Start();

                            Autodesk.Revit.DB.Curve locationCurve = (instance.Location as LocationCurve).Curve;




                            frame.revitCurve = locationCurve;
                            frame.RhinoCurve = RhinoInside.Revit.Convert.Geometry.GeometryDecoder.ToCurve(frame.revitCurve);

                            #region Cosider ProjectElevation

                            // Define the translation vector (0, 0, 500)
                            Vector3d translationVector = new Vector3d(0, 0, Zdiffrence);

                            // Create the translation transformation
                            Rhino.Geometry.Transform translation = Rhino.Geometry.Transform.Translation(translationVector);

                            // Apply the transformation to the curve ////////////////////////////////////////////////////
                            //  frame.RhinoCurve.Transform(translation);
                            #endregion



                            if (true)
                            {
                                #region gePoints


                                Point3d stPt = frame.RhinoCurve.PointAtStart;
                                Point3d endPt = frame.RhinoCurve.PointAtEnd;



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

                                frame.Sec = instance.Symbol.Name;
                                frame.RvtFamily = instance.Symbol.FamilyName;
                                frame.RvtName = instance.Symbol.Name;


                                frame.MemberType = MemberType.Column;


                                try
                                {
                                    frame.MainGroup = instance.LookupParameter("ACM_SE_Zone_Level_ID").AsString();
                                    frame.SubGroup = instance.LookupParameter("ACM_SE_Main Type ID").AsValueString();

                                    //frame.MainGroup = "RevitEleemtnts";
                                    //frame.SubGroup = instance.Symbol.LookupParameter("Type Mark").AsValueString();


                                }


                                catch (Exception ex)
                                {

                                    goto BB;

                                }


                            BB:

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
                                    // frame.Beta = (angle * 180 / Math.PI) + 90;
                                    frame.Beta = (angle * 180 / Math.PI);

                                }

                                #endregion


                                //Autodesk.Revit.DB.Structure.StructuralFramingUtils.DisallowJoinAtEnd(instance, 0);
                                //Autodesk.Revit.DB.Structure.StructuralFramingUtils.DisallowJoinAtEnd(instance, 1);


                                columnFarames.Add(frame);
                            }




                            colStyleParam.Set((int)SlantedOrVerticalColumnType.CT_Vertical);

                        }



                    }


                    catch (Exception ex2)
                    {

                        System.Diagnostics.StackTrace trace = new System.Diagnostics.StackTrace(ex2, true);
                        RevitUtils.Log.Add("Error in getRevitFrames Method in the Columns part... Columninstance no" + instance.Id + ex2.Message + trace.GetFrame(trace.FrameCount - 1).GetFileLineNumber());

                        // TaskDialog.Show("ErrorInGetRevitFramesMethod", "ErrorInRevitBeamID= "+ instance.Id + ex2.Message + trace.GetFrame(trace.FrameCount - 1).GetFileLineNumber());
                        continue;
                    }

                }





                #endregion


                frames = beamFarames.Union(columnFarames).ToList();

                Beams = beamFarames;
                Columns = columnFarames;

                allTrns.Commit();

                return frames;
            }

        }


        // original
        public static void drawRhinoFrames(ForgeTypeId dis, UIDocument revitUiDoc, Document revitDoc, RhinoDoc doc)
        {

            List<Frame> frames = new List<Frame>();
            List<Frame> columns = new List<Frame>();
            List<Frame> beams = new List<Frame>();


            List<Frame> revitFrames = getRevitFrames(revitUiDoc, revitDoc, dis, ref beams, ref columns);

            foreach (var item in revitFrames)
            {

                Rhino.DocObjects.Tables.ObjectTable ot = Rhino.RhinoDoc.ActiveDoc.Objects;
                ObjectAttributes attr = new ObjectAttributes();


                attr.SetUserString("Section", item.Sec);
                attr.SetUserString("Family", item.RvtFamily);
                attr.SetUserString("Angle", item.Beta.ToString());
                attr.SetUserString("MemberType", item.MemberType.ToString());
                attr.SetUserString("ID", item.RvtID);

                ot.AddCurve(item.RhinoCurve, attr);

            }


            #region Create frames Layers
            RhinoObject[] rhinoCurves = doc.Objects.FindByObjectType(ObjectType.Curve).Where(x => x.Attributes.GetUserString("Section") != null).ToArray();


            #region get  list of Sections
            List<string> sections = new List<string>();
            foreach (var item in rhinoCurves)
            {

                ObjectAttributes att = item.Attributes;

                string group = att.GetUserString("Section");

                sections.Add(group);

            }
            sections = sections.Distinct().ToList();
            #endregion


            int FrameLayerIndex;
            Layer FrameLayer;
            FrameLayer = doc.Layers.FindName("RevitFrames");

            if (FrameLayer == null)
            {
                FrameLayer = new Layer();
                FrameLayer.Name = "RevitFrames";
                FrameLayerIndex = doc.Layers.Add(FrameLayer);
            }

            else
            {
                FrameLayerIndex = FrameLayer.Index;
            }



            FrameLayer = doc.Layers.FindName("RevitFrames");

            int i = 0;

            foreach (var sec in sections.Where(x => x != null))
            {
                int red = i * 20 % 256;
                int green = i * 40 % 256;
                int blue = i * 60 % 256;

                System.Drawing.Color color = System.Drawing.Color.FromArgb(255, red, green, blue);

                int subLayerIndex;
                Layer subLayer = doc.Layers.FindName(sec);


                if (subLayer == null)
                {
                    subLayer = new Layer();
                    subLayer.Name = sec;
                    subLayer.Color = color;
                    subLayer.ParentLayerId = FrameLayer.Id;
                    subLayerIndex = doc.Layers.Add(subLayer);
                }

                else
                {
                    subLayer.Name = sec;
                    subLayer.Color = color;
                    subLayer.ParentLayerId = FrameLayer.Id;
                    subLayerIndex = subLayer.Index;
                }




                RhinoObject[] objects = doc.Objects.FindByUserString("Section", sec, true);
                foreach (var item in objects)
                {
                    item.Attributes.LayerIndex = subLayerIndex;
                    item.CommitChanges();
                }

                i++;
            }

            #endregion

        }


        // ForSpecificReason
        public static void drawRhinoFramesForSepcificReason(ForgeTypeId dis, UIDocument revitUiDoc, Document revitDoc, RhinoDoc doc)
        {

            List<Frame> frames = new List<Frame>();
            List<Frame> columns = new List<Frame>();
            List<Frame> beams = new List<Frame>();


            List<Frame> revitFrames = getRevitFramesWithoutZOffsetForRhino(revitUiDoc, revitDoc, dis, ref beams, ref columns);

            RevitUtils.SetNearestbaseAndTopLevels(columns);
            foreach (var item in revitFrames)
            {

                Rhino.DocObjects.Tables.ObjectTable ot = Rhino.RhinoDoc.ActiveDoc.Objects;
                ObjectAttributes attr = new ObjectAttributes();
                if (columns.Contains(item))
                {
                    #region OLD
                    //string topLevel = ((FamilyInstance)revitDoc.GetElement(item.revitEleId)).get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_PARAM).AsValueString();
                    //string BotLevel = ((FamilyInstance)revitDoc.GetElement(item.revitEleId)).get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_PARAM).AsValueString();
                    //double topLvlOffset = ((FamilyInstance)revitDoc.GetElement(item.revitEleId)).get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_OFFSET_PARAM).AsDouble();
                    //topLvlOffset = Math.Round( UnitUtils.ConvertFromInternalUnits(topLvlOffset, dis),0);

                    //double botLvlOffset = ((FamilyInstance)revitDoc.GetElement(item.revitEleId)).get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_OFFSET_PARAM).AsDouble();
                    //botLvlOffset = Math.Round(UnitUtils.ConvertFromInternalUnits(botLvlOffset, dis), 0);


                    //attr.SetUserString("TopLevel", topLevel);
                    //attr.SetUserString("BotLevel", BotLevel);


                    //attr.SetUserString("TopLevelOffset", topLvlOffset.ToString());
                    //attr.SetUserString("BotLevelOffset", botLvlOffset.ToString()); 
                    #endregion

                    #region NEw

                    attr.SetUserString("TopLevel", item.RevitTopLevel);
                    attr.SetUserString("BotLevel", item.RevitBotLevel);


                    attr.SetUserString("TopLevelOffset", item.RevitTopLevelOffset.ToString());
                    attr.SetUserString("BotLevelOffset", item.RevitBotLevelOffset.ToString());
                    #endregion

                }


                attr.SetUserString("Section", item.Sec);
                attr.SetUserString("Family", item.RvtFamily);
                attr.SetUserString("Angle", item.Beta.ToString());
                attr.SetUserString("MemberType", item.MemberType.ToString());
                attr.SetUserString("ID", item.RvtID);
                attr.SetUserString("MainGroup", item.MainGroup);
                attr.SetUserString("SubGroup", item.SubGroup);

                ot.AddCurve(item.RhinoCurve, attr);

            }


            #region Create frames Layers
            RhinoObject[] rhinoCurves = doc.Objects.FindByObjectType(ObjectType.Curve).Where(x => x.Attributes.GetUserString("Family") != null).ToArray();

            #region get  list of Sections
            List<string> Maingroups = new List<string>();

            foreach (var item in rhinoCurves)
            {

                ObjectAttributes att = item.Attributes;

                string Mgroup = att.GetUserString("MainGroup");
                string subgroup = att.GetUserString("SubGroup");
                Maingroups.Add(Mgroup);

            }

            Maingroups = Maingroups.Distinct().ToList();

            Maingroups = Maingroups.Where(x => x != null).ToList();

            #endregion


            foreach (var MG in Maingroups)
            {

                List<string> subGroups = new List<string>();

                rhinoCurves = rhinoCurves.Where(x => x.Attributes.GetUserString("MainGroup") == MG).ToArray();
                foreach (var item in rhinoCurves)
                {
                    ObjectAttributes att = item.Attributes;
                    string subgroup = att.GetUserString("SubGroup");
                    subGroups.Add(subgroup);

                }
                subGroups = subGroups.Distinct().ToList();
                subGroups.Sort();

                int MGLayerIndex;
                Layer MainGroupLayer;
                MainGroupLayer = doc.Layers.FindName(MG);

                if (MainGroupLayer == null)
                {
                    MainGroupLayer = new Layer();
                    MainGroupLayer.Name = MG;
                    MGLayerIndex = doc.Layers.Add(MainGroupLayer);
                }

                else
                {
                    MGLayerIndex = MainGroupLayer.Index;
                }



                MainGroupLayer = doc.Layers.FindName(MG);

                int i = 0;

                foreach (var subGroup in subGroups)
                {
                    //int red = i * 20 % 256;
                    //int green = i * 40 % 256;
                    //int blue = i * 60 % 256;
                    int red = 5;
                    int green = 5;
                    int blue = 5;

                    System.Drawing.Color color = System.Drawing.Color.FromArgb(255, red, green, blue);

                    int subLayerIndex;
                    Layer subLayer = doc.Layers.FindName(subGroup);


                    if (subLayer == null || subLayer.ParentLayerId != MainGroupLayer.Id)
                    {
                        subLayer = new Layer();
                        if (subGroup == null || subGroup.Contains("none"))
                        {
                            subLayer.Name = "none";
                            subLayer.Color = color;
                            subLayer.ParentLayerId = MainGroupLayer.Id;
                            subLayerIndex = doc.Layers.Add(subLayer);
                        }

                        else
                        {
                            subLayer.Name = subGroup;
                            subLayer.Color = color;
                            subLayer.ParentLayerId = MainGroupLayer.Id;
                            subLayerIndex = doc.Layers.Add(subLayer);
                        }

                    }

                    else
                    {
                        subLayer.Name = subGroup;
                        subLayer.Color = color;
                        subLayer.ParentLayerId = MainGroupLayer.Id;
                        subLayerIndex = subLayer.Index;
                    }



                    // Search for objects that satisfy both conditions
                    //RhinoObject[] objects;

                    //objects = doc.Objects.FindByUserString("MainGroup", MG, true);

                    //objects= objects.Where(x => x.Attributes.GetUserString("SubGroup") == subGroup).ToArray();

                    RhinoObject[] objects = doc.Objects.FindByUserString("SubGroup", subGroup, true);

                    objects = objects.Where(x => x.Attributes.GetUserString("Family") != null).ToArray();


                    foreach (var item in objects)
                    {

                        item.Attributes.LayerIndex = subLayerIndex;
                        item.CommitChanges();
                    }

                    i++;
                }

            }

            #endregion

        }


        public static List<Frame> getRhinoFrames()
        {
            List<Frame> rhinoFrames = new List<Frame>();
            Rhino.RhinoDoc doc = RhinoDoc.ActiveDoc;

            List<RhinoObject> rhinoCurves = doc.Objects.FindByObjectType(ObjectType.Curve).ToList();


            #region New
            foreach (var crv in rhinoCurves)
            {
                NameValueCollection n = crv.Attributes.GetUserStrings();

                //List<string> keys = n.AllKeys.ToList().Where(x => (x.Contains("Group") && x != "Groups")).ToList();
                //List<string> newgroups = new List<string>();
                //keys.ForEach(x => newgroups.Add(n[x]));


                Frame rhinoFrame = new Frame();
                rhinoFrame.Sec = crv.Attributes.GetUserString("Section");
                rhinoFrame.Shape = crv.Attributes.GetUserString("Shape");
                rhinoFrame.ID = crv.Attributes.GetUserString("ID");
                double beta = 0.00;

                double.TryParse(crv.Attributes.GetUserString("Angle"), out beta);
                rhinoFrame.Beta = beta;
                rhinoFrame.RhinoID = crv.Id.ToString();

                // string memberType = crv.Attributes.GetUserString("MemberType");

                //if (memberType == "Beam")
                //{
                //    rhinoFrame.MemberType = MemberType.Beam;
                //}
                //else if (memberType == "VLBracing")
                //{
                //    rhinoFrame.MemberType = MemberType.VLBracing;
                //}

                //else if  (memberType == "Column")
                //{

                //    rhinoFrame.MemberType = MemberType.Column;
                //}

                //else
                //{
                //    rhinoFrame.MemberType = MemberType.SlanttedColumn;
                //}





                ///////// very Important
                CurveObject curveObject = crv as Rhino.DocObjects.CurveObject;
                Curve curvGeom = curveObject.CurveGeometry;


                Point startPt = new Point("", curvGeom.PointAtStart.X, curvGeom.PointAtStart.Y, curvGeom.PointAtStart.Z);
                Point endPt = new Point("", curvGeom.PointAtEnd.X, curvGeom.PointAtEnd.Y, curvGeom.PointAtEnd.Z);
                rhinoFrame.StartPoint = startPt;
                rhinoFrame.EndPoint = endPt;

                double MidParameter = curvGeom.Domain.Mid;
                Point3d mid = curvGeom.PointAt(MidParameter);


                rhinoFrame.MidPoint = new Point("", mid.X, mid.Y, mid.Z);
                rhinoFrame.Length = curvGeom.GetLength();

                Point.getMemberType(rhinoFrame);

                rhinoFrame.RhinoCurve = curvGeom;

                // rhinoFrame.revitCurve = RhinoInside.Revit.Convert.Geometry.GeometryEncoder.ToCurve(curvGeom);

                rhinoFrames.Add(rhinoFrame);

            }

            #endregion

            //var groupedFrames = rhinoFrames.GroupBy(x => x.ID);

            var groupedFrames = rhinoFrames
                                 .Where(x => !string.IsNullOrWhiteSpace(x.ID))
                                 .GroupBy(x => x.ID)
                                 .Select(group => group.OrderBy(f => f.SerialNumber).ToList()) // sort by creation order
                                 .ToList();


            foreach (var group in groupedFrames)
            {
                // Get the first item in the group
                var firstItem = group.FirstOrDefault();

                if (firstItem != null)
                {
                    // Set the ID property to null for all other items in the group
                    foreach (var item in group.Skip(1))
                    {
                        item.ParentID = item.ID;

                        item.ID = null;

                    }
                }
            }


            return rhinoFrames;
        }

        public static void updateRevitFrames(UIDocument uidoc, Document revitDoc, ForgeTypeId dis)
        {

            #region getRevitFrames
            List<Frame> Beams = new List<Frame>();
            List<Frame> Columns = new List<Frame>();
            List<Frame> revitFrames = getRevitFramesForRevit22(uidoc, revitDoc, dis, ref Beams, ref Columns);

            #endregion


            List<Frame> rhinoFrames = getRhinoFrames();


            Rhino.RhinoDoc rhinoDoc = RhinoDoc.ActiveDoc;
            List<RhinoObject> rhinoCurves = rhinoDoc.Objects.FindByObjectType(ObjectType.Curve).ToList();


            List<RhinoRevitFrame> AllElements = new List<RhinoRevitFrame>();

            List<Frame> unMappedRevitFrames = new List<Frame>();
            List<Frame> unMappedRhinoFrames = new List<Frame>();
            List<Frame> rhinoUnMappedBeams = new List<Frame>();
            List<Frame> rhinoUnMappedColumns = new List<Frame>();

            List<RhinoRevitFrame> mappedElements = new List<RhinoRevitFrame>();
            List<RhinoRevitFrame> rhinoMappedBeams = new List<RhinoRevitFrame>();
            List<RhinoRevitFrame> rhinoMappedColumns = new List<RhinoRevitFrame>();

            List<RhinoRevitFrame> mappedElementsById = new List<RhinoRevitFrame>();
            List<RhinoRevitFrame> mappedElementsByMid = new List<RhinoRevitFrame>();


            if (revitFrames.Count != 0)
            {
                AllElements = RhinoRevitFrame.getMappedSapRevit(rhinoFrames, revitFrames, 0.05);
                mappedElements = AllElements.AsParallel().Where(x => x.revitFrame != null).ToList();


                mappedElementsById = mappedElements.AsParallel().Where(x => x.mappedBy == "ID").ToList();
                mappedElementsByMid = mappedElements.AsParallel().Where(x => x.mappedBy == "MID").ToList();


                RhinoRevitFrame.getUnMappedEelements(AllElements, rhinoFrames, revitFrames, ref unMappedRhinoFrames, ref unMappedRevitFrames);
            }



            else
            {
                unMappedRhinoFrames = rhinoFrames;
            }


            #region forMappedRhinoFrames

            rhinoMappedBeams = mappedElements.AsParallel().Where(x => x.sapFrame.MemberType == MemberType.Beam || x.sapFrame.MemberType == MemberType.VLBracing).ToList();
            RhinoRevitFrame.updateRevitGeomertyForBeams3(rhinoDoc, revitDoc, rhinoMappedBeams, dis, 0.001);

            rhinoMappedColumns = mappedElements.AsParallel().Where(x => x.sapFrame.MemberType == MemberType.Column || x.sapFrame.MemberType == MemberType.SlanttedColumn).ToList();
            RhinoRevitFrame.updateRevitGeomertyForColumns2BasedOnSlanttedStyle(rhinoDoc, revitDoc, rhinoMappedColumns, dis, 0.001);

            #endregion

            #region unMappedd

            rhinoUnMappedBeams = unMappedRhinoFrames.AsParallel().Where(x => x.MemberType == MemberType.Beam || x.MemberType == MemberType.VLBracing).ToList();
            rhinoUnMappedColumns = unMappedRhinoFrames.AsParallel().Where(x => x.MemberType == MemberType.Column || x.MemberType == MemberType.SlanttedColumn).ToList();


            List<ElementId> newRevitCreatedElementsIds = new List<ElementId>();

            newRevitCreatedElementsIds = CreateAllBeams3(revitDoc, rhinoUnMappedBeams, dis);

            newRevitCreatedElementsIds = newRevitCreatedElementsIds.Union(CreateAllColumns3(revitDoc, rhinoUnMappedColumns, dis)).ToList();



            CreateOrUpdateSelectionFilter(uidoc, newRevitCreatedElementsIds, "NewAddedMembers");
            deleteUnMappedRevitElements(revitDoc, unMappedRevitFrames);


            #endregion

        }


        public static List<RhinoRevitFrame> getMappedRhinoRevit(List<Frame> sapFrames, List<Frame> revitFrames, double margin)
        {
            List<RhinoRevitFrame> allResults = new List<RhinoRevitFrame>();
            


           List<string> sapIDS = sapFrames.Select(x => x.ID).AsParallel().ToList();
            List<string> sapIDParameterFromRvt = revitFrames.Select(x => x.ID).AsParallel().ToList();


            List<Frame> sapElementsMappedByIDS = sapFrames.Where(x => sapIDParameterFromRvt.Contains(x.ID)).AsParallel().OrderBy(x => x.ID).ToList();
            List<Frame> revitElementsMappedByIDS = revitFrames.Where(x => sapIDS.Contains(x.ID)).AsParallel().OrderBy(x => x.ID).ToList();



            for (int i = sapElementsMappedByIDS.Count - 1; i >= 0; i--)
            {

                //if (SapRevitFrame.IsSameType(sapElementsMappedByIDS[i], revitElementsMappedByIDS[i]))
                //{

                RhinoRevitFrame result = new RhinoRevitFrame(sapElementsMappedByIDS[i], revitElementsMappedByIDS[i]);
                result.mappedBy = "ID";
                allResults.Add(result);

                //}

                //else
                //{
                sapElementsMappedByIDS.RemoveAt(i);
                revitElementsMappedByIDS.RemoveAt(i);
                //}

            }


            List<Frame> sapElementsNotMappedByIDS = sapFrames.Except(sapElementsMappedByIDS).ToList();
            List<Frame> revitElementsNotMappedByIDS = revitFrames.Except(revitElementsMappedByIDS).ToList();


            foreach (var sapFrame in sapElementsNotMappedByIDS.AsParallel())
            {
                RhinoRevitFrame result = new RhinoRevitFrame();
                Dictionary<Frame, double> dictionary = new Dictionary<Frame, double>();
                List<double> lengthDiffernce = new List<double>();
                result.sapFrame = sapFrame;
                result.revitFrame = null;



                foreach (Frame revitFrame in revitElementsNotMappedByIDS.AsParallel())
                {
                    double lengthDifferencePersentage = Math.Abs(revitFrame.Length - sapFrame.Length) / sapFrame.Length;
                    if (lengthDifferencePersentage < margin)
                    {
                        dictionary.Add(revitFrame, Point.getDistance(sapFrame.MidPoint, revitFrame.MidPoint) + Point.getDistance(sapFrame.StartPoint, revitFrame.StartPoint));
                    }

                }


                #region getMinValueFromDictionary
                double mainValue = double.MaxValue;
                if (dictionary.Count != 0)
                {
                    mainValue = dictionary.Values.AsParallel().ToList().Min();
                }


                #endregion


                if (mainValue < margin * 2 * sapFrame.Length)
                {
                    int index = dictionary.Values.AsParallel().ToList().IndexOf(dictionary.Values.Min());
                    result.revitFrame = dictionary.Keys.AsParallel().ToList()[index];
                    result.mappedBy = "MID";
                    allResults.Add(result);
                }

                else
                {
                    result.revitFrame = null;
                    allResults.Add(result);
                }

            Next:
                continue;
            }


            //foreach (var sapFrame in sapFrames.AsParallel())
            //{
            //    SapRevitFrame result = new SapRevitFrame();
            //    Dictionary<Frame, double> dictionary = new Dictionary<Frame, double>();
            //    List<double> lengthDiffernce = new List<double>();
            //    result.sapFrame = sapFrame;
            //    result.revitFrame = null;



            //    foreach (Frame revitFrame in revitFrames.AsParallel())
            //    {
            //        if (revitFrame.ID == sapFrame.ID)
            //        {
            //            //double lengthDifferencePersentage = Math.Abs(revitFrame.Length - sapFrame.Length) / sapFrame.Length;
            //            //if (lengthDifferencePersentage < margin)
            //            //{
            //            //    result.revitFrame = revitFrame;
            //            //    result.mappedBy = "ID";
            //            //    allResults.Add(result);
            //            //    goto Next;

            //            //}

            //            result.revitFrame = revitFrame;
            //            result.mappedBy = "ID";
            //            allResults.Add(result);
            //            goto Next;



            //        }

            //        else
            //        {

            //            double lengthDifferencePersentage = Math.Abs(revitFrame.Length - sapFrame.Length) / sapFrame.Length;
            //            if (lengthDifferencePersentage < margin)
            //            {
            //                dictionary.Add(revitFrame, Point.getDistance(sapFrame.MidPoint, revitFrame.MidPoint));
            //            }

            //        }
            //    }


            //    #region getMinValueFromDictionary
            //    double mainValue = 1234567891011;
            //    if (dictionary.Count != 0)
            //    {
            //        mainValue = dictionary.Values.AsParallel().ToList().Min();
            //    }


            //    #endregion


            //    if (mainValue < margin * sapFrame.Length)
            //    {
            //        int index = dictionary.Values.AsParallel().ToList().IndexOf(dictionary.Values.Min());
            //        result.revitFrame = dictionary.Keys.AsParallel().ToList()[index];
            //        result.mappedBy = "MID";
            //        allResults.Add(result);
            //    }

            //    else
            //    {
            //        result.revitFrame = null;
            //        allResults.Add(result);
            //    }

            //Next:
            //    continue;
            //}

            return allResults;

        }



        public static void RevitToRhinoFramesOnlyForEtabs(List<Frame> revitFrames)
        {

            UIDocument revitUiDoc = RhinoInside.Revit.Revit.ActiveUIDocument;
            Document revitDoc = RhinoInside.Revit.Revit.ActiveDBDocument;
            Rhino.RhinoDoc doc = RhinoDoc.ActiveDoc;

            // for revit 2021 or above use (ForgeTypeId dis = UnitTypeId.Millimeters;) however for revit 2020 or lower use ( ForgeTypeId dis = DisplayUnitType.DUT_MILLIMETERS;)
            ForgeTypeId dis = UnitTypeId.Millimeters;

            List<Frame> frames = new List<Frame>();
            List<Frame> columns = new List<Frame>();
            List<Frame> beams = new List<Frame>();


            foreach (var item in revitFrames)
            {

                Rhino.DocObjects.Tables.ObjectTable ot = Rhino.RhinoDoc.ActiveDoc.Objects;
                ObjectAttributes attr = new ObjectAttributes();


                attr.SetUserString("Section", item.Sec);
                attr.SetUserString("Family", item.RvtFamily);
                attr.SetUserString("Angle", item.Beta.ToString());
                attr.SetUserString("MemberType", item.MemberType.ToString());
                attr.SetUserString("ID", item.RvtID);

                ot.AddCurve(item.RhinoCurve, attr);

            }


            #region Create frames Layers
            RhinoObject[] rhinoCurves = doc.Objects.FindByObjectType(ObjectType.Curve).Where(x => x.Attributes.GetUserString("Section") != null).ToArray();


            #region get  list of Sections
            List<string> sections = new List<string>();
            foreach (var item in rhinoCurves)
            {

                ObjectAttributes att = item.Attributes;

                string group = att.GetUserString("Section");

                sections.Add(group);

            }
            sections = sections.Distinct().ToList();
            #endregion


            int FrameLayerIndex;
            Layer FrameLayer;
            FrameLayer = doc.Layers.FindName("RevitFrames");

            if (FrameLayer == null)
            {
                FrameLayer = new Layer();
                FrameLayer.Name = "RevitFrames";
                FrameLayerIndex = doc.Layers.Add(FrameLayer);
            }

            else
            {
                FrameLayerIndex = FrameLayer.Index;
            }



            FrameLayer = doc.Layers.FindName("RevitFrames");

            int i = 0;

            foreach (var sec in sections.Where(x => x != null))
            {
                int red = i * 20 % 256;
                int green = i * 40 % 256;
                int blue = i * 60 % 256;

                System.Drawing.Color color = System.Drawing.Color.FromArgb(255, red, green, blue);

                int subLayerIndex;
                Layer subLayer = doc.Layers.FindName(sec);


                if (subLayer == null)
                {
                    subLayer = new Layer();
                    subLayer.Name = sec;
                    subLayer.Color = color;
                    subLayer.ParentLayerId = FrameLayer.Id;
                    subLayerIndex = doc.Layers.Add(subLayer);
                }

                else
                {
                    subLayer.Name = sec;
                    subLayer.Color = color;
                    subLayer.ParentLayerId = FrameLayer.Id;
                    subLayerIndex = subLayer.Index;
                }




                RhinoObject[] objects = doc.Objects.FindByUserString("Section", sec, true);
                foreach (var item in objects)
                {
                    item.Attributes.LayerIndex = subLayerIndex;
                    item.CommitChanges();
                }

                i++;
            }

            #endregion

        }


        public static List<Rhino.Geometry.Curve> CreateRhinoCurves(ForgeTypeId dis, UIDocument revitUiDoc, Document revitDoc, RhinoDoc doc)
        {
            List<Frame> beams = new List<Frame>();
            List<Frame> columns = new List<Frame>();
            List<Frame> revitFrames = new List<Frame>();

            revitFrames = getRevitFrames(revitUiDoc, revitDoc, dis, ref beams, ref columns);


            List<Curve> curves = new List<Curve>();


            // reate RhinoCurves
            foreach (var rvtFrame in revitFrames)
            {


                Point3d stPoint = new Point3d(rvtFrame.StartPoint.X, rvtFrame.StartPoint.Y, rvtFrame.StartPoint.Z);
                Point3d endPoint = new Point3d(rvtFrame.EndPoint.X, rvtFrame.EndPoint.Y, rvtFrame.EndPoint.Z);
                LineCurve crv = new LineCurve(stPoint, endPoint);
                Rhino.DocObjects.Tables.ObjectTable ot = Rhino.RhinoDoc.ActiveDoc.Objects;
                ObjectAttributes attr = new ObjectAttributes();

                attr.SetUserString("Section", rvtFrame.Sec.ToString());
                //attr.SetUserString("Family", rvtFrame.RvtFamily.ToString());//
                attr.SetUserString("Family", rvtFrame.RvtFamily.ToString());

                attr.SetUserString("MemberType", rvtFrame.MemberType.ToString());

                try
                {
                    attr.SetUserString("SapID", rvtFrame.SapID.ToString());

                }
                catch (Exception)
                {

                }

                // no points ids in Revit
                // attr.SetUserString("StartPointId", rvtFrame.StartPoint.Id);
                //  attr.SetUserString("EndPointId", rvtFrame.EndPoint.Id);

                // attr.SetUserString("Group", rvtFrame.Group);
                attr.SetUserString("Angle", rvtFrame.Beta.ToString());
                attr.SetUserString("Length", rvtFrame.Length.ToString());
                attr.SetUserString("ID", rvtFrame.revitEleId.ToString());



                ot.AddCurve(crv, attr);
                curves.Add(crv);


            }


            #region Create frames Layers
            //RhinoObject[] rhinoCurves = doc.Objects.FindByObjectType(ObjectType.Curve).Where(x => x.Attributes.GetUserString("Section") != null).ToArray();
            RhinoObject[] rhinoCurves = doc.Objects.FindByObjectType(ObjectType.Curve).Where(x => x.Attributes.GetUserString("Family") != null).ToArray();

            #region get  list of Sections
            List<string> sections = new List<string>();
            foreach (var item in rhinoCurves)
            {

                ObjectAttributes att = item.Attributes;

                string sec = att.GetUserString("Section");
                sections.Add(sec);

            }
            sections = sections.Distinct().ToList();
            #endregion

            int FrameLayerIndex;
            Layer FrameLayer;
            FrameLayer = doc.Layers.FindName("RevitFrames");

            if (FrameLayer == null)
            {
                FrameLayer = new Layer();
                FrameLayer.Name = "RevitFrames";
                FrameLayerIndex = doc.Layers.Add(FrameLayer);
            }

            else
            {
                FrameLayerIndex = FrameLayer.Index;
            }



            FrameLayer = doc.Layers.FindName("RevitFrames");

            int i = 0;

            foreach (var sec in sections)
            {
                int red = i * 20 % 256;
                int green = i * 40 % 256;
                int blue = i * 60 % 256;

                System.Drawing.Color color = System.Drawing.Color.FromArgb(255, red, green, blue);

                int subLayerIndex;
                Layer subLayer = doc.Layers.FindName(sec);


                if (subLayer == null)
                {
                    subLayer = new Layer();
                    subLayer.Name = sec;
                    subLayer.Color = color;
                    subLayer.ParentLayerId = FrameLayer.Id;
                    subLayerIndex = doc.Layers.Add(subLayer);
                }

                else
                {
                    subLayer.Name = sec;
                    subLayer.Color = color;
                    subLayer.ParentLayerId = FrameLayer.Id;
                    subLayerIndex = subLayer.Index;
                }



                //RhinoObject[] objects = doc.Objects.FindByUserString("Section", sec, true);

                RhinoObject[] objects = rhinoCurves.Where(x=> x.Attributes.GetUserString("Section") == sec).ToArray();
                foreach (var item in objects)
                {
                    item.Attributes.LayerIndex = subLayerIndex;
                    item.CommitChanges();
                }

                i++;
            }

            #endregion


            return curves;
        }


        public static void CreateRhinoCurvesAsSolides(ForgeTypeId dis, UIDocument revitUiDoc, Document revitDoc, RhinoDoc doc)
        {
            List<Frame> beams = new List<Frame>();
            List<Frame> columns = new List<Frame>();
            List<Frame> revitFrames = new List<Frame>();

            List<Curve> curves = new List<Curve>();

            #region get Selected Elements 
            ICollection<ElementId> selectedIds = revitUiDoc.Selection.GetElementIds();

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
                Element element = revitUiDoc.Document.GetElement(id);

                if (element is Autodesk.Revit.DB.Group group) // Handle the case where the element is a group
                {
                    // Iterate through the elements inside the group and apply the filter
                    foreach (ElementId groupMemberId in group.GetMemberIds())
                    {
                        Element groupMember = revitUiDoc.Document.GetElement(groupMemberId);
                        if (orFilter.PassesFilter(groupMember))
                        {
                            selectedElements.Add(groupMember);
                        }
                    }
                }
                else if (orFilter.PassesFilter(element)) // Apply the filter normally for non-group elements
                {
                    selectedElements.Add(element);
                }
            }


            // Now you can filter the selected elements based on their categories
            List<FamilyInstance> selectedStrucralFraming = selectedElements
                .Where(x => x.Category.Id.IntegerValue == (int)BuiltInCategory.OST_StructuralFraming)
                .Cast<FamilyInstance>().ToList();

            List<FamilyInstance> selectedStrucralColumns = selectedElements
                .Where(x => x.Category.Id.IntegerValue == (int)BuiltInCategory.OST_StructuralColumns)
                .Cast<FamilyInstance>().ToList();

            #endregion


            // reate RhinoCurves

            foreach (var rvtFrame in selectedElements)
            {
                GeometryElement geometryElement = rvtFrame.get_Geometry(new Options());

                // Iterate through the GeometryObjects in the GeometryElement
                foreach (GeometryObject geometryObject in geometryElement)
                {
                    // Check if the geometry object is a Solid
                    if (geometryObject is Solid solid)
                    {
                       
                        // Optionally, convert the solid to Rhino geometry
                        Rhino.Geometry.Brep brep = RhinoInside.Revit.Convert.Geometry.GeometryDecoder.ToBrep(solid);

                        Rhino.RhinoDoc.ActiveDoc.Objects.Add(brep);

                    }

                    else
                    {
                        GeometryElement subGeometryElement = ((FamilyInstance)rvtFrame).Symbol.get_Geometry(new Options());


                        foreach (var item in subGeometryElement)
                        {
                            if (geometryObject is Solid subsolid)
                            {

                                // Optionally, convert the solid to Rhino geometry
                                Rhino.Geometry.Brep brep = RhinoInside.Revit.Convert.Geometry.GeometryDecoder.ToBrep(subsolid);

                                Rhino.RhinoDoc.ActiveDoc.Objects.Add(brep);

                            }



                        }
                    }
                }




               //BoundingBox b = RhinoInside.Revit.Convert.Geometry.GeometryDecoder.ToBoundingBox(rvtFrame.get_BoundingBox(revitDoc.ActiveView));
               // Rhino.Geometry.Box rhinoBox = new Rhino.Geometry.Box(b);


               // Rhino.DocObjects.Tables.ObjectTable ot = Rhino.RhinoDoc.ActiveDoc.Objects;
               // ObjectAttributes attr = new ObjectAttributes();

               // ot.AddBox(rhinoBox);

            }




        }



        #endregion

        #region notUsed
        public static void setRhinoCurve(List<RevitArea> rhinoAreas)
        {
            foreach (var area in rhinoAreas)
            {

                foreach (var item in area.outerCurves)
                {
                    Rhino.DocObjects.Tables.ObjectTable ot = Rhino.RhinoDoc.ActiveDoc.Objects;
                    ObjectAttributes attr = new ObjectAttributes();

                    attr.SetUserString("Type", "Outer");
                    ot.AddCurve(item, attr);

                    //ot.AddPoint(stPoint);
                    //ot.AddPoint(endPoint);


                }

                foreach (var item in area.innerCurves)
                {
                    Rhino.DocObjects.Tables.ObjectTable ot = Rhino.RhinoDoc.ActiveDoc.Objects;
                    ObjectAttributes attr = new ObjectAttributes();

                    attr.SetUserString("Type", "Inner");
                    ot.AddCurve(item, attr);

                }

            }


        }

        public static void updateRevitFloors(UIDocument uidoc, Document revitDoc, ForgeTypeId targetUnitConvertor, List<RevitArea> rhinoAreas)
        {


            FilteredElementCollector collector = new FilteredElementCollector(revitDoc);
            ICollection<Element> revitFloorElements = collector.OfClass(typeof(Floor)).ToElements();


            foreach (var rhinoArea in rhinoAreas)
            {
                Element FloorElment = revitFloorElements.FirstOrDefault(x => x.Id.ToString() == rhinoArea.ID);

                if (FloorElment != null)
                {



                    Floor fl = FloorElment as Floor;



                    Element sketchElement = revitDoc.GetElement(fl.SketchId);

                    if (sketchElement != null)
                    {

                        Sketch sketch = sketchElement as Sketch;

                        #region get old Profile Curves To be Removed

                        // Get sketch elements
                        var sketchEleIds = sketch.GetAllElements();

                        var oldProfileCurves = new List<ElementId>();
                        foreach (var i in sketchEleIds)
                        {

                            oldProfileCurves.Add(i);

                        }
                        #endregion


                        SketchEditScope sketchEditScope = new SketchEditScope(revitDoc, "Edit Sketch");
                        sketchEditScope.Start(sketch.Id);


                        using (Transaction transaction = new Transaction(revitDoc, "Modify sketch"))
                        {
                            transaction.Start();


                            // Delete curves from the current sketch
                            foreach (var curve in oldProfileCurves)
                            {
                                revitDoc.Delete(curve as ElementId);
                            }


                            List<Rhino.Geometry.Curve[]> crvList = new List<Rhino.Geometry.Curve[]>();

                            foreach (Curve i in rhinoArea.outerCurves)
                            {
                                crvList.Add(i.DuplicateSegments());

                            }


                            foreach (Curve i in rhinoArea.innerCurves)
                            {
                                crvList.Add(i.DuplicateSegments());


                            }


                            //IEnumerable < Rhino.Geometry.Curve > flatCrvList = from list in crvList
                            //                                                   from curve in list
                            //                                                   select curve;


                            IEnumerable<Rhino.Geometry.Curve> flatCrvList = crvList.SelectMany(x => x);


                            foreach (var i in flatCrvList)
                            {
                                Autodesk.Revit.DB.Curve c = RhinoInside.Revit.Convert.Geometry.GeometryEncoder.ToCurve(i);
                                revitDoc.Create.NewModelCurve(c, sketch.SketchPlane);
                            }


                            fl.SpanDirectionAngle = 90 * Math.PI / 180;


                            #region Old
                            //foreach (Curve i in rhinoArea.outerCurves)
                            //{
                            //    Autodesk.Revit.DB.Curve c = RhinoInside.Revit.Convert.Geometry.GeometryEncoder.ToCurve(i);
                            //    doc.Create.NewModelCurve(c, sk.SketchPlane);
                            //}

                            //foreach (Curve i in rhinoArea.innerCurves)
                            //{
                            //    Autodesk.Revit.DB.Curve c = RhinoInside.Revit.Convert.Geometry.GeometryEncoder.ToCurve(i);
                            //    doc.Create.NewModelCurve(c, sk.SketchPlane);
                            //} 
                            #endregion

                            transaction.Commit();
                        }


                        //using (Transaction transaction = new Transaction(doc, "Modify spanDirection"))
                        //{
                        //    transaction.Start();
                        //    transaction.Commit();

                        //}

                        sketchEditScope.Commit(new RoomWarningSwallower());
                    }
                }
            }

        }


        #endregion

        #region Classes
        public class RevitArea
        {
            #region  props
            public string ID { get; set; }
            public string Level { get; set; }
            public string rhinoID { get; set; }
            public double hightOffsetFromLevel { get; set; }
            public double targetZ { get; set; }

            public Point3d bottomCenterPoints { get; set; }
            public List<Point> vertices { get; set; }

            public List<string> Groups { get; set; }

            public string Group { get; set; } // Not Used

            public string RvtID { get; set; }

            public double Beta { get; set; }

            public ElementId revitEleId { get; set; }

            public string Sec { get; set; }

            public string RvtName { get; set; }

            public string aa { get; set; }
            public int rowNo { get; set; }

            public int NoOfPts { get; set; }

            public string type { get; set; }

            public List<Rhino.Geometry.Curve> crvList { get; set; }

            public List<Rhino.Geometry.Curve> outerCurves { get; set; }
            public List<Rhino.Geometry.Curve> innerCurves { get; set; }

            public CurveArrArray crvArray { get; set; }




            #endregion

            public RevitArea()
            {
                vertices = new List<Point>();
                crvArray = new CurveArrArray();
                crvList = new List<Rhino.Geometry.Curve>();
            }
            public static Point3d GetAverageOfTwoLowestVertices(Brep brep)
            {
                // Get all vertices of the Brep
                List<Point3d> allVertices = new List<Point3d>();
                foreach (BrepVertex vertex in brep.Vertices)
                {
                    allVertices.Add(vertex.Location);
                }

                // Sort vertices based on Z-coordinate
                List<Point3d> sortedVertices = allVertices.OrderBy(v => v.Z).ToList();

                // Get the two lowest vertices
                List<Point3d> lowestVertices = new List<Point3d>();
                if (sortedVertices.Count >= 2)
                {
                    lowestVertices.Add(sortedVertices[0]);
                    lowestVertices.Add(sortedVertices[1]);
                }

                // Calculate the average of the two lowest vertices
                double avgX = (lowestVertices[0].X + lowestVertices[1].X) / 2.0;
                double avgY = (lowestVertices[0].Y + lowestVertices[1].Y) / 2.0;
                double avgZ = 0;

                return new Point3d(avgX, avgY, avgZ);
            }
        }
        public class RhinoRevitFrame
        {

            #region prop

            public Frame sapFrame { get; set; }
            public Frame revitFrame { get; set; }

            public string mappedBy { get; set; }

            #endregion

            public RhinoRevitFrame()
            {
                sapFrame = new Frame();
                revitFrame = new Frame();

                mappedBy = null;
            }

            public RhinoRevitFrame(Frame sapFrame, Frame revitFrame)
            {
                this.sapFrame = sapFrame;
                this.revitFrame = revitFrame;
                mappedBy = null;
            }

            public static bool IsSameType(Frame sapElement, Frame revitElement)
            {
                if ((sapElement.MemberType == MemberType.Beam || sapElement.MemberType == MemberType.VLBracing) && (revitElement.MemberType == MemberType.Beam || revitElement.MemberType == MemberType.VLBracing))
                {
                    return true;
                }

                else if ((sapElement.MemberType == MemberType.Column || sapElement.MemberType == MemberType.SlanttedColumn) && (revitElement.MemberType == MemberType.Column || revitElement.MemberType == MemberType.SlanttedColumn))
                {
                    return true;
                }
                else
                {
                    return false;
                }

            }

            public static void updateRevitGeomertyForBeams2(RhinoDoc rhinoDoc, Document doc, List<RhinoRevitFrame> MappedBeams, ForgeTypeId dis, double margin)
            {

                XYZ StartPoint = null;
                XYZ EndPoint = null;
                double x, y, z, x1, y1, z1;
                double margin2 = 0.0;

                Level lvl0 = new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().ToList().FirstOrDefault();
                if (lvl0 == null)
                {
                    using (Transaction t = new Transaction(doc, "Create Levels"))
                    {
                        t.Start();
                        lvl0 = Level.Create(doc, 0.0);
                        lvl0.Name = "Hydra-level-1";

                        t.Commit();

                    }
                }


                List<FamilyInstance> revitElements = new List<FamilyInstance>();

                List<FamilyInstance> newlyCreatedElements = new List<FamilyInstance>();


                //RevitServices.Transactions.TransactionManager.Instance.ForceCloseTransaction();

                using (TransactionGroup tg = new TransactionGroup(doc, "aaa"))
                {

                    tg.Start();
                    var chopSize = Frame.chopSize;
                    var MappedbeamChunks = MappedBeams.Chunk(chopSize);


                    foreach (var MappedBeamChunk in MappedbeamChunks)
                    {


                        using (Transaction trns = new Transaction(doc, "trns"))
                        {
                            trns.Start();
                            FailureHandlingOptions failOpt = trns.GetFailureHandlingOptions();
                            failOpt.SetFailuresPreprocessor(new RoomWarningSwallower());
                            trns.SetFailureHandlingOptions(failOpt);


                            foreach (var mappedBeam in MappedBeamChunk)
                            {

                                try
                                {
                                    if (mappedBeam.sapFrame.ID == "523726")
                                    {

                                    }
                                    margin2 = margin * mappedBeam.revitFrame.Length;
                                    double dist1 = Point.getDistance(mappedBeam.revitFrame.StartPoint, mappedBeam.sapFrame.StartPoint);
                                    double dist2 = Point.getDistance(mappedBeam.revitFrame.EndPoint, mappedBeam.sapFrame.EndPoint);
                                    double angle = (mappedBeam.sapFrame.Beta) * Math.PI / 180;

                                    if (dist1 < margin2 && dist2 < margin2)
                                    {

                                        FamilyInstance instance = (FamilyInstance)doc.GetElement(mappedBeam.revitFrame.revitEleId);

                                        instance.LookupParameter("Cross-Section Rotation").Set(angle);


                                        revitElements.Add(instance);

                                        //continue;
                                    }

                                    else
                                    {

                                        //x = UnitUtils.ConvertToInternalUnits(mappedBeam.sapFrame.StartPoint.X, dis);
                                        //y = UnitUtils.ConvertToInternalUnits(mappedBeam.sapFrame.StartPoint.Y, dis);
                                        //z = UnitUtils.ConvertToInternalUnits(mappedBeam.sapFrame.StartPoint.Z, dis);
                                        //StartPoint = new XYZ(x, y, z);

                                        //x1 = UnitUtils.ConvertToInternalUnits(mappedBeam.sapFrame.EndPoint.X, dis);
                                        //y1 = UnitUtils.ConvertToInternalUnits(mappedBeam.sapFrame.EndPoint.Y, dis);
                                        //z1 = UnitUtils.ConvertToInternalUnits(mappedBeam.sapFrame.EndPoint.Z, dis);
                                        //EndPoint = new XYZ(x1, y1, z1);

                                        //Autodesk.Revit.DB.Line line = Autodesk.Revit.DB.Line.CreateBound(StartPoint, EndPoint);


                                        FamilyInstance instance = (FamilyInstance)doc.GetElement(mappedBeam.revitFrame.revitEleId);



                                        #region updateAnalytical
                                        #endregion

                                        #region No change in type
                                        if ((BuiltInCategory)instance.Category.Id.IntegerValue == BuiltInCategory.OST_StructuralFraming)
                                        {
                                            LocationCurve lc = instance.Location as LocationCurve;



                                            // Create a new line with the modified end points

                                            //Line newLine = Line.CreateBound(StartPoint, EndPoint);
                                            //lc.Curve = newLine;

                                            Autodesk.Revit.DB.Curve c = RhinoInside.Revit.Convert.Geometry.GeometryEncoder.ToCurve(mappedBeam.sapFrame.RhinoCurve);


                                            lc.Curve = c;

                                            //Set Parameters
                                            if (angle != 0)
                                            {
                                                instance.LookupParameter("Cross-Section Rotation").Set(angle);

                                            }

                                            double ann = instance.LookupParameter("Cross-Section Rotation").AsDouble();

                                            revitElements.Add(instance);

                                            continue;
                                        }


                                        #endregion

                                        #region Change in type (from Column to beam)
                                        else
                                        {

                                            // handle error in case revit section of new beam not found
                                            Dictionary<string, FamilySymbol> dic = new Dictionary<string, FamilySymbol>();
                                            List<Frame> Sapframes = new List<Frame>() { mappedBeam.sapFrame };


                                            FamilySymbol revitSection = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralFraming).WhereElementIsElementType().Cast<FamilySymbol>().ToList().FirstOrDefault();

                                            if (!revitSection.IsActive)
                                                revitSection.Activate();

                                            Autodesk.Revit.DB.Curve c = RhinoInside.Revit.Convert.Geometry.GeometryEncoder.ToCurve(mappedBeam.sapFrame.RhinoCurve);

                                            FamilyInstance newInstance = doc.Create.NewFamilyInstance(c, revitSection, lvl0, StructuralType.Beam);

                                            if (angle != 0)
                                            {
                                                newInstance.LookupParameter("Cross-Section Rotation").Set(angle);

                                            }

                                            Parameter sss = newInstance.LookupParameter("z Justification");
                                            int zjustf = (int)ZJustification.Center;

                                            Autodesk.Revit.DB.Structure.StructuralFramingUtils.DisallowJoinAtEnd(newInstance, 0);
                                            Autodesk.Revit.DB.Structure.StructuralFramingUtils.DisallowJoinAtEnd(newInstance, 1);


                                            #region setParameters

                                            //Set Parameters
                                            RhinoObject rhinoObj = rhinoDoc.Objects.FindId(new Guid(mappedBeam.sapFrame.RhinoID));
                                            ObjectAttributes attr = rhinoObj.Attributes.Duplicate();
                                            attr.SetUserString("ID", newInstance.Id.ToString());
                                            rhinoObj.Attributes = attr;
                                            rhinoDoc.Objects.ModifyAttributes(rhinoObj, attr, true);

                                            #endregion

                                            doc.Delete(instance.Id);

                                            //doc.Regenerate();

                                            //analyManger.AddAssociation(analytInstance.Id, newInstance.Id);
                                            //continue;

                                        }




                                        #endregion




                                    }
                                }
                                catch (Exception ex2)
                                {
                                    System.Diagnostics.StackTrace trace = new System.Diagnostics.StackTrace(ex2, true);
                                    RevitUtils.Log.Add("Error in updateRevitGeomertyForBeams2 Method  sapBeam no " + mappedBeam.sapFrame.ID + ex2.Message + trace.GetFrame(trace.FrameCount - 1).GetFileLineNumber());
                                    continue;
                                }
                            }
                            trns.Commit();

                        }
                    }


                    tg.Assimilate();



                }






            }

            public static void updateRevitGeomertyForBeams3(RhinoDoc rhinoDoc, Document doc, List<RhinoRevitFrame> MappedBeams, ForgeTypeId dis, double margin)
            {

                XYZ StartPoint = null;
                XYZ EndPoint = null;
                double x, y, z, x1, y1, z1;
                double margin2 = 0.0;

                Level lvl0 = new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().ToList().FirstOrDefault();
                if (lvl0 == null)
                {
                    using (Transaction t = new Transaction(doc, "Create Levels"))
                    {
                        t.Start();
                        lvl0 = Level.Create(doc, 0.0);
                        lvl0.Name = "Hydra-level-1";

                        t.Commit();

                    }
                }


                List<FamilyInstance> revitElements = new List<FamilyInstance>();

                List<FamilyInstance> newlyCreatedElements = new List<FamilyInstance>();


                //RevitServices.Transactions.TransactionManager.Instance.ForceCloseTransaction();

                using (TransactionGroup tg = new TransactionGroup(doc, "aaa"))
                {

                    tg.Start();
                    var chopSize = Frame.chopSize;
                    var MappedbeamChunks = MappedBeams.Chunk(chopSize);


                    foreach (var MappedBeamChunk in MappedbeamChunks)
                    {


                        using (Transaction trns = new Transaction(doc, "trns"))
                        {
                            trns.Start();
                            FailureHandlingOptions failOpt = trns.GetFailureHandlingOptions();
                            failOpt.SetFailuresPreprocessor(new RoomWarningSwallower());
                            trns.SetFailureHandlingOptions(failOpt);


                            foreach (var mappedBeam in MappedBeamChunk)
                            {

                                try
                                {

                                    margin2 = margin * mappedBeam.revitFrame.Length;
                                    double dist1 = Point.getDistance(mappedBeam.revitFrame.StartPoint, mappedBeam.sapFrame.StartPoint);
                                    double dist2 = Point.getDistance(mappedBeam.revitFrame.EndPoint, mappedBeam.sapFrame.EndPoint);
                                    double dist3 = Point.getDistance(mappedBeam.revitFrame.MidPoint, mappedBeam.sapFrame.MidPoint);

                                    double angle = (mappedBeam.sapFrame.Beta) * Math.PI / 180;

                                    if (dist1 < margin2 && dist2 < margin2 && dist3 < margin2)
                                    {

                                        FamilyInstance instance = (FamilyInstance)doc.GetElement(mappedBeam.revitFrame.revitEleId);

                                        instance.LookupParameter("Cross-Section Rotation").Set(angle);


                                        revitElements.Add(instance);

                                        //continue;
                                    }

                                    else
                                    {

                                        //x = UnitUtils.ConvertToInternalUnits(mappedBeam.sapFrame.StartPoint.X, dis);
                                        //y = UnitUtils.ConvertToInternalUnits(mappedBeam.sapFrame.StartPoint.Y, dis);
                                        //z = UnitUtils.ConvertToInternalUnits(mappedBeam.sapFrame.StartPoint.Z, dis);
                                        //StartPoint = new XYZ(x, y, z);

                                        //x1 = UnitUtils.ConvertToInternalUnits(mappedBeam.sapFrame.EndPoint.X, dis);
                                        //y1 = UnitUtils.ConvertToInternalUnits(mappedBeam.sapFrame.EndPoint.Y, dis);
                                        //z1 = UnitUtils.ConvertToInternalUnits(mappedBeam.sapFrame.EndPoint.Z, dis);
                                        //EndPoint = new XYZ(x1, y1, z1);

                                        //Autodesk.Revit.DB.Line line = Autodesk.Revit.DB.Line.CreateBound(StartPoint, EndPoint);


                                        FamilyInstance instance = (FamilyInstance)doc.GetElement(mappedBeam.revitFrame.revitEleId);



                                        #region updateAnalytical
                                        #endregion

                                        #region No change in type
                                        if ((BuiltInCategory)instance.Category.Id.IntegerValue == BuiltInCategory.OST_StructuralFraming)
                                        {
                                            LocationCurve lc = instance.Location as LocationCurve;


                                            // Create a new line with the modified end points

                                            //Line newLine = Line.CreateBound(StartPoint, EndPoint);
                                            //lc.Curve = newLine;

                                            Autodesk.Revit.DB.Curve c = RhinoInside.Revit.Convert.Geometry.GeometryEncoder.ToCurve(mappedBeam.sapFrame.RhinoCurve);



                                            lc.Curve = c;
                                            //Set Parameters
                                            if (angle != 0)
                                            {
                                                instance.LookupParameter("Cross-Section Rotation").Set(angle);

                                            }

                                            double ann = instance.LookupParameter("Cross-Section Rotation").AsDouble();

                                            revitElements.Add(instance);

                                            continue;
                                        }


                                        #endregion

                                        #region Change in type (from Column to beam)
                                        else
                                        {

                                            // handle error in case revit section of new beam not found
                                            Dictionary<string, FamilySymbol> dic = new Dictionary<string, FamilySymbol>();
                                            List<Frame> Sapframes = new List<Frame>() { mappedBeam.sapFrame };


                                            FamilySymbol revitSection = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralFraming).WhereElementIsElementType().Cast<FamilySymbol>().ToList().FirstOrDefault();

                                            if (!revitSection.IsActive)
                                                revitSection.Activate();

                                            Autodesk.Revit.DB.Curve c = RhinoInside.Revit.Convert.Geometry.GeometryEncoder.ToCurve(mappedBeam.sapFrame.RhinoCurve);

                                            FamilyInstance newInstance = doc.Create.NewFamilyInstance(c, revitSection, lvl0, StructuralType.Beam);

                                            if (angle != 0)
                                            {
                                                newInstance.LookupParameter("Cross-Section Rotation").Set(angle);

                                            }

                                            Parameter sss = newInstance.LookupParameter("z Justification");
                                            int zjustf = (int)ZJustification.Center;

                                            Autodesk.Revit.DB.Structure.StructuralFramingUtils.DisallowJoinAtEnd(newInstance, 0);
                                            Autodesk.Revit.DB.Structure.StructuralFramingUtils.DisallowJoinAtEnd(newInstance, 1);


                                            #region setParameters

                                            //Set Parameters
                                            RhinoObject rhinoObj = rhinoDoc.Objects.FindId(new Guid(mappedBeam.sapFrame.RhinoID));
                                            ObjectAttributes attr = rhinoObj.Attributes.Duplicate();
                                            attr.SetUserString("ID", newInstance.Id.ToString());
                                            rhinoObj.Attributes = attr;
                                            rhinoDoc.Objects.ModifyAttributes(rhinoObj, attr, true);

                                            #endregion

                                            doc.Delete(instance.Id);

                                            //doc.Regenerate();

                                            //analyManger.AddAssociation(analytInstance.Id, newInstance.Id);
                                            //continue;

                                        }




                                        #endregion




                                    }
                                }
                                catch (Exception ex2)
                                {
                                    System.Diagnostics.StackTrace trace = new System.Diagnostics.StackTrace(ex2, true);
                                    RevitUtils.Log.Add("Error in updateRevitGeomertyForBeams2 Method  sapBeam no " + mappedBeam.sapFrame.ID + ex2.Message + trace.GetFrame(trace.FrameCount - 1).GetFileLineNumber());
                                    continue;
                                }
                            }
                            trns.Commit();

                        }
                    }


                    tg.Assimilate();



                }






            }


            public static void updateRevitGeomertyForBeams4(RhinoDoc rhinoDoc, Document doc, List<RhinoRevitFrame> MappedBeams, ForgeTypeId dis, double margin)
            {

                XYZ StartPoint = null;
                XYZ EndPoint = null;
                double x, y, z, x1, y1, z1;
                double margin2 = 0.0;

                Level lvl0 = new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().ToList().FirstOrDefault();
                if (lvl0 == null)
                {
                    using (Transaction t = new Transaction(doc, "Create Levels"))
                    {
                        t.Start();
                        lvl0 = Level.Create(doc, 0.0);
                        lvl0.Name = "Hydra-level-1";

                        t.Commit();

                    }
                }


                List<FamilyInstance> revitElements = new List<FamilyInstance>();

                List<FamilyInstance> newlyCreatedElements = new List<FamilyInstance>();


                //RevitServices.Transactions.TransactionManager.Instance.ForceCloseTransaction();

                using (TransactionGroup tg = new TransactionGroup(doc, "aaa"))
                {

                    tg.Start();
                    var chopSize = Frame.chopSize;
                    var MappedbeamChunks = MappedBeams.Chunk(chopSize);


                    foreach (var MappedBeamChunk in MappedbeamChunks)
                    {


                        foreach (var mappedBeam in MappedBeamChunk)
                        {

                            using (Transaction trns = new Transaction(doc, "trns"))
                            {
                                trns.Start();
                                FailureHandlingOptions failOpt = trns.GetFailureHandlingOptions();
                                failOpt.SetFailuresPreprocessor(new RoomWarningSwallower());
                                trns.SetFailureHandlingOptions(failOpt);

                                try
                                {

                                    margin2 = margin * mappedBeam.revitFrame.Length;
                                    double dist1 = Point.getDistance(mappedBeam.revitFrame.StartPoint, mappedBeam.sapFrame.StartPoint);
                                    double dist2 = Point.getDistance(mappedBeam.revitFrame.EndPoint, mappedBeam.sapFrame.EndPoint);
                                    double dist3 = Point.getDistance(mappedBeam.revitFrame.MidPoint, mappedBeam.sapFrame.MidPoint);

                                    double angle = (mappedBeam.sapFrame.Beta) * Math.PI / 180;

                                    if (dist1 < margin2 && dist2 < margin2 && dist3 < margin2)
                                    {

                                        FamilyInstance instance = (FamilyInstance)doc.GetElement(mappedBeam.revitFrame.revitEleId);

                                        instance.LookupParameter("Cross-Section Rotation").Set(angle);


                                        revitElements.Add(instance);

                                        //continue;
                                    }

                                    else
                                    {

                                        //x = UnitUtils.ConvertToInternalUnits(mappedBeam.sapFrame.StartPoint.X, dis);
                                        //y = UnitUtils.ConvertToInternalUnits(mappedBeam.sapFrame.StartPoint.Y, dis);
                                        //z = UnitUtils.ConvertToInternalUnits(mappedBeam.sapFrame.StartPoint.Z, dis);
                                        //StartPoint = new XYZ(x, y, z);

                                        //x1 = UnitUtils.ConvertToInternalUnits(mappedBeam.sapFrame.EndPoint.X, dis);
                                        //y1 = UnitUtils.ConvertToInternalUnits(mappedBeam.sapFrame.EndPoint.Y, dis);
                                        //z1 = UnitUtils.ConvertToInternalUnits(mappedBeam.sapFrame.EndPoint.Z, dis);
                                        //EndPoint = new XYZ(x1, y1, z1);

                                        //Autodesk.Revit.DB.Line line = Autodesk.Revit.DB.Line.CreateBound(StartPoint, EndPoint);


                                        FamilyInstance instance = (FamilyInstance)doc.GetElement(mappedBeam.revitFrame.revitEleId);



                                        #region updateAnalytical
                                        #endregion

                                        #region No change in type
                                        if ((BuiltInCategory)instance.Category.Id.IntegerValue == BuiltInCategory.OST_StructuralFraming)
                                        {
                                            LocationCurve lc = instance.Location as LocationCurve;


                                            // Create a new line with the modified end points

                                            //Line newLine = Line.CreateBound(StartPoint, EndPoint);
                                            //lc.Curve = newLine;

                                            Autodesk.Revit.DB.Curve c = RhinoInside.Revit.Convert.Geometry.GeometryEncoder.ToCurve(mappedBeam.sapFrame.RhinoCurve);


                                            lc.Curve = c;
                                            //Set Parameters
                                            if (angle != 0)
                                            {
                                                instance.LookupParameter("Cross-Section Rotation").Set(angle);

                                            }

                                            double ann = instance.LookupParameter("Cross-Section Rotation").AsDouble();

                                            revitElements.Add(instance);

                                            //continue;
                                        }


                                        #endregion

                                        #region Change in type (from Column to beam)
                                        else
                                        {

                                            // handle error in case revit section of new beam not found
                                            Dictionary<string, FamilySymbol> dic = new Dictionary<string, FamilySymbol>();
                                            List<Frame> Sapframes = new List<Frame>() { mappedBeam.sapFrame };


                                            FamilySymbol revitSection = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralFraming).WhereElementIsElementType().Cast<FamilySymbol>().ToList().FirstOrDefault();

                                            if (!revitSection.IsActive)
                                                revitSection.Activate();

                                            Autodesk.Revit.DB.Curve c = RhinoInside.Revit.Convert.Geometry.GeometryEncoder.ToCurve(mappedBeam.sapFrame.RhinoCurve);

                                            FamilyInstance newInstance = doc.Create.NewFamilyInstance(c, revitSection, lvl0, StructuralType.Beam);

                                            if (angle != 0)
                                            {
                                                newInstance.LookupParameter("Cross-Section Rotation").Set(angle);

                                            }

                                            Parameter sss = newInstance.LookupParameter("z Justification");
                                            int zjustf = (int)ZJustification.Center;

                                            Autodesk.Revit.DB.Structure.StructuralFramingUtils.DisallowJoinAtEnd(newInstance, 0);
                                            Autodesk.Revit.DB.Structure.StructuralFramingUtils.DisallowJoinAtEnd(newInstance, 1);


                                            #region setParameters

                                            //Set Parameters
                                            RhinoObject rhinoObj = rhinoDoc.Objects.FindId(new Guid(mappedBeam.sapFrame.RhinoID));
                                            ObjectAttributes attr = rhinoObj.Attributes.Duplicate();
                                            attr.SetUserString("ID", newInstance.Id.ToString());
                                            rhinoObj.Attributes = attr;
                                            rhinoDoc.Objects.ModifyAttributes(rhinoObj, attr, true);

                                            #endregion

                                            doc.Delete(instance.Id);

                                            //doc.Regenerate();

                                            //analyManger.AddAssociation(analytInstance.Id, newInstance.Id);
                                            //continue;

                                        }




                                        #endregion




                                    }
                                }
                                catch (Exception ex2)
                                {
                                    System.Diagnostics.StackTrace trace = new System.Diagnostics.StackTrace(ex2, true);
                                    RevitUtils.Log.Add("Error in updateRevitGeomertyForBeams2 Method  sapBeam no " + mappedBeam.sapFrame.ID + ex2.Message + trace.GetFrame(trace.FrameCount - 1).GetFileLineNumber());
                                    continue;
                                }


                                trns.Commit();
                            }


                        }
                    }


                    tg.Assimilate();



                }

            }


            public static void updateRevitGeomertyForColumns2(RhinoDoc rhinoDoc, Document doc, List<RhinoRevitFrame> MappedColumns, ForgeTypeId dis, double margin)
            {


                XYZ StartPoint = null;
                XYZ EndPoint = null;
                double x, y, z, x1, y1, z1;
                double margin2 = 0.0;
                Level lvl0;
                Level lvl1;

                lvl0 = new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().FirstOrDefault();
                lvl1 = new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().ElementAtOrDefault(1);

                if (lvl0 == null)
                {
                    using (Transaction t = new Transaction(doc, "Create Levels"))
                    {
                        t.Start();
                        lvl0 = Level.Create(doc, 0.0);
                        lvl0.Name = "Hydra-level-1";

                        t.Commit();

                    }
                }

                if (lvl1 == null)
                {
                    using (Transaction t = new Transaction(doc, "Create Levels"))
                    {
                        t.Start();
                        lvl1 = Level.Create(doc, 10.0);
                        lvl1.Name = "Hydra-level-2";

                        t.Commit();

                    }
                }


                List<FamilyInstance> revitElements = new List<FamilyInstance>();
                List<FamilyInstance> newlyCreatedElements = new List<FamilyInstance>();



                RevitServices.Transactions.TransactionManager.Instance.ForceCloseTransaction();

                using (TransactionGroup tg = new TransactionGroup(doc, "aaa"))
                {

                    tg.Start();
                    var chopSize = Frame.chopSize;
                    var MappedColumnsChunks = MappedColumns.Chunk(chopSize);

                    foreach (var MappedColumnChunk in MappedColumnsChunks)
                    {

                        using (Transaction trns = new Transaction(doc, "trns"))
                        {
                            trns.Start();
                            FailureHandlingOptions failOpt = trns.GetFailureHandlingOptions();
                            failOpt.SetFailuresPreprocessor(new RoomWarningSwallower());
                            trns.SetFailureHandlingOptions(failOpt);


                            foreach (var mappedColumn in MappedColumnChunk.AsParallel())
                            {

                                try
                                {


                                    margin2 = margin * mappedColumn.revitFrame.Length;
                                    double dist1 = Point.getDistance(mappedColumn.revitFrame.StartPoint, mappedColumn.sapFrame.StartPoint);
                                    double dist2 = Point.getDistance(mappedColumn.revitFrame.EndPoint, mappedColumn.sapFrame.EndPoint);
                                    bool isNotMatched = true;


                                    if (dist1 < margin2 && dist2 < margin2)
                                    {

                                        FamilyInstance instance = (FamilyInstance)doc.GetElement(mappedColumn.revitFrame.revitEleId);
                                        revitElements.Add(instance);

                                        #region SetRotationAngle 

                                        x = UnitUtils.ConvertToInternalUnits(mappedColumn.sapFrame.StartPoint.X, dis);
                                        y = UnitUtils.ConvertToInternalUnits(mappedColumn.sapFrame.StartPoint.Y, dis);
                                        z = UnitUtils.ConvertToInternalUnits(mappedColumn.sapFrame.StartPoint.Z, dis);
                                        StartPoint = new XYZ(x, y, z);

                                        x1 = UnitUtils.ConvertToInternalUnits(mappedColumn.sapFrame.EndPoint.X, dis);
                                        y1 = UnitUtils.ConvertToInternalUnits(mappedColumn.sapFrame.EndPoint.Y, dis);
                                        z1 = UnitUtils.ConvertToInternalUnits(mappedColumn.sapFrame.EndPoint.Z, dis);
                                        EndPoint = new XYZ(x1, y1, z1);

                                        double currentRotation = 0;
                                        Parameter colStyleParam = instance.LookupParameter("Column Style");
                                        int ColumnStyle = instance.LookupParameter("Column Style").AsInteger();


                                        if (ColumnStyle == (int)SlantedOrVerticalColumnType.CT_Vertical)
                                        {
                                            LocationPoint pt = instance.Location as LocationPoint;
                                            currentRotation = pt.Rotation;
                                            Line line2 = Line.CreateBound(StartPoint, EndPoint);
                                            double col_angle = (mappedColumn.sapFrame.Beta) * Math.PI / 180;


                                            if (currentRotation != col_angle)
                                            {
                                                // ElementTransformUtils.RotateElement(doc, instance.Id, line2, col_angle - currentRotation);

                                                ElementTransformUtils.RotateElement(doc, instance.Id, line2, col_angle - currentRotation);


                                            }

                                        }



                                        else
                                        {
                                            instance.LookupParameter("Cross-Section Rotation").Set(mappedColumn.sapFrame.Beta * Math.PI / 180);
                                        }

                                        #endregion



                                        continue;
                                    }


                                    else
                                    {
                                        //AnalyticalToPhysicalAssociationManager analyManger = AnalyticalToPhysicalAssociationManager.GetAnalyticalToPhysicalAssociationManager(doc);

                                        x = UnitUtils.ConvertToInternalUnits(mappedColumn.sapFrame.StartPoint.X, dis);
                                        y = UnitUtils.ConvertToInternalUnits(mappedColumn.sapFrame.StartPoint.Y, dis);
                                        z = UnitUtils.ConvertToInternalUnits(mappedColumn.sapFrame.StartPoint.Z, dis);
                                        StartPoint = new XYZ(x, y, z);

                                        x1 = UnitUtils.ConvertToInternalUnits(mappedColumn.sapFrame.EndPoint.X, dis);
                                        y1 = UnitUtils.ConvertToInternalUnits(mappedColumn.sapFrame.EndPoint.Y, dis);
                                        z1 = UnitUtils.ConvertToInternalUnits(mappedColumn.sapFrame.EndPoint.Z, dis);

                                        EndPoint = new XYZ(x1, y1, z1);


                                        Autodesk.Revit.DB.Curve line = Line.CreateBound(StartPoint, EndPoint);

                                        Autodesk.Revit.DB.Curve line3 = Line.CreateBound(EndPoint, StartPoint);


                                        FamilyInstance instance = (FamilyInstance)doc.GetElement(mappedColumn.revitFrame.revitEleId);


                                        #region No change in the type 
                                        if ((BuiltInCategory)instance.Category.Id.IntegerValue == BuiltInCategory.OST_StructuralColumns)
                                        {
                                            revitElements.Add(instance);

                                            Parameter colStyleParam = instance.LookupParameter("Column Style");

                                            int ColumnStyle = instance.LookupParameter("Column Style").AsInteger();


                                            if (ColumnStyle == (int)SlantedOrVerticalColumnType.CT_Vertical && mappedColumn.sapFrame.MemberType == MemberType.Column)
                                            {

                                                isNotMatched = false;

                                                #region handling levels elevations

                                                (instance.Location as LocationPoint).Point = StartPoint;
                                                Level baseLevel = (Level)doc.GetElement(instance.LookupParameter("Base Level").AsElementId());
                                                Level topLevel = (Level)doc.GetElement(instance.LookupParameter("Top Level").AsElementId());

                                                double basePointElevation = Convert.ToDouble(mappedColumn.sapFrame.StartPoint.Z);
                                                double topPointElevation = Convert.ToDouble(mappedColumn.sapFrame.EndPoint.Z);

                                                if (basePointElevation > topPointElevation)
                                                {
                                                    basePointElevation = Convert.ToDouble(mappedColumn.sapFrame.EndPoint.Z);
                                                    topPointElevation = Convert.ToDouble(mappedColumn.sapFrame.StartPoint.Z);
                                                }


                                                double baseLevelElevation = UnitUtils.ConvertFromInternalUnits(baseLevel.ProjectElevation, dis);
                                                double topLevelElevation = UnitUtils.ConvertFromInternalUnits(topLevel.ProjectElevation, dis);


                                                double baseOffset = basePointElevation - baseLevelElevation;
                                                double topOffset = topPointElevation - topLevelElevation;


                                                double base_offset = UnitUtils.ConvertToInternalUnits(baseOffset, dis);
                                                double top_offset = UnitUtils.ConvertToInternalUnits(topOffset, dis);


                                                if ((topLevelElevation + topOffset) == baseLevelElevation)
                                                {
                                                    instance.LookupParameter("Base Offset").Set(base_offset);
                                                    instance.LookupParameter("Top Offset").Set(top_offset);

                                                }
                                                else
                                                {
                                                    instance.LookupParameter("Top Offset").Set(top_offset);
                                                    instance.LookupParameter("Base Offset").Set(base_offset);
                                                }



                                                #endregion

                                                #region SetRotationAngle 
                                                LocationPoint pt = instance.Location as LocationPoint;
                                                double currentRotation = pt.Rotation;



                                                Line line2 = Line.CreateBound(StartPoint, EndPoint);
                                                double col_angle = (mappedColumn.sapFrame.Beta) * Math.PI / 180;


                                                if (currentRotation != col_angle)
                                                {
                                                    // to be checked 30/5/2023
                                                    //  ElementTransformUtils.RotateElement(doc, instance.Id, line2, col_angle - currentRotation);

                                                }



                                                #endregion

                                                //doc.Regenerate();
                                            }


                                            else if (ColumnStyle == (int)SlantedOrVerticalColumnType.CT_EndPoint && mappedColumn.sapFrame.MemberType == MemberType.SlanttedColumn)
                                            {

                                                isNotMatched = false;

                                                if (z > z1)
                                                {
                                                    (instance.Location as LocationCurve).Curve = line3;
                                                }
                                                else
                                                {
                                                    (instance.Location as LocationCurve).Curve = line;
                                                }



                                                Line line2 = Line.CreateBound(StartPoint, EndPoint);
                                                double col_angle = (mappedColumn.sapFrame.Beta) * Math.PI / 180;
                                                // ElementTransformUtils.RotateElement(doc, instance.Id, line2, col_angle);
                                                // doc.Regenerate();

                                            }


                                        }

                                        #endregion



                                        #region change in the type

                                        if (isNotMatched)
                                        {
                                            // case of VrColumn (to vr column)
                                            if (mappedColumn.sapFrame.MemberType == MemberType.Column)
                                            {

                                                Dictionary<string, FamilySymbol> dic = new Dictionary<string, FamilySymbol>();
                                                List<Frame> Sapframes = new List<Frame>() { mappedColumn.sapFrame };

                                                FamilySymbol revitSection = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralColumns).WhereElementIsElementType().Cast<FamilySymbol>().ToList().FirstOrDefault();
                                                if (!revitSection.IsActive)
                                                    revitSection.Activate();


                                                FamilyInstance newInstance = doc.Create.NewFamilyInstance(StartPoint, revitSection, lvl0, StructuralType.Column);

                                                newlyCreatedElements.Add(newInstance);
                                                revitElements.Add(instance);

                                                //analyManger.AddAssociation(analytInstance.Id, newInstance.Id);

                                                #region handling levels elevations
                                                Parameter pbase = newInstance.LookupParameter("Base Level");
                                                Parameter ptop = newInstance.LookupParameter("Top Level");

                                                //
                                                pbase.Set(lvl0.Id);
                                                ptop.Set(lvl1.Id);

                                                Level baselvl = doc.GetElement(pbase.AsElementId()) as Level;
                                                Level toplvl = doc.GetElement(ptop.AsElementId()) as Level;


                                                double base_elevation = UnitUtils.ConvertFromInternalUnits(baselvl.ProjectElevation, dis);
                                                double top_elevation = UnitUtils.ConvertFromInternalUnits(toplvl.ProjectElevation, dis);

                                                double base_elev = Convert.ToDouble(mappedColumn.sapFrame.StartPoint.Z);
                                                double top_elv = Convert.ToDouble(mappedColumn.sapFrame.EndPoint.Z);


                                                double base_offset = UnitUtils.ConvertToInternalUnits(base_elev - base_elevation, dis);
                                                double top_offset = UnitUtils.ConvertToInternalUnits(top_elv - top_elevation, dis);

                                                if (base_elev - base_elevation > top_elv - top_elevation)
                                                {
                                                    base_elev = Convert.ToDouble(mappedColumn.sapFrame.EndPoint.Z);
                                                    top_elv = Convert.ToDouble(mappedColumn.sapFrame.StartPoint.Z);

                                                    base_offset = UnitUtils.ConvertToInternalUnits(base_elev - base_elevation, dis);
                                                    top_offset = UnitUtils.ConvertToInternalUnits(top_elv - top_elevation, dis);
                                                }

                                                newInstance.LookupParameter("Top Offset").Set(top_offset);
                                                newInstance.LookupParameter("Base Offset").Set(base_offset);



                                                #endregion

                                                #region Rotation angle
                                                Line line2 = Line.CreateBound(StartPoint, EndPoint);
                                                double col_angle = (mappedColumn.sapFrame.Beta) * Math.PI / 180;

                                                // ElementTransformUtils.RotateElement(doc, newInstance.Id, line2, col_angle);


                                                #endregion

                                                #region setParameters

                                                //Set Parameters
                                                RhinoObject rhinoObj = rhinoDoc.Objects.FindId(new Guid(mappedColumn.sapFrame.RhinoID));
                                                ObjectAttributes attr = rhinoObj.Attributes.Duplicate();
                                                attr.SetUserString("ID", newInstance.Id.ToString());
                                                rhinoObj.Attributes = attr;

                                                #endregion

                                                doc.Delete(instance.Id);

                                                //doc.Regenerate();



                                            }


                                            // case of SlanttedColumn (to slantted column)
                                            if (mappedColumn.sapFrame.MemberType == MemberType.SlanttedColumn)
                                            {
                                                FamilySymbol revitSection = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralColumns).WhereElementIsElementType().Cast<FamilySymbol>().ToList().FirstOrDefault(); ;

                                                if (!revitSection.IsActive)
                                                    revitSection.Activate();


                                                FamilyInstance newInstance = doc.Create.NewFamilyInstance(StartPoint, revitSection, lvl0, StructuralType.Column);
                                                newlyCreatedElements.Add(newInstance);
                                                //analyManger.AddAssociation(analytInstance.Id, newInstance.Id);
                                                revitElements.Add(instance);


                                                Parameter param = newInstance.LookupParameter("Column Style");
                                                param.Set((int)SlantedOrVerticalColumnType.CT_EndPoint);
                                                LocationCurve instanceCurve = (LocationCurve)newInstance.Location;

                                                if (instanceCurve != null)
                                                {
                                                    if (z > z1)
                                                    {
                                                        instanceCurve.Curve = line3;
                                                    }
                                                    else
                                                    {
                                                        instanceCurve.Curve = line;

                                                    }

                                                }



                                                #region Rotation angle

                                                double col_angle = (mappedColumn.sapFrame.Beta) * Math.PI / 180;
                                                // newInstance.LookupParameter("Cross-Section Rotation").Set(col_angle);


                                                #endregion

                                                #region setParameters

                                                //Set Parameters
                                                RhinoObject rhinoObj = rhinoDoc.Objects.FindId(new Guid(mappedColumn.sapFrame.RhinoID));
                                                ObjectAttributes attr = rhinoObj.Attributes.Duplicate();
                                                attr.SetUserString("ID", newInstance.Id.ToString());
                                                rhinoObj.Attributes = attr;

                                                #endregion


                                                doc.Delete(instance.Id);
                                                //doc.Regenerate();





                                            }
                                        }




                                        #endregion






                                    }

                                }


                                catch (Exception ex2)
                                {
                                    System.Diagnostics.StackTrace trace = new System.Diagnostics.StackTrace(ex2, true);
                                    RevitUtils.Log.Add("Error in updateRevitGeomertyForColumns2 Method  sapCol no " + mappedColumn.sapFrame.ID + ex2.Message + trace.GetFrame(trace.FrameCount - 1).GetFileLineNumber());
                                    continue;
                                }


                            }


                            trns.Commit();

                        }
                    }

                    tg.Assimilate();


                }

            }

            public static void updateRevitGeomertyForColumns2BasedOnSlanttedStyle(RhinoDoc rhinoDoc, Document doc, List<RhinoRevitFrame> MappedColumns, ForgeTypeId dis, double margin)
            {


                XYZ StartPoint = null;
                XYZ EndPoint = null;
                double x, y, z, x1, y1, z1;
                double margin2 = 0.0;
                Level lvl0;
                Level lvl1;

                lvl0 = new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().FirstOrDefault();
                lvl1 = new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().ElementAtOrDefault(1);

                if (lvl0 == null)
                {
                    using (Transaction t = new Transaction(doc, "Create Levels"))
                    {
                        t.Start();
                        lvl0 = Level.Create(doc, 0.0);
                        lvl0.Name = "Hydra-level-1";

                        t.Commit();

                    }
                }

                if (lvl1 == null)
                {
                    using (Transaction t = new Transaction(doc, "Create Levels"))
                    {
                        t.Start();
                        lvl1 = Level.Create(doc, 10.0);
                        lvl1.Name = "Hydra-level-2";

                        t.Commit();

                    }
                }


                List<FamilyInstance> revitElements = new List<FamilyInstance>();
                List<FamilyInstance> newlyCreatedElements = new List<FamilyInstance>();



                RevitServices.Transactions.TransactionManager.Instance.ForceCloseTransaction();

                using (TransactionGroup tg = new TransactionGroup(doc, "aaa"))
                {

                    tg.Start();
                    var chopSize = Frame.chopSize;
                    var MappedColumnsChunks = MappedColumns.Chunk(chopSize);

                    foreach (var MappedColumnChunk in MappedColumnsChunks)
                    {

                        using (Transaction trns = new Transaction(doc, "trns"))
                        {
                            trns.Start();
                            FailureHandlingOptions failOpt = trns.GetFailureHandlingOptions();
                            failOpt.SetFailuresPreprocessor(new RoomWarningSwallower());
                            trns.SetFailureHandlingOptions(failOpt);


                            foreach (var mappedColumn in MappedColumnChunk.AsParallel())
                            {

                                try
                                {


                                    margin2 = margin * mappedColumn.revitFrame.Length;
                                    double dist1 = Point.getDistance(mappedColumn.revitFrame.StartPoint, mappedColumn.sapFrame.StartPoint);
                                    double dist2 = Point.getDistance(mappedColumn.revitFrame.EndPoint, mappedColumn.sapFrame.EndPoint);
                                    bool isNotMatched = true;


                                    if (dist1 < margin2 && dist2 < margin2)
                                    {

                                        FamilyInstance instance = (FamilyInstance)doc.GetElement(mappedColumn.revitFrame.revitEleId);
                                        revitElements.Add(instance);

                                        #region SetRotationAngle 

                                        x = UnitUtils.ConvertToInternalUnits(mappedColumn.sapFrame.StartPoint.X, dis);
                                        y = UnitUtils.ConvertToInternalUnits(mappedColumn.sapFrame.StartPoint.Y, dis);
                                        z = UnitUtils.ConvertToInternalUnits(mappedColumn.sapFrame.StartPoint.Z, dis);
                                        StartPoint = new XYZ(x, y, z);

                                        x1 = UnitUtils.ConvertToInternalUnits(mappedColumn.sapFrame.EndPoint.X, dis);
                                        y1 = UnitUtils.ConvertToInternalUnits(mappedColumn.sapFrame.EndPoint.Y, dis);
                                        z1 = UnitUtils.ConvertToInternalUnits(mappedColumn.sapFrame.EndPoint.Z, dis);
                                        EndPoint = new XYZ(x1, y1, z1);

                                        double currentRotation = 0;
                                        Parameter colStyleParam = instance.LookupParameter("Column Style");
                                        int ColumnStyle = instance.LookupParameter("Column Style").AsInteger();


                                        if (ColumnStyle == (int)SlantedOrVerticalColumnType.CT_Vertical)
                                        {
                                            LocationPoint pt = instance.Location as LocationPoint;
                                            currentRotation = pt.Rotation;
                                            Line line2 = Line.CreateBound(StartPoint, EndPoint);
                                            double col_angle = (mappedColumn.sapFrame.Beta) * Math.PI / 180;


                                            if (currentRotation != col_angle)
                                            {
                                                // ElementTransformUtils.RotateElement(doc, instance.Id, line2, col_angle - currentRotation);

                                                ElementTransformUtils.RotateElement(doc, instance.Id, line2, col_angle - currentRotation);


                                            }

                                        }



                                        else
                                        {
                                            instance.LookupParameter("Cross-Section Rotation").Set(mappedColumn.sapFrame.Beta * Math.PI / 180);
                                        }

                                        #endregion



                                        continue;
                                    }


                                    else
                                    {
                                        //AnalyticalToPhysicalAssociationManager analyManger = AnalyticalToPhysicalAssociationManager.GetAnalyticalToPhysicalAssociationManager(doc);

                                        x = UnitUtils.ConvertToInternalUnits(mappedColumn.sapFrame.StartPoint.X, dis);
                                        y = UnitUtils.ConvertToInternalUnits(mappedColumn.sapFrame.StartPoint.Y, dis);
                                        z = UnitUtils.ConvertToInternalUnits(mappedColumn.sapFrame.StartPoint.Z, dis);
                                        StartPoint = new XYZ(x, y, z);

                                        x1 = UnitUtils.ConvertToInternalUnits(mappedColumn.sapFrame.EndPoint.X, dis);
                                        y1 = UnitUtils.ConvertToInternalUnits(mappedColumn.sapFrame.EndPoint.Y, dis);
                                        z1 = UnitUtils.ConvertToInternalUnits(mappedColumn.sapFrame.EndPoint.Z, dis);

                                        EndPoint = new XYZ(x1, y1, z1);


                                        Autodesk.Revit.DB.Curve line = Line.CreateBound(StartPoint, EndPoint);

                                        Autodesk.Revit.DB.Curve line3 = Line.CreateBound(EndPoint, StartPoint);


                                        FamilyInstance instance = (FamilyInstance)doc.GetElement(mappedColumn.revitFrame.revitEleId);


                                        #region No change in the type 
                                        if ((BuiltInCategory)instance.Category.Id.IntegerValue == BuiltInCategory.OST_StructuralColumns)
                                        {
                                            revitElements.Add(instance);

                                            Parameter colStyleParam = instance.LookupParameter("Column Style");

                                            int ColumnStyle = instance.LookupParameter("Column Style").AsInteger();


                                            if (ColumnStyle == (int)SlantedOrVerticalColumnType.CT_Vertical && mappedColumn.sapFrame.MemberType == MemberType.Column)
                                            {

                                                isNotMatched = false;
                                                colStyleParam.Set((int)SlantedOrVerticalColumnType.CT_EndPoint);

                                                if (z > z1)
                                                {
                                                    alignXYForBothPoints(line3);
                                                    (instance.Location as LocationCurve).Curve = line3;
                                                }
                                                else
                                                {
                                                    alignXYForBothPoints(line);
                                                    (instance.Location as LocationCurve).Curve = line;
                                                }

                                                colStyleParam.Set((int)SlantedOrVerticalColumnType.CT_Vertical);

                                                #region SetRotationAngle 
                                                LocationPoint pt = instance.Location as LocationPoint;
                                                double currentRotation = pt.Rotation;



                                                Line line2 = Line.CreateBound(StartPoint, EndPoint);
                                                double col_angle = (mappedColumn.sapFrame.Beta) * Math.PI / 180;


                                                if (currentRotation != col_angle)
                                                {
                                                    // to be checked 30/5/2023
                                                    //  ElementTransformUtils.RotateElement(doc, instance.Id, line2, col_angle - currentRotation);

                                                }



                                                #endregion

                                                //doc.Regenerate();
                                            }

                                            else if (ColumnStyle == (int)SlantedOrVerticalColumnType.CT_EndPoint && mappedColumn.sapFrame.MemberType == MemberType.SlanttedColumn)
                                            {

                                                isNotMatched = false;

                                                if (z > z1)
                                                {
                                                    (instance.Location as LocationCurve).Curve = line3;
                                                }
                                                else
                                                {
                                                    (instance.Location as LocationCurve).Curve = line;
                                                }



                                                Line line2 = Line.CreateBound(StartPoint, EndPoint);
                                                double col_angle = (mappedColumn.sapFrame.Beta) * Math.PI / 180;
                                                // ElementTransformUtils.RotateElement(doc, instance.Id, line2, col_angle);
                                                // doc.Regenerate();

                                            }


                                        }

                                        #endregion



                                        #region change in the type

                                        if (isNotMatched)
                                        {
                                            // case of VrColumn (to vr column)
                                            if (mappedColumn.sapFrame.MemberType == MemberType.Column)
                                            {

                                                Dictionary<string, FamilySymbol> dic = new Dictionary<string, FamilySymbol>();
                                                List<Frame> Sapframes = new List<Frame>() { mappedColumn.sapFrame };

                                                FamilySymbol revitSection = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralColumns).WhereElementIsElementType().Cast<FamilySymbol>().ToList().FirstOrDefault();
                                                if (!revitSection.IsActive)
                                                    revitSection.Activate();


                                                FamilyInstance newInstance = doc.Create.NewFamilyInstance(StartPoint, revitSection, lvl0, StructuralType.Column);

                                                newlyCreatedElements.Add(newInstance);
                                                revitElements.Add(instance);

                                                //analyManger.AddAssociation(analytInstance.Id, newInstance.Id);

                                                #region handling levels elevations
                                                Parameter pbase = newInstance.LookupParameter("Base Level");
                                                Parameter ptop = newInstance.LookupParameter("Top Level");

                                                //
                                                pbase.Set(lvl0.Id);
                                                ptop.Set(lvl1.Id);

                                                Level baselvl = doc.GetElement(pbase.AsElementId()) as Level;
                                                Level toplvl = doc.GetElement(ptop.AsElementId()) as Level;


                                                double base_elevation = UnitUtils.ConvertFromInternalUnits(baselvl.ProjectElevation, dis);
                                                double top_elevation = UnitUtils.ConvertFromInternalUnits(toplvl.ProjectElevation, dis);

                                                double base_elev = Convert.ToDouble(mappedColumn.sapFrame.StartPoint.Z);
                                                double top_elv = Convert.ToDouble(mappedColumn.sapFrame.EndPoint.Z);


                                                double base_offset = UnitUtils.ConvertToInternalUnits(base_elev - base_elevation, dis);
                                                double top_offset = UnitUtils.ConvertToInternalUnits(top_elv - top_elevation, dis);

                                                if (base_elev - base_elevation > top_elv - top_elevation)
                                                {
                                                    base_elev = Convert.ToDouble(mappedColumn.sapFrame.EndPoint.Z);
                                                    top_elv = Convert.ToDouble(mappedColumn.sapFrame.StartPoint.Z);

                                                    base_offset = UnitUtils.ConvertToInternalUnits(base_elev - base_elevation, dis);
                                                    top_offset = UnitUtils.ConvertToInternalUnits(top_elv - top_elevation, dis);
                                                }

                                                newInstance.LookupParameter("Top Offset").Set(top_offset);
                                                newInstance.LookupParameter("Base Offset").Set(base_offset);



                                                #endregion

                                                #region Rotation angle
                                                Line line2 = Line.CreateBound(StartPoint, EndPoint);
                                                double col_angle = (mappedColumn.sapFrame.Beta) * Math.PI / 180;

                                                // ElementTransformUtils.RotateElement(doc, newInstance.Id, line2, col_angle);


                                                #endregion

                                                #region setParameters

                                                //Set Parameters
                                                RhinoObject rhinoObj = rhinoDoc.Objects.FindId(new Guid(mappedColumn.sapFrame.RhinoID));
                                                ObjectAttributes attr = rhinoObj.Attributes.Duplicate();
                                                attr.SetUserString("ID", newInstance.Id.ToString());
                                                rhinoObj.Attributes = attr;

                                                #endregion

                                                doc.Delete(instance.Id);

                                                //doc.Regenerate();



                                            }


                                            // case of SlanttedColumn (to slantted column)
                                            if (mappedColumn.sapFrame.MemberType == MemberType.SlanttedColumn)
                                            {
                                                FamilySymbol revitSection = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralColumns).WhereElementIsElementType().Cast<FamilySymbol>().ToList().FirstOrDefault(); ;

                                                if (!revitSection.IsActive)
                                                    revitSection.Activate();


                                                FamilyInstance newInstance = doc.Create.NewFamilyInstance(StartPoint, revitSection, lvl0, StructuralType.Column);
                                                newlyCreatedElements.Add(newInstance);
                                                //analyManger.AddAssociation(analytInstance.Id, newInstance.Id);
                                                revitElements.Add(instance);


                                                Parameter param = newInstance.LookupParameter("Column Style");
                                                param.Set((int)SlantedOrVerticalColumnType.CT_EndPoint);
                                                LocationCurve instanceCurve = (LocationCurve)newInstance.Location;

                                                if (instanceCurve != null)
                                                {
                                                    if (z > z1)
                                                    {
                                                        instanceCurve.Curve = line3;
                                                    }
                                                    else
                                                    {
                                                        instanceCurve.Curve = line;

                                                    }

                                                }



                                                #region Rotation angle

                                                double col_angle = (mappedColumn.sapFrame.Beta) * Math.PI / 180;
                                                // newInstance.LookupParameter("Cross-Section Rotation").Set(col_angle);


                                                #endregion

                                                #region setParameters

                                                //Set Parameters
                                                RhinoObject rhinoObj = rhinoDoc.Objects.FindId(new Guid(mappedColumn.sapFrame.RhinoID));
                                                ObjectAttributes attr = rhinoObj.Attributes.Duplicate();
                                                attr.SetUserString("ID", newInstance.Id.ToString());
                                                rhinoObj.Attributes = attr;

                                                #endregion


                                                doc.Delete(instance.Id);
                                                //doc.Regenerate();





                                            }
                                        }




                                        #endregion






                                    }

                                }


                                catch (Exception ex2)
                                {
                                    System.Diagnostics.StackTrace trace = new System.Diagnostics.StackTrace(ex2, true);
                                    RevitUtils.Log.Add("Error in updateRevitGeomertyForColumns2 Method  sapCol no " + mappedColumn.sapFrame.ID + ex2.Message + trace.GetFrame(trace.FrameCount - 1).GetFileLineNumber());
                                    continue;
                                }


                            }


                            trns.Commit();

                        }
                    }

                    tg.Assimilate();


                }

            }



            public static List<RhinoRevitFrame> getMappedSapRevit(List<Frame> sapFrames, List<Frame> revitFrames, double margin)
            {
                List<RhinoRevitFrame> allResults = new List<RhinoRevitFrame>();


                List<string> sapIDS = sapFrames.Select(x => x.ID).AsParallel().ToList();
                List<string> sapIDParameter = revitFrames.Select(x => x.ID).AsParallel().ToList();


                List<Frame> sapElementsMappedByIDS = sapFrames.Where(x => sapIDParameter.Contains(x.ID)).AsParallel().OrderBy(x => x.ID).ToList();
                List<Frame> revitElementsMappedByIDS = revitFrames.Where(x => sapIDS.Contains(x.ID)).AsParallel().OrderBy(x => x.ID).ToList();



                for (int i = sapElementsMappedByIDS.Count - 1; i >= 0; i--)
                {

                    //if (sapElementsMappedByIDS[i].MemberType == revitElementsMappedByIDS[i].MemberType)
                    //{

                    RhinoRevitFrame result = new RhinoRevitFrame(sapElementsMappedByIDS[i], revitElementsMappedByIDS[i]);
                    result.mappedBy = "ID";
                    allResults.Add(result);

                    //double lengthDifferencePersentage = Math.Abs(revitElementsMappedByIDS[i].Length - sapElementsMappedByIDS[i].Length) / sapElementsMappedByIDS[i].Length;
                    //if (lengthDifferencePersentage < margin)
                    //{
                    //    SapRevitFrame result = new SapRevitFrame(sapElementsMappedByIDS[i], revitElementsMappedByIDS[i]);
                    //    result.mappedBy = "ID";
                    //    allResults.Add(result);
                    //}

                    //else
                    //{
                    //    sapElementsMappedByIDS.Remove(sapElementsMappedByIDS[i]);
                    //    revitElementsMappedByIDS.Remove(revitElementsMappedByIDS[i]);
                    //}


                    // }

                    //else
                    //{
                    //    sapElementsMappedByIDS.RemoveAt(i);
                    //    revitElementsMappedByIDS.RemoveAt(i);
                    //}

                }


                List<Frame> sapElementsNotMappedByIDS = sapFrames.Except(sapElementsMappedByIDS).ToList();
                List<Frame> revitElementsNotMappedByIDS = revitFrames.Except(revitElementsMappedByIDS).ToList();


                foreach (var sapFrame in sapElementsNotMappedByIDS.AsParallel())
                {
                    RhinoRevitFrame result = new RhinoRevitFrame();
                    Dictionary<Frame, double> dictionary = new Dictionary<Frame, double>();
                    List<double> lengthDiffernce = new List<double>();
                    result.sapFrame = sapFrame;
                    result.revitFrame = null;



                    foreach (Frame revitFrame in revitElementsNotMappedByIDS.AsParallel())
                    {
                        double lengthDifferencePersentage = Math.Abs(revitFrame.Length - sapFrame.Length) / sapFrame.Length;
                        if (lengthDifferencePersentage < margin)
                        {
                            dictionary.Add(revitFrame, Point.getDistance(sapFrame.MidPoint, revitFrame.MidPoint) + Point.getDistance(sapFrame.StartPoint, revitFrame.StartPoint));
                        }

                    }


                    #region getMinValueFromDictionary
                    double mainValue = double.MaxValue;
                    if (dictionary.Count != 0)
                    {
                        mainValue = dictionary.Values.AsParallel().ToList().Min();
                    }


                    #endregion


                    if (mainValue < margin * 2 * sapFrame.Length)
                    {
                        int index = dictionary.Values.AsParallel().ToList().IndexOf(dictionary.Values.Min());
                        result.revitFrame = dictionary.Keys.AsParallel().ToList()[index];
                        result.mappedBy = "MID";
                        allResults.Add(result);
                    }

                    else
                    {
                        result.revitFrame = null;
                        allResults.Add(result);
                    }

                Next:
                    continue;
                }


                return allResults;

            }

            public static List<RhinoRevitFrame> getMappedSapRevit2(List<Frame> sapFrames, List<Frame> revitFrames, double margin)
            {
                List<RhinoRevitFrame> allResults = new List<RhinoRevitFrame>();


                List<string> sapIDS = sapFrames.Select(x => x.ID).AsParallel().ToList();
                List<string> sapIDParameter = revitFrames.Select(x => x.ID).AsParallel().ToList();


                List<Frame> sapElementsMappedByIDS = sapFrames.Where(x => sapIDParameter.Contains(x.ID)).AsParallel().OrderBy(x => x.ID).ToList();
                List<Frame> revitElementsMappedByIDS = revitFrames.Where(x => sapIDS.Contains(x.ID)).AsParallel().OrderBy(x => x.ID).ToList();



                for (int i = 0; i < sapElementsMappedByIDS.Count; i++)
                {

                    RhinoRevitFrame result = new RhinoRevitFrame(sapElementsMappedByIDS[i], revitElementsMappedByIDS[i]);
                    result.mappedBy = "ID";
                    allResults.Add(result);

                    //double lengthDifferencePersentage = Math.Abs(revitElementsMappedByIDS[i].Length - sapElementsMappedByIDS[i].Length) / sapElementsMappedByIDS[i].Length;
                    //if (lengthDifferencePersentage < margin)
                    //{
                    //    SapRevitFrame result = new SapRevitFrame(sapElementsMappedByIDS[i], revitElementsMappedByIDS[i]);
                    //    result.mappedBy = "ID";
                    //    allResults.Add(result);
                    //}

                    //else
                    //{
                    //    sapElementsMappedByIDS.Remove(sapElementsMappedByIDS[i]);
                    //    revitElementsMappedByIDS.Remove(revitElementsMappedByIDS[i]);
                    //}


                }


                List<Frame> sapElementsNotMappedByIDS = sapFrames.Except(sapElementsMappedByIDS).ToList();
                List<Frame> revitElementsNotMappedByIDS = revitFrames.Except(revitElementsMappedByIDS).ToList();


                foreach (var sapFrame in sapElementsNotMappedByIDS.AsParallel())
                {
                    RhinoRevitFrame result = new RhinoRevitFrame();
                    Dictionary<Frame, double> dictionary = new Dictionary<Frame, double>();
                    List<double> lengthDiffernce = new List<double>();
                    result.sapFrame = sapFrame;
                    result.revitFrame = null;



                    foreach (Frame revitFrame in revitElementsNotMappedByIDS.AsParallel())
                    {
                        double lengthDifferencePersentage = Math.Abs(revitFrame.Length - sapFrame.Length) / sapFrame.Length;
                        if (lengthDifferencePersentage < margin)
                        {
                            dictionary.Add(revitFrame, Point.getDistance(sapFrame.MidPoint, revitFrame.MidPoint) + Point.getDistance(sapFrame.StartPoint, revitFrame.StartPoint));
                        }

                    }


                    #region getMinValueFromDictionary
                    double mainValue = double.MaxValue;
                    if (dictionary.Count != 0)
                    {
                        mainValue = dictionary.Values.AsParallel().ToList().Min();
                    }


                    #endregion


                    if (mainValue < margin * 2 * sapFrame.Length)
                    {
                        int index = dictionary.Values.AsParallel().ToList().IndexOf(dictionary.Values.Min());
                        result.revitFrame = dictionary.Keys.AsParallel().ToList()[index];
                        result.mappedBy = "MID";
                        allResults.Add(result);
                    }

                    else
                    {
                        result.revitFrame = null;
                        allResults.Add(result);
                    }

                Next:
                    continue;
                }


                //foreach (var sapFrame in sapFrames.AsParallel())
                //{
                //    SapRevitFrame result = new SapRevitFrame();
                //    Dictionary<Frame, double> dictionary = new Dictionary<Frame, double>();
                //    List<double> lengthDiffernce = new List<double>();
                //    result.sapFrame = sapFrame;
                //    result.revitFrame = null;



                //    foreach (Frame revitFrame in revitFrames.AsParallel())
                //    {
                //        if (revitFrame.ID == sapFrame.ID)
                //        {
                //            //double lengthDifferencePersentage = Math.Abs(revitFrame.Length - sapFrame.Length) / sapFrame.Length;
                //            //if (lengthDifferencePersentage < margin)
                //            //{
                //            //    result.revitFrame = revitFrame;
                //            //    result.mappedBy = "ID";
                //            //    allResults.Add(result);
                //            //    goto Next;

                //            //}

                //            result.revitFrame = revitFrame;
                //            result.mappedBy = "ID";
                //            allResults.Add(result);
                //            goto Next;



                //        }

                //        else
                //        {

                //            double lengthDifferencePersentage = Math.Abs(revitFrame.Length - sapFrame.Length) / sapFrame.Length;
                //            if (lengthDifferencePersentage < margin)
                //            {
                //                dictionary.Add(revitFrame, Point.getDistance(sapFrame.MidPoint, revitFrame.MidPoint));
                //            }

                //        }
                //    }


                //    #region getMinValueFromDictionary
                //    double mainValue = 1234567891011;
                //    if (dictionary.Count != 0)
                //    {
                //        mainValue = dictionary.Values.AsParallel().ToList().Min();
                //    }


                //    #endregion


                //    if (mainValue < margin * sapFrame.Length)
                //    {
                //        int index = dictionary.Values.AsParallel().ToList().IndexOf(dictionary.Values.Min());
                //        result.revitFrame = dictionary.Keys.AsParallel().ToList()[index];
                //        result.mappedBy = "MID";
                //        allResults.Add(result);
                //    }

                //    else
                //    {
                //        result.revitFrame = null;
                //        allResults.Add(result);
                //    }

                //Next:
                //    continue;
                //}

                return allResults;

            }

            public static List<RhinoRevitFrame> getMappedSapRevitByIdOnly(List<Frame> sapFrames, List<Frame> revitFrames)
            {
                List<RhinoRevitFrame> allResults = new List<RhinoRevitFrame>();

                foreach (var sapFrame in sapFrames.AsParallel())
                {
                    RhinoRevitFrame result = new RhinoRevitFrame();
                    Dictionary<Frame, double> dictionary = new Dictionary<Frame, double>();
                    List<double> lengthDiffernce = new List<double>();
                    result.sapFrame = sapFrame;
                    result.revitFrame = null;

                    foreach (Frame revitFrame in revitFrames.AsParallel())
                    {
                        if (revitFrame.ID == sapFrame.ID)
                        {
                            result.revitFrame = revitFrame;
                            result.mappedBy = "ID";
                            goto Next;

                        }

                    }

                Next:
                    allResults.Add(result);
                    continue;
                }

                return allResults;

            }


            public static void getUnMappedEelements(List<RhinoRevitFrame> sapRevitFrames, List<Frame> sapFrames, List<Frame> revitFrames, ref List<Frame> unMappedSapFrames, ref List<Frame> UnMappedRevitFrames)
            {

                unMappedSapFrames = sapRevitFrames.AsParallel().Where(x => x.revitFrame == null).ToList().Select(x => x.sapFrame).ToList();
                UnMappedRevitFrames = revitFrames.Except(sapRevitFrames.Select(x => x.revitFrame).ToList()).ToList();
            }

            public static void setMarkInRevit(List<RhinoRevitFrame> mappedElements, Document doc)
            {
                foreach (var item in mappedElements)
                {
                    if (item.revitFrame != null)
                    {
                        FamilyInstance instance = doc.GetElement(item.revitFrame.revitEleId) as FamilyInstance;
                        Parameter mark = instance.LookupParameter("Mark");


                        using (Transaction trns = new Transaction(doc, "aa"))
                        {

                            trns.Start();
                            FailureHandlingOptions failOpt = trns.GetFailureHandlingOptions();
                            failOpt.SetFailuresPreprocessor(new RoomWarningSwallower());
                            trns.SetFailureHandlingOptions(failOpt);

                            mark.Set(item.sapFrame.ID);
                            trns.Commit();

                        }
                    }


                }


            }
            public static View3D createView(List<ElementId> ids, ExternalCommandData commandData, string groupName)
            {
                Document doc = commandData.Application.ActiveUIDocument.Document;

                using (Transaction trns = new Transaction(doc, "aa"))
                {

                    trns.Start();
                    ViewFamilyType viewType = new FilteredElementCollector(doc).OfClass(typeof(ViewFamilyType)).Cast<ViewFamilyType>().First(x => x.ViewFamily == ViewFamily.ThreeDimensional);
                    View3D unMappedView = new FilteredElementCollector(doc).OfClass(typeof(View3D)).Cast<View3D>().FirstOrDefault(x => x.Name == groupName);
                    if (unMappedView == null)
                    {
                        unMappedView = View3D.CreateIsometric(doc, viewType.Id);
                        unMappedView.Name = groupName;
                    }

                    TemporaryViewMode tempView = TemporaryViewMode.TemporaryHideIsolate;
                    unMappedView.DisableTemporaryViewMode(tempView);

                    UIDocument uiDoc = commandData.Application.ActiveUIDocument;
                    unMappedView.IsolateElementsTemporary(ids);

                    trns.Commit();
                    //commandData.Application.ActiveUIDocument.ActiveView = unMappedView;
                    //commandData.Application.ActiveUIDocument.RequestViewChange(unMappedView);
                    return unMappedView;
                }



            }



            public static View3D createView(List<ElementId> ids, Document doc, UIDocument uiDoc, string groupName)
            {


                using (Transaction trns = new Transaction(doc, "aa"))
                {

                    trns.Start();
                    ViewFamilyType viewType = new FilteredElementCollector(doc).OfClass(typeof(ViewFamilyType)).Cast<ViewFamilyType>().First(x => x.ViewFamily == ViewFamily.ThreeDimensional);
                    View3D unMappedView = new FilteredElementCollector(doc).OfClass(typeof(View3D)).Cast<View3D>().FirstOrDefault(x => x.Name == groupName);
                    if (unMappedView == null)
                    {
                        unMappedView = View3D.CreateIsometric(doc, viewType.Id);
                        unMappedView.Name = groupName;
                    }

                    TemporaryViewMode tempView = TemporaryViewMode.TemporaryHideIsolate;
                    unMappedView.DisableTemporaryViewMode(tempView);


                    unMappedView.IsolateElementsTemporary(ids);

                    trns.Commit();
                    //commandData.Application.ActiveUIDocument.ActiveView = unMappedView;
                    //commandData.Application.ActiveUIDocument.RequestViewChange(unMappedView);
                    return unMappedView;
                }



            }
            public static void alignXYForBothPoints(Autodesk.Revit.DB.Curve c)
            {
                XYZ st = c.GetEndPoint(0);
                XYZ end = new XYZ(c.GetEndPoint(0).X, c.GetEndPoint(0).Y, c.GetEndPoint(1).Z);
                Autodesk.Revit.DB.Curve newAlignedCurve = Autodesk.Revit.DB.Line.CreateBound(st, end);

                c = newAlignedCurve;

            }

        }

        public class RoomWarningSwallower : IFailuresPreprocessor
        {

            public FailureProcessingResult PreprocessFailures(FailuresAccessor a)
            {


                #region New 2


                var preprocessorMessages
                                           = a.GetFailureMessages(FailureSeverity.Error)
                                                  .Where(x => x.HasResolutionOfType(FailureResolutionType.DeleteElements)
                                                    || x.HasResolutionOfType(FailureResolutionType.DetachElements))
                                                   .ToList();

                var preprocessorMessagesOfWarnings
                                 = a.GetFailureMessages(FailureSeverity.Warning).ToList();


                if (preprocessorMessages.Count == 0 && preprocessorMessagesOfWarnings.Count == 0)
                    return FailureProcessingResult.Continue;

                foreach (var failureAccessor in preprocessorMessages)
                {
                    if (failureAccessor.GetSeverity() == FailureSeverity.Error)
                    {
                        failureAccessor.SetCurrentResolutionType(
                     failureAccessor.HasResolutionOfType(FailureResolutionType.DetachElements)
                       ? FailureResolutionType.DetachElements
                       : FailureResolutionType.DeleteElements);

                        a.ResolveFailure(failureAccessor);
                    }


                }

                foreach (var failureAccessor in preprocessorMessagesOfWarnings)
                {
                    if (failureAccessor.GetSeverity() == FailureSeverity.Warning)
                    {
                        // notice here, when i used delete all warnings i should add below it directly the return 
                        //because when i use delete all drawigns it will not see the second return found out of the braces  

                        a.DeleteAllWarnings();
                        return FailureProcessingResult.ProceedWithCommit;



                        // or i can use delete warning only and it will see the return found out of the braces  (Normal case) 
                        // a.DeleteWarning(failureAccessor);



                    }
                }


                return FailureProcessingResult.ProceedWithCommit;
                #endregion

                // inside event handler, get all warnings

                IList<FailureMessageAccessor> failures = a.GetFailureMessages();


                foreach (FailureMessageAccessor f in failures)
                {

                    List<string> elementsIds2 = f.GetFailingElementIds().Select(x => x.ToString()).ToList();
                    if (elementsIds2.Contains("486501"))
                    {
                        int aa = 0;
                    }

                    FailureDefinitionId id = f.GetFailureDefinitionId();
                    string jj = f.GetDescriptionText();

                    if (jj.ToLower().Contains("keep") && jj.ToLower().Contains("join"))
                    {

                        List<string> elementsIds = f.GetFailingElementIds().Select(x => x.ToString()).ToList();
                        elementsIds.Add(f.GetFailureDefinitionId().ToString());

                        string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                        string filePath = Path.Combine(desktopPath, "ColumnselementIds.txt");
                        WriteListToFile(filePath, elementsIds);

                    }


                    else if (jj.ToLower().Contains("make") && jj.ToLower().Contains("type"))
                    {
                        List<string> elementsIds = f.GetFailingElementIds().Select(x => x.ToString()).ToList();
                        elementsIds.Add(f.GetFailureDefinitionId().ToString());


                        string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                        string filePath = Path.Combine(desktopPath, "BeamselementIds.txt");
                        WriteListToFile(filePath, elementsIds);
                    }


                    #region old
                    //if (id == BuiltInFailures.FamilyFailures.CannotMakeTypeCB)
                    //{
                    //    if (f.HasResolutions())
                    //    {
                    //        a.ResolveFailure(f);
                    //        a.DeleteAllWarnings();
                    //        a.DeleteWarning(f);
                    //        return FailureProcessingResult.Continue;
                    //    }
                    //    a.DeleteWarning(f);
                    //    return FailureProcessingResult.Continue;


                    //}

                    //if (id == BuiltInFailures.FamilyFailures.SlantedColumnEndPlanesHaveFlipped)
                    //{
                    //    List<ElementId> allelementsw = f.GetFailingElementIds().ToList();
                    //    string sapId = a.GetDocument().GetElement(allelementsw[0]).LookupParameter("SapID").AsString();

                    //    a.DeleteElements(allelementsw);
                    //    //a.DeleteWarning(f);
                    //    return FailureProcessingResult.ProceedWithCommit;
                    //}
                    //if (id == BuiltInFailures.CurveFailures.LineTooShortError)
                    //{
                    //    if (f.HasResolutions())
                    //    {
                    //        a.ResolveFailure(f);
                    //        a.DeleteWarning(f);
                    //        return FailureProcessingResult.Continue;
                    //    }
                    //    a.DeleteWarning(f);
                    //    return FailureProcessingResult.Continue;


                    //}

                    //if (id == BuiltInFailures.GeneralFailures.ErrorInFamilyResolved || id == BuiltInFailures.GeneralFailures.ErrorInSymbolFamilyResolved || id == BuiltInFailures.GeneralFailures.ErrorResolved)
                    //{
                    //    if (f.HasResolutions())
                    //    {
                    //        a.ResolveFailure(f);
                    //        a.DeleteWarning(f);
                    //        return FailureProcessingResult.Continue;
                    //    }
                    //    a.DeleteWarning(f);
                    //    return FailureProcessingResult.Continue;


                    //}

                    //if (id == BuiltInFailures.OverlapFailures.DuplicateInstances)
                    //{
                    //    List<ElementId> allelementsw = f.GetFailingElementIds().ToList();
                    //    List<ElementId> deletedelements = new List<ElementId>();
                    //    deletedelements.Add(allelementsw[1]);
                    //    a.DeleteElements(deletedelements);
                    //    a.DeleteWarning(f);
                    //    return FailureProcessingResult.ProceedWithCommit;
                    //}


                    //if (f.HasResolutions())
                    //{
                    //    a.ResolveFailure(f);
                    //    a.DeleteWarning(f);
                    //    a.DeleteAllWarnings();

                    //    return FailureProcessingResult.ProceedWithCommit;
                    //}

                    //else
                    //{
                    //    a.DeleteWarning(f);
                    //    return FailureProcessingResult.ProceedWithCommit;
                    //}


                    #endregion

                    #region new


                    if (id == BuiltInFailures.FamilyFailures.CannotMakeTypeCB)
                    {
                        if (f.HasResolutions())
                        {
                            a.ResolveFailure(f);
                            a.DeleteAllWarnings();
                            a.DeleteWarning(f);
                            return FailureProcessingResult.Continue;
                        }
                        a.DeleteWarning(f);
                        return FailureProcessingResult.Continue;


                    }

                    if (id == BuiltInFailures.FamilyFailures.SlantedColumnEndPlanesHaveFlipped)
                    {
                        List<ElementId> allelementsw = f.GetFailingElementIds().ToList();
                        string sapId = a.GetDocument().GetElement(allelementsw[0]).LookupParameter("SapID").AsString();

                        a.DeleteElements(allelementsw);
                        a.DeleteWarning(f);
                        return FailureProcessingResult.ProceedWithCommit;
                    }
                    if (id == BuiltInFailures.CurveFailures.LineTooShortError)
                    {
                        if (f.HasResolutions())
                        {
                            a.ResolveFailure(f);
                            a.DeleteWarning(f);
                            return FailureProcessingResult.Continue;
                        }
                        a.DeleteWarning(f);
                        return FailureProcessingResult.Continue;


                    }

                    if (id == BuiltInFailures.GeneralFailures.ErrorInFamilyResolved || id == BuiltInFailures.GeneralFailures.ErrorInSymbolFamilyResolved || id == BuiltInFailures.GeneralFailures.ErrorResolved)
                    {
                        if (f.HasResolutions())
                        {
                            a.ResolveFailure(f);
                            a.DeleteWarning(f);
                            return FailureProcessingResult.Continue;
                        }
                        a.DeleteWarning(f);
                        return FailureProcessingResult.Continue;


                    }

                    if (id == BuiltInFailures.OverlapFailures.DuplicateInstances)
                    {
                        List<ElementId> allelementsw = f.GetFailingElementIds().ToList();
                        List<ElementId> deletedelements = new List<ElementId>();
                        deletedelements.Add(allelementsw[1]);
                        a.DeleteElements(deletedelements);
                        a.DeleteWarning(f);
                        return FailureProcessingResult.ProceedWithCommit;
                    }


                    if (f.GetSeverity() == FailureSeverity.Warning)
                    {
                        if (f.HasResolutions())
                        {
                            a.ResolveFailure(f);
                            a.DeleteWarning(f);
                            a.DeleteAllWarnings();

                            return FailureProcessingResult.ProceedWithCommit;
                        }

                        else
                        {
                            a.DeleteWarning(f);
                            return FailureProcessingResult.ProceedWithCommit;
                        }
                    }

                    else if (f.GetSeverity() == FailureSeverity.Error)
                    {
                        List<ElementId> allelementsw = f.GetFailingElementIds().ToList();
                        List<ElementId> deletedelements = new List<ElementId>();
                        deletedelements.Add(allelementsw[1]);
                        a.DeleteElements(allelementsw);



                        a.DeleteWarning(f);
                        return FailureProcessingResult.ProceedWithCommit;
                    }

                    if (BuiltInFailures.InaccurateFailures.InaccurateBeamOrBrace == id || BuiltInFailures.InaccurateFailures.InaccurateAreaLine == id || BuiltInFailures.InaccurateFailures.InaccurateSketchLine == id || BuiltInFailures.InaccurateFailures.InaccurateCurveBasedFamily == id || BuiltInFailures.InaccurateFailures.InaccurateDriveCurve == id || BuiltInFailures.InaccurateFailures.InaccurateMassingSketchLine == id || BuiltInFailures.InaccurateFailures.InaccurateGrid == id || BuiltInFailures.InaccurateFailures.InaccurateLevel == id || BuiltInFailures.InaccurateFailures.InaccurateLine == id || BuiltInFailures.InaccurateFailures.InaccurateRefPlane == id || BuiltInFailures.InaccurateFailures.InaccurateRoomSeparation == id)
                    {

                        a.DeleteWarning(f);
                        return FailureProcessingResult.ProceedWithCommit;


                    }


                    if (id == BuiltInFailures.JoinElementsFailures.CannotJoinElements || id == BuiltInFailures.JoinElementsFailures.CannotJoinElementsError || id == BuiltInFailures.JoinElementsFailures.CannotJoinElementsStructural || id == BuiltInFailures.JoinElementsFailures.CannotJoinElementsWarn || id == BuiltInFailures.JoinElementsFailures.CannotKeepJoined || id == BuiltInFailures.JoinElementsFailures.CannotCutJoinedGeometry || id == BuiltInFailures.JoinElementsFailures.CannotJoinDesignOptions || id == BuiltInFailures.JoinElementsFailures.CannotKeepReferenceOfJoinedElement || id == BuiltInFailures.JoinElementsFailures.CannotKeepReferenceOfJoinedElementWarning || id == BuiltInFailures.JoinElementsFailures.ElementWillBeDetachedFromPlane || id == BuiltInFailures.JoinElementsFailures.CannotCope)
                    {

                        bool isResolved = false;

                        if (f.HasResolutions())
                        {
                            a.ResolveFailure(f);
                            isResolved = true;
                        }


                        if (isResolved)
                        {

                            return FailureProcessingResult.ProceedWithCommit;
                        }

                        return FailureProcessingResult.Continue;

                    }


                    #endregion


                    //#region MyRegion
                    //if (id == BuiltInFailures.OverlapFailures.DuplicateInstances)
                    //{
                    //    List<ElementId> allelementsw = f.GetFailingElementIds().ToList();
                    //    List<ElementId> deletedelements = new List<ElementId>();
                    //    deletedelements.Add(allelementsw[1]);
                    //    a.DeleteElements(deletedelements);
                    //    a.DeleteWarning(f);
                    //    return FailureProcessingResult.ProceedWithCommit;
                    //}
                    //else
                    //{
                    //    a.DeleteAllWarnings();
                    //    return FailureProcessingResult.Continue;

                    //}
                    //#endregion


                    // return FailureProcessingResult.Continue;
                }
                return FailureProcessingResult.Continue;
            }

            public static void WriteListToFile(string filePath, List<string> elementsIds)
            {
                try
                {
                    bool fileExists = File.Exists(filePath);

                    using (StreamWriter writer = new StreamWriter(filePath, append: fileExists))
                    {
                        if (fileExists)
                        {
                            writer.WriteLine("--------------------"); // Separator
                        }

                        foreach (string elementId in elementsIds)
                        {
                            writer.WriteLine(elementId);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"An error occurred: {ex.Message}");
                }
            }
        }


        #endregion


        public static List<ElementId> CreateAllBeams3(Document doc, List<Frame> beams, ForgeTypeId dis)
        {
            // AnalyticalToPhysicalAssociationManager analyManger = AnalyticalToPhysicalAssociationManager.GetAnalyticalToPhysicalAssociationManager(doc);

            List<ElementId> result = new List<ElementId>();
            List<FamilyInstance> revitInstances = new List<FamilyInstance>();

            #region defintions for creation of beams and columns 
            FamilySymbol revitSection = null;
            double angle = 0;
            Level lvl0 = new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().ToList().FirstOrDefault();

            if (lvl0 == null)
            {
                using (Transaction t = new Transaction(doc, "Create Levels"))
                {
                    t.Start();
                    lvl0 = Level.Create(doc, 0.0);
                    lvl0.Name = "Hydra-level-1";

                    t.Commit();

                }
            }
        #endregion

        Beammodeling:


            RevitServices.Transactions.TransactionManager.Instance.ForceCloseTransaction();

            using (TransactionGroup tg = new TransactionGroup(doc, "aaa"))
            {

                tg.Start();

                var chopSize = Frame.chopSize;
                var beamChunks = beams.Chunk(chopSize).ToList();

                foreach (var beamChunk in beamChunks)
                {

                    using (Transaction trns = new Transaction(doc, "trns"))
                    {


                        trns.Start();
                        FailureHandlingOptions failOpt = trns.GetFailureHandlingOptions();
                        failOpt.SetFailuresPreprocessor(new RoomWarningSwallower());
                        trns.SetFailureHandlingOptions(failOpt);

                        foreach (var bm in beamChunk)
                        {

                            var x = UnitUtils.ConvertToInternalUnits(bm.StartPoint.X, dis);
                            var y = UnitUtils.ConvertToInternalUnits(bm.StartPoint.Y, dis);
                            var z = UnitUtils.ConvertToInternalUnits(bm.StartPoint.Z, dis);
                            var StartPoint = new XYZ(x, y, z);

                            var x1 = UnitUtils.ConvertToInternalUnits(bm.EndPoint.X, dis);
                            var y1 = UnitUtils.ConvertToInternalUnits(bm.EndPoint.Y, dis);
                            var z1 = UnitUtils.ConvertToInternalUnits(bm.EndPoint.Z, dis);
                            var EndPoint = new XYZ(x1, y1, z1);



                            try
                            {
                                Autodesk.Revit.DB.Curve line = Line.CreateBound(StartPoint, EndPoint);
                                angle = (bm.Beta) * Math.PI / 180;

                                revitSection = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralFraming).WhereElementIsElementType().Cast<FamilySymbol>().ToList().FirstOrDefault(a => a.Name == bm.Sec);

                                if (revitSection == null)
                                {
                                    revitSection = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralFraming).WhereElementIsElementType().Cast<FamilySymbol>().ToList().FirstOrDefault();

                                }


                                if (!revitSection.IsActive)
                                    revitSection.Activate();


                                Autodesk.Revit.DB.Curve c = RhinoInside.Revit.Convert.Geometry.GeometryEncoder.ToCurve(bm.RhinoCurve);



                                FamilyInstance instance = doc.Create.NewFamilyInstance(c, revitSection, lvl0, StructuralType.Beam);
                                //AnalyticalMember analytInstance = AnalyticalMember.Create(doc, line);
                                // analyManger.AddAssociation(analytInstance.Id, instance.Id);

                                instance.LookupParameter("Cross-Section Rotation").Set(angle);
                                Parameter zJustification = instance.LookupParameter("z Justification");
                                Parameter yJustification = instance.LookupParameter("y Justification");


                                int zjustf = 0;
                                int yjustf = 0;


                                #region setInsPoint

                                switch (bm.InsPt)
                                {
                                    case InsertionPoint.BottomLeft:
                                        zjustf = (int)ZJustification.Bottom;
                                        yjustf = (int)YJustification.Left;

                                        break;
                                    case InsertionPoint.BottomCenter:
                                        zjustf = (int)ZJustification.Bottom;
                                        yjustf = (int)YJustification.Origin;

                                        break;
                                    case InsertionPoint.BottomRight:
                                        zjustf = (int)ZJustification.Bottom;
                                        yjustf = (int)YJustification.Right;

                                        break;
                                    case InsertionPoint.MiddleLeft:
                                        zjustf = (int)ZJustification.Origin;
                                        yjustf = (int)YJustification.Left;

                                        break;
                                    case InsertionPoint.MiddleCenter:
                                        zjustf = (int)ZJustification.Origin;
                                        yjustf = (int)YJustification.Origin;

                                        break;
                                    case InsertionPoint.MiddleRight:
                                        zjustf = (int)ZJustification.Origin;
                                        yjustf = (int)YJustification.Right;

                                        break;
                                    case InsertionPoint.TopLeft:
                                        zjustf = (int)ZJustification.Top;
                                        yjustf = (int)YJustification.Left;

                                        break;
                                    case InsertionPoint.TopCenter:
                                        zjustf = (int)ZJustification.Top;
                                        yjustf = (int)YJustification.Origin;

                                        break;
                                    case InsertionPoint.TopRight:
                                        zjustf = (int)ZJustification.Top;
                                        yjustf = (int)YJustification.Right;

                                        break;
                                    case InsertionPoint.Centroid:
                                        zjustf = (int)ZJustification.Origin;
                                        yjustf = (int)YJustification.Origin;

                                        break;
                                    case InsertionPoint.ShearCenter:
                                        zjustf = (int)ZJustification.Bottom;
                                        yjustf = (int)YJustification.Left;
                                        break;
                                    default:
                                        zjustf = (int)ZJustification.Origin;
                                        yjustf = (int)YJustification.Origin;
                                        break;
                                }

                                switch (bm.InsPt)
                                {


                                }



                                #endregion

                                zJustification.Set(zjustf);
                                yJustification.Set(yjustf);

                                Autodesk.Revit.DB.Structure.StructuralFramingUtils.DisallowJoinAtEnd(instance, 0);
                                Autodesk.Revit.DB.Structure.StructuralFramingUtils.DisallowJoinAtEnd(instance, 1);



                                revitInstances.Add(instance);
                                result.Add(instance.Id);

                                bm.ID = instance.Id.ToString();

                                RhinoDoc.ActiveDoc.Objects.FindId(new Guid(bm.RhinoID)).Attributes.SetUserString("ID", instance.Id.ToString());
                                RhinoDoc.ActiveDoc.Objects.FindId(new Guid(bm.RhinoID)).Attributes.SetUserString("Section", instance.Symbol.Name.ToString());
                                RhinoDoc.ActiveDoc.Objects.FindId(new Guid(bm.RhinoID)).Attributes.SetUserString("Family", instance.Symbol.FamilyName.ToString());
                                RhinoDoc.ActiveDoc.Objects.FindId(new Guid(bm.RhinoID)).Attributes.SetUserString("MemberType", "Beam");
                                RhinoDoc.ActiveDoc.Objects.FindId(new Guid(bm.RhinoID)).Attributes.SetUserString("Angle", bm.Beta.ToString());



                                #region Here Handling of get same parameters as Parent element for SplittedElements
                                if (bm.ParentID != null)
                                {
                                    Element parentElement = null;

                                    // Element parentElement = doc.GetElement(bm.ParentID);

                                    FilteredElementCollector collector = new FilteredElementCollector(doc);
                                    parentElement = null;

                                    // Check if bm.ParentID is a valid integer (ElementId)
                                    if (int.TryParse(bm.ParentID, out int elementIdValue))
                                    {
                                        parentElement = collector
                                            .WhereElementIsNotElementType()
                                            .FirstOrDefault(e => e.Id.IntegerValue == elementIdValue);
                                    }

                                    #region Parameters
                                    // Handle Y Justification
                                    Parameter YJustParentParam = parentElement.get_Parameter(BuiltInParameter.Y_JUSTIFICATION);
                                    Parameter YJustInstanceParam = instance.get_Parameter(BuiltInParameter.Y_JUSTIFICATION);
                                    if (YJustParentParam != null && YJustInstanceParam != null && !YJustInstanceParam.IsReadOnly)
                                    {
                                        YJustInstanceParam.Set(YJustParentParam.AsInteger());
                                    }

                                    // Handle Z Justification
                                    Parameter ZJustParentParam = parentElement.get_Parameter(BuiltInParameter.Z_JUSTIFICATION);
                                    Parameter ZJustInstanceParam = instance.get_Parameter(BuiltInParameter.Z_JUSTIFICATION);
                                    if (ZJustParentParam != null && ZJustInstanceParam != null && !ZJustInstanceParam.IsReadOnly)
                                    {
                                        ZJustInstanceParam.Set(ZJustParentParam.AsInteger());
                                    }

                                    // Handle Y Offset Value
                                    Parameter YOffsetParentParam = parentElement.get_Parameter(BuiltInParameter.Y_OFFSET_VALUE);
                                    Parameter YOffsetInstanceParam = instance.get_Parameter(BuiltInParameter.Y_OFFSET_VALUE);
                                    if (YOffsetParentParam != null && YOffsetInstanceParam != null && !YOffsetInstanceParam.IsReadOnly)
                                    {
                                        YOffsetInstanceParam.Set(YOffsetParentParam.AsDouble());
                                    }

                                    // Handle Z Offset Value
                                    Parameter ZOffsetParentParam = parentElement.get_Parameter(BuiltInParameter.Z_OFFSET_VALUE);
                                    Parameter ZOffsetInstanceParam = instance.get_Parameter(BuiltInParameter.Z_OFFSET_VALUE);
                                    if (ZOffsetParentParam != null && ZOffsetInstanceParam != null && !ZOffsetInstanceParam.IsReadOnly)
                                    {
                                        ZOffsetInstanceParam.Set(ZOffsetParentParam.AsDouble());
                                    }


                                    // Handle Cross-Section Rotation
                                    Parameter RotationParentParam = parentElement.LookupParameter("Cross-Section Rotation");
                                    Parameter RotationInstanceParam = instance.LookupParameter("Cross-Section Rotation");
                                    if (RotationParentParam != null && RotationInstanceParam != null && !RotationInstanceParam.IsReadOnly)
                                    {
                                        RotationInstanceParam.Set(RotationParentParam.AsDouble());
                                    }
                                    #endregion


                                    instance.LookupParameter("Start Extension").Set(0);
                                    instance.LookupParameter("End Extension").Set(0);

                                }



                                #endregion

                            }


                            catch (Exception ex2)

                            {
                                System.Diagnostics.StackTrace trace = new System.Diagnostics.StackTrace(ex2, true);
                                RevitUtils.Log.Add("Error in CreateAllBeams3 Method... beam no" + bm.ID + ex2.Message + trace.GetFrame(trace.FrameCount - 1).GetFileLineNumber());

                                //TaskDialog.Show("Error", "error in beam no" + bm.ID + ex2.Message + trace.GetFrame(trace.FrameCount - 1).GetFileLineNumber());

                                continue;

                            }


                        }

                        trns.Commit();


                    }

                }


                #region Handling Parameters of Broken Elements
                //var grouppedBrokendMembers =  beams.Where(x => x.ParentID != null).GroupBy(x => x.ParentID);

                //using (Transaction t = new Transaction(doc, "Copy Parent Parameters to Children"))
                //{
                //    t.Start();

                //    foreach (var group in grouppedBrokendMembers)
                //    {
                //        List<Frame> groupOfMembers = group.ToList();
                //        string parentRevitID = group.Key;

                //        try
                //        {
                //            int idValue = int.Parse(parentRevitID);
                //            ElementId elementId = new ElementId(idValue);
                //            Element parentRevitElement = doc.GetElement(elementId);

                //            if (parentRevitElement == null)
                //                continue;

                //            foreach (var childElement in groupOfMembers)
                //            {
                //                int childIdValue = int.Parse(childElement.ID);
                //                ElementId childElementId = new ElementId(childIdValue);
                //                Element childRevitElement = doc.GetElement(childElementId);

                //                if (childRevitElement == null)
                //                    continue;

                //                foreach (var parameterName in HardCodedData.allowedParams)
                //                {
                //                    try
                //                    {
                //                        Parameter parentParam = parentRevitElement.LookupParameter(parameterName);
                //                        Parameter childParam = childRevitElement.LookupParameter(parameterName);

                //                        if (parentParam == null || childParam == null)
                //                            continue;

                //                        if (!parentParam.HasValue || childParam.IsReadOnly)
                //                            continue;

                //                        if (parentParam.StorageType != childParam.StorageType)
                //                            continue;

                //                        switch (parentParam.StorageType)
                //                        {
                //                            case StorageType.Double:
                //                                childParam.Set(parentParam.AsDouble());
                //                                break;

                //                            case StorageType.Integer:
                //                                childParam.Set(parentParam.AsInteger());
                //                                break;

                //                            case StorageType.String:
                //                                childParam.Set(parentParam.AsString());
                //                                break;

                //                            case StorageType.ElementId:
                //                                childParam.Set(parentParam.AsElementId());
                //                                break;
                //                        }
                //                    }
                //                    catch
                //                    {
                //                        continue;
                //                    }
                //                }
                //            }
                //        }
                //        catch
                //        {
                //            continue;
                //        }
                //    }

                //    t.Commit();
                //}
                #endregion

                #region Handling Parameters of Broken Elements
                var grouppedBrokendMembers = beams.Where(x => x.ParentID != null).GroupBy(x => x.ParentID);

                using (Transaction t = new Transaction(doc, "Copy Identity Data Parameters"))
                {
                    t.Start();

                    foreach (var group in grouppedBrokendMembers)
                    {
                        List<Frame> groupOfMembers = group.ToList();
                        string parentRevitID = group.Key;

                        try
                        {
                            int idValue = int.Parse(parentRevitID);
                            ElementId elementId = new ElementId(idValue);
                            Element parentRevitElement = doc.GetElement(elementId);

                            if (parentRevitElement == null)
                                continue;

                            foreach (var childElement in groupOfMembers)
                            {
                                int childIdValue = int.Parse(childElement.ID);
                                ElementId childElementId = new ElementId(childIdValue);
                                Element childRevitElement = doc.GetElement(childElementId);

                                if (childRevitElement == null)
                                    continue;

                                // 🔹 Loop through ALL parent parameters
                                foreach (Parameter parentParam in parentRevitElement.Parameters)
                                {
                                    try
                                    {
                                        // 🔹 Filter: only Identity Data
                                        if (parentParam.Definition.GetGroupTypeId() != GroupTypeId.IdentityData)
                                            continue;

                                        if (!parentParam.HasValue)
                                            continue;

                                        Parameter childParam = null;

                                        // 🔹 Safer matching: GUID for shared params
                                        if (parentParam.IsShared)
                                        {
                                            childParam = childRevitElement.get_Parameter(parentParam.GUID);
                                        }
                                        else
                                        {
                                            childParam = childRevitElement.LookupParameter(parentParam.Definition.Name);
                                        }

                                        if (childParam == null || childParam.IsReadOnly)
                                            continue;

                                        if (parentParam.StorageType != childParam.StorageType)
                                            continue;

                                        // 🔹 Copy value safely
                                        switch (parentParam.StorageType)
                                        {
                                            case StorageType.Double:
                                                childParam.Set(parentParam.AsDouble());
                                                break;

                                            case StorageType.Integer:
                                                childParam.Set(parentParam.AsInteger());
                                                break;

                                            case StorageType.String:
                                                childParam.Set(parentParam.AsString());
                                                break;

                                            case StorageType.ElementId:
                                                childParam.Set(parentParam.AsElementId());
                                                break;
                                        }
                                    }
                                    catch
                                    {
                                        continue;
                                    }
                                }
                            }
                        }
                        catch
                        {
                            continue;
                        }
                    }

                    t.Commit();
                }
                #endregion



                tg.Assimilate();


            }

            return result;
        }

        public static List<ElementId> CreateAllColumns3(Document doc, List<Frame> columns, ForgeTypeId dis)
        {
            // AnalyticalToPhysicalAssociationManager analyManger = AnalyticalToPhysicalAssociationManager.GetAnalyticalToPhysicalAssociationManager(doc);

            List<ElementId> result = new List<ElementId>();
            List<FamilyInstance> revitEelemtns = new List<FamilyInstance>();

            #region defintions for creation of beams and columns 
            FamilySymbol revitSection = null;
            double x, y, z, x1, y1, z1;
            XYZ StartPoint = null;
            XYZ EndPoint = null;
            Autodesk.Revit.DB.Curve line;
            Autodesk.Revit.DB.Curve line3;

            double angle = 0;
            Level lvl0;
            Level lvl1;
            Parameter param;
            LocationCurve instanceCurve;
            Parameter pbase;
            Parameter ptop;
            Level baselvl;
            Level toplvl;
            double basePointElevation;
            double topPointElevation;
            double baseLevelElevation;
            double topLevelElevation;
            double baseOffset;
            double topOffset;
            double base_offset;
            double top_offset;
            Line line2;
            double col_angle;


            lvl0 = new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().FirstOrDefault();
            lvl1 = new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().ElementAtOrDefault(1);

            if (lvl0 == null)
            {
                using (Transaction t = new Transaction(doc, "Create Levels"))
                {
                    t.Start();
                    lvl0 = Level.Create(doc, 0.0);
                    lvl0.Name = "Hydra-level-1";

                    t.Commit();

                }
            }

            if (lvl1 == null)
            {
                using (Transaction t = new Transaction(doc, "Create Levels"))
                {
                    t.Start();
                    lvl1 = Level.Create(doc, 10.0);
                    lvl1.Name = "Hydra-level-2";

                    t.Commit();

                }
            }


            if (lvl0.Elevation > lvl1.Elevation)
            {
                lvl0 = new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().ToList()[1];
                lvl1 = new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().ToList().First();
            }

            #endregion

            RevitServices.Transactions.TransactionManager.Instance.ForceCloseTransaction();

            using (TransactionGroup tg = new TransactionGroup(doc, "aaa"))

            {
                tg.Start();

                var chopSize = Frame.chopSize;
                var columnChuncks = columns.Chunk(chopSize).ToList();

                foreach (var columnChunck in columnChuncks)
                {

                    using (Transaction trns = new Transaction(doc, "trns"))
                    {
                        trns.Start();
                        FailureHandlingOptions failOpt = trns.GetFailureHandlingOptions();
                        failOpt.SetFailuresPreprocessor(new RoomWarningSwallower());
                        trns.SetFailureHandlingOptions(failOpt);


                        foreach (var cl in columnChunck)
                        {

                            x = UnitUtils.ConvertToInternalUnits(cl.StartPoint.X, dis);
                            y = UnitUtils.ConvertToInternalUnits(cl.StartPoint.Y, dis);
                            z = UnitUtils.ConvertToInternalUnits(cl.StartPoint.Z, dis);
                            StartPoint = new XYZ(x, y, z);

                            x1 = UnitUtils.ConvertToInternalUnits(cl.EndPoint.X, dis);
                            y1 = UnitUtils.ConvertToInternalUnits(cl.EndPoint.Y, dis);
                            z1 = UnitUtils.ConvertToInternalUnits(cl.EndPoint.Z, dis);
                            EndPoint = new XYZ(x1, y1, z1);


                            try
                            {
                                line = Line.CreateBound(StartPoint, EndPoint);
                                line3 = Line.CreateBound(EndPoint, StartPoint);

                                revitSection = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralColumns).WhereElementIsElementType().Cast<FamilySymbol>().ToList().FirstOrDefault(a => a.Name == cl.Sec);

                                if (revitSection == null)
                                {
                                    revitSection = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralColumns).WhereElementIsElementType().Cast<FamilySymbol>().ToList().FirstOrDefault();

                                }

                                if (!revitSection.IsActive)
                                    revitSection.Activate();

                                FamilyInstance instance = doc.Create.NewFamilyInstance(StartPoint, revitSection, lvl0, StructuralType.Column);
                                // AnalyticalMember analytInstance = AnalyticalMember.Create(doc, line);
                                //analyManger.AddAssociation(analytInstance.Id, instance.Id);

                                if (cl.MemberType == MemberType.SlanttedColumn)
                                {
                                    param = instance.LookupParameter("Column Style");
                                    param.Set((int)SlantedOrVerticalColumnType.CT_EndPoint);

                                    instanceCurve = (LocationCurve)instance.Location;
                                    if (instanceCurve != null)
                                    {
                                        if (z > z1)
                                        {
                                            instanceCurve.Curve = line3;
                                        }
                                        else
                                        {
                                            instanceCurve.Curve = line;

                                        }

                                    }

                                }

                                else
                                {
                                    #region handling levels elevations
                                    pbase = instance.LookupParameter("Base Level");
                                    ptop = instance.LookupParameter("Top Level");

                                    //
                                    pbase.Set(lvl0.Id);
                                    ptop.Set(lvl1.Id);


                                    baselvl = doc.GetElement(pbase.AsElementId()) as Level;
                                    toplvl = doc.GetElement(ptop.AsElementId()) as Level;


                                    basePointElevation = Convert.ToDouble(cl.StartPoint.Z);
                                    topPointElevation = Convert.ToDouble(cl.EndPoint.Z);

                                    if (basePointElevation > topPointElevation)
                                    {
                                        basePointElevation = Convert.ToDouble(cl.EndPoint.Z);
                                        topPointElevation = Convert.ToDouble(cl.StartPoint.Z);

                                    }


                                    baseLevelElevation = UnitUtils.ConvertFromInternalUnits(baselvl.ProjectElevation, dis);
                                    topLevelElevation = UnitUtils.ConvertFromInternalUnits(toplvl.ProjectElevation, dis);


                                    baseOffset = basePointElevation - baseLevelElevation;
                                    topOffset = topPointElevation - topLevelElevation;

                                    base_offset = UnitUtils.ConvertToInternalUnits(basePointElevation - baseLevelElevation, dis);
                                    top_offset = UnitUtils.ConvertToInternalUnits(topPointElevation - topLevelElevation, dis);




                                    if ((topLevelElevation + topOffset) == baseLevelElevation)
                                    {
                                        instance.LookupParameter("Base Offset").Set(base_offset);
                                        instance.LookupParameter("Top Offset").Set(top_offset);

                                    }
                                    else
                                    {
                                        instance.LookupParameter("Top Offset").Set(top_offset);
                                        instance.LookupParameter("Base Offset").Set(base_offset);
                                    }


                                    #endregion

                                }


                                line2 = Line.CreateBound(StartPoint, EndPoint);
                                col_angle = (cl.Beta) * Math.PI / 180;
                                ElementTransformUtils.RotateElement(doc, instance.Id, line2, col_angle);

                                // set Parameters 

                                cl.ID = instance.Id.ToString();
                                RhinoDoc.ActiveDoc.Objects.FindId(new Guid(cl.RhinoID)).Attributes.SetUserString("ID", instance.Id.ToString());
                                RhinoDoc.ActiveDoc.Objects.FindId(new Guid(cl.RhinoID)).Attributes.SetUserString("Section", instance.Symbol.Name.ToString());
                                RhinoDoc.ActiveDoc.Objects.FindId(new Guid(cl.RhinoID)).Attributes.SetUserString("Family", instance.Symbol.FamilyName.ToString());
                                RhinoDoc.ActiveDoc.Objects.FindId(new Guid(cl.RhinoID)).Attributes.SetUserString("MemberType", "Column");
                                RhinoDoc.ActiveDoc.Objects.FindId(new Guid(cl.RhinoID)).Attributes.SetUserString("Angle", cl.Beta.ToString());


                                try
                                {
                                    revitEelemtns.Add(instance);
                                    result.Add(instance.Id);

                                }

                                catch (Exception)
                                {
                                    continue;
                                }

                            }


                            catch (Exception ex2)

                            {
                                System.Diagnostics.StackTrace trace = new System.Diagnostics.StackTrace(ex2, true);
                                RevitUtils.Log.Add("Error in CreateAllColumns3 Method... Col no" + cl.ID + ex2.Message + trace.GetFrame(trace.FrameCount - 1).GetFileLineNumber());

                                //TaskDialog.Show("Error", ex2.Message + " ,line of the error is " + trace.GetFrame(trace.FrameCount - 1).GetFileLineNumber());
                                continue;
                            }



                        }

                        trns.Commit();


                    }

                }

                #region Handling Parameters of Broken Elements

                //var grouppedBrokendMembers = columns.Where(m => m.ParentID != null).GroupBy(q => q.ParentID);

                //using (Transaction t = new Transaction(doc, "Copy Parent Parameters to Children"))
                //{
                //    t.Start();

                //    foreach (var group in grouppedBrokendMembers)
                //    {
                //        List<Frame> groupOfMembers = group.ToList();
                //        string parentRevitID = group.Key;

                //        try
                //        {
                //            int idValue = int.Parse(parentRevitID);
                //            ElementId elementId = new ElementId(idValue);
                //            Element parentRevitElement = doc.GetElement(elementId);

                //            if (parentRevitElement == null)
                //                continue;

                //            foreach (var childElement in groupOfMembers)
                //            {
                //                int childIdValue = int.Parse(childElement.ID);
                //                ElementId childElementId = new ElementId(childIdValue);
                //                Element childRevitElement = doc.GetElement(childElementId);

                //                if (childRevitElement == null)
                //                    continue;

                //                foreach (var parameterName in HardCodedData.allowedParams)
                //                {
                //                    try
                //                    {
                //                        Parameter parentParam = parentRevitElement.LookupParameter(parameterName);
                //                        Parameter childParam = childRevitElement.LookupParameter(parameterName);

                //                        if (parentParam == null || childParam == null)
                //                            continue;

                //                        if (!parentParam.HasValue || childParam.IsReadOnly)
                //                            continue;

                //                        if (parentParam.StorageType != childParam.StorageType)
                //                            continue;

                //                        switch (parentParam.StorageType)
                //                        {
                //                            case StorageType.Double:
                //                                childParam.Set(parentParam.AsDouble());
                //                                break;

                //                            case StorageType.Integer:
                //                                childParam.Set(parentParam.AsInteger());
                //                                break;

                //                            case StorageType.String:
                //                                childParam.Set(parentParam.AsString());
                //                                break;

                //                            case StorageType.ElementId:
                //                                childParam.Set(parentParam.AsElementId());
                //                                break;
                //                        }
                //                    }
                //                    catch
                //                    {
                //                        continue;
                //                    }
                //                }
                //            }
                //        }
                //        catch
                //        {
                //            continue;
                //        }
                //    }

                //    t.Commit();
                //}
                #endregion


                #region Handling Parameters of Broken Elements
                var grouppedBrokendMembers = columns.Where(m => m.ParentID != null).GroupBy(q => q.ParentID);

                using (Transaction t = new Transaction(doc, "Copy Identity Data Parameters"))
                {
                    t.Start();

                    foreach (var group in grouppedBrokendMembers)
                    {
                        List<Frame> groupOfMembers = group.ToList();
                        string parentRevitID = group.Key;

                        try
                        {
                            int idValue = int.Parse(parentRevitID);
                            ElementId elementId = new ElementId(idValue);
                            Element parentRevitElement = doc.GetElement(elementId);

                            if (parentRevitElement == null)
                                continue;

                            foreach (var childElement in groupOfMembers)
                            {
                                int childIdValue = int.Parse(childElement.ID);
                                ElementId childElementId = new ElementId(childIdValue);
                                Element childRevitElement = doc.GetElement(childElementId);

                                if (childRevitElement == null)
                                    continue;

                                // 🔹 Loop through ALL parent parameters
                                foreach (Parameter parentParam in parentRevitElement.Parameters)
                                {
                                    try
                                    {
                                        // 🔹 Filter: only Identity Data
                                        if (parentParam.Definition.GetGroupTypeId() != GroupTypeId.IdentityData)
                                            continue;

                                        if (!parentParam.HasValue)
                                            continue;

                                        Parameter childParam = null;

                                        // 🔹 Safer matching: GUID for shared params
                                        if (parentParam.IsShared)
                                        {
                                            childParam = childRevitElement.get_Parameter(parentParam.GUID);
                                        }
                                        else
                                        {
                                            childParam = childRevitElement.LookupParameter(parentParam.Definition.Name);
                                        }

                                        if (childParam == null || childParam.IsReadOnly)
                                            continue;

                                        if (parentParam.StorageType != childParam.StorageType)
                                            continue;

                                        // 🔹 Copy value safely
                                        switch (parentParam.StorageType)
                                        {
                                            case StorageType.Double:
                                                childParam.Set(parentParam.AsDouble());
                                                break;

                                            case StorageType.Integer:
                                                childParam.Set(parentParam.AsInteger());
                                                break;

                                            case StorageType.String:
                                                childParam.Set(parentParam.AsString());
                                                break;

                                            case StorageType.ElementId:
                                                childParam.Set(parentParam.AsElementId());
                                                break;
                                        }
                                    }
                                    catch
                                    {
                                        continue;
                                    }
                                }
                            }
                        }
                        catch
                        {
                            continue;
                        }
                    }

                    t.Commit();
                }
                #endregion


                tg.Assimilate();

            }

            return result;

        }

        public static void deleteUnMappedRevitElements(Document doc, List<Frame> revitFreamstoBeDeleted)
        {

            List<ElementId> idsToBeDeleted = new List<ElementId>();

            foreach (var item in revitFreamstoBeDeleted)
            {

                FamilyInstance instance = (FamilyInstance)doc.GetElement(item.revitEleId);
                idsToBeDeleted.Add(instance.Id);


            }

            using (Transaction trns = new Transaction(doc))
            {
                trns.Start("deleteElement");
                doc.Delete(idsToBeDeleted);
                trns.Commit();

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


                                // if (topPoint.Z > topPt.Z)
                                // {

                                topPt.X = topPoint.X;
                                topPt.Y = topPoint.Y;
                                topPt.Z = topPoint.Z;

                                // }
                            }

                            else if (normal.IsAlmostEqualTo(-XYZ.BasisZ))
                            {
                                // This is the bottom face
                                BoundingBoxUV bbox = planarFace.GetBoundingBox();
                                UV midPoint = (bbox.Min + bbox.Max) / 2;
                                XYZ bottomPoint = planarFace.Evaluate(midPoint);
                                // if (bottomPoint.Z < bottomPt.Z)
                                //{
                                bottomPt.X = bottomPoint.X;
                                bottomPt.Y = bottomPoint.Y;
                                bottomPt.Z = bottomPoint.Z;
                                // }
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


        public static void CreateOrUpdateSelectionFilter(
                                                          UIDocument uidoc,
                                                          ICollection<ElementId> elementIds,
                                                          string filterName,
                                                          Color lineColor = null)
        {
            Document doc = uidoc.Document;
            View view = doc.ActiveView; // always use active view

            // Default color (blue) if not passed
            Color finalColor = lineColor ?? new Color(0x00, 0x00, 0xFF);

            using (Transaction t = new Transaction(doc, "Create/Update SelectionFilterElement"))
            {
                t.Start();

                // Check if a filter with the same name already exists
                SelectionFilterElement filterElement = new FilteredElementCollector(doc)
                    .OfClass(typeof(SelectionFilterElement))
                    .Cast<SelectionFilterElement>()
                    .FirstOrDefault(f => f.Name.Equals(filterName, StringComparison.OrdinalIgnoreCase));

                // If not found, create new
                if (filterElement == null)
                {
                    filterElement = SelectionFilterElement.Create(doc, filterName);
                }

                // Update element ids
                filterElement.SetElementIds(elementIds);

                //// Get its Id
                //ElementId filterId = filterElement.Id;

                //// If the view does not already contain this filter, add it
                //if (!view.GetFilters().Contains(filterId))
                //{
                //    view.AddFilter(filterId);
                //}

                //// Apply overrides
                //OverrideGraphicSettings overrideSettings = view.GetFilterOverrides(filterId);
                //overrideSettings.SetProjectionLineColor(finalColor);
                //view.SetFilterOverrides(filterId, overrideSettings);

                t.Commit();
            }
        }

        #region test

        public static void RhinoToRhino(IGH_DataAccess DA)
        {

            Rhino.RhinoDoc doc = RhinoDoc.ActiveDoc;

            RhinoObject[] rhinoCurves = doc.Objects.FindByObjectType(ObjectType.Curve).Where(x => x.Attributes.GetUserString("MemberType") != null).ToArray();

            foreach (var item in rhinoCurves)
            {
                item.Attributes.SetUserString("RevitSection", item.Attributes.GetUserString("Section"));
                item.CommitChanges();

            }

        }


        #region safe and Revit Integration

        private static CSiAPIv1.cSapModel getCSapModel()
        {
            CSiAPIv1.cOAPI mySafeObject = (CSiAPIv1.cOAPI)ComActiveObject.Get("CSI.SAFE.API.ETABSObject");

            CSiAPIv1.cSapModel mySapModel = mySafeObject.SapModel;

            mySafeObject.SapModel.SetPresentUnits(CSiAPIv1.eUnits.Ton_mm_C);
            return mySapModel;
        }

        public static void drawSafePoints()
        {

            CSiAPIv1.cSapModel mySapModel = getCSapModel();

            #region defintions 
            string[] PointIDs = null;
            int noOfPoints = 0;
            int ret;

            #endregion

            Rhino.RhinoDoc doc = Rhino.RhinoDoc.ActiveDoc;

            int number = 0;
            int[] objectType = null;
            string[] ids = new string[0];

            mySapModel.SelectObj.GetSelected(ref number, ref objectType, ref ids);

            ret = mySapModel.PointObj.GetNameList(ref noOfPoints, ref PointIDs);



            double x, y, z;
            x = y = z = 0;

            bool s = true;
            string springName = null;
            foreach (var item in ids)
            {
                mySapModel.PointObj.GetCoordCartesian(item, ref x, ref y, ref z);
                mySapModel.PointObj.GetSpringAssignment(item, ref springName);





                Point3d rhinoPt = new Point3d(x, y, z);

                ObjectAttributes att = new ObjectAttributes();
                att.SetUserString("SafePtID", item);
                att.SetUserString("SpringName", springName);
                doc.Objects.AddPoint(rhinoPt, att);


            }


        }

        public static void PilesFromRevitToRhino()
        {

            // for revit 2021 or above use (ForgeTypeId dis = UnitTypeId.Millimeters;) however for revit 2020 or lower use ( ForgeTypeId dis = DisplayUnitType.DUT_MILLIMETERS;)
            ForgeTypeId dis = UnitTypeId.Millimeters;

            UIDocument revitUiDoc = RhinoInside.Revit.Revit.ActiveUIDocument;
            Document revitDoc = RhinoInside.Revit.Revit.ActiveDBDocument;

            Rhino.RhinoDoc doc = RhinoDoc.ActiveDoc;

            #region get Selected Elements 
            ICollection<ElementId> selectedIds = revitUiDoc.Selection.GetElementIds();

            // Create a filter for Structural Framing and Structural Beams categories
            List<BuiltInCategory> categories = new List<BuiltInCategory>();
            categories.Add(BuiltInCategory.OST_StructuralFoundation);

            // Create an ElementCategoryFilter for each category and combine them with a LogicalOrFilter
            List<ElementFilter> categoryFilters = new List<ElementFilter>();
            foreach (BuiltInCategory category in categories)
            {
                categoryFilters.Add(new ElementCategoryFilter(category));
            }
            //
            LogicalOrFilter orFilter = new LogicalOrFilter(categoryFilters);

            // Get the selected elements that match the filter
            List<Element> selectedElements = new List<Element>();
            foreach (ElementId id in selectedIds)
            {
                Element element = revitUiDoc.Document.GetElement(id);


                FilteredElementCollector tagCollector = new FilteredElementCollector(revitDoc).OfCategory(BuiltInCategory.OST_Tags);

                if (orFilter.PassesFilter(element))
                {
                    selectedElements.Add(element);
                }
            }

            
            List<FamilyInstance> AllFoundations = selectedElements
                .Where(x =>
                    x.Category != null &&
                    x.Category.Id.IntegerValue == (int)BuiltInCategory.OST_StructuralFoundation &&
                    x.Category.Id.IntegerValue != (int)BuiltInCategory.OST_Floors)
                .OfType<FamilyInstance>()
                .ToList();

            #endregion

            if (AllFoundations.Count == 0)
            {
                AllFoundations = new FilteredElementCollector(revitDoc)
                    .OfCategory(BuiltInCategory.OST_StructuralFoundation)
                    .WhereElementIsNotElementType()
                    .OfType<FamilyInstance>()
                    .ToList();
            }





            foreach (var item in AllFoundations)
            {
                XYZ pt = (item.Location as LocationPoint).Point;

                Point3d PileLocation = RhinoInside.Revit.Convert.Geometry.GeometryDecoder.ToPoint3d(pt);
                ObjectAttributes att = new ObjectAttributes();

                var symbol = item.Symbol;
                if (symbol.Family.Name.Contains("Cap"))
                {
                    continue;
                }

                #region getParameters
                //string MaxFx = item.LookupParameter("Max_Fx").AsValueString();
                //string MaxFy = item.LookupParameter("Max_Fy").AsValueString();
                //string MaxFz = item.LookupParameter("Max_Fz").AsValueString();
                //string MinFx = item.LookupParameter("Min_Fx").AsValueString();
                //string MinFy = item.LookupParameter("Min_Fy").AsValueString();
                //string MinFz = item.LookupParameter("Min_Fz").AsValueString();

                //string PileDesignGroup = item.LookupParameter("WME Pile Design Grouping Mark").AsValueString();

                //string PileNumber = item.LookupParameter("WME_Pile Number").AsValueString();

                //string PileType = item.Symbol.Name;

                #endregion



                att.SetUserString("RevitPileID", item.Id.ToString());
                //att.SetUserString("MaxFx", MaxFx);
                //att.SetUserString("MaxFy", MaxFy);
                //att.SetUserString("MaxFz", MaxFz);
                //att.SetUserString("MinFx", MinFx);
                //att.SetUserString("MinFy", MinFy);
                //att.SetUserString("MinFz", MinFz);
                //att.SetUserString("PileDesignGroup", PileDesignGroup);
                //att.SetUserString("PileNumber", PileNumber);
                //att.SetUserString("PileType", PileType);


                doc.Objects.AddPoint(PileLocation, att);

            }


        }

        public static void MapBetweenSafeAndRevit()
        {
            Rhino.RhinoDoc doc = Rhino.RhinoDoc.ActiveDoc;

            RhinoObject[] rhinoPoints = doc.Objects.FindByObjectType(ObjectType.Point).Where(x => x.Attributes.GetUserString("SafePtID") != null).ToArray();

            RhinoObject[] revitPoints = doc.Objects.FindByObjectType(ObjectType.Point).Where(x => x.Attributes.GetUserString("RevitPileID") != null).ToArray();




            foreach (var rhinoPoint in rhinoPoints)
            {

                Point3d rhinoPtLocation = (rhinoPoint as PointObject).PointGeometry.Location;

                RhinoObject closestRevitPoint = null;

                double closestDistance = 500;

                foreach (var revitPoint in revitPoints)
                {
                    Point3d revitPtLocation = (revitPoint as PointObject).PointGeometry.Location;
                    double distance = rhinoPtLocation.DistanceTo(revitPtLocation);

                    if (distance < closestDistance)
                    {
                        closestDistance = distance;
                        closestRevitPoint = revitPoint;
                    }
                }

                if (closestRevitPoint != null)
                {
                    closestRevitPoint.Attributes.SetUserString("CrosspSafeId", rhinoPoint.Attributes.GetUserString("SafePtID"));
                    closestRevitPoint.Attributes.SetUserString("CrosspSafeSpring", rhinoPoint.Attributes.GetUserString("SpringName"));
                    closestRevitPoint.CommitChanges();



                    rhinoPoint.Attributes.SetUserString("CrosspRevitId", closestRevitPoint.Attributes.GetUserString("RevitPileID"));
                    rhinoPoint.Attributes.SetUserString("CrossRevitPileType", closestRevitPoint.Attributes.GetUserString("RevitPileType"));
                    rhinoPoint.CommitChanges();

                }
            }

        }

        public static void ExtractToExcel()
        {


            #region Tables
            List<string> type = new List<string> { "P1A", "P1B", "P1C", "P1D", "P1E", "P1F", "P1G", "P2A", "P2B", "P2C", "P2D", "P2E", "P3A", "P3B", "P3C", "P4A", "P4B", "P4C", "P4D", "P4E", "P5A", "P5B", "P5C", "P5D", "P5E", "P6A", "P7A", "P7B", "P7C", "P7D", "P7E", "P8A", "P8B" };
            List<int> nos = new List<int> { 54, 210, 291, 136, 77, 58, 26, 42, 644, 392, 6, 51, 240, 224, 27, 49, 65, 11, 43, 36, 22, 33, 10, 31, 8, 29, 20, 65, 139, 50, 2, 128, 18 };
            List<int> dia = new List<int> { 900, 900, 900, 900, 900, 900, 900, 1000, 1000, 1000, 1000, 1000, 1200, 1200, 1200, 900, 900, 900, 900, 900, 1000, 1000, 1000, 1000, 1000, 900, 900, 900, 900, 900, 900, 900, 900 };
            List<int> fc = new List<int> { 60, 60, 60, 60, 60, 60, 60, 60, 60, 60, 60, 60, 60, 60, 60, 60, 60, 60, 60, 60, 60, 60, 60, 60, 60, 60, 60, 60, 60, 60, 60, 60, 60 };
            List<double> RAD = new List<double> { -6.565, -6.565, -6.565, -6.565, -6.565, -6.565, -6.565, -6.565, -6.565, -6.565, -8.915, -9.065, -9.415, -7.065, -10.415, -6.565, -6.565, -6.565, -6.565, -6.565, -6.565, -6.565, -6.565, -9.565, -10.565, -6.565, -6.565, -6.565, -6.565, -6.565, -6.565, -9.565, -9.565 };
            List<int> length = new List<int> { 37, 37, 40, 40, 40, 40, 40, 42, 42, 42, 42, 42, 42, 42, 42, 40, 40, 40, 40, 40, 40, 40, 40, 40, 40, 40, 37, 37, 40, 40, 40, 40, 40 };
            List<int> max_axial = new List<int> { 1396, 2976, 3117, 3988, 4081, 4470, 4470, 5178, 5570, 5570, 5570, 5570, 6910, 6910, 6910, 3472, 4470, 3699, 4470, 3502, 3870, 4808, 5070, 5070, 4748, 1968, 1998, 2967, 3960, 4470, 4470, 4470, 4470 };
            List<int> corres_shear1 = new List<int> { 115, 132, 148, 138, 132, 145, 85, 118, 120, 136, 81, 100, 97, 143, 185, 170, 232, 236, 218, 124, 138, 215, 260, 219, 166, 136, 86, 81, 130, 125, 75, 215, 120 };
            List<int> min_axial = new List<int> { -359, -621, -1040, -888, -1162, -944, -1198, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, -1088, -1229, -525, -1320, -1982, -2154, 0, 0, 0, 0, 0, -1447, -1733 };
            List<int> corrs_shear2 = new List<int> { 112, 126, 124, 106, 115, 108, 105, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 149, 153, 156, 132, 97, 80, 0, 0, 0, 0, 0, 116, 78 };
            List<int> transient_axial = new List<int> { -783, -1406, -1897, -1691, -2203, -2590, -2590, 0, -885, -1432, -89, -476, -3029, -819, -683, -1994, -1100, -1345, -929, -1022, -2298, -2575, -1764, -2880, -2520, -2389, -616, -983, -1375, -1822, -576, -2569, -2459 };
            List<int> corres_shear3 = new List<int> { 109, 82, 98, 161, 132, 142, 72, 144, 57, 115, 84, 100, 126, 163, 126, 250, 153, 211, 208, 134, 171, 170, 186, 236, 152, 142, 86, 77, 107, 80, 75, 137, 134 };
            List<int> max_shear = new List<int> { 154, 205, 203, 180, 198, 219, 186, 154, 211, 207, 84, 153, 210, 240, 207, 261, 241, 236, 225, 189, 283, 250, 260, 240, 169, 177, 163, 176, 230, 212, 75, 215, 137 };
            List<int> corres_axial = new List<int> { 1286, 2707, 2928, 3221, 3135, 2202, 1415, 4272, 3223, 3153, 5216, 4341, 5947, 5904, 5785, 2504, 3941, 40, 4470, 1564, 3260, 3780, 40, 4409, 4399, 1331, 1659, 1601, 2791, 3645, 4470, 40, 3946 };
            List<int> compn = new List<int> { 3820, 3820, 4470, 4470, 4470, 4470, 4470, 5570, 5570, 5570, 5570, 5570, 6910, 6910, 6910, 4470, 4470, 4470, 4470, 4470, 5070, 5070, 5070, 5070, 5070, 4470, 3820, 3820, 4470, 4470, 4470, 4470, 4470 };
            List<int> utilization1 = new List<int> { 37, 78, 70, 89, 91, 100, 100, 93, 100, 100, 100, 100, 100, 100, 100, 78, 100, 83, 100, 78, 76, 95, 100, 100, 94, 44, 52, 77, 95, 100, 100, 100, 100 };
            List<int> utilization2 = new List<int> { 40, 40, 41, 44, 44, 47, 52, 43, 56, 54, 40, 40, 44, 52, 47, 40, 44, 40, 44, 40, 55, 55, 44, 55, 55, 40, 40, 40, 44, 44, 44, 50, 44 };
            #endregion
            Rhino.RhinoDoc doc = Rhino.RhinoDoc.ActiveDoc;

            RhinoObject[] rhinoPoints = doc.Objects.FindByObjectType(ObjectType.Point).Where(x => x.Attributes.GetUserString("SafePtID") != null).ToArray();
            RhinoObject[] revitPiles = doc.Objects.FindByObjectType(ObjectType.Point).Where(x => x.Attributes.GetUserString("RevitPileID") != null).ToArray();



            List<object> records = new List<object>();

            foreach (var pt in revitPiles)
            {
                string safePtID = pt.Attributes.GetUserString("CrosspSafeId");
                string SafeSpring = pt.Attributes.GetUserString("CrosspSafeSpring");

                int index = type.IndexOf(SafeSpring.Substring(0, 3));

                string RevitDesignGroup = pt.Attributes.GetUserString("PileDesignGroup");
                string RevitpileNumber = pt.Attributes.GetUserString("PileNumber");
                string PileProfile = pt.Attributes.GetUserString("PileType");
                string maxFxString = pt.Attributes.GetUserString("MaxFx");
                string maxFyString = pt.Attributes.GetUserString("MaxFy");
                string maxFzString = pt.Attributes.GetUserString("MaxFz");
                string minFxString = pt.Attributes.GetUserString("MinFx");
                string minFyString = pt.Attributes.GetUserString("MinFy");
                string minFzString = pt.Attributes.GetUserString("MinFz");


                var rec = new
                {
                    RevitpileNumber = RevitpileNumber,
                    RevitDesignGroup = RevitDesignGroup,
                    SAFESpring = SafeSpring.Substring(0, 3),
                    SAFEPtID = safePtID,

                    //Number = mainGroups.IndexOf(group),
                    //GroupNAme = group,
                    //Section = string.Join(",", sections),
                };

                records.Add(rec);
            }


            #region add in excel

            // Create Excel application
            Application excelApp = new Application();
            excelApp.Visible = true;

            // Create a new workbook
            Workbook workbook = excelApp.Workbooks.Add();

            // Add data to the first worksheet

            ExcelUtils.AddDataToWorksheet2(workbook, records, "SectionsDiscr");

            #endregion


        }


        public static void AssignSafeDataToRevitItem(Document Revitdoc)
        {
            Rhino.RhinoDoc doc = Rhino.RhinoDoc.ActiveDoc;

            RhinoObject[] rhinoPoints = doc.Objects.FindByObjectType(ObjectType.Point).Where(x => x.Attributes.GetUserString("SafePtID") != null).ToArray();
            RhinoObject[] revitPiles = doc.Objects.FindByObjectType(ObjectType.Point).Where(x => x.Attributes.GetUserString("RevitPileID") != null).ToArray();


            var chopSize = 100;
            var revitPilesChunks = revitPiles.Chunk(chopSize);

            foreach (var chunk in revitPilesChunks)
            {

                using (Transaction trns = new Transaction(Revitdoc, "trns"))
                {

                    trns.Start();

                    foreach (var pt in chunk)
                    {
                        string safePtID = pt.Attributes.GetUserString("CrosspSafeId");

                        // Extract only digits, keep as string (leading zeros included)
                        string PileNumber = Regex.Replace(safePtID, @"\D", "");

                        string SafeSpring = pt.Attributes.GetUserString("CrosspSafeSpring");
                        string PileType = SafeSpring.Split('-')[0];

                        string RevitId = pt.Attributes.GetUserString("RevitPileID");
                        int idInt = int.Parse(RevitId);


                        // create ElementId
                        ElementId elemId = new ElementId(idInt);

                        FamilyInstance instance = (FamilyInstance)Revitdoc.GetElement(elemId);

                        instance.LookupParameter("ACM_SE_Main Type ID").Set(PileType);
                        instance.LookupParameter("ACM_SE_Element Number").Set((PileNumber));

                    }

                    trns.Commit();
                }



               
            }


          

        }
        #endregion


        public static void RevitSchdulesManupulation(string scheduleName, List<Record> records)
        {

            UIDocument uidoc = RhinoInside.Revit.Revit.ActiveUIDocument;
            Document doc = RhinoInside.Revit.Revit.ActiveDBDocument;

            // Retrieve the key schedule by name
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            ViewSchedule keySchedule = collector.OfClass(typeof(ViewSchedule))
                                                .Cast<ViewSchedule>()
                                                .FirstOrDefault(vs => vs.Name == scheduleName);



            //// Collect all elements (rows) in the key schedule
            FilteredElementCollector scheduleCollector = new FilteredElementCollector(doc, keySchedule.Id);


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

            #region Assign new Element to schedule
            //using (Transaction trans = new Transaction(doc, "Modify Key Schedule"))
            //{

            //    trans.Start();
            //    TableData tableData = keySchedule.GetTableData();

            //    // Get the section data for the body
            //    TableSectionData sectionData = tableData.GetSectionData(SectionType.Body);


            //    for (int i = 0; i < 10; i++)
            //    {
            //        sectionData.InsertRow(sectionData.LastRowNumber);
            //        Element Newelement = scheduleCollector.ToElements().Last();
            //        Newelement.LookupParameter("Key Name").Set(i.ToString());
            //    }
            //    trans.Commit();
            //}

            #endregion

            #region Get All fileds in a Schedule
            List<string> filedsNAmes = new List<string>();
            IList<SchedulableField> schedulableFields = keySchedule.Definition.GetSchedulableFields();
            for (int i = 0; i < schedulableFields.Count; i++)
            {
                SchedulableField field = schedulableFields[i];
                string fieldName = field.GetName(doc);
                filedsNAmes.Add(fieldName);
            }

            int a = 1;

            #endregion

            #region DeleteAllElementsInSchedule

            //IList<Element> allElements = scheduleCollector.ToElements();

            //using (Transaction trans = new Transaction(doc, "Modify Key Schedule"))
            //{

            //    trans.Start();
            //    foreach (var item in allElements)
            //    {
            //        doc.Delete(item.Id);

            //    }

            //    trans.Commit();
            //}



            #endregion



            #region UpdateAllElementsInSchedule

            IList<Element> allElements = scheduleCollector.ToElements();

            using (Transaction trans = new Transaction(doc, "Modify Key Schedule"))
            {

                trans.Start();
                foreach (var item in allElements)
                {
                    doc.Delete(item.Id);

                }

                trans.Commit();
            }


            using (Transaction trans = new Transaction(doc, "Modify Key Schedule"))
            {

                trans.Start();
                TableData tableData = keySchedule.GetTableData();

                // Get the section data for the body
                TableSectionData sectionData = tableData.GetSectionData(SectionType.Body);
                // Insert a new row at the first row number


                for (int i = 0; i < records.Count - 1; i++)
                {
                    sectionData.InsertRow(sectionData.LastRowNumber);
                    Element Newelement = scheduleCollector.ToElements().Last();
                    Newelement.LookupParameter(HardCodedData.ColumnMarkDic["Mark"]).Set(records[i].Mark);
                    Newelement.LookupParameter(HardCodedData.ColumnMarkDic["SubMark"]).Set(records[i].SubMark);
                    Newelement.LookupParameter(HardCodedData.ColumnMarkDic["Depth"]).Set(records[i].Depth);
                    Newelement.LookupParameter(HardCodedData.ColumnMarkDic["Width"]).Set(records[i].Width);
                    Newelement.LookupParameter(HardCodedData.ColumnMarkDic["Designation"]).Set(records[i].Designation);



                }
                trans.Commit();
            }

            #endregion

        }

        #region SetAnalysisIDs


        public static void SetAnalysisIDForMappedInRevitElements(Document revitDoc)
        {
            string SapIDParameter = "ACM_SE_SBM_SapID";
            string EtabsIDParametr = "ACM_SE_SBM_EtabsID";  //ACM_SE_SBM_EtabsID


            string ParamterToBeUSed = SapIDParameter;
            string ParamterToBeUSed2 = EtabsIDParametr;



            List<Frame> rhinoFrames = new List<Frame>();
            Rhino.RhinoDoc doc = RhinoDoc.ActiveDoc;

            List<RhinoObject> allRhinoCurves = doc.Objects.FindByObjectType(ObjectType.Curve).ToList();
            List<RhinoObject> MappedrhinoCurves = allRhinoCurves.Where(x => x.Attributes.GetUserString("Family") != null && x.Attributes.GetUserString("IsMapped") == "True").ToList();

            List<RhinoObject> UnMappedrhinoCurves = allRhinoCurves.Where(x => x.Attributes.GetUserString("Family") != null && x.Attributes.GetUserString("IsMapped") == "False").ToList();


            #region New
            using (Transaction trns = new Transaction(revitDoc, "AssignElements"))
            {
                trns.Start();
                foreach (var crv in MappedrhinoCurves)
                {

                    NameValueCollection n = crv.Attributes.GetUserStrings();
                    string RevitID = crv.Attributes.GetUserString("ID");
                    string AnalysisId = crv.Attributes.GetUserString("EtabsCropssID");

                    ElementId RevitEleID = new ElementId(int.Parse(RevitID));
                    Element revitElement = revitDoc.GetElement(RevitEleID);
                    if (revitElement != null)
                    {
                        revitElement.LookupParameter(ParamterToBeUSed).Set(AnalysisId);
                        revitElement.LookupParameter(ParamterToBeUSed2).Set(AnalysisId);


                    }



                }

                foreach (var crv in UnMappedrhinoCurves)
                {

                    NameValueCollection n = crv.Attributes.GetUserStrings();
                    string RevitID = crv.Attributes.GetUserString("ID");
                    ElementId RevitEleID = new ElementId(int.Parse(RevitID));

                    Element revitElement = revitDoc.GetElement(RevitEleID);
                    if (revitElement != null)
                    {
                        revitElement.LookupParameter(ParamterToBeUSed).Set(" ");

                    }



                }

                trns.Commit();

            }

            #endregion


        }



        public static void SetParameterFromRevitElementToAnother(Document revitDoc)
        {

            List<ElementId> elementsWithSizesChanged = new List<ElementId>();

            List<string> parameters = new List<string>()
                {
                    "ACM_SE_BasePlate_ID",
                    "ACM_SE_ConnectionType",
                    "ACM_SE_Design Section",
                    "ACM_SE_Element Remarks",
                    "ACM_SE_EndMajorM_Release",
                    "ACM_SE_EndMinorM_Release",
                    "ACM_SE_Main Type ID",
                    "ACM_SE_Member Usage",
                    "ACM_SE_SBM_AnalysisMaterial",
                    "ACM_SE_SBM_EtabsID",
                    "ACM_SE_SBM_SapID",
                    "ACM_SE_StartMajorM_Release",
                    "ACM_SE_StartMinorM_Release",
                    "ACM_SE_SubZone_Level_Name",
                    "ACM_SE_Zone_Level_ID",
                    "ACM_SE_Zone_Level_Name",
                    "BEAM TYPE",

                    "ACM_SE_Member Force_Fx",
                    "ACM_SE_Member Force_Fy",
                    "ACM_SE_Member Force_Fz",
                    "ACM_SE_Member Force_Mx",
                    "ACM_SE_Member Force_My",
                    "ACM_SE_Member Force_Mz",
                    "ACM_SE_Member Force_Fx2",
                    "ACM_SE_Member Force_Fy2",
                    "ACM_SE_Member Force_Fz2",
                    "ACM_SE_Member Force_Mx2",
                    "ACM_SE_Member Force_My2",
                    "ACM_SE_Member Force_Mz2",
                    "ACM_SE_Member Force_Fx3",
                    "ACM_SE_Member Force_Fy3",
                    "ACM_SE_Member Force_Fz3",
                    "ACM_SE_Member Force_Mx3",
                    "ACM_SE_Member Force_My3",
                    "ACM_SE_Member Force_Mz3",
                };
           

            RhinoDoc doc = RhinoDoc.ActiveDoc;

            if (doc == null)
                return;


            // Get all Rhino curves
            List<RhinoObject> allRhinoCurves =
                doc.Objects
                   .FindByObjectType(ObjectType.Curve)
                   .ToList();


            // Get mapped Rhino curves
            List<RhinoObject> mappedRhinoCurves =
                allRhinoCurves
                    .Where(x =>
                        x.Attributes.GetUserString("Family") != null &&
                        x.Attributes.GetUserString("IsMapped") == "True")
                    .ToList();


            using (Transaction trns =
                   new Transaction(revitDoc, "Assign Elements"))
            {
                trns.Start();


                foreach (RhinoObject crv in mappedRhinoCurves)
                {
                    string targetRevitID =
                        crv.Attributes.GetUserString("ID");

                    string sourceRevitID =
                        crv.Attributes.GetUserString("EtabsCropssID");


                    // Validate IDs
                    if (!int.TryParse(targetRevitID, out int targetId))
                        continue;

                    if (!int.TryParse(sourceRevitID, out int sourceId))
                        continue;


                    // Get Revit elements
                    ElementId targetRevitEleID =
                        new ElementId(targetId);

                    Element targetRevitElement =
                        revitDoc.GetElement(targetRevitEleID);


                    ElementId sourceRevitEleID =
                        new ElementId(sourceId);

                    Element sourceRevitElement =
                        revitDoc.GetElement(sourceRevitEleID);


                    if (targetRevitElement == null ||
                        sourceRevitElement == null)
                    {
                        continue;
                    }

                    FamilySymbol sourceSymbol = ((FamilyInstance)sourceRevitElement).Symbol;

                    FamilySymbol targetSymbol = ((FamilyInstance)targetRevitElement).Symbol;

                    if (sourceSymbol.Name != targetSymbol.Name)
                    {
                        ((FamilyInstance)targetRevitElement).Symbol = sourceSymbol;
                        elementsWithSizesChanged.Add(targetRevitElement.Id);
                    }



                    // Copy parameters
                    foreach (string parameterName in parameters)
                    {
                        CopyParameter(
                            sourceRevitElement,
                            targetRevitElement,
                            parameterName);
                    }
                }


                CreateOrUpdateSelectionFilterForGivenElementsID("SectionChangedRhino", elementsWithSizesChanged);

                trns.Commit();
            }
        }


        private static bool CopyParameter(
    Element sourceElement,
    Element targetElement,
    string parameterName)
        {
            Parameter sourceParameter =
                sourceElement.LookupParameter(parameterName);

            Parameter targetParameter =
                targetElement.LookupParameter(parameterName);


            // Source parameter doesn't exist
            if (sourceParameter == null)
                return false;


            // Target parameter doesn't exist
            if (targetParameter == null)
                return false;


            // Target parameter cannot be modified
            if (targetParameter.IsReadOnly)
                return false;


            // Make sure both parameters have the same storage type
            if (sourceParameter.StorageType !=
                targetParameter.StorageType)
            {
                return false;
            }


            try
            {
                switch (sourceParameter.StorageType)
                {
                    case StorageType.String:
                        {
                            string value =
                                sourceParameter.AsString();

                            targetParameter.Set(value ?? string.Empty);

                            break;
                        }


                    case StorageType.Double:
                        {
                            double value =
                                sourceParameter.AsDouble();

                            targetParameter.Set(value);

                            break;
                        }


                    case StorageType.Integer:
                        {
                            int value =
                                sourceParameter.AsInteger();

                            targetParameter.Set(value);

                            break;
                        }


                    case StorageType.ElementId:
                        {
                            ElementId value =
                                sourceParameter.AsElementId();

                            targetParameter.Set(value);

                            break;
                        }


                    case StorageType.None:
                    default:
                        return false;
                }


                return true;
            }
            catch
            {
                return false;
            }
        }

        #endregion

        #region FlippingConnectionSymbols

        public static void FlippingConnectionSymbols(Document revitDoc)
        {





        }


        #endregion

        #endregion

        #region 3alSari3

        public static void createGroupOfShownElementsInEtabs()
        {


            ETABSv1.cSapModel MyEtabs = EtabsUtils.Get_sap_mode();

            Rhino.RhinoDoc doc = RhinoDoc.ActiveDoc;

            List<RhinoObject> allRhinoCurves = doc.Objects.FindByObjectType(ObjectType.Curve).ToList();

            List<string> etabsIDS = new List<string>();
            foreach (var crv in allRhinoCurves)
            {
                NameValueCollection n = crv.Attributes.GetUserStrings();
                string AnalysisId = crv.Attributes.GetUserString("ID");
                etabsIDS.Add(AnalysisId);
            }

            EtabsUtils.createGroupinSap(etabsIDS, MyEtabs, "InEtabsNotInRevit");

        }


        public static void createGroupOfShownElementsInREvit()
        {
            #region GetRevitIDs of ShownElementInRhino


            Rhino.RhinoDoc Rhinodoc = RhinoDoc.ActiveDoc;

            List<RhinoObject> allRhinoCurves = Rhinodoc.Objects.FindByObjectType(ObjectType.Curve).ToList();

            List<string> RevitIDS = new List<string>();
            foreach (var crv in allRhinoCurves)
            {
                NameValueCollection n = crv.Attributes.GetUserStrings();
                string AnalysisId = crv.Attributes.GetUserString("ID");
                RevitIDS.Add(AnalysisId);
            }

            #endregion


            Document doc = RhinoInside.Revit.Revit.ActiveDBDocument;


            List<FamilyInstance> TargetRevitElements = new List<FamilyInstance>();
            foreach (var item in RevitIDS)
            {
                ElementId RevitEleID = new ElementId(int.Parse(item));
                FamilyInstance revitElement = doc.GetElement(RevitEleID) as FamilyInstance;
                TargetRevitElements.Add(revitElement);
            }


            using (Transaction trans = new Transaction(doc, "Create Group"))
            {
                trans.Start();

                try
                {
                    // Create a group from the selected element IDs
                    Autodesk.Revit.DB.Group newGroup = doc.Create.NewGroup(TargetRevitElements.Select(x => x.Id).ToList());

                    // Check if a group with the desired name already exists using FilteredElementCollector
                    string desiredGroupName = "InRevitNotInEtabs";

                    // Create a filtered collector to search for groups in the document
                    FilteredElementCollector collector = new FilteredElementCollector(doc);
                    collector.OfClass(typeof(GroupType)); // Only look for GroupType elements

                    // Check if any group has the desired name
                    if (collector.Cast<GroupType>().Any(g => g.Name == desiredGroupName))
                    {
                        // Optionally, handle the case where the group name exists
                        int counter = 1;
                        string newName = desiredGroupName + "_" + counter;

                        // Keep incrementing until we find a unique name
                        while (collector.Cast<GroupType>().Any(g => g.Name == newName))
                        {
                            counter++;
                            newName = desiredGroupName + "_" + counter;
                        }

                        // Set the group name to the unique one
                        newGroup.GroupType.Name = newName;
                    }
                    else
                    {
                        // No conflict, set the group name
                        newGroup.GroupType.Name = desiredGroupName;
                    }

                    trans.Commit();
                }
                catch (Exception ex)
                {
                    trans.RollBack();
                    // Optionally, log or show the exception
                    TaskDialog.Show("Error", ex.Message);
                }
            }
        }


        public static void CreateOrUpdateSelectionFilterForShownElements(string filterName)
        {
            #region GetRevitIDs of ShownElementInRhino


            Rhino.RhinoDoc Rhinodoc = RhinoDoc.ActiveDoc;

            List<RhinoObject> allRhinoCurves = Rhinodoc.Objects.FindByObjectType(ObjectType.Curve).ToList();

            List<string> RevitIDS = new List<string>();
            foreach (var crv in allRhinoCurves)
            {
                NameValueCollection n = crv.Attributes.GetUserStrings();
                string AnalysisId = crv.Attributes.GetUserString("ID");
                RevitIDS.Add(AnalysisId);
            }

            #endregion


            Document doc = RhinoInside.Revit.Revit.ActiveDBDocument;


            List<FamilyInstance> TargetRevitElements = new List<FamilyInstance>();
            foreach (var item in RevitIDS)
            {
                ElementId RevitEleID = new ElementId(int.Parse(item));
                FamilyInstance revitElement = doc.GetElement(RevitEleID) as FamilyInstance;
                TargetRevitElements.Add(revitElement);
            }


            using (Transaction trans = new Transaction(doc, "Set Filter"))
            {
                trans.Start();

                try
                {

                    // Check if a group with the desired name already exists using FilteredElementCollector
                    string desiredGroupName = filterName;

                    // Check if a filter with the same name already exists
                    SelectionFilterElement filterElement = new FilteredElementCollector(doc)
                        .OfClass(typeof(SelectionFilterElement))
                        .Cast<SelectionFilterElement>()
                        .FirstOrDefault(f => f.Name.Equals(desiredGroupName, StringComparison.OrdinalIgnoreCase));

                    // If not found, create new
                    if (filterElement == null)
                    {
                        filterElement = SelectionFilterElement.Create(doc, desiredGroupName);
                    }

                    // Update element ids
                    filterElement.SetElementIds(TargetRevitElements.Select(x=> x.Id).ToList());



                    trans.Commit();
                }
                catch (Exception ex)
                {
                    trans.RollBack();
                    // Optionally, log or show the exception
                    TaskDialog.Show("Error", ex.Message);
                }
            }
        }


        public static void CreateOrUpdateSelectionFilterForGivenElementsID(string filterName, List<ElementId> revitIDS)
        {
            #region GetRevitIDs of ShownElementInRhino


            Rhino.RhinoDoc Rhinodoc = RhinoDoc.ActiveDoc;


            Document doc = RhinoInside.Revit.Revit.ActiveDBDocument;


            #endregion


            try
            {

                // Check if a group with the desired name already exists using FilteredElementCollector

                // Check if a filter with the same name already exists
                SelectionFilterElement filterElement = new FilteredElementCollector(doc)
                    .OfClass(typeof(SelectionFilterElement))
                    .Cast<SelectionFilterElement>()
                    .FirstOrDefault(f => f.Name.Equals(filterName, StringComparison.OrdinalIgnoreCase));

                // If not found, create new
                if (filterElement == null)
                {
                    filterElement = SelectionFilterElement.Create(doc, filterName);
                }

                // Update element ids
                filterElement.SetElementIds(revitIDS);



            }
            catch (Exception ex)
            {
                // Optionally, log or show the exception
            }
            }
        }



        #endregion
    }

