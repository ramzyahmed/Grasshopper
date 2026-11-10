using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Numerics;
using Parameter = Autodesk.Revit.DB.Parameter;

namespace BIM_Integration
{
    [Transaction(TransactionMode.Manual)]
    internal class PlaceFamilyInstance : IExternalCommand
    {

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                #region get Revit Data
                UIDocument uidoc = commandData.Application.ActiveUIDocument;
                Document doc = uidoc.Document;

                // for Revit 2021 or above use (ForgeTypeId dis = UnitTypeId.Millimeters;),
                // however, for Revit 2020 or lower use (ForgeTypeId dis = DisplayUnitType.DUT_MILLIMETERS;)
                ForgeTypeId dis = UnitTypeId.Millimeters;



                #endregion

                double gap = UnitUtils.ConvertToInternalUnits(350, dis);



                // List to store information about each scope box and its adjacent scope boxes
                List<ScopeBoxInfo> scopeBoxInfoList = new List<ScopeBoxInfo>();
                getScopeBoxWithAdjacent(scopeBoxInfoList, doc, gap);


                // create Matching Lines in the current active Veiw
                createMatchingLines(doc, gap);


                // create viewRefrences
                createViewRefrences2(uidoc,doc, scopeBoxInfoList, gap);


                return Result.Succeeded;
            }
            catch (Exception ex2)
            {
                System.Diagnostics.StackTrace trace = new System.Diagnostics.StackTrace(ex2, true);
                TaskDialog.Show("Error", "Error in adding views: " + ex2.Message + " Line: " + trace.GetFrame(trace.FrameCount - 1).ToString());
                return Result.Failed;
            }

        }


        private void createMatchingLines(Document doc, double gap, string MatchingLineSyombol = "STR_DT_BREAK_LINE_QUADRANTI_OPQ")
        {


            // get matchingLine symbol
            FamilySymbol MatchLineSymbol = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).WhereElementIsElementType().Cast<FamilySymbol>().FirstOrDefault(x => x.Name == MatchingLineSyombol);

            // get list of scopeBoxes
            List<Element> scopeBoxes = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_VolumeOfInterest).Cast<Element>().ToList();
            #region get dependent veiws 

            // get active View
            Autodesk.Revit.DB.View activeView = doc.ActiveView;

            List<ElementId> dependentviewsIDS = activeView.GetDependentViewIds().ToList();
            List<ViewPlan> dependentviews = new List<ViewPlan>();
            dependentviewsIDS.ForEach(x => dependentviews.Add(doc.GetElement(x) as ViewPlan));


            List<ViewPlan> childViews = new List<ViewPlan>();
            foreach (var viewId in activeView.GetDependentViewIds())
            {
                childViews.Add(doc.GetElement(viewId) as ViewPlan);
            }

            #endregion




            using (Transaction trns = new Transaction(doc, "Create matchingLines"))
            {
                trns.Start();


                foreach (Element element in scopeBoxes)
                {

                    BoundingBoxXYZ scopeBoundingBox = element.get_BoundingBox(null);
                    XYZ scopeCenter = (scopeBoundingBox.Max + scopeBoundingBox.Min) / 2;

                    double width = Math.Abs(scopeBoundingBox.Max.X - scopeBoundingBox.Min.X) - 2 * gap;
                    double height = Math.Abs(scopeBoundingBox.Max.Y - scopeBoundingBox.Min.Y) - 2 * gap;

                    MatchLineSymbol.Activate();
                    FamilyInstance matchLine = doc.Create.NewFamilyInstance(scopeCenter, MatchLineSymbol, activeView);
                    XYZ offset = new XYZ(-width * 0.5, -height * 0.5, 0);
                    matchLine.Location.Move(offset);
                    #region Parameters 
                    matchLine.LookupParameter("A").Set(width);
                    matchLine.LookupParameter("B").Set(height);

                    matchLine.LookupParameter("A_LONG").Set(gap);
                    matchLine.LookupParameter("A_SHORT").Set(gap);

                    matchLine.LookupParameter("B_LONG").Set(gap);
                    matchLine.LookupParameter("B_SHORT").Set(gap);


                    // visibility
                    matchLine.LookupParameter("1").Set(1);
                    matchLine.LookupParameter("2").Set(1);
                    matchLine.LookupParameter("3").Set(1);
                    matchLine.LookupParameter("4").Set(1);

                    ViewPlan adjacentView = childViews.Where(x => x.get_Parameter(BuiltInParameter.VIEWER_VOLUME_OF_INTEREST_CROP).AsElementId() == element.Id).FirstOrDefault();


                    matchLine.LookupParameter("View Name").Set(adjacentView.Name);
                    matchLine.LookupParameter("Sheet Number").Set(adjacentView.Name);

                    #endregion




                }

                trns.Commit();

            }

        }


        private void createViewRefrences(UIDocument uiDoc, Document doc, List<ScopeBoxInfo> scopeBoxInfoList, double gap)
        {


            Element baseViewRef = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ReferenceViewer).ToList().FirstOrDefault();
            XYZ basePoint = (baseViewRef.get_BoundingBox(doc.ActiveView).Max + baseViewRef.get_BoundingBox(doc.ActiveView).Min) / 2;


            // get ViewRefrenceSymbols
            FamilySymbol viewRefLeft = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).OfCategory(BuiltInCategory.OST_ReferenceViewerSymbol).Cast<FamilySymbol>().ToList().FirstOrDefault(x => x.Name == "Vertical LEFT");
            FamilySymbol viewRefRight = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).OfCategory(BuiltInCategory.OST_ReferenceViewerSymbol).Cast<FamilySymbol>().ToList().FirstOrDefault(x => x.Name == "Vertical RIGHT");
            FamilySymbol viewRefTop = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).OfCategory(BuiltInCategory.OST_ReferenceViewerSymbol).Cast<FamilySymbol>().ToList().FirstOrDefault(x => x.Name == "Horizontal UP");
            FamilySymbol viewRefBottom = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).OfCategory(BuiltInCategory.OST_ReferenceViewerSymbol).Cast<FamilySymbol>().ToList().FirstOrDefault(x => x.Name == "Horizontal DOWN");


            #region get dependent veiws 

            // get active View
            Autodesk.Revit.DB.View activeView = doc.ActiveView;

            List<ElementId> dependentviewsIDS = activeView.GetDependentViewIds().ToList();
            List<ViewPlan> dependentviews = new List<ViewPlan>();
            dependentviewsIDS.ForEach(x => dependentviews.Add(doc.GetElement(x) as ViewPlan));


            List<ViewPlan> childViews = new List<ViewPlan>();
            foreach (var viewId in activeView.GetDependentViewIds())
            {
                childViews.Add(doc.GetElement(viewId) as ViewPlan);
            }

            #endregion

            using (Transaction trns = new Transaction(doc, "Create ViewRefrences"))
            {
                trns.Start();

                foreach (var item in scopeBoxInfoList)
                {

                    if (item.TopAdjacent != null)
                    {
                        viewRefTop.Activate();
                        XYZ topMiddlewithGap = new XYZ(item.TopMiddle.X, item.TopMiddle.Y - gap, item.TopMiddle.Z);
                        //FamilyInstance viewRef = doc.Create.NewFamilyInstance(topMiddlewithGap, viewRefTop, activeView);

                        ViewPlan adjacentView = childViews.Where(x => x.get_Parameter(BuiltInParameter.VIEWER_VOLUME_OF_INTEREST_CROP).AsElementId() == item.ScopeBox.Id).FirstOrDefault();
                        ElementId copiedElementId = ElementTransformUtils.CopyElement(doc, baseViewRef.Id, topMiddlewithGap - basePoint).ToList().FirstOrDefault() ;
                         doc.GetElement(copiedElementId).get_Parameter(BuiltInParameter.REFERENCE_VIEWER_UI_TARGET_VIEW).Set(adjacentView.Id);
                        // doc.GetElement(copiedElementId).LookupParameter("View Reference Tag").Set(viewRefTop.Id);

                    }

                    if (item.BottomAdjacent != null)
                    {
                        viewRefBottom.Activate();
                        XYZ bottomMiddlewithGap = new XYZ(item.BottomMiddle.X, item.BottomMiddle.Y + gap, item.BottomMiddle.Z);
                        //FamilyInstance viewRef = doc.Create.NewFamilyInstance(bottomMiddlewithGap, viewRefBottom, activeView);

                        ViewPlan adjacentView = childViews.Where(x => x.get_Parameter(BuiltInParameter.VIEWER_VOLUME_OF_INTEREST_CROP).AsElementId() == item.ScopeBox.Id).FirstOrDefault();
                        ElementId copiedElementId = ElementTransformUtils.CopyElement(doc, baseViewRef.Id, bottomMiddlewithGap - basePoint).ToList().FirstOrDefault();
                        doc.GetElement(copiedElementId).get_Parameter(BuiltInParameter.REFERENCE_VIEWER_UI_TARGET_VIEW).Set(adjacentView.Id);
                        // doc.GetElement(copiedElementId).LookupParameter("View Reference Tag").Set(viewRefBottom.Id);

                    }

                    if (item.LeftAdjacent != null)
                    {
                        viewRefLeft.Activate();
                        XYZ leftMiddlewithGap = new XYZ(item.LeftMiddle.X + gap, item.LeftMiddle.Y, item.LeftMiddle.Z);
                        // FamilyInstance viewRef = doc.Create.NewFamilyInstance(leftMiddlewithGap, viewRefLeft, activeView);

                        ViewPlan adjacentView = childViews.Where(x => x.get_Parameter(BuiltInParameter.VIEWER_VOLUME_OF_INTEREST_CROP).AsElementId() == item.ScopeBox.Id).FirstOrDefault();
                        ElementId copiedElementId = ElementTransformUtils.CopyElement(doc, baseViewRef.Id, leftMiddlewithGap - basePoint).ToList().FirstOrDefault();
                        doc.GetElement(copiedElementId).get_Parameter(BuiltInParameter.REFERENCE_VIEWER_UI_TARGET_VIEW).Set(adjacentView.Id);
                       // doc.GetElement(copiedElementId).LookupParameter("View Reference Tag").Set(viewRefLeft.Id);

                    }

                    if (item.RightAdjacent != null)
                    {
                        viewRefRight.Activate();
                        XYZ rightMiddlewithGap = new XYZ(item.RightMiddle.X - gap, item.RightMiddle.Y, item.RightMiddle.Z);
                        // FamilyInstance viewRef = doc.Create.NewFamilyInstance(rightMiddlewithGap, viewRefRight, activeView);

                        ViewPlan adjacentView = childViews.Where(x => x.get_Parameter(BuiltInParameter.VIEWER_VOLUME_OF_INTEREST_CROP).AsElementId() == item.ScopeBox.Id).FirstOrDefault();
                        ElementId copiedElementId = ElementTransformUtils.CopyElement(doc, baseViewRef.Id, rightMiddlewithGap - basePoint).ToList().FirstOrDefault();
                        doc.GetElement(copiedElementId).get_Parameter(BuiltInParameter.REFERENCE_VIEWER_UI_TARGET_VIEW).Set(adjacentView.Id);
                       // doc.GetElement(copiedElementId).LookupParameter("View Reference Tag").Set(viewRefRight.Id);

                    }



                }


                trns.Commit();

            }

        }

        private void createViewRefrences2(UIDocument uiDoc, Document doc, List<ScopeBoxInfo> scopeBoxInfoList, double gap)
        {



            FamilySymbol viewRefLeft = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).OfCategory(BuiltInCategory.OST_GenericAnnotation).Cast<FamilySymbol>().ToList().FirstOrDefault(x => x.Name == "Left");
            FamilySymbol viewRefRight = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).OfCategory(BuiltInCategory.OST_GenericAnnotation).Cast<FamilySymbol>().ToList().FirstOrDefault(x => x.Name == "Right");
            FamilySymbol viewRefTop = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).OfCategory(BuiltInCategory.OST_GenericAnnotation).Cast<FamilySymbol>().ToList().FirstOrDefault(x => x.Name == "Top");
            FamilySymbol viewRefBottom = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).OfCategory(BuiltInCategory.OST_GenericAnnotation).Cast<FamilySymbol>().ToList().FirstOrDefault(x => x.Name == "Bottom");






            #region get dependent veiws 

            // get active View
            Autodesk.Revit.DB.View activeView = doc.ActiveView;

            List<ElementId> dependentviewsIDS = activeView.GetDependentViewIds().ToList();
            List<ViewPlan> dependentviews = new List<ViewPlan>();
            dependentviewsIDS.ForEach(x => dependentviews.Add(doc.GetElement(x) as ViewPlan));


            List<ViewPlan> childViews = new List<ViewPlan>();
            foreach (var viewId in activeView.GetDependentViewIds())
            {
                childViews.Add(doc.GetElement(viewId) as ViewPlan);
            }

            #endregion

            using (Transaction trns = new Transaction(doc, "Create ViewRefrences"))
            {
                trns.Start();

                foreach (var item in scopeBoxInfoList)
                {

                    if (item.TopAdjacent != null)
                    {
                        viewRefTop.Activate();
                        XYZ topMiddlewithGap = new XYZ(item.TopMiddle.X, item.TopMiddle.Y , item.TopMiddle.Z);
                        FamilyInstance viewRef = doc.Create.NewFamilyInstance(topMiddlewithGap, viewRefTop, activeView);
                        ViewPlan adjacentView = childViews.Where(x => x.get_Parameter(BuiltInParameter.VIEWER_VOLUME_OF_INTEREST_CROP).AsElementId() == item.TopAdjacent.ScopeBox.Id).FirstOrDefault();
                        viewRef.LookupParameter("AdjacentViewName").Set(adjacentView.Name);



                    }

                    if (item.BottomAdjacent != null)
                    {
                        viewRefBottom.Activate();
                        XYZ bottomMiddlewithGap = new XYZ(item.BottomMiddle.X, item.BottomMiddle.Y , item.BottomMiddle.Z);
                        FamilyInstance viewRef = doc.Create.NewFamilyInstance(bottomMiddlewithGap, viewRefBottom, activeView);
                        ViewPlan adjacentView = childViews.Where(x => x.get_Parameter(BuiltInParameter.VIEWER_VOLUME_OF_INTEREST_CROP).AsElementId() == item.BottomAdjacent.ScopeBox.Id).FirstOrDefault();
                        viewRef.LookupParameter("AdjacentViewName").Set(adjacentView.Name);


                    }

                    if (item.LeftAdjacent != null)
                    {
                        viewRefLeft.Activate();
                        XYZ leftMiddlewithGap = new XYZ(item.LeftMiddle.X , item.LeftMiddle.Y, item.LeftMiddle.Z);
                        FamilyInstance viewRef = doc.Create.NewFamilyInstance(leftMiddlewithGap, viewRefLeft, activeView);
                        ViewPlan adjacentView = childViews.Where(x => x.get_Parameter(BuiltInParameter.VIEWER_VOLUME_OF_INTEREST_CROP).AsElementId() == item.LeftAdjacent.ScopeBox.Id).FirstOrDefault();
                        viewRef.LookupParameter("AdjacentViewName").Set(adjacentView.Name);
                    }

                    if (item.RightAdjacent != null)
                    {
                        viewRefRight.Activate();
                        XYZ rightMiddlewithGap = new XYZ(item.RightMiddle.X , item.RightMiddle.Y, item.RightMiddle.Z);
                        FamilyInstance viewRef = doc.Create.NewFamilyInstance(rightMiddlewithGap, viewRefRight, activeView);
                        ViewPlan adjacentView = childViews.Where(x => x.get_Parameter(BuiltInParameter.VIEWER_VOLUME_OF_INTEREST_CROP).AsElementId() == item.RightAdjacent.ScopeBox.Id).FirstOrDefault();
                        viewRef.LookupParameter("AdjacentViewName").Set(adjacentView.Name);

                    }



                }


                trns.Commit();

            }

        }


        private void getScopeBoxWithAdjacent(List<ScopeBoxInfo> scopeBoxInfoList, Document doc, double gap)
        {
            // get scopeBoxes
            List<Element> scopeBoxes = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_VolumeOfInterest)
                .Cast<Element>()
                .ToList();

            foreach (Element element in scopeBoxes)
            {
                BoundingBoxXYZ scopeBoundingBox = element.get_BoundingBox(null);
                XYZ scopeCenter = (scopeBoundingBox.Max + scopeBoundingBox.Min) / 2;

                // Create an instance of ScopeBoxInfo for the current scope box
                ScopeBoxInfo scopeBoxInfo = new ScopeBoxInfo(element);

                List<Element> otherScopeBoxes = scopeBoxes.Where(x => x.Id != element.Id).ToList();

                // assign the adjacents
                AssignAdjacentproperties(scopeBoxInfo, otherScopeBoxes, gap);

                // Add the scopeBoxInfo to the list
                scopeBoxInfoList.Add(scopeBoxInfo);
            }
        }



        private void AssignAdjacentproperties(ScopeBoxInfo originalScope, List<Element> otherScopeBoxes, double gap)
        {


            List<ScopeBoxInfo> otherScopeBoxesModified = new List<ScopeBoxInfo>();
            otherScopeBoxes.ForEach(o => { otherScopeBoxesModified.Add(new ScopeBoxInfo(o)); });

            List<ScopeBoxInfo> rightAdjacents = new List<ScopeBoxInfo>();
            List<ScopeBoxInfo> leftAdjacents = new List<ScopeBoxInfo>();
            List<ScopeBoxInfo> TopAdjacents = new List<ScopeBoxInfo>();
            List<ScopeBoxInfo> bottomAdjacents = new List<ScopeBoxInfo>();



            foreach (var item in otherScopeBoxesModified)
            {
                XYZ diff = item.Center - originalScope.Center;
                Vector3 diffVec = ToVector(diff);
                diffVec.X = Math.Abs(diffVec.X);
                diffVec.Y = Math.Abs(diffVec.Y);
                diffVec.Z = Math.Abs(diffVec.Z);


                if (Vector3.Normalize(diffVec).X > 0.5)
                {
                    if (diff.X > 0)
                    {
                        rightAdjacents.Add(item);
                    }
                    else
                    {
                        leftAdjacents.Add(item);

                    }
                }

                if (Vector3.Normalize(diffVec).Y > 0.5)
                {

                    if (diff.Y > 0)
                    {
                        TopAdjacents.Add(item);

                    }
                    else
                    {
                        bottomAdjacents.Add(item);

                    }
                }


            }

            ScopeBoxInfo rightAdjacent = rightAdjacents.OrderBy(x => x.LeftMiddle.DistanceTo(originalScope.RightMiddle)).FirstOrDefault();
            ScopeBoxInfo leftAdjacent = leftAdjacents.OrderBy(x => x.RightMiddle.DistanceTo(originalScope.LeftMiddle)).FirstOrDefault();
            ScopeBoxInfo topAdjacent = TopAdjacents.OrderBy(x => x.BottomMiddle.DistanceTo(originalScope.TopMiddle)).FirstOrDefault();
            ScopeBoxInfo bottomAdjacent = bottomAdjacents.OrderBy(x => x.TopMiddle.DistanceTo(originalScope.BottomMiddle)).FirstOrDefault();


            if (rightAdjacent != null && rightAdjacent.LeftMiddle.DistanceTo(originalScope.RightMiddle) < 10000 * gap)
            {
                originalScope.RightAdjacent = rightAdjacent;
            }


            if (leftAdjacent != null && leftAdjacent.RightMiddle.DistanceTo(originalScope.LeftMiddle) < 10000 * gap)
            {
                originalScope.LeftAdjacent = leftAdjacent;
            }



            if (topAdjacent != null && topAdjacent.BottomMiddle.DistanceTo(originalScope.TopMiddle) < 10000 * gap)
            {
                originalScope.TopAdjacent = topAdjacent;
            }



            if (bottomAdjacent != null && bottomAdjacent.TopMiddle.DistanceTo(originalScope.BottomMiddle) < 10000 * gap)
            {
                originalScope.BottomAdjacent = bottomAdjacent;
            }

        }


        private RelativePosition GetRelativePosition(BoundingBoxXYZ box1, BoundingBoxXYZ box2)
        {
            double tolerance = UnitUtils.ConvertToInternalUnits(10, UnitTypeId.Millimeters);

            if (Math.Abs(box1.Max.X - box2.Min.X) < tolerance && Math.Abs(box1.Max.Y - box2.Min.Y) < tolerance)
            {
                // box2 is on the right of box1
                return RelativePosition.Right;
            }
            else if (Math.Abs(box1.Min.X - box2.Max.X) < tolerance && Math.Abs(box1.Max.Y - box2.Min.Y) < tolerance)
            {
                // box2 is on the left of box1
                return RelativePosition.Left;
            }
            else if (Math.Abs(box1.Max.X - box2.Min.X) < tolerance && Math.Abs(box1.Min.Y - box2.Max.Y) < tolerance)
            {
                // box2 is below box1
                return RelativePosition.Bottom;
            }
            else if (Math.Abs(box1.Max.X - box2.Min.X) < tolerance && Math.Abs(box1.Min.Y - box2.Max.Y) < tolerance)
            {
                // box2 is above box1
                return RelativePosition.Top;
            }

            return RelativePosition.None; // No relative position
        }

        public class ScopeBoxInfo
        {
            public Element ScopeBox { get; set; }
            //  public List<AdjacentScopeBoxInfo> AdjacentScopeBoxes { get; set; }
            public XYZ TopLeftCorner { get; private set; }
            public XYZ TopRightCorner { get; private set; }
            public XYZ BottomLeftCorner { get; private set; }
            public XYZ BottomRightCorner { get; private set; }
            public XYZ TopMiddle { get; private set; }
            public XYZ BottomMiddle { get; private set; }
            public XYZ RightMiddle { get; private set; }
            public XYZ LeftMiddle { get; private set; }
            public XYZ Center { get; private set; }

            public ScopeBoxInfo RightAdjacent { get; set; }
            public ScopeBoxInfo LeftAdjacent { get; set; }
            public ScopeBoxInfo TopAdjacent { get; set; }
            public ScopeBoxInfo BottomAdjacent { get; set; }



            public ScopeBoxInfo(Element scopeBox)
            {
                ScopeBox = scopeBox;

                BoundingBoxXYZ scopeBoundingBox = scopeBox.get_BoundingBox(null);
                TopLeftCorner = new XYZ(scopeBoundingBox.Min.X, scopeBoundingBox.Max.Y, scopeBoundingBox.Min.Z);
                TopRightCorner = new XYZ(scopeBoundingBox.Max.X, scopeBoundingBox.Max.Y, scopeBoundingBox.Min.Z);
                BottomLeftCorner = new XYZ(scopeBoundingBox.Min.X, scopeBoundingBox.Min.Y, scopeBoundingBox.Min.Z);
                BottomRightCorner = new XYZ(scopeBoundingBox.Max.X, scopeBoundingBox.Min.Y, scopeBoundingBox.Min.Z);

                TopMiddle = new XYZ((TopLeftCorner.X + TopRightCorner.X) / 2, (TopLeftCorner.Y + TopRightCorner.Y) / 2, scopeBoundingBox.Min.Z);
                BottomMiddle = new XYZ((BottomLeftCorner.X + BottomRightCorner.X) / 2, (BottomLeftCorner.Y + BottomRightCorner.Y) / 2, scopeBoundingBox.Min.Z);
                RightMiddle = new XYZ((TopRightCorner.X + BottomRightCorner.X) / 2, (TopRightCorner.Y + BottomRightCorner.Y) / 2, scopeBoundingBox.Min.Z);
                LeftMiddle = new XYZ((TopLeftCorner.X + BottomLeftCorner.X) / 2, (TopLeftCorner.Y + BottomLeftCorner.Y) / 2, scopeBoundingBox.Min.Z);

                Center = (scopeBoundingBox.Max + scopeBoundingBox.Min) / 2;
            }
        }

        public class AdjacentScopeBoxInfo
        {
            public Element ScopeBox { get; set; }
            public RelativePosition Position { get; set; }

            public AdjacentScopeBoxInfo(Element scopeBox, RelativePosition position)
            {
                ScopeBox = scopeBox;
                Position = position;
            }
        }

        public enum RelativePosition
        {
            Left,
            Right,
            Top,
            Bottom,
            None
        }

        public Vector3 ToVector(XYZ pt)
        {
            return new Vector3((float)pt.X, (float)pt.Y, (float)pt.Z);
        }

    }
}