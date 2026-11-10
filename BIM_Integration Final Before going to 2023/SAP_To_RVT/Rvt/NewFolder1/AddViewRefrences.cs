using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

namespace BIM_Integration
{
    [Transaction(TransactionMode.Manual)]
    internal class AddviewRefernces : IExternalCommand
    {

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                UIDocument uidoc = commandData.Application.ActiveUIDocument;
                Document doc = uidoc.Document;
                ForgeTypeId dis = UnitTypeId.Millimeters;


                // List to store information about each scope box and its adjacent scope boxes
                List<ScopeBoxInfo> scopeBoxInfoList = new List<ScopeBoxInfo>();




                getScopeBoxWithAdjacent(scopeBoxInfoList, doc);



                return Result.Succeeded;
            }
            catch (Exception ex2)
            {
                System.Diagnostics.StackTrace trace = new System.Diagnostics.StackTrace(ex2, true);
                TaskDialog.Show("Error", "Error in adding views: " + ex2.Message + " Line: " + trace.GetFrame(trace.FrameCount - 1).ToString());
                return Result.Failed;
            }

        }



        private void getScopeBoxWithAdjacent(List<ScopeBoxInfo> scopeBoxInfoList, Document doc)
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

                // Find adjacent scope boxes
                foreach (Element otherElement in scopeBoxes)
                {
                    if (element.Id != otherElement.Id) // Skip the same scope box
                    {
                        BoundingBoxXYZ otherBoundingBox = otherElement.get_BoundingBox(null);

                        // Check if otherBoundingBox is adjacent to scopeBoundingBox
                        RelativePosition position = GetRelativePosition(scopeBoundingBox, otherBoundingBox);

                        if (position != RelativePosition.None)
                        {
                            scopeBoxInfo.AdjacentScopeBoxes.Add(new AdjacentScopeBoxInfo(otherElement, position));
                        }
                    }
                }

                // Add the scopeBoxInfo to the list
                scopeBoxInfoList.Add(scopeBoxInfo);
            }
        }

        private RelativePosition GetRelativePosition(BoundingBoxXYZ box1, BoundingBoxXYZ box2)
        {
            double tolerance = 0.1; // You may need to adjust this based on your specific case

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
            public List<AdjacentScopeBoxInfo> AdjacentScopeBoxes { get; set; }

            public ScopeBoxInfo(Element scopeBox)
            {
                ScopeBox = scopeBox;
                AdjacentScopeBoxes = new List<AdjacentScopeBoxInfo>();
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

    }
}