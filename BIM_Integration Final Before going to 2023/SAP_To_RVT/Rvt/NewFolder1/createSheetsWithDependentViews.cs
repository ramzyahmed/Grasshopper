using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

namespace BIM_Integration
{
    [Transaction(TransactionMode.Manual)]
    internal class createSheetsWithDependentViews : IExternalCommand
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
                ForgeTypeId dis = UnitTypeId.Feet;
                #endregion


                FamilySymbol titleBlock = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_TitleBlocks).Cast<FamilySymbol>().FirstOrDefault();

                Autodesk.Revit.DB.Outline titleBlockOutline = RevitUtils.TitleBlockDrawingAreaBoundary(titleBlock);


                XYZ upperLeftPoint = new XYZ(titleBlockOutline.MinimumPoint.X, titleBlockOutline.MaximumPoint.Y, 0);
                XYZ lowerLeftPoint = titleBlockOutline.MinimumPoint;
                XYZ upperRightPoint = titleBlockOutline.MaximumPoint;




                using (Transaction trns = new Transaction(doc, "Distribute Views"))
                {

                    Autodesk.Revit.DB.View parentView = doc.ActiveView;
                    List<Element> scopeBoxes = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_VolumeOfInterest).Cast<Element>().ToList();
                    trns.Start();
                    int i = 1;
                    foreach (Element element in scopeBoxes)
                    {

                        //LocationPoint scopeBoxPt=   element.Location as LocationPoint;
                        //XYZ pt = new XYZ(scopeBoxPt.Point.X, scopeBoxPt.Point.Y, 0);

                        BoundingBoxXYZ scopeBoundingBox = element.get_BoundingBox(null);
                        XYZ scopeCenter = (scopeBoundingBox.Max + scopeBoundingBox.Min) / 2;
                        scopeCenter = new XYZ(scopeCenter.X, scopeCenter.Y, 0);





                        // Duplicate the view as dependent
                        ElementId duplicateView = parentView.Duplicate(ViewDuplicateOption.AsDependent);
                        ViewPlan dependentView = doc.GetElement(duplicateView) as ViewPlan;

                        dependentView.Name = parentView.Name + "Part " + i + "SB " + element.Name;
                        dependentView.LookupParameter("Scope Box").Set(element.Id);


                        ViewSheet sht = ViewSheet.Create(doc, titleBlock.Id);
                        sht.Name = "ViewSheet" + parentView.Name + "Part " + i;



                        #region ViewCenter
                        XYZ viewCenter = new XYZ(
                    (dependentView.Outline.Min.U + dependentView.Outline.Max.U) / 2,
                    (dependentView.Outline.Min.V + dependentView.Outline.Max.V) / 2,
                    0);

                        XYZ viewCenterModified = new XYZ(UnitUtils.ConvertToInternalUnits(viewCenter.X, dis), UnitUtils.ConvertToInternalUnits(viewCenter.Y, dis), 0);





                        ViewCropRegionShapeManager vv = dependentView.GetCropRegionShapeManager();
                        List<CurveLoop> curveLoops = (List<CurveLoop>)vv.GetCropShape();

                        XYZ viewCenter2 = GetCenterOfCropShape(curveLoops);
                        #endregion


                        #region ٍSheetCenter
                        XYZ sheetCenter = new XYZ(
                                                   (titleBlockOutline.MinimumPoint.X + titleBlockOutline.MaximumPoint.X) / 2,
                                                   (titleBlockOutline.MinimumPoint.Y + titleBlockOutline.MaximumPoint.Y) / 2,
                                                   0);
                        #endregion


                        XYZ offset = sheetCenter - viewCenterModified;


                        Viewport viewport = Viewport.Create(doc, sht.Id, dependentView.Id, new XYZ());


                        //Outline viewportOutline = viewport.GetBoxOutline();


                        //// Extract the center point from the bounding box
                        //XYZ centerPoint = new XYZ(
                        //    (viewportOutline.MinimumPoint.X + viewportOutline.MaximumPoint.X) / 2,
                        //    (viewportOutline.MinimumPoint.Y + viewportOutline.MaximumPoint.Y) / 2,
                        //    (viewportOutline.MinimumPoint.Z + viewportOutline.MaximumPoint.Z) / 2);


                        //viewport.Location.Move(sheetCenter- centerPoint); 

                        doc.Regenerate();
                        i++;
                    }


                    trns.Commit();

                }

                return Result.Succeeded;
            }
            catch (Exception ex2)
            {
                System.Diagnostics.StackTrace trace = new System.Diagnostics.StackTrace(ex2, true);
                TaskDialog.Show("Error", "Error in adding views: " + ex2.Message + " Line: " + trace.GetFrame(trace.FrameCount - 1).ToString());
                return Result.Failed;
            }

        }


        public static XYZ GetCenterOfCropShape(List<CurveLoop> cropShape)
        {
            // Initialize variables for calculating centroid
            double totalArea = 0.0;
            XYZ centroid = XYZ.Zero;

            // Iterate over each curve loop in the crop shape
            foreach (CurveLoop curveLoop in cropShape)
            {
                // Iterate over each curve in the curve loop
                foreach (Curve curve in curveLoop)
                {
                    // Compute the length of the curve as an approximation of its area
                    double curveLength = curve.Length;

                    // Update the total area and centroid based on the current curve
                    totalArea += curveLength;
                    centroid += curveLength * curve.Evaluate(0.5, true);
                }
            }

            // Calculate the centroid by dividing the sum by the total area
            if (totalArea > 0.0)
            {
                centroid /= totalArea;
            }

            return centroid;
        }


    }
}