using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;


namespace BIM_Integration
{
    [Transaction(TransactionMode.Manual)]
    internal class SheetCreation3 : IExternalCommand
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

                using (Transaction trns = new Transaction(doc, "Distribute Views"))
                {


                    ViewSheet viewSheet = new FilteredElementCollector(doc)
                        .OfClass(typeof(ViewSheet))
                        .Cast<ViewSheet>()
                        .FirstOrDefault(x => x.Title == "Sheet: S101 - TYP");

                    FamilyInstance titleBlock = RevitUtils.getTitleBlockInViewSHeet(doc, viewSheet);
                    Outline titleBlockOutline = RevitUtils.TitleBlockDrawingAreaBoundary(titleBlock.Symbol);
                    XYZ lowerLeftPoint = titleBlockOutline.MinimumPoint;
                    XYZ upperRightPoint = titleBlockOutline.MaximumPoint;


                    trns.Start();
                    // List of views to distribute
                    List<ViewPlan> viewsToDistribute = new FilteredElementCollector(doc)
                        .OfClass(typeof(ViewPlan))
                        .Cast<ViewPlan>()
                        .Where(x => x.Name.Contains("ST-00-GF2") && x.ViewType == ViewType.FloorPlan)
                        .ToList();
                    // Fill the viewsToDistribute list with your desired views

                    double spacing = UnitUtils.ConvertToInternalUnits(10, dis);
                    double xOffset = lowerLeftPoint.X;
                    double yOffset = lowerLeftPoint.Y;

                    int viewIndex = 0;
                    int sheetCounter = 1;

                    while (viewIndex < viewsToDistribute.Count)
                    {
                        ViewPlan view = viewsToDistribute[viewIndex];

                        Viewport viewport = Viewport.Create(doc, viewSheet.Id, view.Id, lowerLeftPoint);
                        BoundingBoxXYZ boundingBox = viewport.get_BoundingBox(viewSheet);
                        double width = Math.Abs(boundingBox.Max.X - boundingBox.Min.X);
                        double height = Math.Abs(boundingBox.Max.Y - boundingBox.Min.Y);

                        if (xOffset + width > upperRightPoint.X) // If the view exceeds the sheet's width
                        {
                            xOffset = lowerLeftPoint.X; // Reset xOffset
                            yOffset += height + spacing; // Add some spacing between views
                            if (yOffset + height > upperRightPoint.Y) // If the view exceeds the sheet's height
                            {
                                // delete the created view in the full filled viewSheet
                                viewSheet.DeleteViewport(viewport);


                                // Sheet is full, duplicate and reset offsets
                                string newSheetName = viewSheet.Name + "_" + viewIndex;
                                string newSheetNumber = viewSheet.SheetNumber + "_" + viewIndex;
                                viewSheet = ViewSheet.Create(doc, titleBlock.GetTypeId());

                                viewSheet.Name = newSheetName;
                                viewSheet.SheetNumber = newSheetNumber;


                                // create the viewPoirt in the new sheet
                                viewport = Viewport.Create(doc, viewSheet.Id, view.Id, lowerLeftPoint);

                                xOffset = lowerLeftPoint.X;
                                yOffset = lowerLeftPoint.Y;
                            }
                        }

                        XYZ newLocationPoint = new XYZ(xOffset + width / 2, yOffset + height / 2, 0);
                        viewport.SetBoxCenter(newLocationPoint);

                        // Adjust the xOffset for the next view
                        xOffset += width + spacing; // Add some spacing between views
                        viewIndex++;
                    }

                    doc.Regenerate();

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

    }
}