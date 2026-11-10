using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Numerics;


[Transaction(TransactionMode.Manual)]
public class ExtendTrimGridsCommand : IExternalCommand
{
    public Result Execute(
      ExternalCommandData commandData,
      ref string message,
      ElementSet elements)
    {
        // Get the active document
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;

        // Ask the user to select grids
        List<Grid> selectedGrids = SelectGrids(uidoc);

        if (selectedGrids.Count == 0)
        {
            TaskDialog.Show("Error", "No grids selected. Operation canceled.");
            return Result.Cancelled;
        }

        // Ask the user to select a reference curve
        Element referenceCurve = SelectReferenceCurve(uidoc);

        if (referenceCurve == null)
        {
            TaskDialog.Show("Error", "No reference curve selected. Operation canceled.");
            return Result.Cancelled;
        }

        // Ask the user for extension value and swap option
        double extensionValue = 1000;
        bool swap = false;

        if (!GetUserInput(out extensionValue, out swap))
        {
            TaskDialog.Show("Error", "Invalid input. Operation canceled.");
            return Result.Cancelled;
        }

        // Extend/trim selected grids based on the reference curve

        LocationCurve crv =  referenceCurve.Location as LocationCurve ;
        ModifyGrids(selectedGrids, crv, extensionValue, swap, doc);

        return Result.Succeeded;
    }

    private List<Grid> SelectGrids(UIDocument uidoc)
    {
        // Use PickObjects to allow the user to select multiple grids
        ElementId pickedRef = uidoc.Selection.PickObject(ObjectType.Element, "Select grids")?.ElementId;

        // Retrieve the selected grids
        Element selectedElement = uidoc.Document.GetElement(pickedRef);
        if (selectedElement is Grid)
        {
            return new List<Grid> { selectedElement as Grid };
        }
        else if (selectedElement is Group)
        {
            // If a group is selected, extract grids from the group
            Group group = selectedElement as Group;
            return new FilteredElementCollector(uidoc.Document, group.GetMemberIds())
                .OfClass(typeof(Grid))
                .Cast<Grid>()
                .ToList();
        }
        else
        {
            return new List<Grid>();
        }
    }

    private Element SelectReferenceCurve(UIDocument uidoc)
    {
        // Use PickObject to allow the user to select a reference curve
        Reference pickedRef = uidoc.Selection.PickObject(ObjectType.Element, "Select a reference curve");

        return pickedRef != null ? uidoc.Document.GetElement(pickedRef) as Element : null;
    }

    private bool GetUserInput(out double extensionValue, out bool swap)
    {
        // TODO: Implement logic to get user input for extension value and swap option
        // For simplicity, we'll use hardcoded values here
        extensionValue = 10.0;
        swap = true;

        return true;
    }

    private void ModifyGrids(List<Grid> grids, LocationCurve referenceCurve, double extensionValue, bool swap, Document doc)
    {
        using (Transaction transaction = new Transaction(doc, "Extend/Trim Grids"))
        {
            transaction.Start();



            foreach (Grid grid in grids)
            {
                // Get the curve of the grid
                Curve GridCurve = grid.Curve;

                // Determine which end of the grid to change and which end to maintain
                XYZ toChangePt, toMaintainPt;
                bool inverted = false;





               double distFromGridStart =  DistancePointToLine(ToVector(GridCurve.GetEndPoint(0)), ToVector(referenceCurve.Curve.GetEndPoint(0)), ToVector(referenceCurve.Curve.GetEndPoint(1)));
               double distFromGridEnd = DistancePointToLine(ToVector(GridCurve.GetEndPoint(1)), ToVector(referenceCurve.Curve.GetEndPoint(0)), ToVector(referenceCurve.Curve.GetEndPoint(1)));




                if (distFromGridStart < distFromGridEnd)
                {
                    toChangePt = GridCurve.GetEndPoint(0);
                    toMaintainPt = GridCurve.GetEndPoint(1);

                    if (swap)
                    {

                    }
                }
                else
                {
                    toChangePt = GridCurve.GetEndPoint(1);
                    toMaintainPt = GridCurve.GetEndPoint(0);
                    inverted = true;

                    if (swap)
                    {

                    }
                }

                // Create a line parallel to the to_change end with a specified extension value
                Line parralelLine = Line.CreateBound(toChangePt, toChangePt + (toChangePt - toMaintainPt) * extensionValue);



                // Intersect the created line with the reference curve


                var intersects = GridCurve.Intersect(parralelLine, out var intersections);

                if (intersects != SetComparisonResult.Disjoint)
                {
                    var intersection = intersections.Cast<IntersectionResult>().First();
                    var point = intersection.XYZPoint;


                    // Create a new grid line based on the determined ends and intersection point

                    Line newGridLine;
                    if (inverted)
                    {
                         newGridLine = Line.CreateBound(toMaintainPt, point);
                    }
                    else
                    {
                        newGridLine = Line.CreateBound(point, toMaintainPt);
                    }


                   string oldGridname = grid.Name+"3";

                    doc.Delete(grid.Id);
                  
                  Grid newGrid =   Grid.Create(doc,newGridLine);
                  newGrid.Name = oldGridname;


                    // Set the modified curve to the grid
                   // (grid.Location as LocationCurve).Curve = GridCurve;
                }
            }

            // Commit the transaction
            transaction.Commit();
        }
    }




    public XYZ ToXYz ( Vector3 vec )
    {
        return new XYZ (vec.X, vec.Y, vec.Z);   
    }


    public Vector3 ToVector(XYZ pt)
    {
        return new Vector3((float)pt.X, (float) pt.Y,(float) pt.Z);
    }


    public Vector3 curveToVector (LocationCurve crv)
    {
        Vector3 stPt = new Vector3((float)crv.Curve.GetEndPoint(0).X, (float)crv.Curve.GetEndPoint(0).Y, (float)crv.Curve.GetEndPoint(0).Z);


        Vector3 endPt = new Vector3((float)crv.Curve.GetEndPoint(1).X, (float)crv.Curve.GetEndPoint(1).Y, (float)crv.Curve.GetEndPoint(1).Z);



        Vector3 result = endPt - stPt;
        return result;  
    }



  public  static float DistancePointToLine(Vector3 point, Vector3 lineStart, Vector3 lineEnd)
    {
        // Calculate the vectors representing the line segments
        Vector3 line = lineEnd - lineStart;
        Vector3 pointToLineStart = point - lineStart;

        // Calculate the perpendicular distance using the cross product
        float distance = Vector3.Cross(line, pointToLineStart).Length() / line.Length();

        return distance;
    }


    static Vector3 CalculateIntersection(Vector3 line1Start, Vector3 line1End, Vector3 line2Start, Vector3 line2End)
    {
        // Define vectors for the two lines
        Vector3 line1 = line1End - line1Start;
        Vector3 line2 = line2End - line2Start;

        // Calculate the intersection point using vector cross and dot products
        Vector3 cross1 = Vector3.Cross(line1Start - line2Start, line2);
        Vector3 cross2 = Vector3.Cross(line1, line2);

        float t = Vector3.Dot(cross1, cross2) / cross2.LengthSquared();

        Vector3 intersection = line1Start + t * line1;

        return intersection;
    }
}



