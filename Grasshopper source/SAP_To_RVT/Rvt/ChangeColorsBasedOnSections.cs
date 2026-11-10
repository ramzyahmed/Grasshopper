using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;


namespace BIM_Integration
{
    [Transaction(TransactionMode.Manual)]
    class ChangeColorsBasedOnSections : IExternalCommand
    {
        private Random random = new Random();
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)

        {

            try
            {
                UIApplication uiApp = commandData.Application;
                UIDocument uiDoc = uiApp.ActiveUIDocument;
                Document doc = uiDoc.Document;

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

                return Result.Succeeded;
            }

            catch (Exception ex2)

            {

                System.Diagnostics.StackTrace trace = new System.Diagnostics.StackTrace(ex2, true);
                TaskDialog.Show("Error", ex2.Message + " ,line of the error is " + trace.GetFrame(trace.FrameCount - 1).GetFileLineNumber());
                return Result.Failed;
            }


        }

        private Color GetRandomColor()
        {
 
            byte red = (byte)random.Next(256);
            byte green = (byte)random.Next(256);
            byte blue = (byte)random.Next(256);
            return new Color (red, green, blue);
        }

    }
}
