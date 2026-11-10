using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Numerics;

namespace BIM_Integration
{
    [Transaction(TransactionMode.Manual)]
    internal class DivideCurvedBEams: IExternalCommand
    {
        public string ModelType = "";
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {

            try
            {

                #region get Revit Data
                UIDocument uidoc = commandData.Application.ActiveUIDocument;
                Document doc = uidoc.Document;

                // for revit 2021 or above use (ForgeTypeId dis = UnitTypeId.Millimeters;) however for revit 2020 or lower use ( ForgeTypeId dis = DisplayUnitType.DUT_MILLIMETERS;)
                ForgeTypeId dis = UnitTypeId.Millimeters;



                #region get get Selected Beams
                List<ElementId> selectedElements = uidoc.Selection.GetElementIds().ToList();

                List<FamilyInstance> selectedCurvedBeams = new List<FamilyInstance>();

                foreach (ElementId elementId in selectedElements)
                {
                    Element element = doc.GetElement(elementId);

                    // Check if the element belongs to the Structural Columns or Structural Framing category
                    if ((BuiltInCategory)element.Category.Id.IntegerValue == BuiltInCategory.OST_StructuralFraming)
                    {
                        selectedCurvedBeams.Add(element as FamilyInstance);
                    }
                }


                #endregion
                #endregion

                #region handling curvedBeams (not implemented yet)
                List<FamilyInstance> strightBeams = new List<FamilyInstance>();
                List<FamilyInstance> curvedBeams = new List<FamilyInstance>();
                List<FamilyInstance> StrightElementsSplittedFromCurved = new List<FamilyInstance>();


                foreach (var instance in selectedCurvedBeams)
                {


                    SapRevitFrame.createView(curvedBeams.Select(x => x.Id).ToList(), doc, uidoc, "curvedBeams");

                 

                    StrightElementsSplittedFromCurved.Add(instance);


                 
                   LocationCurve locCurve = instance.Location as LocationCurve;
                    Curve curve = locCurve.Curve;
                    double totalLength = curve.Length;

                    double segmentLength = UnitUtils.ConvertToInternalUnits(3000.0, dis);

                    using (Transaction trns = new Transaction(doc, "aa"))

                    {
                        trns.Start();

                        while ((instance.Location as LocationCurve).Curve.Length > segmentLength)
                        {
                            totalLength = (instance.Location as LocationCurve).Curve.Length;
                            StrightElementsSplittedFromCurved.Add(doc.GetElement(instance.Split(1 - segmentLength / totalLength)) as FamilyInstance);

                        }

                    aa:
                        trns.Commit();

                    }

                }




                // delete Beams
                using (Transaction trns = new Transaction(doc, "aa"))
                {
                    trns.Start();
                    selectedCurvedBeams.ForEach(x => doc.Delete(x.Id));
                    trns.Commit();
                }
                #endregion

                return Result.Succeeded;

            }


            catch (Exception ex2)
            {

                System.Diagnostics.StackTrace trace = new System.Diagnostics.StackTrace(ex2, true);
                TaskDialog.Show("Error", "error in beam no" + ex2.Message + trace.GetFrame(trace.FrameCount - 1).GetFileLineNumber());

                return Result.Failed;
            }


        }
    }

}