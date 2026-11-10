using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;


namespace BIM_Integration
{
    [Transaction(TransactionMode.Manual)]
    public class DownGradeBeams : IExternalCommand
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
                #endregion

                List<RevitBeam> beams = new List<RevitBeam>();  
                List<RevitLevel> levels = new List<RevitLevel>();

                RevitBeam.getRevitBeams2(doc, uidoc,ref beams, dis);
                RevitLevel.getRevitLevels(doc, uidoc,ref levels, dis);



                string BeamPath = @"C:\Users\Ahmed.Ramzy\Desktop\New Text Document.txt";
                string LevelPath = @"C:\Users\Ahmed.Ramzy\Desktop\New Text Document2.txt";

               RevitBeam.SerializeRevitBeamsToJsonWithSkip(beams, BeamPath);
               RevitBeam.SerializeLevelsToJsonWithSkip(levels,LevelPath);


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