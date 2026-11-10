using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;


namespace BIM_Integration
{
    [Transaction(TransactionMode.Manual)]
    public class SapToEtabs : IExternalCommand
    {
        public string ModelType = "";
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {

            try
            {

                SapUtils.DataBase.ExportSapToEtabs("");
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