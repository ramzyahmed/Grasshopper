using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using ETABSv1;
using System.Collections.Generic;
using System.Linq;
using System.Numerics;

using Excel = Microsoft.Office.Interop.Excel;


namespace BIM_Integration
{
    [Transaction(TransactionMode.Manual)]
    internal class M_ConvertFromSegementsToLinesEtabs : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {

            try
            {

                #region get all Sap Data from Excel 
                List<Column> sapColumns = new List<Column>();
                List<Beam> sapBeams = new List<Beam>();
                List<Frame> sapFrames = new List<Frame>();

                // get all beams, columns and all members
                sapFrames = SapUtils.getSapFrames(ref sapBeams, ref sapColumns);

                ChangeSapSectionNames.GetSectionsDimentions2(SapUtils.Get_sap_mode(), ref sapFrames);

                #endregion

                SapUtils.isSameVector(sapFrames[0], sapFrames[1]);

                SAP2000v1.cSapModel mySapModel = SapUtils.Get_sap_mode();
                List<List<Frame>> sapFrames2 = SapUtils.GroupFrames(sapFrames);




                foreach (var item in sapFrames2)
                {
                    foreach (var item2 in item)
                    {
                        mySapModel.FrameObj.SetSelected(item2.ID, true);
                    }

                    int aaa = 11;
                }

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