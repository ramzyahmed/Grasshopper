using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using BIM_Integration;
using System;
using System.IO;
using SAP2000v1;
using System.Linq;


namespace BIM_Integration
{
    [Transaction(TransactionMode.Manual)]
    public class BtnUpdateSections : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp = new UIApplication(commandData.Application.Application);
            Document doc = commandData.Application.ActiveUIDocument.Document;
            BtnUpdateSections.commandData1 = commandData;

            uiApp.Idling += idleUpdate;

            return Result.Succeeded;
        }

        public static void idleUpdate(object sender, IdlingEventArgs e)
        {

            try
            {
                string message = "";
                ElementSet elements = new ElementSet();

                string firstLine;
                string userName = Environment.UserName;
                string fullPath = @"C:\Users\" + userName + @"\Desktop\H Y D R A _temp\ExcDataWithRvt\HYDRA Integration.txt";


                using (StreamReader reader = new StreamReader(fullPath))
                {
                    firstLine = reader.ReadLine();
                }

                if (firstLine == "RvtToSap")
                {
                    File.WriteAllText(fullPath, string.Empty);
                    M_RvtToSapUpdateSections instance = new M_RvtToSapUpdateSections();
                    instance.Execute(BtnUpdateSections.commandData1, ref message, elements);
                    return;

                }

                else if (firstLine == "SapToRvt")
                {
                    File.WriteAllText(fullPath, string.Empty);
                    M_SapToRvtUpdateSections instance = new M_SapToRvtUpdateSections();
                    instance.Execute(BtnUpdateSections.commandData1, ref message, elements);
                    return;
                }



                else if (firstLine == "StdToRvt")
                {
                    File.WriteAllText(fullPath, string.Empty);
                    M_StaadToRvtUpdateSections instance = new M_StaadToRvtUpdateSections();
                    instance.Execute(BtnUpdateSections.commandData1, ref message, elements);
                    return;

                }

                else if (firstLine == "RvtToStd")
                {
                    File.WriteAllText(fullPath, string.Empty);
                    M_RvtToStaadUpdateSections instance = new M_RvtToStaadUpdateSections();
                    instance.Execute(BtnUpdateSections.commandData1, ref message, elements);
                    return;
                }


               else if (firstLine == "RvtToEtabs")
                {
                    File.WriteAllText(fullPath, string.Empty);
                    M_RvtToEtabsUpdateSections instance = new M_RvtToEtabsUpdateSections();
                    instance.Execute(BtnUpdateSections.commandData1, ref message, elements);
                    return;

                }

                else if (firstLine == "EtabsToRvt")
                {
                    File.WriteAllText(fullPath, string.Empty);
                    M_EtabsToRvtUpdateSections instance = new M_EtabsToRvtUpdateSections();
                    instance.Execute(BtnUpdateSections.commandData1, ref message, elements);
                    return;
                }

            }

            catch (Exception ex2)
            {

                BtnUpdateSections.commandData1.Application.Idling -= idleUpdate;
                System.Diagnostics.StackTrace trace = new System.Diagnostics.StackTrace(ex2, true);
                TaskDialog.Show("Error", ex2.Message + trace.GetFrame(trace.FrameCount - 1).GetFileLineNumber());
                return;
            }

        }

        public static ExternalCommandData commandData1 = null;

        ~BtnUpdateSections()
        {


        }

    }

}