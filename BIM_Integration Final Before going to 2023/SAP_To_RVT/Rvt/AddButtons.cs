using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.IO;
using System.Reflection;
using System.Windows.Media.Imaging;

namespace BIM_Integration
{
    [Transaction(TransactionMode.Manual)]
    public class AddButtons : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication application)
        {

            application.CreateRibbonTab("HYDRA");
            string path = Assembly.GetExecutingAssembly().Location;
            string ExecutingAssemblyPath = System.Reflection.Assembly.GetExecutingAssembly().Location;
            string ExecutingAssemblyName = System.Reflection.Assembly.GetExecutingAssembly().GetName().Name;



            PushButtonData automationButton = new PushButtonData("AutomationBtn", "Automation", ExecutingAssemblyPath, ExecutingAssemblyName + ".BtnAutomation");
            PushButtonData sectionsButton = new PushButtonData("Update Sections", "Update Sections", ExecutingAssemblyPath, ExecutingAssemblyName + ".BtnUpdateSections");
            PushButtonData geometrysButton = new PushButtonData("Update Geometry", "Update Geometry", ExecutingAssemblyPath, ExecutingAssemblyName + ".BtnUpdateGeometry");



            RibbonPanel HydraPanel = application.CreateRibbonPanel("HYDRA", "BIM Automation");
            PushButton automationPushButton = HydraPanel.AddItem(automationButton) as PushButton;
            PushButton sectionsPushButton = HydraPanel.AddItem(sectionsButton) as PushButton;
            PushButton geometryPushButton = HydraPanel.AddItem(geometrysButton) as PushButton;



            automationPushButton.LargeImage = new BitmapImage(new Uri("D:\\Hydra\\Azure Repos (BIM Integration)\\BIM_Integration\\Sap_To_RVT\\Resources\\HYDRA.jpg"));
            sectionsPushButton.LargeImage = new BitmapImage(new Uri("D:\\Hydra\\Azure Repos (BIM Integration)\\BIM_Integration\\Sap_To_RVT\\Resources\\HYDRA.jpg"));
            geometryPushButton.LargeImage = new BitmapImage(new Uri("D:\\Hydra\\Azure Repos (BIM Integration)\\BIM_Integration\\Sap_To_RVT\\Resources\\HYDRA.jpg"));



            automationPushButton.ToolTipImage = new BitmapImage(new Uri("D:\\Hydra\\Azure Repos (BIM Integration)\\BIM_Integration\\Sap_To_RVT\\Resources\\HYDRA_1.jpg"));
            sectionsPushButton.ToolTipImage = new BitmapImage(new Uri("D:\\Hydra\\Azure Repos (BIM Integration)\\BIM_Integration\\Sap_To_RVT\\Resources\\HYDRA_1.jpg"));
            geometryPushButton.ToolTipImage = new BitmapImage(new Uri("D:\\Hydra\\Azure Repos (BIM Integration)\\BIM_Integration\\Sap_To_RVT\\Resources\\HYDRA_1.jpg"));



            // for divide Beams
            PushButtonData DivideCurvedBeamsButton = new PushButtonData("DivideCurvedBeams", "DivideCurvedBeams", ExecutingAssemblyPath, ExecutingAssemblyName + ".DivideCurvedBEams");
            PushButton DivideCurvedBeamsPushButton = HydraPanel.AddItem(DivideCurvedBeamsButton) as PushButton;
            DivideCurvedBeamsPushButton.LargeImage = new BitmapImage(new Uri("D:\\Hydra\\Azure Repos (BIM Integration)\\BIM_Integration\\Sap_To_RVT\\Resources\\HYDRA.jpg"));
            DivideCurvedBeamsPushButton.ToolTipImage = new BitmapImage(new Uri("D:\\Hydra\\Azure Repos (BIM Integration)\\BIM_Integration\\Sap_To_RVT\\Resources\\HYDRA_1.jpg"));

            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }





    }

}