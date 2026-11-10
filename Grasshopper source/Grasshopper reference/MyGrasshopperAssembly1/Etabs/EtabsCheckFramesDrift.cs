using Eto.Forms;
using Grasshopper.Kernel;
using System;


namespace MyGrasshopperAssembly1
{
    public class CheckFramesDriftEtabs : GH_Component
    {
        public CheckFramesDriftEtabs()
          : base("CheckFramesDriftEtabs", "CheckFramesDriftEtabs",
            "CheckFramesDriftEtabs",
           "Curve", "Util")
        {
        }


        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            // pManager.AddBooleanParameter("run", "run","True for attach to the Active instance of  Sap otherwise new sap model will be created", GH_ParamAccess.item);
            pManager.AddBooleanParameter("Run", "Run", "", GH_ParamAccess.item);
            pManager.AddNumberParameter("factorForWind", "factorForWind", "", GH_ParamAccess.item);
            pManager.AddNumberParameter("factorForSeismic", "factorForSeismic", "", GH_ParamAccess.item);
            pManager.AddTextParameter("ColumnGroupSName", "ColumnGroupSName", "", GH_ParamAccess.list);
            pManager.AddTextParameter("LoadCaseKeyNAmes", "LoadCaseKeyNAmes", "", GH_ParamAccess.list);


        }

        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {



        }


        protected override void SolveInstance(IGH_DataAccess DA)
        {

            bool run = false;
            DA.GetData(0, ref run);

            if (run)
            {

                string assembleName = "BIM_Integration.dll";
                string FullClassName = "MyGrasshopperAssembly2.RhinoEtabs";
                string methodName = "EtabsCheckFramesDrift";


                ReflectionFrom inst = new ReflectionFrom();
                inst.InvokeMethod(DA, assembleName, FullClassName, methodName);
            }
        }


        protected override System.Drawing.Bitmap Icon => null;
       

        public override Guid ComponentGuid => new Guid("94EF4E95-EED2-4AF8-9DE0-A1B468E31C9C");
    }


}