using Eto.Forms;
using Grasshopper.Kernel;
using System;



namespace MyGrasshopperAssembly1
{
    public class RhinoToEtabs : GH_Component
    {

        public RhinoToEtabs()
          : base("RhinoToEtabs", "RhinoToEtabs",
            "Description",
            "Curve", "Util")
        {
        }

        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddBooleanParameter("bool", "bool", "", GH_ParamAccess.item);
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
                string methodName = "RhinoToEtabs";


                ReflectionFrom inst = new ReflectionFrom();
                inst.InvokeMethod(DA, assembleName, FullClassName, methodName);
            }
        }

        protected override System.Drawing.Bitmap Icon => null;

        public override Guid ComponentGuid => new Guid("0EB931DA-664D-4481-BDFF-DF25A0ED2398");
    }


}