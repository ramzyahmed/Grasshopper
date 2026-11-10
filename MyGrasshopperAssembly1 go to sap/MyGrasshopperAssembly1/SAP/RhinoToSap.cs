using Grasshopper.Kernel;
using System;



namespace MyGrasshopperAssembly1
{
    public class RhinoToSAP : GH_Component
    {

        public RhinoToSAP()
          : base("RhinoToSAP", "RhinoToSAP",
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

            bool aa = false;
            DA.GetData(0, ref aa);
            if (aa)
            {
                string assembleName = "BIM_Integration.dll";
                string FullClassName = "MyGrasshopperAssembly2.RhinoSap";
                string methodName = "RhinoToSap";


                ReflectionFrom inst = new ReflectionFrom();
                inst.InvokeMethod(DA, assembleName, FullClassName, methodName);


            }
        }


        protected override System.Drawing.Bitmap Icon => null;

        public override Guid ComponentGuid => new Guid("72062C69-D370-4D93-9BEF-F69EA227AE4A");
    }


}