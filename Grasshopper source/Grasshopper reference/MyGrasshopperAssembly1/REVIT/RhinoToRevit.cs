using Grasshopper.Kernel;
using System;


namespace MyGrasshopperAssembly1
{
    public class RhinoToRevit : GH_Component
    {

        public RhinoToRevit()
          : base("RhinoToRevit", "RhinoToRevit",
            "Description",
            "Curve", "Util")
        {
        }



        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddBooleanParameter("Run", "Run", "", GH_ParamAccess.item);


        }

        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {



        }


        protected override void SolveInstance(IGH_DataAccess DA)
        {


            string assembleName = "BIM_Integration.dll";
            string FullClassName = "MyGrasshopperAssembly2.RhinoRevit";
            string methodName = "RhinoToRevit";


            ReflectionFrom inst = new ReflectionFrom();
            inst.InvokeMethod(DA, assembleName, FullClassName, methodName);



        }




        protected override System.Drawing.Bitmap Icon => null;

        public override Guid ComponentGuid => new Guid("888a7f62-f768-4bef-af8c-ea39bd714c1e");
    }


}