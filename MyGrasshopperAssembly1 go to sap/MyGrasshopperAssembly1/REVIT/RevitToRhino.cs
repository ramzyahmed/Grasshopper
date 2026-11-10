using Grasshopper.Kernel;

using System;

namespace MyGrasshopperAssembly1
{
    public class RevitToRhino : GH_Component
    {

        public RevitToRhino()
          : base("RevitToRhino", "RevitToRhino",
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
            string methodName = "RevitToRhino";


            ReflectionFrom inst = new ReflectionFrom();
            inst.InvokeMethod(DA, assembleName, FullClassName, methodName);
        }

       


        protected override System.Drawing.Bitmap Icon => null;

        public override Guid ComponentGuid => new Guid("c4b3537d-5002-4f56-a9c2-05041f89669c");
    }


}