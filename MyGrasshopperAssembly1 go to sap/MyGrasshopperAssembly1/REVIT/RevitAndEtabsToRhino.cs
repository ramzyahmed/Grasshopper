using Grasshopper.Kernel;

using System;

namespace MyGrasshopperAssembly1
{
    public class RevitAndEtabsToRhino : GH_Component
    {

        public RevitAndEtabsToRhino()
          : base("RevitAndEtabsToRhino", "RevitAndEtabsToRhino",
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
            string methodName = "RevitAndEtabsToRhino";


            ReflectionFrom inst = new ReflectionFrom();
            inst.InvokeMethod(DA, assembleName, FullClassName, methodName);
        }

       


        protected override System.Drawing.Bitmap Icon => null;

        public override Guid ComponentGuid => new Guid("7e9e7ac8-9d5f-4169-a4e1-ab2b465efcaf");
    }


}