using Grasshopper.Kernel;
using System;


namespace MyGrasshopperAssembly1
{
    public class RhinoToRhino : GH_Component
    {

        public RhinoToRhino()
          : base("RhinoToRhino", "RhinoToRhino",
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
            string methodName = "RhinoToRhino";


            ReflectionFrom inst = new ReflectionFrom();
            inst.InvokeMethod(DA, assembleName, FullClassName, methodName);

        }

        protected override System.Drawing.Bitmap Icon => null;

        public override Guid ComponentGuid => new Guid("c31e1952-a466-4426-9b54-721165441f98");
    }


}