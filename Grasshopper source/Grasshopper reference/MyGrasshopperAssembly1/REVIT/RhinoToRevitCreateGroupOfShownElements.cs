using Grasshopper.Kernel;
using System;


namespace MyGrasshopperAssembly1
{
    public class RhinoToRevitCreateGroupOfShownElements : GH_Component
    {
        public RhinoToRevitCreateGroupOfShownElements()
         : base("RhinoToRevitCreateGroupOfShownElements", "RhinoToRevitCreateGroupOfShownElements",
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
            string methodName = "RhinoToRevitCreateGroupOfShownElemenets";


            ReflectionFrom inst = new ReflectionFrom();
            inst.InvokeMethod(DA, assembleName, FullClassName, methodName);



        }



        protected override System.Drawing.Bitmap Icon => null;

        public override Guid ComponentGuid => new Guid("051B8FEF-4BA6-4665-87C6-1A81C0DEB64F");
    }
}
