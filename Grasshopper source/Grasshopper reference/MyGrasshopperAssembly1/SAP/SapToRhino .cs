using Grasshopper.Kernel;
using Rhino.DocObjects;
using Rhino.Geometry;
using System;
using System.Collections.Generic;
using System.Linq;


namespace MyGrasshopperAssembly1
{
    public class SapToRhino : GH_Component
    {

        public SapToRhino()
          : base("SapToRhino", "SapToRhino",
            "Description",
            "Curve", "Util")
        {
        }


        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddBooleanParameter("Do", "Do", "", GH_ParamAccess.item);
        }

        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddCurveParameter("Curves", "crvs", "", GH_ParamAccess.list);
            pManager.AddBrepParameter("Breabs", "Brebs", "", GH_ParamAccess.list);
            pManager.AddPointParameter("Points", "pts", "", GH_ParamAccess.list);


        }


        protected override void SolveInstance(IGH_DataAccess DA)
        {
            bool aa = false;
            DA.GetData(0, ref aa);

            if (aa)
            {
                string assembleName = "BIM_Integration.dll";
                string FullClassName = "MyGrasshopperAssembly2.RhinoSap";
                string methodName = "SapToRhino";


                ReflectionFrom inst = new ReflectionFrom();
                inst.InvokeMethod(DA, assembleName, FullClassName, methodName);

            }



        }


        protected override System.Drawing.Bitmap Icon => null;


        public override Guid ComponentGuid => new Guid("17b03553-41d4-4191-8bda-678f099e93b4");
    }


}