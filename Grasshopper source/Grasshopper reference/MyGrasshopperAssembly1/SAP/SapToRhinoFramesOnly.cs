using Grasshopper.Kernel;
using Rhino.DocObjects;
using Rhino.Geometry;
using System;
using System.Collections.Generic;
using System.Linq;


namespace MyGrasshopperAssembly1
{
    public class RhinoToSapFramesOnly : GH_Component
    {

        public RhinoToSapFramesOnly()
          : base("RhinoToSapFramesOnly", "RhinoToSapFramesOnly",
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
        }


        protected override void SolveInstance(IGH_DataAccess DA)
        {
            bool aa = false;
            DA.GetData(0, ref aa);

            if (aa)
            {
                string assembleName = "BIM_Integration.dll";
                string FullClassName = "MyGrasshopperAssembly2.RhinoSap";
                string methodName = "RhinoToSapFramesOnly";


                ReflectionFrom inst = new ReflectionFrom();
                inst.InvokeMethod(DA, assembleName, FullClassName, methodName);

            }



        }


        protected override System.Drawing.Bitmap Icon => null;


        public override Guid ComponentGuid => new Guid("5838295c-9b83-4f91-9fab-0b6bf546e44e");
    }


}