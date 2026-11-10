using Grasshopper.Kernel;
using System;
using System.Collections.Generic;
using System.Linq;
namespace MyGrasshopperAssembly1.GH
{
    public class SapGroupingCheck : GH_Component
    {


        public SapGroupingCheck()
          : base("SapGrouping", "SapGrouping",
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
                string methodName = "SapGroupingCheck";


                ReflectionFrom inst = new ReflectionFrom();
                inst.InvokeMethod(DA, assembleName, FullClassName, methodName);

            }
        }

        public override Guid ComponentGuid
        {
            get { return new Guid("945dc13c-0ec4-4a01-b5c8-93d0509ff987"); }
        }

      

    }
}