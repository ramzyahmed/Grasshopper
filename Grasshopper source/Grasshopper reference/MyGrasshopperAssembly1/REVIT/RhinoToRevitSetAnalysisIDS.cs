using Grasshopper.Kernel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace MyGrasshopperAssembly1
{
    public class RhinoToRevitSetAnalysisIDS : GH_Component
    {
        public RhinoToRevitSetAnalysisIDS()
         : base("RhinoToRevitSetAnalysisIDS", "RhinoToRevitSetAnalysisIDS",
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
            string methodName = "RhinoToRevitSetAnalysisIDs";


            ReflectionFrom inst = new ReflectionFrom();
            inst.InvokeMethod(DA, assembleName, FullClassName, methodName);



        }



        protected override System.Drawing.Bitmap Icon => null;

        public override Guid ComponentGuid => new Guid("5C056C14-A55C-4914-B2E8-4CEC56C838E6");
    }
}
