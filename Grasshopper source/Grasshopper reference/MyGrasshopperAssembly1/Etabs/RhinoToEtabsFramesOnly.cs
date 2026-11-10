using Eto.Forms;
using Grasshopper.Kernel;
using Rhino.DocObjects;
using Rhino.Geometry;
using System;
using System.Collections.Generic;
using System.Linq;


namespace MyGrasshopperAssembly1
{
    public class RhinoToEtabsFramesOnly : GH_Component
    {
        /// <summary>
        /// Each implementation of GH_Component must provide a public 
        /// constructor without any arguments.
        /// Category represents the Tab in which the component will appear, 
        /// Subcategory the panel. If you use non-existing tab or panel names, 
        /// new tabs/panels will automatically be created.
        /// </summary>
        public RhinoToEtabsFramesOnly()
          : base("RhinoToEtabsFramesOnly", "RhinoToEtabsFramesOnly",
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


                string assembleName = "BIM_Integration.dll";
                string FullClassName = "MyGrasshopperAssembly2.RhinoEtabs";
                string methodName = "RhinoToEtabsFramesOnly";

                ReflectionFrom inst = new ReflectionFrom();
                inst.InvokeMethod(DA, assembleName, FullClassName, methodName);
            

        }


        protected override System.Drawing.Bitmap Icon => null;


        public override Guid ComponentGuid => new Guid("e67cdabc-eeed-4c8c-b055-bfc2b6ea76d5");
    }


}