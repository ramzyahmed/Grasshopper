using Eto.Forms;
using Grasshopper.Kernel;
using Rhino.DocObjects;
using Rhino.Geometry;
using System;
using System.Collections.Generic;
using System.Linq;


namespace MyGrasshopperAssembly1
{
    public class EtabsToRhino : GH_Component
    {
        /// <summary>
        /// Each implementation of GH_Component must provide a public 
        /// constructor without any arguments.
        /// Category represents the Tab in which the component will appear, 
        /// Subcategory the panel. If you use non-existing tab or panel names, 
        /// new tabs/panels will automatically be created.
        /// </summary>
        public EtabsToRhino()
          : base("EtabsToRhino", "EtabsToRhino",
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


                string assembleName = "BIM_Integration.dll";
                string FullClassName = "MyGrasshopperAssembly2.RhinoEtabs";
                string methodName = "EtabsToRhino";


                ReflectionFrom inst = new ReflectionFrom();
                inst.InvokeMethod(DA, assembleName, FullClassName, methodName);
            

        }


        protected override System.Drawing.Bitmap Icon => null;


        public override Guid ComponentGuid => new Guid("96936EAD-1297-4A78-AE91-EFEE057242E7");
    }


}