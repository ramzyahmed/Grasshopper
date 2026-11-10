using Eto.Forms;
using Grasshopper.Kernel;
using System;
using System.Collections.Generic;
using System.Linq;

namespace MyGrasshopperAssembly1.GH
{
    public class SapGroupingCheckEtabs : GH_Component
    {


        public SapGroupingCheckEtabs()
          : base("SapGroupingCheckEtabs", "SapGroupingCheckEtabs",
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

            bool run = false;
            DA.GetData(0, ref run);

            if (run)
            {

                string assembleName = "BIM_Integration.dll";
                string FullClassName = "MyGrasshopperAssembly2.RhinoEtabs";
                string methodName = "EtabsGroupingCheck";


                ReflectionFrom inst = new ReflectionFrom();
                inst.InvokeMethod(DA, assembleName, FullClassName, methodName);
            }
        }


        public override Guid ComponentGuid
        {
            get { return new Guid("A9D94F9E-AF96-4766-93AC-D13C3895AD49"); }
        }

    }
}