using Grasshopper.Kernel;
using System;


namespace MyGrasshopperAssembly1.GH
{
    public class setPointlocalAxes : GH_Component
    {


        public setPointlocalAxes()
          : base("setPointLocalAxes", "setPointLocalAxes",
            "Description",
            "Curve", "Util")
        {
        }

        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddBooleanParameter("bool", "bool", "", GH_ParamAccess.item);
            pManager.AddTextParameter("ColumnsGroupName", "ColumnsGroupName", "", GH_ParamAccess.item);

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
                string methodName = "SetPointLocalAxesBasedOnConnectedColumn";


                ReflectionFrom inst = new ReflectionFrom();
                inst.InvokeMethod(DA, assembleName, FullClassName, methodName);

            }
        }

        public override Guid ComponentGuid
        {
            get { return new Guid("1ed72d64-d413-418f-a9f4-cbc90181a0ee"); }
        }


    }
}