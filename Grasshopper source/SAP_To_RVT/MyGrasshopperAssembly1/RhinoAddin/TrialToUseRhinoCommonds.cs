using BIM_Integration;
using Grasshopper.Kernel;
using Rhino.Commands;
using Rhino;
using Rhino.DocObjects;
using Rhino.Geometry;
using System;
using System.Collections.Generic;
using System.Linq;
using ClassesDefintions = BIM_Integration;
using Rhino.Input.Custom;

namespace MyGrasshopperAssembly2
{
    public class MyCommand : Rhino.Commands.Command
    {
        public override string EnglishName => "MyCommand";

        protected override Result RunCommand(RhinoDoc doc, RunMode mode)
        {
            // Prompt the user to select the source Brep
            var sourceBrep = default(Brep);
            var sourceBrepObjRef = new GetObject();
            sourceBrepObjRef.SetCommandPrompt("Select source Brep");
            sourceBrepObjRef.GeometryFilter = ObjectType.Brep;
            sourceBrepObjRef.Get();
            if (sourceBrepObjRef.CommandResult() != Result.Success)
                return sourceBrepObjRef.CommandResult();
            sourceBrep = sourceBrepObjRef.Object(0).Brep();

            // Prompt the user to select the target Brep
            var targetBrep = default(Brep);
            var targetBrepObjRef = new GetObject();
            targetBrepObjRef.SetCommandPrompt("Select target Brep");
            targetBrepObjRef.GeometryFilter = ObjectType.Brep;
            targetBrepObjRef.Get();
            if (targetBrepObjRef.CommandResult() != Result.Success)
                return targetBrepObjRef.CommandResult();
            targetBrep = targetBrepObjRef.Object(0).Brep();

            // Get the ID attribute from the target Brep
            var targetID = default(string);
            if (!targetBrep.UserDictionary.TryGetString("ID", out targetID))
                return Result.Failure;

            // Set the ID attribute on the source Brep to the value from the target Brep
            sourceBrep.UserDictionary.Set("ID", targetID);

            // TODO: Add your command code here

            return Result.Success;
        }


    }


}