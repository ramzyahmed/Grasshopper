using Rhino;
using Rhino.Geometry;
using Rhino.DocObjects;
using Rhino.Input;
using Rhino.Commands;
using System;

public class GetAndApplyTransformationMatrix : Command
{
    public GetAndApplyTransformationMatrix()
    {
        Instance = this;
    }

    public static GetAndApplyTransformationMatrix Instance { get; private set; }

    public override string EnglishName => "GetAndApplyTransformationMatrix";

    protected override Result RunCommand(RhinoDoc doc, RunMode mode)
    {
        // Get the initial two curves
        ObjRef objRef1;
        ObjRef objRef2;

        Result result = RhinoGet.GetOneObject("Select first curve", false, ObjectType.Curve, out objRef1);
        if (result != Result.Success) return result;

        result = RhinoGet.GetOneObject("Select second curve", false, ObjectType.Curve, out objRef2);
        if (result != Result.Success) return result;

        Curve curve1 = objRef1.Curve();
        Curve curve2 = objRef2.Curve();

        // Get start and end points of the first curve
        Point3d start1 = curve1.PointAtStart;
        Point3d end1 = curve1.PointAtEnd;

        // Get start and end points of the second curve
        Point3d start2 = curve2.PointAtStart;
        Point3d end2 = curve2.PointAtEnd;

        // Create vectors from the points
        Vector3d vector1 = end1 - start1;
        Vector3d vector2 = end2 - start2;

        // Create a translation vector from start1 to start2
        Vector3d translation = start2 - start1;

        // Create a rotation transformation from vector1 to vector2
        Transform rotation = Transform.Identity;
        rotation =  Transform.Rotation(vector1, vector2, Point3d.Origin);

        // Create the translation transformation
        Transform translationTransform = Transform.Translation(translation);

        // Combine the rotation and translation transformations
        Transform transform = translationTransform * rotation;

        // Output the transformation matrix
        RhinoApp.WriteLine("Transformation Matrix:");
        RhinoApp.WriteLine(transform.ToString());

        // Get additional curves to transform
        ObjRef[] objRefs;
        result = RhinoGet.GetMultipleObjects("Select curves to transform", false, ObjectType.Curve, out objRefs);
        if (result != Result.Success) return result;

        foreach (var objRef in objRefs)
        {
            Curve curve = objRef.Curve();
            if (curve != null)
            {
                Curve transformedCurve = curve.DuplicateCurve();
                transformedCurve.Transform(transform);
                doc.Objects.AddCurve(transformedCurve);
            }
        }

        doc.Views.Redraw();
        return Result.Success;
    }
}