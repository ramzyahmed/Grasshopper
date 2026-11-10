using System;
using System.Collections.Generic;
using System.Linq;
using System.Numerics;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI.Selection;

using Autodesk.Revit.UI;

namespace BIM_Integration

{
    public class VectorHelper
    {
        private static Matrix4x4 GetRotationalMatrix(double rotationAngleInDegree)
        {
            // Convert the rotation angle from degrees to radians
            float rotationAngleInRadians = (float)ToRadians(rotationAngleInDegree);

            // Create quaternion from axis angle
            Quaternion rotationQuaternion = Quaternion.CreateFromAxisAngle(Vector3.UnitZ, rotationAngleInRadians);

            // Get matrix for rotation 
            return Matrix4x4.CreateFromQuaternion(rotationQuaternion);
        }

        public static void GetTransformedFramesForRotationOnly(ref List<Frame> sapFrames, double rotationAngleInDegree, bool UseTheInverse = false)
        {
            Matrix4x4 rotationalMatrix = GetRotationalMatrix(rotationAngleInDegree);

            if (UseTheInverse)
            {
                Matrix4x4.Invert(rotationalMatrix, out rotationalMatrix);
            }



            // Update Beta for relevant frames
            sapFrames.Where(x => x.MemberType == MemberType.Column || x.MemberType == MemberType.SlanttedColumn)
                     .ToList()
                     .ForEach(x => x.Beta += rotationAngleInDegree);

            // Transform each frame
            foreach (Frame frame in sapFrames)
            {
                Vector3 startPt = Point.ToVector(frame.StartPoint);
                startPt = Vector3.Transform(startPt, rotationalMatrix);
                frame.StartPoint = new Point(startPt);

                Vector3 endPt = Point.ToVector(frame.EndPoint);
                endPt = Vector3.Transform(endPt, rotationalMatrix);
                frame.EndPoint = new Point(endPt);

                Vector3 midPt = Point.ToVector(frame.MidPoint);
                midPt = Vector3.Transform(midPt, rotationalMatrix);
                frame.MidPoint = new Point(midPt);
            }
        }

        private static double ToRadians(double degrees)
        {
            return degrees * (Math.PI / 180.0);
        }
    }


    #region SelectionClasses
    public class RevitStructuralFramingSelector
    {
        private readonly UIDocument _uidoc;

        public RevitStructuralFramingSelector(UIDocument uidoc)
        {
            _uidoc = uidoc;
        }

        public ElementId SelectStructuralFraming()
        {
            ElementId selectedFramingId = null;

            while (selectedFramingId == null || !IsStructuralFraming(selectedFramingId))
            {
                try
                {
                    // Prompt user to select an element
                    Reference pickedRef = _uidoc.Selection.PickObject(ObjectType.Element, new StructuralFramingSelectionFilter(), "Please select a structural framing element.");

                    selectedFramingId = pickedRef.ElementId;
                }
                catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                {
                    // User canceled the selection
                    return null;
                }

                if (selectedFramingId != null && !IsStructuralFraming(selectedFramingId))
                {
                    // Show dialog to inform the user that the selected element is not a structural framing element
                    TaskDialog.Show("Invalid Selection", "The selected element is not a structural framing element. Please select a structural framing element.");
                }
            }

            return selectedFramingId;
        }

        public List<ElementId> SelectStructuralFramesAndColumns()
        {
            List<ElementId> selectedFramingIds = new List<ElementId>();

            try
            {
                // Prompt user to select an element
                List<Reference> pickedRefs = _uidoc.Selection.PickObjects(ObjectType.Element, new StructuralElementsSelectionFilter(), "Please select  Beams and Columns").ToList();

                selectedFramingIds = pickedRefs.Select(x => x.ElementId).ToList();
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                // User canceled the selection
                return null;
            }

            return selectedFramingIds;
        }

        private bool IsStructuralFraming(ElementId elementId)
        {
            Element element = _uidoc.Document.GetElement(elementId);
            return element.Category != null && element.Category.Id.IntegerValue == (int)BuiltInCategory.OST_StructuralFraming;
        }
    }

    public class StructuralFramingSelectionFilter : ISelectionFilter
    {
        public bool AllowElement(Element elem)
        {
            // Filter only structural framing elements
            return elem.Category != null && elem.Category.Id.IntegerValue == (int)BuiltInCategory.OST_StructuralFraming;
        }

        public bool AllowReference(Reference reference, XYZ position)
        {
            return true;
        }
    }

    public class StructuralElementsSelectionFilter : ISelectionFilter
    {
        public bool AllowElement(Element elem)
        {
            // Filter only structural framing and structural columns elements
            if (elem.Category != null)
            {
                int categoryId = elem.Category.Id.IntegerValue;
                return categoryId == (int)BuiltInCategory.OST_StructuralFraming ||
                       categoryId == (int)BuiltInCategory.OST_StructuralColumns;
            }
            return false;
        }

        public bool AllowReference(Reference reference, XYZ position)
        {
            return true;
        }
    }
    #endregion

    public class TransformationHelper
    {

        public static void getTransfromedFrames(Matrix4x4 mat, ref List<Frame> sapFrames, Matrix4x4 trans, double rotatinAngle)
        {

            //sapFrames.Where(x => x.MemberType == MemberType.Column || x.MemberType == MemberType.SlanttedColumn).ToList().ForEach(x => x.Beta += rotatinAngle);

            sapFrames.Where(x => x.MemberType == MemberType.Column ).ToList().ForEach(x => x.Beta += rotatinAngle);

            foreach (Frame frame in sapFrames)
            {
                Vector3 startPt = Point.ToVector(frame.StartPoint);
                startPt = Vector3.Transform(startPt, mat);
                startPt = Vector3.Transform(startPt, trans);
                frame.StartPoint = new Point(startPt);

                Vector3 endPt = Point.ToVector(frame.EndPoint);
                endPt = Vector3.Transform(endPt, mat);
                endPt = Vector3.Transform(endPt, trans);
                frame.EndPoint = new Point(endPt);

                Vector3 midPt = Point.ToVector(frame.MidPoint);
                midPt = Vector3.Transform(midPt, mat);
                midPt = Vector3.Transform(midPt, trans);
                frame.MidPoint = new Point(midPt);

            }

        }

        public static void getTransfromedFrames2(Matrix4x4 mat, ref List<Frame> revitFrames)
        {

            foreach (Frame frame in revitFrames)
            {
                Vector3 startPt = Point.ToVector(frame.StartPoint);
                startPt = Vector3.Transform(startPt, mat);
                // startPt = Vector3.Transform(startPt, trans);
                frame.StartPoint = new Point(startPt);

                Vector3 endPt = Point.ToVector(frame.EndPoint);
                endPt = Vector3.Transform(endPt, mat);
                // endPt = Vector3.Transform(endPt, trans);
                frame.EndPoint = new Point(endPt);

                Vector3 midPt = Point.ToVector(frame.MidPoint);
                midPt = Vector3.Transform(midPt, mat);
                // midPt = Vector3.Transform(midPt, trans);
                frame.MidPoint = new Point(midPt);

            }

        }

        public static void getTransfromedFramesForRotationOnly(Matrix4x4 rotationalMatrix, ref List<Frame> sapFrames, double rotatinAngleInDegree)
        {

            sapFrames.Where(x => x.MemberType == MemberType.Column || x.MemberType == MemberType.SlanttedColumn).ToList().ForEach(x => x.Beta += rotatinAngleInDegree);


            foreach (Frame frame in sapFrames)
            {
                Vector3 startPt = Point.ToVector(frame.StartPoint);
                startPt = Vector3.Transform(startPt, rotationalMatrix);
                frame.StartPoint = new Point(startPt);

                Vector3 endPt = Point.ToVector(frame.EndPoint);
                endPt = Vector3.Transform(endPt, rotationalMatrix);
                frame.EndPoint = new Point(endPt);

                Vector3 midPt = Point.ToVector(frame.MidPoint);
                midPt = Vector3.Transform(midPt, rotationalMatrix);
                frame.MidPoint = new Point(midPt);

            }

        }

        public static Matrix4x4 get_transformation_masterMatrix(Frame sapLine, Frame revitLine, ref Matrix4x4 transMatrix, ref double rotationAngle)
        {
            Matrix4x4 rotationMatrix = new Matrix4x4();

            Vector3 sapStartPoint = Point.ToVector(sapLine.StartPoint);
            Vector3 revitStartPoint = Point.ToVector(revitLine.StartPoint);


            Vector3 sapVectorNonNormalized = new Vector3((float)(sapLine.EndPoint.X - sapLine.StartPoint.X), (float)(sapLine.EndPoint.Y - sapLine.StartPoint.Y), (float)(sapLine.EndPoint.Z - sapLine.StartPoint.Z));
            Vector3 sapVector = Vector3.Normalize(sapVectorNonNormalized);

            Vector3 revitVectorNonNormalized = new Vector3((float)(revitLine.EndPoint.X - revitLine.StartPoint.X), (float)(revitLine.EndPoint.Y - revitLine.StartPoint.Y), (float)(revitLine.EndPoint.Z - revitLine.StartPoint.Z));
            Vector3 revitVector = Vector3.Normalize(revitVectorNonNormalized);



            // get quatranion
            float roation = Vector3.Dot(revitVector, sapVector);
            Vector3 normalVector = Vector3.Cross(sapVector, revitVector);

            int signForRotation = 1;

            if (normalVector.Z < 0)
            {
                signForRotation = -1;
            }

            if (roation >1)
            {
                roation = 1;
            }

            if (roation <-1)
            {
                roation = -1;
            }
            Quaternion myquat = Quaternion.CreateFromAxisAngle(Vector3.UnitZ, signForRotation * (float)Math.Acos(roation));



            // get matrix for rotation 
            rotationMatrix = Matrix4x4.CreateFromQuaternion(myquat);

            Vector3 sapStPointTransfromed = Vector3.Transform(sapStartPoint, myquat);
            //Vector3 difference = Vector3.Subtract(sapStartPoint, revitStPointTransfromed);
            Vector3 difference = Vector3.Subtract(revitStartPoint, sapStPointTransfromed);
            Matrix4x4 trans = Matrix4x4.CreateTranslation(difference);


            transMatrix = trans;


            // rotationAngle = roation * 180.0 / Math.PI;

            rotationAngle = Math.Acos(Vector3.Dot(revitVector, sapVector)) * 180 / Math.PI;

            // If you want to ensure rotation is always positive (0 to 180 degrees)
            if (rotationAngle > 180f)
            {
                rotationAngle = 360f - rotationAngle;
            }



            return rotationMatrix;
        }


    }
}
