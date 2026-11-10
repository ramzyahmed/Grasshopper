using System;
using System.Collections.Generic;
using System.Linq;
using System.Numerics;

namespace BIM_Integration_Hydra
{

        #region Classes_declerations

        public class Frame
        {
            // List of Revit Families

            public static string BM_BUS_I = "BM_STL_BUS-I_DAR";
            public static string BM_HEA = "BM_STL_HEA_DAR";
            public static string BM_HEAA = "BM_STL_HEAA_DAR";
            public static string BM_HEB = "BM_STL_HEB_DAR";
            public static string BM_HEM = "BM_STL_HEM_DAR";
            public static string BM_IPE = "BM_STL_IPE_DAR";
            public static string BM_IPEA = "BM_STL_IPEA_DAR";
            public static string BM_IPEO = "BM_STL_IPEO_DAR";
            public static string BM_L_EQUAL = "BM_STL_L_EQUAL_DAR";
            public static string BM_2L_EQUALANGLE = "BM_STL_2L_EQUAL_DAR";
            public static string BM_L_UNEQUAL = "BM_STL_L_UNEQUAL_DAR";
            public static string BM_PIPE = "BM_STL_PIPE_DAR";
            public static string BM_TUBE = "BM_STL_TUBE_DAR";
            public static string BM_UPN = "BM_STL_UPN_DAR";
            public static string BM_BUS_I_Tapered = "BM_STL_BUS-I-Tapered_DAR";
            public static string BM_BUS_PIPE_Tapered = "BM_STL_BUS-PIPE-Tapered_DAR";
            public static string BM_BUS_TUBE_Tapered = "BM_STL_BUS-TUBE-Tapered_DAR";
            public static string BM_W = "BM_STL_W_DAR";
            public static string BM_UB = "BM_STL_UB_DAR";
            public static string BM_UC = "BM_STL_UC_DAR";
            public static string BM_Dummy = "Beam_Dummy_DAR";



            public static string COL_BUS_I = "COL_STL_BUS-I_DAR";
            public static string COL_HEA = "COL_STL_HEA_DAR";
            public static string COL_HEAA = "COL_STL_HEAA_DAR";
            public static string COL_HEB = "COL_STL_HEB_DAR";
            public static string COL_HEM = "COL_STL_HEM_DAR";
            public static string COL_IPE = "COL_STL_IPE_DAR";
            public static string COL_IPEA = "COL_STL_IPEA_DAR";
            public static string COL_IPEO = "COL_STL_IPEO_DAR";
            public static string COL_L_EQUAL = "COL_STL_L_EQUAL_DAR";
            public static string COL_L_UNEQUAL = "COL_STL_L_UNEQUAL_DAR";
            public static string COL_2L_EQUALANGLE = "COL_STL_2L_EQUAL_DAR";
            public static string COL_PIPE = "COL_STL_PIPE_DAR";
            public static string COL_TUBE = "COL_STL_TUBE_DAR";
            public static string COL_UPN = "COL_STL_UPN_DAR";
            public static string COL_BUS_I_Tapered = "COL_STL_BUS-I-Tapered_DAR";
            public static string COL_BUS_PIPE_Tapered = "COL_STL_BUS-PIPE-Tapered_DAR";
            public static string COL_BUS_TUBE_Tapered = "COL_STL_BUS-TUBE-Tapered_DAR";
            public static string COL_W = "COL_STL_W_DAR";
            public static string COL_UB = "COL_STL_UB_DAR";
            public static string COL_UC = "COL_STL_UC_DAR";
            public static string COL_Dummy = "COL_Dummy_DAR";



            #region  props
            public string ID { get; set; }
            public string RvtID { get; set; }
            public Point StartPoint { get; set; }
            public Point EndPoint { get; set; }
            public Point MidPoint { get; set; }
            public double Beta { get; set; }
            public string Shape { get; set; }
            public Section_type SecType { get; set; }

            public SectType FamilyType { get; set; }

            //public ElementId revitEleId { get; set; }

            public string Sec { get; set; }

            public bool IsBUS { get; set; }
            public bool IsCable { get; set; }
            public string RvtFamily
            {
                get; set;
            }
            public bool IsLink { get; set; }
            public bool IsConc { get; set; }


            public string RvtName { get; set; }

            public string Lib { get; set; }
            public string Material { get; set; }

            public string Group { get; set; } // Not Used

            public double Length { get; set; }


            public MemberType MemberType { get; set; }

            public int rowNo { get; set; }

            #endregion

            public Frame()
            {
                StartPoint = new Point();
                EndPoint = new Point();
                MidPoint = new Point();
            }



            public static int CountCharsUsingIndex(string source, char toFind)
            {
                int count = 0;
                int n = 0;
                while ((n = source.IndexOf(toFind, n) + 1) != 0)
                {
                    n++;
                    count++;
                }
                return count;
            }


            public static int IndexOfNth(string str, string value, int nth)
            {
                if (nth < 0)
                    throw new ArgumentException("Can not find a negative index of substring in string. Must start with 0");

                int offset = str.IndexOf(value);
                for (int i = 0; i < nth; i++)
                {
                    if (offset == -1) return -1;
                    offset = str.IndexOf(value, offset + 1);
                }

                return offset;
            }



        }

    
        public class Beam
        {
            #region  props

            public string ID { get; set; }
            public string RvtID { get; set; }
            public Point StartPoint { get; set; }
            public Point MidPoint { get; set; }
            public Point EndPoint { get; set; }
            public double Beta { get; set; }
            public string Shape { get; set; }
            public Section_type SecType { get; set; }
            public SectType FamilyType { get; set; }
            public string Sec { get; set; }

            public bool IsBUS { get; set; }
            public bool IsCable { get; set; }
            public bool IsLink { get; set; }
            public bool IsConc { get; set; }


            public string RvtName { get; set; }

            public string RvtFamily { get; set; }
            public string Lib { get; set; }
            public string Material { get; set; }

            public string Group { get; set; } // Not Used

            public double Length { get; set; }
            public MemberType MemberType { get; set; }

            public int rowNo { get; set; }

            public Beam(string ID, Point StartPoint, Point EndPoint, Point midPoint, double Beta, string Shape, Section_type SecType, string Sec, bool IsBUS, bool IsCable, bool IsLink,
              bool IsConc, string RvtName, string Lib, string Material, string Group, double Length, MemberType memberType, int rowNo)

            {
                this.ID = ID;
                this.StartPoint = StartPoint;
                this.EndPoint = EndPoint;
                this.Beta = Beta;
                this.Shape = Shape;
                this.SecType = SecType;
                this.Sec = Sec;
                this.IsBUS = IsBUS;
                this.IsCable = IsCable;
                this.IsLink = IsLink;
                this.IsConc = IsConc;
                this.RvtName = RvtName;
                this.Lib = Lib;
                this.Material = Material;
                this.Group = Group;
                this.Length = Length;
                this.MemberType = memberType;
                this.rowNo = rowNo;
                this.MidPoint = midPoint;

            }


            public Beam(Frame frame)
            {
                this.Beta = frame.Beta;
                this.EndPoint = frame.EndPoint;
                this.ID = frame.ID;
                this.IsBUS = frame.IsBUS;
                this.IsCable = frame.IsCable;
                this.IsConc = frame.IsConc;
                this.IsLink = frame.IsLink;
                this.Length = frame.Length;
                this.Lib = frame.Lib;
                this.Material = frame.Material;
                this.MemberType = frame.MemberType;
                this.rowNo = frame.rowNo;
                this.RvtFamily = frame.RvtFamily;
                this.RvtID = frame.RvtID;
                this.RvtName = frame.RvtName;
                this.Sec = frame.Sec;
                this.SecType = frame.SecType;
                this.Shape = frame.Shape;
                this.StartPoint = frame.StartPoint;
                this.MidPoint = frame.MidPoint;

            }


            #endregion


            public static void updateStaadSections(List<SapRevitFrame> MappedBeams, string type)
            {


                Dictionary<string, int> revitSTaadSectionDic = new Dictionary<string, int>();
                List<Frame> FilteredRevitFrameByName = MappedBeams.Select(x => x.revitFrame).ToList().GroupBy(x => x.RvtName).Select(x => x.FirstOrDefault()).ToList();


                if (type == "BM")
                {
                    revitSTaadSectionDic =BIM_Integration_Hydra.StaadUtils_HYDRA.createStaadSectionForBeams(FilteredRevitFrameByName);

                }

                else if (type == "COL")
                {
                    revitSTaadSectionDic =  BIM_Integration_Hydra.StaadUtils_HYDRA.createStaadSectionForColumns(FilteredRevitFrameByName);

                }

                foreach (var mappedBeam in MappedBeams)
                {

                    // Update section 
                    StaadUtils_HYDRA.property.AssignBeamProperty(mappedBeam.sapFrame.ID, revitSTaadSectionDic[mappedBeam.revitFrame.RvtName]);
                    mappedBeam.sapFrame.Sec =  StaadUtils_HYDRA.property.GetBeamSectionName(mappedBeam.sapFrame.ID);
                    mappedBeam.sapFrame.rowNo = mappedBeam.revitFrame.rowNo;

                    // update rotation
                    StaadUtils_HYDRA.property.AssignBetaAngle(mappedBeam.sapFrame.ID, -mappedBeam.revitFrame.Beta);


                    // update rotation and section in Excel 
                    continue;


                }


            }

            public static void updateStaadGeomerty(List<SapRevitFrame> MappedBeams, double margin)
            {
                double margin2 = 0.0;
                string propName = null;
                string Sauto = null;
                string materialName = null;
                int ret = 0;
                string stPoint = null;
                string endPoint = null;

                double lengthFactor = 0.0;
                double forceFactor = 0.0;
                double momentFactor = 0.0;




                foreach (var mappedBeam in MappedBeams)
                {


                    margin2 = margin * mappedBeam.revitFrame.Length;
                    double dist1 = Point.getDistance(mappedBeam.revitFrame.StartPoint, mappedBeam.sapFrame.StartPoint);
                    double dist2 = Point.getDistance(mappedBeam.revitFrame.EndPoint, mappedBeam.sapFrame.EndPoint);


                    if (dist1 < margin2 && dist2 < margin2)
                    {

                        // Update rotation 
                         StaadUtils_HYDRA.property.AssignBetaAngle(mappedBeam.sapFrame.ID, -mappedBeam.revitFrame.Beta);
                        mappedBeam.sapFrame.rowNo = mappedBeam.revitFrame.rowNo;
                        // update Rotation in Excel to be modified





                        continue;
                    }

                    else
                    {

                         StaadUtils_HYDRA.getUnitFactorsFactorToStaad(ref forceFactor, ref momentFactor, ref lengthFactor);


                        int newStartNodeId =  StaadUtils_HYDRA.gometry.AddNode(mappedBeam.revitFrame.StartPoint.X * lengthFactor, mappedBeam.revitFrame.StartPoint.Z * lengthFactor, mappedBeam.revitFrame.StartPoint.Y * lengthFactor * -1);
                        int newEndtNodeId =  StaadUtils_HYDRA.gometry.AddNode(mappedBeam.revitFrame.EndPoint.X * lengthFactor, mappedBeam.revitFrame.EndPoint.Z * lengthFactor, mappedBeam.revitFrame.EndPoint.Y * lengthFactor * -1);

                         StaadUtils_HYDRA.gometry.SetNodeCoordinate(mappedBeam.sapFrame.StartPoint.Id, mappedBeam.revitFrame.StartPoint.X * lengthFactor, mappedBeam.revitFrame.StartPoint.Z * lengthFactor, mappedBeam.revitFrame.StartPoint.Y * lengthFactor * -1);
                         StaadUtils_HYDRA.gometry.SetNodeCoordinate(mappedBeam.sapFrame.EndPoint.Id, mappedBeam.revitFrame.EndPoint.X * lengthFactor, mappedBeam.revitFrame.EndPoint.Z * lengthFactor, mappedBeam.revitFrame.EndPoint.Y * lengthFactor * -1);



                        // Update rotation 
                         StaadUtils_HYDRA.property.AssignBetaAngle(mappedBeam.sapFrame.ID, -mappedBeam.revitFrame.Beta);
                        mappedBeam.sapFrame.rowNo = mappedBeam.revitFrame.rowNo;

                        // update Geometry in Excel to be modified
                        continue;
                    }


                }


            }


        }


        public class Column
        {
            #region  props

            public string ID { get; set; }
            public string RvtID { get; set; }
            public Point StartPoint { get; set; }
            public Point MidPoint { get; set; }
            public Point EndPoint { get; set; }
            public double Beta { get; set; }
            public string Shape { get; set; }
            public Section_type SecType { get; set; }

            public string Sec { get; set; }

            public bool IsBUS { get; set; }
            public bool IsCable { get; set; }
            public bool IsLink { get; set; }
            public bool IsConc { get; set; }


            public string RvtName { get; set; }

            public string RvtFamily { get; set; }
            public string Lib { get; set; }
            public string Material { get; set; }

            public string Group { get; set; } // Not Used

            public double Length { get; set; }
            public MemberType MemberType { get; set; }

            public int rowNo { get; set; }

            public Column(string ID, Point StartPoint, Point EndPoint, Point MidPoint, double Beta, string Shape, Section_type SecType, string Sec, bool IsBUS, bool IsCable, bool IsLink,
              bool IsConc, string RvtName, string Lib, string Material, string Group, double Length, MemberType memberTypE, int rowNo)

            {
                this.ID = ID;
                this.StartPoint = StartPoint;
                this.EndPoint = EndPoint;
                this.Beta = Beta;
                this.Shape = Shape;
                this.SecType = SecType;
                this.Sec = Sec;
                this.IsBUS = IsBUS;
                this.IsCable = IsCable;
                this.IsLink = IsLink;
                this.IsConc = IsConc;
                this.RvtName = RvtName;
                this.Lib = Lib;
                this.Material = Material;
                this.Group = Group;
                this.Length = Length;
                this.MemberType = memberTypE;
                this.rowNo = rowNo;
                this.MidPoint = MidPoint;
            }


            public Column(Frame frame)
            {
                this.Beta = frame.Beta;
                this.EndPoint = frame.EndPoint;
                this.ID = frame.ID;
                this.IsBUS = frame.IsBUS;
                this.IsCable = frame.IsCable;
                this.IsConc = frame.IsConc;
                this.IsLink = frame.IsLink;
                this.Length = frame.Length;
                this.Lib = frame.Lib;
                this.Material = frame.Material;
                this.MemberType = frame.MemberType;
                this.rowNo = frame.rowNo;
                this.RvtFamily = frame.RvtFamily;
                this.RvtID = frame.RvtID;
                this.RvtName = frame.RvtName;
                this.Sec = frame.Sec;
                this.SecType = frame.SecType;
                this.Shape = frame.Shape;
                this.StartPoint = frame.StartPoint;
                this.MidPoint = frame.MidPoint;

            }


            #endregion



        }


        public class Point
        {
            #region prop

            public string Id { get; set; }
            public double X { get; set; }
            public double Y { get; set; }
            public double Z { get; set; }

            public Point(string id, double x, double y, double z)
            {
                Id = id;
                X = x;
                Y = y;
                Z = z;
            }

            #endregion



            public Point()
            {

            }

            public Point(Vector3 vec)
            {
                X = vec.X;
                Y = vec.Y;
                Z = vec.Z;
            }


            public static Vector3 ToVector(Point p)
            {
                return new Vector3((float)p.X, (float)p.Y, (float)p.Z);
            }


            public static Point getMidPoint(Point P1, Point P2)
            {
                return new Point(0.ToString(), (P1.X + P2.X) / 2, (P1.Y + P2.Y) / 2, (P1.Z + P2.Z) / 2);

            }

            public static double getDistance(Point p1, Point p2)
            {
                double result = 0;

                double sumOFX = Math.Pow(p1.X - p2.X, 2);
                double sumOFY = Math.Pow(p1.Y - p2.Y, 2);
                double sumOFZ = Math.Pow(p1.Z - p2.Z, 2);

                result = Math.Sqrt(sumOFX + sumOFY + sumOFZ);


                return result;
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

                Quaternion myquat = Quaternion.CreateFromAxisAngle(Vector3.UnitZ, signForRotation * (float)Math.Acos(roation));



                // get matrix for rotation 
                rotationMatrix = Matrix4x4.CreateFromQuaternion(myquat);

                Vector3 sapStPointTransfromed = Vector3.Transform(sapStartPoint, myquat);
                //Vector3 difference = Vector3.Subtract(sapStartPoint, revitStPointTransfromed);
                Vector3 difference = Vector3.Subtract(revitStartPoint, sapStPointTransfromed);
                Matrix4x4 trans = Matrix4x4.CreateTranslation(difference);


                transMatrix = trans;
                rotationAngle = roation * 180.0 / Math.PI;
                return rotationMatrix;
            }

            public static void getMemberType(Frame F)
            {

                double y_dif = 0;
                double x_dif = 0;
                double z_dif = 0;
                double incl_angle = 0;
                double incl_angle2 = 0;
                double inclRefColumn = 4;
                double inclRefVRColumn = 1 / 1000.00;
                double inclRefVLBracing = 20;


                #region filter between columns and beams



                if (F.StartPoint.X == F.EndPoint.X && F.StartPoint.Y == F.EndPoint.Y)
                {
                    F.MemberType = MemberType.Column; // vertical column
                }



                else
                {
                    y_dif = Math.Abs(F.EndPoint.Y - F.StartPoint.Y);
                    x_dif = Math.Abs(F.EndPoint.X - F.StartPoint.X);
                    z_dif = Math.Abs(F.EndPoint.Z - F.StartPoint.Z);
                    incl_angle = Math.Atan(y_dif / z_dif) * 180 / Math.PI;
                    incl_angle2 = Math.Atan(x_dif / z_dif) * 180 / Math.PI;



                    // column or SlanttedColumn
                    if (incl_angle < inclRefColumn && incl_angle2 < inclRefColumn)
                    {
                        //VRcolumn
                        if (incl_angle < inclRefVRColumn && incl_angle2 < inclRefVRColumn)

                        {
                            F.MemberType = MemberType.Column; // vrcolumn

                        }

                        //SlanttedColumn
                        else
                        {
                            F.MemberType = MemberType.SlanttedColumn; // slanted column

                        }

                    }





                    // Beam of Vrbracing
                    else

                    {

                        if ((incl_angle < inclRefVLBracing && incl_angle2 < inclRefVLBracing))
                        {
                            F.MemberType = MemberType.VLBracing;   //vr_bracing
                        }

                        else
                        {
                            F.MemberType = MemberType.Beam;    //BeaM
                        }

                    }


                }



                #endregion


            }



        }


        public class SapRevitFrame
        {

            #region prop

            public Frame sapFrame { get; set; }
            public Frame revitFrame { get; set; }

            public string mappedBy { get; set; }

            #endregion

            public SapRevitFrame()
            {
                sapFrame = new Frame();
                revitFrame = new Frame();

                mappedBy = null;
            }
            public SapRevitFrame(Frame sapFrame, Frame revitFrame)
            {
                this.sapFrame = sapFrame;
                this.revitFrame = revitFrame;
                mappedBy = null;
            }
            public static List<SapRevitFrame> getMappedSapRevit(List<Frame> sapFrames, List<Frame> revitFrames, double margin)
            {
                List<SapRevitFrame> allResults = new List<SapRevitFrame>();


                List<string> sapIDS = sapFrames.Select(x => x.ID).ToList();
                List<string> sapIDParameter = revitFrames.Select(x => x.ID).ToList();


                List<Frame> sapElementsMappedByIDS = sapFrames.Where(x => sapIDParameter.Contains(x.ID)).OrderBy(x => x.ID).ToList();
                List<Frame> revitElementsMappedByIDS = revitFrames.Where(x => sapIDS.Contains(x.ID)).OrderBy(x => x.ID).ToList();



                for (int i = 0; i < sapElementsMappedByIDS.Count; i++)
                {

                    SapRevitFrame result = new SapRevitFrame(sapElementsMappedByIDS[i], revitElementsMappedByIDS[i]);
                    result.mappedBy = "ID";
                    allResults.Add(result);


                    //double lengthDifferencePersentage = Math.Abs(revitElementsMappedByIDS[i].Length - sapElementsMappedByIDS[i].Length) / sapElementsMappedByIDS[i].Length;
                    //if (lengthDifferencePersentage < margin)
                    //{
                    //    SapRevitFrame result = new SapRevitFrame(sapElementsMappedByIDS[i], revitElementsMappedByIDS[i]);
                    //    result.mappedBy = "ID";
                    //    allResults.Add(result);
                    //}

                    //else
                    //{
                    //    sapElementsMappedByIDS.Remove(sapElementsMappedByIDS[i]);
                    //    revitElementsMappedByIDS.Remove(revitElementsMappedByIDS[i]);
                    //}


                }

                List<Frame> sapElementsNotMappedByIDS = sapFrames.Except(sapElementsMappedByIDS).ToList();
                List<Frame> revitElementsNotMappedByIDS = revitFrames.Except(revitElementsMappedByIDS).ToList();


                foreach (var sapFrame in sapElementsNotMappedByIDS)
                {
                    SapRevitFrame result = new SapRevitFrame();
                    Dictionary<Frame, double> dictionary = new Dictionary<Frame, double>();
                    List<double> lengthDiffernce = new List<double>();
                    result.sapFrame = sapFrame;
                    result.revitFrame = null;



                    foreach (Frame revitFrame in revitElementsNotMappedByIDS)
                    {
                        double lengthDifferencePersentage = Math.Abs(revitFrame.Length - sapFrame.Length) / sapFrame.Length;
                        if (lengthDifferencePersentage < margin)
                        {
                            dictionary.Add(revitFrame, Point.getDistance(sapFrame.MidPoint, revitFrame.MidPoint));
                        }

                    }


                    #region getMinValueFromDictionary
                    double mainValue = 1234567891011;
                    if (dictionary.Count != 0)
                    {
                        mainValue = dictionary.Values.ToList().Min();
                    }


                    #endregion


                    if (mainValue < margin * sapFrame.Length)
                    {
                        int index = dictionary.Values.ToList().IndexOf(dictionary.Values.Min());
                        result.revitFrame = dictionary.Keys.ToList()[index];
                        result.mappedBy = "MID";
                        allResults.Add(result);
                    }

                    else
                    {
                        result.revitFrame = null;
                        allResults.Add(result);
                    }

                Next:
                    continue;
                }


                //foreach (var sapFrame in sapFrames)
                //{
                //    SapRevitFrame result = new SapRevitFrame();
                //    Dictionary<Frame, double> dictionary = new Dictionary<Frame, double>();
                //    List<double> lengthDiffernce = new List<double>();
                //    result.sapFrame = sapFrame;
                //    result.revitFrame = null;



                //    foreach (Frame revitFrame in revitFrames)
                //    {
                //        if (revitFrame.ID == sapFrame.ID)
                //        {
                //            //double lengthDifferencePersentage = Math.Abs(revitFrame.Length - sapFrame.Length) / sapFrame.Length;
                //            //if (lengthDifferencePersentage < margin)
                //            //{
                //            //    result.revitFrame = revitFrame;
                //            //    result.mappedBy = "ID";
                //            //    allResults.Add(result);
                //            //    goto Next;

                //            //}

                //            result.revitFrame = revitFrame;
                //            result.mappedBy = "ID";
                //            allResults.Add(result);
                //            goto Next;



                //        }

                //        else
                //        {

                //            double lengthDifferencePersentage = Math.Abs(revitFrame.Length - sapFrame.Length) / sapFrame.Length;
                //            if (lengthDifferencePersentage < margin)
                //            {
                //                dictionary.Add(revitFrame, Point.getDistance(sapFrame.MidPoint, revitFrame.MidPoint));
                //            }

                //        }
                //    }


                //    #region getMinValueFromDictionary
                //    double mainValue = 1234567891011;
                //    if (dictionary.Count != 0)
                //    {
                //        mainValue = dictionary.Values.ToList().Min();
                //    }


                //    #endregion


                //    if (mainValue < margin * sapFrame.Length)
                //    {
                //        int index = dictionary.Values.ToList().IndexOf(dictionary.Values.Min());
                //        result.revitFrame = dictionary.Keys.ToList()[index];
                //        result.mappedBy = "MID";
                //        allResults.Add(result);
                //    }

                //    else
                //    {
                //        result.revitFrame = null;
                //        allResults.Add(result);
                //    }

                //Next:
                //    continue;
                //}

                return allResults;

            }
            public static List<SapRevitFrame> getMappedSapRevitByIdOnly(List<Frame> sapFrames, List<Frame> revitFrames)
            {
                List<SapRevitFrame> allResults = new List<SapRevitFrame>();

                foreach (var sapFrame in sapFrames)
                {
                    SapRevitFrame result = new SapRevitFrame();
                    Dictionary<Frame, double> dictionary = new Dictionary<Frame, double>();
                    List<double> lengthDiffernce = new List<double>();
                    result.sapFrame = sapFrame;
                    result.revitFrame = null;

                    foreach (Frame revitFrame in revitFrames)
                    {
                        if (revitFrame.ID == sapFrame.ID)
                        {
                            result.revitFrame = revitFrame;
                            result.mappedBy = "ID";
                            goto Next;

                        }

                    }

                Next:
                    allResults.Add(result);
                    continue;
                }

                return allResults;

            }
            public static void getUnMappedEelements(List<SapRevitFrame> sapRevitFrames, List<Frame> sapFrames, List<Frame> revitFrames, ref List<Frame> unMappedSapFrames, ref List<Frame> UnMappedRevitFrames)
            {

                unMappedSapFrames = sapRevitFrames.Where(x => x.revitFrame == null).ToList().Select(x => x.sapFrame).ToList();
                UnMappedRevitFrames = revitFrames.Except(sapRevitFrames.Select(x => x.revitFrame).ToList()).ToList();
            }

        }

        #endregion

        #region other_delerations 

        public enum Section_type
        {
            Pipe_H,
            Pipe_B,
            Pipe_T,

            Tube_H,
            Tube_B,
            Tube_T,

            HEA,
            HEAA,
            HEB,
            HEM,
            IPE,
            IPEA,
            IPEO,

            I_B,         // BUS
            I_T,

            C_H,

            DblAngle,
            EQAngle,
            UQAngle,

            I_SECTION_B,
            I_SECTION_H,
            Hollow_box_B,
            Hollow_box_H,
            Channel_B,
            Channel_H,
            D_Angle_B,
            D_Angle_H,
            S_Angle_B,
            S_Angle_H,
            T_Section_B,
            T_Section_H,
            Hollow_pipe_B,
            Hollow_pipe_H,
            I_Tapered,
            Box_Tapered,
            Pipe_Tapered,
            Solid_pipe,
            Solid_box,
            Concrete_rect,
            Concrete_circ,
            z_cold_Section,
            Channel_cold,
            Box_cold,
            Section_Designer,
            Dummy,
        }

        public enum SectType
        {
            BM_BUS_I,
            BM_HEA,
            BM_HEAA,
            BM_HEB,
            BM_HEM,
            BM_IPE,
            BM_IPEA,
            BM_IPEO,
            BM_L_EQUAL,
            BM_2L_EQUALANGLE,
            BM_L_UNEQUAL,
            BM_PIPE,
            BM_TUBE,
            BM_UPN,
            BM_BUS_I_Tapered,
            BM_BUS_PIPE_Tapered,
            BM_BUS_TUBE_Tapered,
            BM_W,
            BM_UB,
            BM_UC,

            COL_BUS_I,
            COL_HEA,
            COL_HEAA,
            COL_HEB,
            COL_HEM,
            COL_IPE,
            COL_IPEA,
            COL_IPEO,
            COL_L_EQUAL,
            COL_L_UNEQUAL,
            COL_2L_EQUALANGLE,
            COL_PIPE,
            COL_TUBE,
            COL_UPN,
            COL_BUS_I_Tapered,
            COL_BUS_PIPE_Tapered,
            COL_BUS_TUBE_Tapered,
            COL_W,
            COL_UB,
            COL_UC
        }



        public enum MemberType

        {
            Beam = 0,
            VLBracing = 1,
            Column = 2,
            SlanttedColumn = 3,

        }



        #endregion



    
}