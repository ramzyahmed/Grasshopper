using Autodesk.Revit.Creation;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using OpenSTAADUI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;


namespace BIM_Integration.STAAD
{
   public class StaadUtils
    {
        public static OpenSTAAD staadobject = Marshal.GetActiveObject("StaadPro.OpenStaad") as OpenSTAAD;
        public static IOSGeometryUI gometry = staadobject.Geometry;
        public static IOSCommandsUI command = staadobject.Command;
        public static IOSDesignUI design = staadobject.Design;
        public static IOSLoadUI loads = staadobject.Load;
        public static IOSOutputUI output = staadobject.Output;
        public static IOSPropertyUI property = staadobject.Property;
        public static IOSSupportUI supportUI = staadobject.Support;
        public static IOSTableUI tabels = staadobject.Table;
        public static IOSViewUI views = staadobject.View;
        public static IStaadProWindow window = staadobject.Window;



        
        public static void getUnitFactorsFactorFromStaad(ref double forceFactor, ref double momentFactor, ref double LengthFactor)
        {

            int base_unit = staadobject.GetBaseUnit();   // english = 1, metric = 2
            if (base_unit == 1)
            {
                forceFactor = 4.44822162;
                momentFactor = 0.112985 * 1000;
                LengthFactor = 25.4;
            }

            
            else
            {
                forceFactor = 1;
                momentFactor = 1000.0;
                LengthFactor = 1000.0;
            }

        }

        public static void getUnitFactorsFactorToStaad(ref double forceFactor, ref double momentFactor, ref double LengthFactor)
        {

            int base_unit = staadobject.GetBaseUnit();   // english = 1, metric = 2
            if (base_unit == 1)
            {
                forceFactor = 4.44822162;
                momentFactor = 0.112985;
                LengthFactor = 25.4 / 1000;
            }


            else
            {
                forceFactor = 1;
                momentFactor = 1.0 / 1000.0;
                LengthFactor = 1.0 / 1000.0;
            }

        }

        public static List<int> getStaadIDS()
        {

            int count = gometry.GetMemberCount();
            object beamsList = new int[count];

            gometry.GetBeamList(ref beamsList);

            return ((int[])beamsList).ToList();

        }

        public static List<int> getStaadSelectedFrameIDS()

        {
            List<int> result = new List<int>();

            int noOfSelectedMembers = gometry.GetNoOfSelectedBeams();
            object FrameIds = new int[noOfSelectedMembers];
            gometry.GetSelectedBeams(ref FrameIds, 1);

            for (int i = 0; i < noOfSelectedMembers; i++)
            {
                result.Add(((int[])FrameIds)[i]);

            }





            return result;
        }

        public static List<Frame> getStaadFrames(ref List<Beam> beams, ref List<Column> columns)
        {
            List<int> frameIds = getStaadIDS();

            if (getStaadSelectedFrameIDS().Count > 0)
            {
                frameIds = getStaadSelectedFrameIDS();
            }

            List<Frame> staadFrames = new List<Frame>();

            #region defintions 

            object x1, y1, z1, x2, y2, z2;
            x1 = y1 = z1 = x2 = y2 = z2 = 0.0;

            double lengthFactor = 0;
            double forceFactor = 0;
            double momentFactor = 0;
            getUnitFactorsFactorFromStaad(ref forceFactor, ref momentFactor, ref lengthFactor);


            object OstPointID = 0;


            object OendPointid = 0;

            double rotation = 0;

            Dictionary<int, String> shapeDic = new Dictionary<int, string>();
            Dictionary<int, String> secNameDic = new Dictionary<int, string>();
            Dictionary<int, String> LibDic = new Dictionary<int, string>();
            Dictionary<int, bool> isBUSDIC = new Dictionary<int, bool>();

            #endregion


            List<int> sections = getSections();

            getDictionaries(lengthFactor, sections, ref shapeDic, ref secNameDic, ref LibDic, ref isBUSDIC);

            foreach (int id in frameIds)
            {

                Frame frame = new Frame();

                #region get POints and length 

                gometry.GetMemberIncidence(id, ref OstPointID, ref OendPointid);

                gometry.GetNodeCoordinates(OstPointID, ref x1, ref y1, ref z1);
                gometry.GetNodeCoordinates(OendPointid, ref x2, ref y2, ref z2);




                string stPointID = OstPointID.ToString();
                string endPointID = OendPointid.ToString();



                frame.ID = id.ToString();
                frame.StartPoint = new Point(stPointID, (double)x1 * lengthFactor, (double)z1 * lengthFactor * -1, (double)y1 * lengthFactor);
                frame.EndPoint = new Point(endPointID, (double)x2 * lengthFactor, (double)z2 * lengthFactor * -1, (double)y2 * lengthFactor);
                frame.MidPoint = Point.getMidPoint(frame.StartPoint, frame.EndPoint);
                frame.Length = System.Numerics.Vector3.Distance(Point.ToVector(frame.StartPoint), Point.ToVector(frame.EndPoint));
                #endregion

                #region get rotation angle 

                rotation = property.GetBetaAngle(id);


                if (frame.MemberType == MemberType.SlanttedColumn || frame.MemberType == MemberType.Beam || frame.MemberType == MemberType.VLBracing)
                {
                    frame.Beta = -rotation;
                }


                else
                {
                    frame.Beta = rotation + 90.00;
                }

                #endregion

                #region get sectionName and shapetype and bus and library


                int sectionRefNo = property.GetBeamSectionPropertyRefNo(id);
                int sectionTypeNo = property.GetBeamSectionPropertyTypeNo(id);

                try
                {
                    frame.Shape = shapeDic[sectionRefNo];
                    frame.Sec = secNameDic[sectionRefNo];
                    frame.Lib = LibDic[sectionRefNo];
                    frame.IsBUS = isBUSDIC[sectionRefNo];

                    #region getMaterial name 

                    frame.Material = property.GetBeamMaterialName(id);

                    #endregion
                }

                catch (Exception)
                {
                    frame.Sec = "null";
                    goto EE;
                }


            #endregion


            EE:
                Point.getMemberType(frame);
                staadFrames.Add(frame);

                if (frame.MemberType == MemberType.Column || frame.MemberType == MemberType.SlanttedColumn)
                {

                    Column CL = new Column(frame.ID, frame.StartPoint, frame.EndPoint, frame.MidPoint, frame.Beta, frame.Shape, frame.SecType, frame.Sec, frame.IsBUS, frame.IsCable, frame.IsLink, frame.IsConc, frame.RvtName, frame.Lib, frame.Material, frame.Group, frame.Length, frame.MemberType, frame.rowNo);
                    columns.Add(CL);
                }

                else
                {
                    Beam BM = new Beam(frame.ID, frame.StartPoint, frame.EndPoint, frame.MidPoint, frame.Beta, frame.Shape, frame.SecType, frame.Sec, frame.IsBUS, frame.IsCable, frame.IsLink, frame.IsConc, frame.RvtName, frame.Lib, frame.Material, frame.Group, frame.Length, frame.MemberType, frame.rowNo);
                    beams.Add(BM);
                }



            }

            return staadFrames;
        }

        public static List<int> getSections()
        {

            // get count of sections 
            var sections_count = property.GetSectionPropertyCount();

            // get section refernces list
            List<int> sections = new List<int>();
            List<string> section_names = new List<string>();
            object sectionName = "";

            // 20 + count (in order to avoid unsequential numbers of sections)
            for (int i = 1; i <= sections_count + 20; i++)
            {

                property.GetSectionPropertyName(i, ref sectionName);
                if (sectionName != "")
                {
                    sections.Add(i);
                    section_names.Add((string)sectionName);
                }
            }

            return sections;

        }

        public static void getDictionaries(double lengthFactor, List<int> sections, ref Dictionary<int, string> shapeDic, ref Dictionary<int, string> secNameDic, ref Dictionary<int, string> LibDic, ref Dictionary<int, bool> isBUSDIC)
        {


            #region get lists of sections

            List<int> BoxH = new List<int>();
            List<int> BoxB = new List<int>();
            List<int> BoxT = new List<int>();
            List<int> BoxUpt = new List<int>();

            List<int> PipeH = new List<int>();
            List<int> PipeB = new List<int>();
            List<int> PipeT = new List<int>();
            List<int> PipeUpt = new List<int>();


            List<int> IH = new List<int>();
            List<int> IB = new List<int>();
            List<int> IT = new List<int>();
            List<int> IUptWideFlange = new List<int>();
            List<int> IUptISections = new List<int>();



            List<int> Angles = new List<int>();
            List<int> DblAngles = new List<int>();
            List<int> StarAngles = new List<int>();



            List<int> dummySections = new List<int>();



            List<int> channel = new List<int>();




            for (int i = 0; i < sections.Count; i++)
            {

                var refNo = property.GetSectionPropertyType(sections[i]);


                // box H 
                if (refNo == 654 || refNo == 650)
                {
                    BoxH.Add(sections[i]);
                }
                // box B
                else if (refNo == 651)
                {
                    BoxB.Add(sections[i]);
                }
                // box UPT
                else if (refNo == 696)
                {
                    BoxUpt.Add(sections[i]);
                }




                // pipe H 
                else if (refNo == 655 || refNo == 660)
                {
                    PipeH.Add(sections[i]);
                }

                // pipe B
                else if (refNo == 661)
                {
                    PipeB.Add(sections[i]);
                }

                // pipe UPT
                else if (refNo == 695)
                {
                    PipeUpt.Add(sections[i]);
                }





                // box and pipe Tapered
                else if (refNo == 675)
                {
                    object secName = "aaaa";
                    property.GetSectionPropertyName(sections[i], ref secName);

                    if (secName.ToString().Contains("SQU"))
                    {
                        BoxT.Add(sections[i]);
                    }

                    else if (secName.ToString().Contains("Round"))
                    {
                        PipeT.Add(sections[i]);
                    }


                }





                // IBeam H
                else if (refNo == 610)
                {
                    IH.Add(sections[i]);
                }


                // IBeam B and Tapered
                else if (refNo == 680)
                {
                    object dimentions = new double[24];
                    object varPropType = "aaa";
                    int n = 0;

                    n = property.GetSectionPropertyValuesEx(sections[i], ref varPropType, ref dimentions);
                    double[] dimentions2 = (double[])dimentions;

                    if (dimentions2[0] == dimentions2[2])
                    {
                        IB.Add(sections[i]);
                    }

                    else
                    {
                        IT.Add(sections[i]);
                    }

                }


                // ibeam wideflange 
                else if (refNo == 690)
                {
                    IUptWideFlange.Add(sections[i]);
                }


                // ibeam isection
                else if (refNo == 698)
                {
                    IUptISections.Add(sections[i]);

                }



                // single angle
                else if (refNo == 640)
                {
                    Angles.Add(sections[i]);
                }

                // dbl Angle btb
                else if (refNo == 643 || refNo == 642)
                {
                    DblAngles.Add(sections[i]);
                }

                // dbl start shape
                else if (refNo == 646)
                {
                    StarAngles.Add(sections[i]);
                }



                else if (refNo == 630)
                {
                    channel.Add(sections[i]);
                }


                else
                {
                    dummySections.Add(sections[i]);
                }

            }


            #endregion


            #region Create the dictionaries 
            foreach (var sec in sections)
            {


                #region definitions 
                String shape = null;
                string secName = null;
                string libraryName = null;
                bool isBUS = false;


                object orirginalName = "";
                property.GetSectionPropertyName(sec, ref orirginalName);


                object Odimentions = new double[24];
                double[] dimentions = null;
                object varPropType = "aaa";
                int ret = 0;

                #endregion


                #region set Library 

                int c = property.GetSectionPropertyCountry(sec);
                property.GetSectionPropertyCountry(c);
                if (c == 1) libraryName = "American";
                if (c == 2) libraryName = ("Australian");
                if (c == 3) libraryName = "British";
                if (c == 4) libraryName = "Canadian";
                if (c == 5) libraryName = "Chinese";
                if (c == 6) libraryName = "Dutch";
                if (c == 7) libraryName = "European";
                if (c == 8) libraryName = "French";
                if (c == 9) libraryName = "German";
                if (c == 10) libraryName = "Indian";
                if (c == 11) libraryName = ("Japanese");
                if (c == 12) libraryName = ("Russian");
                if (c == 13) libraryName = ("SouthAfrican");
                if (c == 14) libraryName = ("Spanish");
                if (c == 15) libraryName = ("Venezuelan");
                if (c == 16) libraryName = ("Korean");
                if (c == 17) libraryName = "Aluminum";
                if (c == 18) libraryName = "USColdformed";
                if (c == 19) libraryName = "ISColdformed";

                if (libraryName == null)
                {
                    libraryName = "N.A";
                }


                #endregion

                if (PipeH.Contains(sec))
                {
                    shape = "Pipe";
                    isBUS = false;
                    ret = property.GetSectionPropertyValuesEx(sec, ref varPropType, ref Odimentions);
                    dimentions = (double[])Odimentions;

                    double diameter = Math.Round(dimentions[1] * lengthFactor, 1);
                    double thick = Math.Round(dimentions[2] * lengthFactor, 1);

                    secName = "PIPE " + diameter + "X" + thick;


                }

                else if (PipeB.Contains(sec))
                {
                    shape = "Pipe";
                    isBUS = true;
                    ret = property.GetSectionPropertyValuesEx(sec, ref varPropType, ref Odimentions);
                    dimentions = (double[])Odimentions;

                    double Odiameter = Math.Round(dimentions[4] * lengthFactor, 1);
                    double Idiameter = Math.Round(dimentions[5] * lengthFactor, 1);
                    double thick = Math.Round(((Odiameter - Idiameter) / 2), 1);

                    secName = "PIPE " + Odiameter + "X" + thick;
                }

                else if (PipeT.Contains(sec))
                {
                    shape = "VARPipe";
                    isBUS = true;
                    ret = property.GetSectionPropertyValuesEx(sec, ref varPropType, ref Odimentions);
                    dimentions = (double[])Odimentions;

                    double stDiameter = Math.Round(dimentions[4] * lengthFactor, 1);
                    double endDiameter = Math.Round(dimentions[5] * lengthFactor, 1);
                    double thick = Math.Round(dimentions[6] * lengthFactor, 1);

                    double largeDiameter = 0;
                    double smallDiameter = 0;

                    if (stDiameter > endDiameter)
                    {
                        largeDiameter = stDiameter;
                        smallDiameter = endDiameter;
                    }

                    else
                    {
                        largeDiameter = endDiameter;
                        smallDiameter = stDiameter;

                    }

                    secName = "PIPE " + largeDiameter + "/" + smallDiameter + "X" + thick;

                }

                else if (PipeUpt.Contains(sec))
                {
                    shape = "Pipe";
                    isBUS = true;
                    ret = property.GetSectionPropertyValuesEx(sec, ref varPropType, ref Odimentions);
                    dimentions = (double[])Odimentions;
                    // check these values 
                    double Odiameter = Math.Round(dimentions[4] * lengthFactor, 1);
                    double Idiameter = Math.Round(dimentions[5] * lengthFactor, 1);
                    double thick = Math.Round(((Odiameter - Idiameter) / 2), 1);

                    secName = "Dummy";
                }



                else if (BoxH.Contains(sec))
                {
                    shape = "Box";
                    isBUS = false;
                    ret = property.GetSectionPropertyValuesEx(sec, ref varPropType, ref Odimentions);
                    dimentions = (double[])Odimentions;
                    double depth = Math.Round(dimentions[1] * lengthFactor, 1);
                    double width = Math.Round(dimentions[2] * lengthFactor, 1);
                    double thick = Math.Round(dimentions[3] * lengthFactor, 1);

                    secName = "TUBE " + depth + "X" + width + "X" + thick;

                }

                else if (BoxB.Contains(sec))
                {
                    shape = "Box";
                    isBUS = true;
                    ret = property.GetSectionPropertyValuesEx(sec, ref varPropType, ref Odimentions);
                    dimentions = (double[])Odimentions;

                    double depth = Math.Round(dimentions[6] * lengthFactor, 1);
                    double width = Math.Round(dimentions[5] * lengthFactor, 1);
                    double thick = Math.Round(dimentions[4] * lengthFactor, 1);

                    secName = "TUBE " + depth + "X" + width + "X" + thick;

                }

                else if (BoxT.Contains(sec))
                {
                    shape = "VARBox";
                    isBUS = true;
                    ret = property.GetSectionPropertyValuesEx(sec, ref varPropType, ref Odimentions);
                    dimentions = (double[])Odimentions;

                    // width = depth 
                    double stDepth = Math.Round(dimentions[4] * lengthFactor, 1);
                    double endDepth = Math.Round(dimentions[5] * lengthFactor, 1);
                    double thick = Math.Round(dimentions[6] * lengthFactor, 1);

                    double largeDepth = 0;
                    double smallDepth = 0;

                    if (stDepth > endDepth)
                    {
                        largeDepth = stDepth;
                        smallDepth = endDepth;
                    }

                    else
                    {
                        largeDepth = endDepth;
                        smallDepth = stDepth;

                    }


                    secName = "TUBE " + largeDepth + "/" + smallDepth + "X" + thick;
                }

                else if (BoxUpt.Contains(sec))
                {
                    shape = "Box";
                    isBUS = true;
                    ret = property.GetSectionPropertyValuesEx(sec, ref varPropType, ref Odimentions);
                    dimentions = (double[])Odimentions;

                    double depth = Math.Round(dimentions[6] * lengthFactor, 1);
                    double width = Math.Round(dimentions[5] * lengthFactor, 1);
                    double thick = Math.Round(dimentions[4] * lengthFactor, 1);

                    secName = "Dummy";

                }




                else if (IH.Contains(sec))
                {
                    shape = "I";
                    isBUS = false;
                    secName = orirginalName.ToString();

                }

                else if (IB.Contains(sec))
                {
                    shape = "I";
                    isBUS = true;
                    ret = property.GetSectionPropertyValuesEx(sec, ref varPropType, ref Odimentions);
                    dimentions = (double[])Odimentions;


                    double Depth = Math.Round(dimentions[0] * lengthFactor, 1);
                    double webThick = Math.Round(dimentions[1] * lengthFactor, 1);
                    double flgWidth = Math.Round(dimentions[3] * lengthFactor, 1);
                    double flgThick = Math.Round(dimentions[4] * lengthFactor, 1);

                    double WebDepth = Math.Round(Depth - 2 * flgThick, 1);

                    secName = "2FL. PL. " + flgWidth + "X" + flgThick + " +WEB PL. " + WebDepth + "X" + webThick;
                }

                else if (IT.Contains(sec))
                {
                    shape = "VARI";
                    isBUS = true;
                    ret = property.GetSectionPropertyValuesEx(sec, ref varPropType, ref Odimentions);
                    dimentions = (double[])Odimentions;


                    double stDepth = Math.Round(dimentions[0] * lengthFactor, 1);
                    double endDepth = Math.Round(dimentions[2] * lengthFactor, 1);

                    double webThick = Math.Round(dimentions[1] * lengthFactor, 1);
                    double flgWidth = Math.Round(dimentions[3] * lengthFactor, 1);
                    double flgThick = Math.Round(dimentions[4] * lengthFactor, 1);


                    double largeDepth = 0;
                    double smallDepth = 0;


                    if (stDepth > endDepth)
                    {
                        largeDepth = stDepth;
                        smallDepth = endDepth;
                    }

                    else
                    {
                        largeDepth = endDepth;
                        smallDepth = stDepth;

                    }


                    double largeWebDepth = Math.Round(largeDepth - 2 * flgThick, 1);
                    double smallWebDepth = Math.Round(smallDepth - 2 * flgThick, 1);



                    secName = "2FL. PL. " + flgWidth + "X" + flgThick + " +WEB PL. " + largeWebDepth + "/" + smallWebDepth + "X" + webThick;


                }

                else if (IUptWideFlange.Contains(sec))
                {
                    shape = "I";
                    isBUS = true;
                    ret = property.GetSectionPropertyValuesEx(sec, ref varPropType, ref Odimentions);
                    dimentions = (double[])Odimentions;


                    double Depth = Math.Round(dimentions[0] * lengthFactor, 1);
                    double webThick = Math.Round(dimentions[1] * lengthFactor, 1);
                    double flgWidth = Math.Round(dimentions[3] * lengthFactor, 1);
                    double flgThick = Math.Round(dimentions[4] * lengthFactor, 1);

                    double WebDepth = Math.Round(Depth - 2 * flgThick, 1);

                    secName = "Dummy";
                }

                else if (IUptISections.Contains(sec))
                {
                    shape = "I";
                    isBUS = true;
                    ret = property.GetSectionPropertyValuesEx(sec, ref varPropType, ref Odimentions);
                    dimentions = (double[])Odimentions;


                    double stDepth = Math.Round(dimentions[0] * lengthFactor, 1);
                    double endDepth = Math.Round(dimentions[2] * lengthFactor, 1);

                    double webThick = Math.Round(dimentions[1] * lengthFactor, 1);
                    double flgWidth = Math.Round(dimentions[3] * lengthFactor, 1);
                    double flgThick = Math.Round(dimentions[4] * lengthFactor, 1);


                    double largeDepth = 0;
                    double smallDepth = 0;


                    if (stDepth > endDepth)
                    {
                        largeDepth = stDepth;
                        smallDepth = endDepth;
                    }

                    else
                    {
                        largeDepth = endDepth;
                        smallDepth = stDepth;

                    }


                    double largeWebDepth = Math.Round(largeDepth - 2 * flgThick, 1);
                    double smallWebDepth = Math.Round(smallDepth - 2 * flgThick, 1);



                    secName = "Dummy";


                }




                else if (Angles.Contains(sec))
                {
                    shape = "Angle";
                    isBUS = false;
                    secName = orirginalName.ToString();

                }

                else if (DblAngles.Contains(sec))
                {
                    shape = "DblAngle";
                    isBUS = false;
                    secName = orirginalName.ToString();

                }

                else if (StarAngles.Contains(sec))
                {
                    shape = "DblAnglestar";
                    isBUS = false;
                    secName = orirginalName.ToString();
                }




                else if (channel.Contains(sec))
                {
                    shape = "Channel";
                    isBUS = false;
                    secName = orirginalName.ToString();
                }


                else
                {
                    shape = "dummy";
                    isBUS = true;
                    secName = "dummy";
                }


                shapeDic.Add(sec, shape);
                secNameDic.Add(sec, secName);
                isBUSDIC.Add(sec, isBUS);

                LibDic.Add(sec, libraryName);
            }
            #endregion



        }

        public static string createStaadBeam(Frame revitFrame, Autodesk.Revit.DB.Document doc, Dictionary<string, int> revitStaadSecDic)
        {
            #region defintions


            int ret;
            string id = revitFrame.RvtID;
            string staadSection = null;
            string materialName = null;



            double forceFactor = 0.0;
            double momentFactor = 0.0;
            double lengthFactor = 0.0;


            getUnitFactorsFactorToStaad(ref forceFactor, ref momentFactor, ref lengthFactor);

            #endregion


            try
            {


                int stId = gometry.AddNode(revitFrame.StartPoint.X * lengthFactor, revitFrame.StartPoint.Z * lengthFactor, revitFrame.StartPoint.Y * -1 * lengthFactor);

                int endId = gometry.AddNode(revitFrame.EndPoint.X * lengthFactor, revitFrame.EndPoint.Z * lengthFactor, revitFrame.EndPoint.Y * -1 * lengthFactor);

                int beamId = gometry.AddBeam(stId, endId);

                int staadSecNo = revitStaadSecDic[revitFrame.RvtName];

                property.AssignBeamProperty(beamId, staadSecNo);
                property.AssignBetaAngle(beamId, revitFrame.Beta);

                property.AssignMaterialToMember("STEEL", beamId);




                // set mark in revit 
                FamilyInstance instance = (FamilyInstance)doc.GetElement(revitFrame.revitEleId);
                using (Transaction trns = new Transaction(doc, "aa"))
                {

                    trns.Start();
                    Frame staadFrame = new Frame();
                    staadFrame.ID = beamId.ToString();
                    staadFrame.Material = materialName;



                    FailureHandlingOptions failOpt = trns.GetFailureHandlingOptions();
                    RevitUtils.setParametersSTAAD(instance, staadFrame);
                    trns.Commit();
                }


                return beamId.ToString();
            }


            catch (Exception ex2)
            {

                System.Diagnostics.StackTrace trace = new System.Diagnostics.StackTrace(ex2, true);
                TaskDialog.Show("Error", "error in beam no " + revitFrame.revitEleId.ToString() + ex2.Message + trace.GetFrame(trace.FrameCount - 1).GetFileLineNumber());
                return null;
            }


        }

        public static void createGroup(List<string> ids, string groupName)
        {

            int[] ids2 = new int[ids.Count];

            for (int i = 0; i < ids.Count; i++)
            {
                ids2[i] = int.Parse(ids[i]);
            }

            object ids3 = ids2;




            int noOfGroups = gometry.GetGroupCount(2);
            object groupNames = new string[noOfGroups];
            gometry.GetGroupNames(2, ref groupNames);




            if (!((string[])groupNames).Contains(groupName))
            {
                gometry.CreateGroupEx(2, groupName, ids.Count, ids3);
            }


            else
            {
                gometry.UpdateGroup(groupName, 0, ids.Count, ids3);
            }





        }

        public static void deleteGroup(string groupName)
        {


            gometry.DeleteGroup(groupName);

        }


        #region RelatedToSections
        public static Dictionary<string, int> createStaadSectionForBeams(List<Frame> revitFrames)
        {

            Dictionary<string, int> result = new Dictionary<string, int>();
            #region defintions

            int ret;
            char cc;
            int lastSpacecharacterIndex;
            string dimentionswithX;
            string sec;
            string[] dimentions;
            double lengthFactor = 1.00 / 1000.00;
            string revitSectionName = null;

            Dictionary<int, string> staadExistingSecNameDic = getModifiedNameOfAllDefinedSectionsInStaad();

            List<string> StaadExistingSecNames = staadExistingSecNameDic.Values.ToList();


            #endregion

            // group RevitFrames By its sectionName
            revitFrames = revitFrames.GroupBy(x => x.RvtName).Select(x => x.FirstOrDefault()).ToList();

            foreach (var revitFrame in revitFrames)
            {
                revitSectionName = revitFrame.RvtName;


                sec = StaadExistingSecNames.AsParallel().FirstOrDefault(x => x == revitSectionName);

                if (sec != null)
                {
                    revitFrame.Sec = sec;
                    result.Add(revitSectionName, staadExistingSecNameDic.FirstOrDefault(x => x.Value == sec).Key);
                    continue;
                }



                try
                {
                    switch (revitFrame.FamilyType)
                    {


                        #region I_B

                        case SectType.BM_BUS_I:

                            #region getDimentions 
                            int seperatorIndex = revitSectionName.LastIndexOf("+WEB");

                            string webString = revitSectionName.Substring(seperatorIndex + 4);
                            string flgString = revitSectionName.Substring(0, seperatorIndex - 1);


                            string[] webstring2 = webString.Split(("X".ToCharArray()));
                            string webThick = webstring2[1];


                            string[] flgString2 = flgString.Split(("X".ToCharArray()));
                            string flgThick = flgString2[1];


                            string consWebDepth = null;
                            string consFlgDepth = null;


                            consWebDepth = (Convert.ToDouble(new string((webstring2[0].SkipWhile(c => !char.IsDigit(c))
                                    .TakeWhile(c => char.IsDigit(c)).ToArray()
                                   ))) + 2 * Convert.ToDouble(flgThick)).ToString();




                            consFlgDepth = new string((flgString2[0].Substring(2).SkipWhile(c => !char.IsDigit(c))
                                   .TakeWhile(c => char.IsDigit(c)).ToArray()
                                   ));


                            #endregion

                            double[] dimentionsArray = new double[7];
                            dimentionsArray[0] = double.Parse(consWebDepth) * lengthFactor;
                            dimentionsArray[1] = double.Parse(webThick) * lengthFactor;
                            dimentionsArray[2] = double.Parse(consWebDepth) * lengthFactor;
                            dimentionsArray[3] = double.Parse(consFlgDepth) * lengthFactor;
                            dimentionsArray[4] = double.Parse(flgThick) * lengthFactor;
                            dimentionsArray[5] = double.Parse(consFlgDepth) * lengthFactor;
                            dimentionsArray[6] = double.Parse(flgThick) * lengthFactor;

                            object Odarray = dimentionsArray;
                            string staadSec = "2FL. PL. " + consFlgDepth + "X" + flgThick + " +WEB PL. " + consWebDepth + "X" + webThick;

                            if (StaadExistingSecNames.Contains(staadSec))
                            {
                                ret = staadExistingSecNameDic.FirstOrDefault(x => x.Value == staadSec).Key;
                            }
                            else
                            {
                                ret = property.CreateTaperedIProperty(dimentionsArray);
                                if (ret == 0 || ret < 0)
                                {
                                    ret = -1;
                                }
                            }

                            result.Add(revitSectionName, ret);
                            continue;
                            #endregion

                            #region I_H
                            string numeric = null;
                            string FamilyType = null;

                        case SectType.BM_HEA:
                            FamilyType = "HEA";
                            goto IBeamFamilySelection;



                        case SectType.BM_HEAA:
                            FamilyType = "HEAA";
                            goto IBeamFamilySelection;



                        case SectType.BM_HEB:
                            FamilyType = "HEB";
                            goto IBeamFamilySelection;


                        case SectType.BM_HEM:
                            FamilyType = "HEM";
                            goto IBeamFamilySelection;


                        case SectType.BM_IPE:
                            FamilyType = "IPE";
                            goto IBeamFamilySelection;


                        case SectType.BM_IPEA:
                            FamilyType = "IPEA";
                            goto IBeamFamilySelection;


                        case SectType.BM_IPEO:
                            FamilyType = "IPEO";
                            goto IBeamFamilySelection;

                        #region SelectFamilyType

                        IBeamFamilySelection:

                            switch (FamilyType)
                            {
                                case "HEAA":
                                    numeric = new String(revitSectionName.Where(Char.IsDigit).ToArray());
                                    staadSec = "HE" + numeric + "AA";


                                    if (StaadExistingSecNames.Contains(staadSec))
                                    {

                                        ret = staadExistingSecNameDic.FirstOrDefault(x => x.Value == staadSec).Key;
                                    }

                                    else
                                    {
                                        ret = property.CreateBeamPropertyFromTable(7, staadSec, 0, 0.0, 0.0);
                                    }
                                    result.Add(revitSectionName, ret);
                                    continue;



                                case "HEA":
                                    numeric = new String(revitSectionName.Where(Char.IsDigit).ToArray());
                                    staadSec = "HE" + numeric + "A";
                                    if (StaadExistingSecNames.Contains(staadSec))
                                    {

                                        ret = staadExistingSecNameDic.FirstOrDefault(x => x.Value == staadSec).Key;
                                    }

                                    else
                                    {
                                        ret = property.CreateBeamPropertyFromTable(7, staadSec, 0, 0.0, 0.0);
                                    }
                                    result.Add(revitSectionName, ret);
                                    continue;

                                case "HEB":
                                    numeric = new String(revitSectionName.Where(Char.IsDigit).ToArray());
                                    staadSec = "HE" + numeric + "B";
                                    if (StaadExistingSecNames.Contains(staadSec))
                                    {

                                        ret = staadExistingSecNameDic.FirstOrDefault(x => x.Value == staadSec).Key;
                                    }

                                    else
                                    {
                                        ret = property.CreateBeamPropertyFromTable(7, staadSec, 0, 0.0, 0.0);
                                    }
                                    result.Add(revitSectionName, ret);
                                    continue;


                                case "HEM":
                                    numeric = new String(revitSectionName.Where(Char.IsDigit).ToArray());
                                    staadSec = "HE" + numeric + "M";
                                    if (StaadExistingSecNames.Contains(staadSec))
                                    {

                                        ret = staadExistingSecNameDic.FirstOrDefault(x => x.Value == staadSec).Key;
                                    }

                                    else
                                    {
                                        ret = property.CreateBeamPropertyFromTable(7, staadSec, 0, 0.0, 0.0);
                                    }
                                    result.Add(revitSectionName, ret);
                                    continue;



                                case "IPE":
                                    numeric = new String(revitSectionName.Where(Char.IsDigit).ToArray());
                                    staadSec = "IPE" + numeric;
                                    if (StaadExistingSecNames.Contains(staadSec))
                                    {

                                        ret = staadExistingSecNameDic.FirstOrDefault(x => x.Value == staadSec).Key;
                                    }

                                    else
                                    {
                                        ret = property.CreateBeamPropertyFromTable(7, staadSec, 0, 0.0, 0.0);
                                    }
                                    result.Add(revitSectionName, ret);
                                    continue;

                                case "IPEA":
                                    numeric = new String(revitSectionName.Where(Char.IsDigit).ToArray());
                                    staadSec = "IPEA" + numeric;
                                    if (StaadExistingSecNames.Contains(staadSec))
                                    {

                                        ret = staadExistingSecNameDic.FirstOrDefault(x => x.Value == staadSec).Key;
                                    }

                                    else
                                    {
                                        ret = property.CreateBeamPropertyFromTable(7, staadSec, 0, 0.0, 0.0);
                                    }
                                    result.Add(revitSectionName, ret);
                                    continue;

                                case "IPEO":
                                    numeric = new String(revitSectionName.Where(Char.IsDigit).ToArray());
                                    staadSec = "IPEO" + numeric;
                                    if (StaadExistingSecNames.Contains(staadSec))
                                    {

                                        ret = staadExistingSecNameDic.FirstOrDefault(x => x.Value == staadSec).Key;
                                    }

                                    else
                                    {
                                        ret = property.CreateBeamPropertyFromTable(7, staadSec, 0, 0.0, 0.0);
                                    }
                                    result.Add(revitSectionName, ret);
                                    continue;


                                default:

                                    result.Add(revitSectionName, -1);
                                    continue;
                            }

                        #endregion



                        #endregion

                        #region pipe
                        case SectType.BM_PIPE:
                            cc = revitSectionName.SkipWhile(c => !char.IsDigit(c)).ToList()[0];
                            lastSpacecharacterIndex = revitSectionName.IndexOf(cc);
                            dimentionswithX = revitSectionName.Substring(lastSpacecharacterIndex);
                            dimentions = dimentionswithX.Split(("X".ToCharArray()));

                            staadSec = dimentions[0] + "x" + dimentions[1] + "CHS";
                            string staadSec2 = "PIPE " + dimentions[0] + "X" + dimentions[1];

                            if (StaadExistingSecNames.Contains(staadSec))
                            {

                                ret = staadExistingSecNameDic.FirstOrDefault(x => x.Value == staadSec).Key;
                            }


                            else if (StaadExistingSecNames.Contains(staadSec2))
                            {
                                ret = staadExistingSecNameDic.FirstOrDefault(x => x.Value == staadSec2).Key;
                            }


                            else
                            {
                                ret = property.CreatePipePropertyFromTable(7, staadSec, 0, 0.0, 0.0);
                                if (ret == 0 || ret < 0)
                                {
                                    double outerDiameter = double.Parse(dimentions[0]) * lengthFactor;
                                    double InnerDiametr = outerDiameter - 2 * double.Parse(dimentions[1]) * lengthFactor;

                                    try
                                    {
                                        ret = property.CreatePipePropertyFromTable(7, "", -1, outerDiameter, InnerDiametr);
                                        if (ret == 0 || ret < 0)
                                        {
                                            ret = -1;
                                        }
                                    }

                                    catch (Exception ex2)
                                    {

                                        ret = -1;

                                    }


                                }

                            }

                            result.Add(revitSectionName, ret);
                            continue;



                        #endregion

                        #region BOX(TUBE)

                        case SectType.BM_TUBE:
                            cc = revitSectionName.SkipWhile(c => !char.IsDigit(c)).ToList()[0];
                            lastSpacecharacterIndex = revitSectionName.IndexOf(cc);
                            dimentionswithX = revitSectionName.Substring(lastSpacecharacterIndex);
                            dimentions = dimentionswithX.Split(("X".ToCharArray()));
                            int noOfX = revitSectionName.Count(a => a == 'X');
                            double thick = 0;


                            if (dimentions[0] == dimentions[1])
                            {
                                staadSec = dimentions[0] + "x" + dimentions[2] + "SHS";
                            }
                            else
                            {
                                staadSec = dimentions[0] + "x" + dimentions[1] + "x" + dimentions[2] + "RHS";
                            }


                            staadSec2 = "TUBE " + dimentions[0] + "X" + dimentions[1] + "X" + dimentions[2];


                            if (StaadExistingSecNames.Contains(staadSec))
                            {

                                ret = staadExistingSecNameDic.FirstOrDefault(x => x.Value == staadSec).Key;
                            }


                            else if (StaadExistingSecNames.Contains(staadSec2))
                            {
                                ret = staadExistingSecNameDic.FirstOrDefault(x => x.Value == staadSec2).Key;
                            }



                            else
                            {
                                ret = property.CreateTubePropertyFromTable(7, staadSec, 0, 0.0, 0.0, 0.0);

                                if (ret == 0 || ret < 0)
                                {

                                    thick = double.Parse(dimentions[2]) * lengthFactor;


                                    double width = double.Parse(dimentions[1]) * lengthFactor;
                                    double depth = double.Parse(dimentions[0]) * lengthFactor;

                                    try
                                    {
                                        ret = property.CreateTubePropertyFromTable(7, "", -1, thick, width, depth);
                                    }

                                    catch (Exception ex2)
                                    {

                                        ret = -1;


                                    }


                                }
                            }

                            result.Add(revitSectionName, ret);
                            continue;



                        #endregion

                        #region UPN
                        case SectType.BM_UPN:

                            staadSec = new string(revitSectionName.ToCharArray()
                            .Where(c => !Char.IsWhiteSpace(c))
                            .ToArray());



                            if (StaadExistingSecNames.Contains(staadSec))
                            {

                                ret = staadExistingSecNameDic.FirstOrDefault(x => x.Value == staadSec).Key;
                            }

                            else
                            {
                                ret = property.CreateChannelPropertyFromTable(7, staadSec, 0, 0.0);
                                if (ret == 0 || ret < 0)
                                {
                                    ret = property.CreateChannelPropertyFromTable(8, staadSec, 0, 0.0);

                                }
                                if (ret == 0 || ret < 0)
                                {
                                    ret = property.CreateChannelPropertyFromTable(9, staadSec, 0, 0.0);

                                }

                                if (ret == 0 || ret < 0)
                                {
                                    ret = -1;
                                }
                            }

                            result.Add(revitSectionName, ret);
                            continue;



                        #endregion

                        #region W
                        case SectType.BM_W:
                            cc = revitFrame.RvtName.SkipWhile(c => !char.IsDigit(c)).ToList()[0];
                            lastSpacecharacterIndex = revitFrame.RvtName.IndexOf(cc);
                            dimentionswithX = revitFrame.RvtName.Substring(lastSpacecharacterIndex);
                            staadSec = "W" + dimentionswithX;


                            if (StaadExistingSecNames.Contains(staadSec))
                            {

                                ret = staadExistingSecNameDic.FirstOrDefault(x => x.Value == staadSec).Key;
                            }

                            else
                            {
                                ret = property.CreateBeamPropertyFromTable(1, staadSec, 0, 0.0, 0.0);

                                if (ret == 0 || ret < 0)
                                {
                                    ret = -1;
                                }
                            }

                            result.Add(revitSectionName, ret);
                            continue;

                        #endregion

                        #region UB
                        case SectType.BM_UB:
                            cc = revitFrame.RvtName.SkipWhile(c => !char.IsDigit(c)).ToList()[0];
                            lastSpacecharacterIndex = revitFrame.RvtName.IndexOf(cc);
                            dimentionswithX = revitFrame.RvtName.Substring(lastSpacecharacterIndex);
                            staadSec = "UB" + dimentionswithX;



                            if (StaadExistingSecNames.Contains(staadSec))
                            {

                                ret = staadExistingSecNameDic.FirstOrDefault(x => x.Value == staadSec).Key;
                            }

                            else
                            {
                                ret = property.CreateBeamPropertyFromTable(3, staadSec, 0, 0.0, 0.0);
                                if (ret <= 0)
                                {
                                    ret = -1;
                                }
                            }


                            result.Add(revitSectionName, ret);
                            continue;
                        #endregion

                        #region UC
                        case SectType.BM_UC:
                            cc = revitFrame.RvtName.SkipWhile(c => !char.IsDigit(c)).ToList()[0];
                            lastSpacecharacterIndex = revitFrame.RvtName.IndexOf(cc);
                            dimentionswithX = revitFrame.RvtName.Substring(lastSpacecharacterIndex);
                            staadSec = "UC" + dimentionswithX;

                            if (StaadExistingSecNames.Contains(staadSec))
                            {

                                ret = staadExistingSecNameDic.FirstOrDefault(x => x.Value == staadSec).Key;
                            }

                            else
                            {
                                ret = property.CreateBeamPropertyFromTable(3, staadSec, 0, 0.0, 0.0);
                                if (ret <= 0)
                                {
                                    ret = -1;
                                }
                            }

                            result.Add(revitSectionName, ret);
                            continue;
                        #endregion

                        #region I_T
                        case SectType.BM_BUS_I_Tapered:

                            #region getDimentions 

                            seperatorIndex = revitSectionName.LastIndexOf("+WEB");

                            webString = revitSectionName.Substring(seperatorIndex + 4);
                            flgString = revitSectionName.Substring(0, seperatorIndex - 1);


                            webstring2 = webString.Split(("X".ToCharArray()));
                            webThick = webstring2[1];


                            flgString2 = flgString.Split(("X".ToCharArray()));
                            flgThick = flgString2[1];

                            string largeWebDepth = null;
                            string smallWebDepth = null;


                            consFlgDepth = null;


                            // get largeDepth ,smalldepth of the web and constWidth of the flange
                            string[] webstring3 = webstring2[0].Split(("/".ToCharArray()));


                            largeWebDepth = (Convert.ToDouble(new string((webstring3[0].SkipWhile(c => !char.IsDigit(c))
                                   .TakeWhile(c => char.IsDigit(c)).ToArray()
                                  ))) + 2 * Convert.ToDouble(flgThick)).ToString();



                            smallWebDepth = (Convert.ToDouble(new string((webstring3[1].SkipWhile(c => !char.IsDigit(c))
                                   .TakeWhile(c => char.IsDigit(c)).ToArray()
                                  ))) + 2 * Convert.ToDouble(flgThick)).ToString();




                            consFlgDepth = new string((flgString2[0].Substring(2).SkipWhile(c => !char.IsDigit(c))
                                   .TakeWhile(c => char.IsDigit(c)).ToArray()
                                   ));

                            #endregion


                            staadSec = "2FL. PL. " + consFlgDepth + "X" + flgThick + " +WEB PL. " + largeWebDepth + "/" + smallWebDepth + "X" + webThick;


                            if (StaadExistingSecNames.Contains(staadSec))
                            {

                                ret = staadExistingSecNameDic.FirstOrDefault(x => x.Value == staadSec).Key;
                            }

                            else
                            {
                                dimentionsArray = new double[7];
                                dimentionsArray[0] = double.Parse(largeWebDepth) * lengthFactor;
                                dimentionsArray[1] = double.Parse(webThick) * lengthFactor;
                                dimentionsArray[2] = double.Parse(smallWebDepth) * lengthFactor;
                                dimentionsArray[3] = double.Parse(consFlgDepth) * lengthFactor;
                                dimentionsArray[4] = double.Parse(flgThick) * lengthFactor;
                                dimentionsArray[5] = double.Parse(consFlgDepth) * lengthFactor;
                                dimentionsArray[6] = double.Parse(flgThick) * lengthFactor;

                                Odarray = dimentionsArray;
                                ret = property.CreateTaperedIProperty(dimentionsArray);
                                if (ret <= 0)
                                {
                                    ret = -1;
                                }
                            }


                            result.Add(revitSectionName, ret);
                            continue;
                        #endregion

                        #region Pipe_T
                        case SectType.BM_BUS_PIPE_Tapered:

                            int lastXIndex = revitSectionName.LastIndexOf("X");
                            string Diameters = revitSectionName.Substring(0, lastXIndex);

                            int lastseperetorIndex = Diameters.LastIndexOf("/");



                            string LargDiamterValue = new string((Diameters.SkipWhile(c => !char.IsDigit(c))
                                 .TakeWhile(c => char.IsDigit(c)).ToArray()
                                ));

                            string SmallDiamaterValue = Diameters.Substring(lastseperetorIndex + 1);
                            string PipeThicknessValue = revitSectionName.Substring(lastXIndex + 1);


                            staadSec = "PIPE " + LargDiamterValue + "/" + SmallDiamaterValue + "X" + PipeThicknessValue;


                            if (StaadExistingSecNames.Contains(staadSec))
                            {

                                ret = staadExistingSecNameDic.FirstOrDefault(x => x.Value == staadSec).Key;
                            }

                            else
                            {
                                double largeDiameter = double.Parse(LargDiamterValue) * lengthFactor;
                                double smallDiamater = double.Parse(SmallDiamaterValue) * lengthFactor;
                                thick = double.Parse(PipeThicknessValue) * lengthFactor;

                                ret = property.CreateTaperedTubeProperty(0, largeDiameter, smallDiamater, thick);
                                if (ret <= 0)
                                {
                                    ret = -1;
                                }
                            }

                            result.Add(revitSectionName, ret);
                            continue;



                        #endregion

                        #region TUBE_T
                        case SectType.BM_BUS_TUBE_Tapered:
                            #region GetDimentions 

                            largeWebDepth = null;
                            smallWebDepth = null;
                            consFlgDepth = null;
                            webThick = null;


                            string largeFlgWidth = null;
                            string smallFlgWidth = null;
                            string constFlgWidth = null;
                            flgThick = null;

                            bool isFlgTapered = false;



                            int no_of_x = Frame.CountCharsUsingIndex(revitSectionName, 'X');
                            int no_of_slash = Frame.CountCharsUsingIndex(revitSectionName, '/');


                            //  800/500X20
                            if (no_of_x == 1 && no_of_slash == 1)
                            {

                                largeWebDepth = largeFlgWidth = new string(revitSectionName.SkipWhile(c => !char.IsDigit(c)).TakeWhile(c => char.IsDigit(c)).ToArray());
                                smallWebDepth = smallFlgWidth = new string(revitSectionName.Substring(revitSectionName.IndexOf("/")).SkipWhile(c => !char.IsDigit(c)).TakeWhile(c => char.IsDigit(c)).ToArray());



                                webThick = flgThick = new string(revitSectionName.Substring(revitSectionName.IndexOf('X')).SkipWhile(c => !char.IsDigit(c)).TakeWhile(c => char.IsDigit(c)).ToArray());


                            }

                            //  800/500X300X20
                            else if (no_of_x == 2 && no_of_slash == 1)
                            {

                                largeWebDepth = new string(revitSectionName.SkipWhile(c => !char.IsDigit(c)).TakeWhile(c => char.IsDigit(c)).ToArray());
                                smallWebDepth = new string(revitSectionName.Substring(revitSectionName.IndexOf("/")).SkipWhile(c => !char.IsDigit(c)).TakeWhile(c => char.IsDigit(c)).ToArray());


                                largeFlgWidth = smallFlgWidth = new string(revitSectionName.Substring(revitSectionName.IndexOf('X')).SkipWhile(c => !char.IsDigit(c)).TakeWhile(c => char.IsDigit(c)).ToArray());
                                webThick = flgThick = new string(revitSectionName.Substring(revitSectionName.LastIndexOf('X')).SkipWhile(c => !char.IsDigit(c)).TakeWhile(c => char.IsDigit(c)).ToArray());

                            }

                            //  800/500X400/300X20
                            else if (no_of_x == 2 && no_of_slash == 2)
                            {


                                isFlgTapered = true;
                                int smallFlgWidthIndex = Frame.IndexOfNth(revitSectionName, "/", 1);



                                largeWebDepth = new string(revitSectionName.SkipWhile(c => !char.IsDigit(c)).TakeWhile(c => char.IsDigit(c)).ToArray());
                                smallWebDepth = new string(revitSectionName.Substring(revitSectionName.IndexOf("/")).SkipWhile(c => !char.IsDigit(c)).TakeWhile(c => char.IsDigit(c)).ToArray());



                                largeFlgWidth = new string(revitSectionName.Substring(revitSectionName.IndexOf('X')).SkipWhile(c => !char.IsDigit(c)).TakeWhile(c => char.IsDigit(c)).ToArray());
                                smallFlgWidth = new string(revitSectionName.Substring(smallFlgWidthIndex).SkipWhile(c => !char.IsDigit(c)).TakeWhile(c => char.IsDigit(c)).ToArray());


                                webThick = flgThick = new string(revitSectionName.Substring(revitSectionName.LastIndexOf("X")).SkipWhile(c => !char.IsDigit(c)).TakeWhile(c => char.IsDigit(c)).ToArray());



                            }

                            #endregion
                            double largeDim;
                            double smallDim;
                            if (double.Parse(largeWebDepth) < double.Parse(largeFlgWidth))
                            {
                                largeDim = double.Parse(largeWebDepth) * lengthFactor;
                            }
                            else
                            {
                                largeDim = double.Parse(largeFlgWidth) * lengthFactor;
                            }

                            if (double.Parse(smallWebDepth) < double.Parse(smallFlgWidth))
                            {
                                smallDim = double.Parse(smallWebDepth) * lengthFactor;
                            }

                            else
                            {
                                smallDim = double.Parse(smallFlgWidth) * lengthFactor;
                            }


                            thick = double.Parse(webThick) * lengthFactor;



                            staadSec = "TUBE " + largeDim / lengthFactor + "/" + smallDim / lengthFactor + "X" + thick / lengthFactor;
                            if (StaadExistingSecNames.Contains(staadSec))
                            {

                                ret = staadExistingSecNameDic.FirstOrDefault(x => x.Value == staadSec).Key;
                            }

                            else
                            {
                                ret = property.CreateTaperedTubeProperty(5, largeDim, smallDim, thick);
                                if (ret <= 0)
                                {
                                    ret = -1;
                                }
                            }


                            result.Add(revitSectionName, ret);
                            continue;
                        #endregion

                        #region equal single ANgle

                        case SectType.BM_L_EQUAL:

                            string firstDimention = null;
                            string secondDimention = null;

                            noOfX = Frame.CountCharsUsingIndex(revitSectionName, 'X');
                            dimentionswithX = new string(revitSectionName.SkipWhile(c => !char.IsDigit(c)).ToArray());
                            firstDimention = new string(dimentionswithX.TakeWhile(c => char.IsDigit(c)).ToArray());



                            if (noOfX == 2)
                            {
                                staadSec = "L" + dimentionswithX;

                                if (StaadExistingSecNames.Contains(staadSec))
                                {

                                    ret = staadExistingSecNameDic.FirstOrDefault(x => x.Value == staadSec).Key;
                                }

                                else
                                {
                                    ret = property.CreateAnglePropertyFromTable(7, staadSec, 0, 0.0);
                                    if (ret <= 0)
                                    {
                                        ret = -1;
                                    }
                                }



                            }

                            else
                            {
                                string dimentionswithXDublicated = dimentionswithX.Insert(firstDimention.IndexOf("X") + 1, firstDimention + "X"); ;
                                staadSec = "L" + dimentionswithXDublicated;

                                if (StaadExistingSecNames.Contains(staadSec))
                                {

                                    ret = staadExistingSecNameDic.FirstOrDefault(x => x.Value == staadSec).Key;
                                }

                                else
                                {
                                    ret = property.CreateAnglePropertyFromTable(7, staadSec, 0, 0.0);
                                    if (ret <= 0)
                                    {
                                        ret = -1;
                                    }
                                }


                            }
                            result.Add(revitSectionName, ret);
                            continue;

                        #endregion

                        #region UNEqual single ANgle

                        case SectType.BM_L_UNEQUAL:

                            firstDimention = null;
                            secondDimention = null;

                            noOfX = Frame.CountCharsUsingIndex(revitSectionName, 'X');
                            dimentionswithX = new string(revitSectionName.SkipWhile(c => !char.IsDigit(c)).ToArray());
                            firstDimention = new string(dimentionswithX.TakeWhile(c => char.IsDigit(c)).ToArray());


                            staadSec = "L" + dimentionswithX;


                            if (StaadExistingSecNames.Contains(staadSec))
                            {

                                ret = staadExistingSecNameDic.FirstOrDefault(x => x.Value == staadSec).Key;
                            }

                            else
                            {
                                ret = property.CreateAnglePropertyFromTable(7, staadSec, 0, 0.0);
                                if (ret <= 0)
                                {
                                    ret = -1;
                                }
                            }




                            result.Add(revitSectionName, ret);
                            continue;

                        #endregion

                        #region  dbl ANgle

                        case SectType.BM_2L_EQUALANGLE:
                            noOfX = Frame.CountCharsUsingIndex(revitFrame.RvtName, 'X');
                            string revitSectionModified = revitFrame.RvtName.Substring(2);
                            dimentionswithX = new string(revitSectionModified.SkipWhile(c => !char.IsDigit(c)).ToArray());

                            int indexOfSecondX = Frame.IndexOfNth(dimentionswithX, "X", 0);
                            int lastXindex = revitSectionModified.LastIndexOf("X");

                            string angleThick = revitSectionModified.Substring(lastXindex + 1);

                            firstDimention = new string(dimentionswithX.TakeWhile(c => char.IsDigit(c)).ToArray());

                            staadSec = "L" + firstDimention + "X" + firstDimention + "X" + angleThick;



                            if (StaadExistingSecNames.Contains(staadSec))
                            {

                                ret = staadExistingSecNameDic.FirstOrDefault(x => x.Value == staadSec).Key;
                            }

                            else
                            {
                                ret = property.CreateAnglePropertyFromTable(7, staadSec, 3, 0.0);

                                if (ret <= 0)
                                {
                                    ret = -1;
                                }
                            }


                            result.Add(revitSectionName, ret);
                            continue;


                        #endregion


                        default:
                            result.Add(revitSectionName, -1);
                            continue;





                    }
                }


                catch (Exception)
                {


                    result.Add(revitSectionName, -1);
                    continue;

                }


            }

            #region hadling of revit section not found in staad 
            int dummySection = 0;

            dummySection = staadExistingSecNameDic.FirstOrDefault(x => x.Value == "PIPE 40X10").Key;

            if (dummySection == 0)
            {

                double outerDiameter = double.Parse("40") * lengthFactor;
                double InnerDiametr = outerDiameter - 2 * double.Parse("10") * lengthFactor;

                ret = property.CreatePipePropertyFromTable(7, "", -1, outerDiameter, InnerDiametr);
                if (ret == 0 || ret < 0)
                {
                    ret = -1;
                }

                staadExistingSecNameDic.Add(ret, "PIPE 40X10");
            }


            List<string> keysToUpdate = new List<string>();
            foreach (var kvp in result)
            {
                if (kvp.Value == -1)
                {
                    keysToUpdate.Add(kvp.Key);
                }
            }

            foreach (var key in keysToUpdate)
            {
                result[key] = staadExistingSecNameDic.FirstOrDefault(x => x.Value == "PIPE 40X10").Key;
            }



            #endregion

            return result;


        }


        public static Dictionary<string, int> createStaadSectionForColumns(List<Frame> revitFrames)
        {

            Dictionary<string, int> result = new Dictionary<string, int>();
            #region defintions

            int ret;
            char cc;
            int lastSpacecharacterIndex;
            string dimentionswithX;
            string sec;
            string[] dimentions;
            double lengthFactor = 1.00 / 1000.00;
            string revitSectionName = null;

            Dictionary<int, string> staadExistingSecNameDic = getModifiedNameOfAllDefinedSectionsInStaad();

            List<string> StaadExistingSecNames = staadExistingSecNameDic.Values.ToList();


            #endregion

            // group RevitFrames By its sectionName
            revitFrames = revitFrames.GroupBy(x => x.RvtName).Select(x => x.FirstOrDefault()).ToList();

            foreach (var revitFrame in revitFrames)
            {
                revitSectionName = revitFrame.RvtName;


                sec = StaadExistingSecNames.AsParallel().FirstOrDefault(x => x == revitSectionName);

                if (sec != null)
                {
                    revitFrame.Sec = sec;
                    result.Add(revitSectionName, staadExistingSecNameDic.FirstOrDefault(x => x.Value == sec).Key);
                    continue;
                }



                try
                {
                    switch (revitFrame.FamilyType)
                    {


                        #region I_B

                        case SectType.COL_BUS_I:

                            #region getDimentions 
                            int seperatorIndex = revitSectionName.LastIndexOf("+WEB");

                            string webString = revitSectionName.Substring(seperatorIndex + 4);
                            string flgString = revitSectionName.Substring(0, seperatorIndex - 1);


                            string[] webstring2 = webString.Split(("X".ToCharArray()));
                            string webThick = webstring2[1];


                            string[] flgString2 = flgString.Split(("X".ToCharArray()));
                            string flgThick = flgString2[1];


                            string consWebDepth = null;
                            string consFlgDepth = null;


                            consWebDepth = (Convert.ToDouble(new string((webstring2[0].SkipWhile(c => !char.IsDigit(c))
                                    .TakeWhile(c => char.IsDigit(c)).ToArray()
                                   ))) + 2 * Convert.ToDouble(flgThick)).ToString();




                            consFlgDepth = new string((flgString2[0].Substring(2).SkipWhile(c => !char.IsDigit(c))
                                   .TakeWhile(c => char.IsDigit(c)).ToArray()
                                   ));


                            #endregion

                            double[] dimentionsArray = new double[7];
                            dimentionsArray[0] = double.Parse(consWebDepth) * lengthFactor;
                            dimentionsArray[1] = double.Parse(webThick) * lengthFactor;
                            dimentionsArray[2] = double.Parse(consWebDepth) * lengthFactor;
                            dimentionsArray[3] = double.Parse(consFlgDepth) * lengthFactor;
                            dimentionsArray[4] = double.Parse(flgThick) * lengthFactor;
                            dimentionsArray[5] = double.Parse(consFlgDepth) * lengthFactor;
                            dimentionsArray[6] = double.Parse(flgThick) * lengthFactor;

                            object Odarray = dimentionsArray;
                            string staadSec = "2FL. PL. " + consFlgDepth + "X" + flgThick + " +WEB PL. " + consWebDepth + "X" + webThick;

                            if (StaadExistingSecNames.Contains(staadSec))
                            {
                                ret = staadExistingSecNameDic.FirstOrDefault(x => x.Value == staadSec).Key;
                            }
                            else
                            {
                                ret = property.CreateTaperedIProperty(dimentionsArray);
                                if (ret == 0 || ret < 0)
                                {
                                    ret = -1;
                                }
                            }

                            result.Add(revitSectionName, ret);
                            continue;
                            #endregion

                            #region I_H
                            string numeric = null;
                            string FamilyType = null;

                        case SectType.COL_HEA:
                            FamilyType = "HEA";
                            goto IBeamFamilySelection;



                        case SectType.COL_HEAA:
                            FamilyType = "HEAA";
                            goto IBeamFamilySelection;



                        case SectType.COL_HEB:
                            FamilyType = "HEB";
                            goto IBeamFamilySelection;


                        case SectType.COL_HEM:
                            FamilyType = "HEM";
                            goto IBeamFamilySelection;


                        case SectType.COL_IPE:
                            FamilyType = "IPE";
                            goto IBeamFamilySelection;


                        case SectType.COL_IPEA:
                            FamilyType = "IPEA";
                            goto IBeamFamilySelection;


                        case SectType.COL_IPEO:
                            FamilyType = "IPEO";
                            goto IBeamFamilySelection;

                        #region SelectFamilyType

                        IBeamFamilySelection:

                            switch (FamilyType)
                            {
                                case "HEAA":
                                    numeric = new String(revitSectionName.Where(Char.IsDigit).ToArray());
                                    staadSec = "HE" + numeric + "AA";


                                    if (StaadExistingSecNames.Contains(staadSec))
                                    {

                                        ret = staadExistingSecNameDic.FirstOrDefault(x => x.Value == staadSec).Key;
                                    }

                                    else
                                    {
                                        ret = property.CreateBeamPropertyFromTable(7, staadSec, 0, 0.0, 0.0);
                                    }
                                    result.Add(revitSectionName, ret);
                                    continue;



                                case "HEA":
                                    numeric = new String(revitSectionName.Where(Char.IsDigit).ToArray());
                                    staadSec = "HE" + numeric + "A";
                                    if (StaadExistingSecNames.Contains(staadSec))
                                    {

                                        ret = staadExistingSecNameDic.FirstOrDefault(x => x.Value == staadSec).Key;
                                    }

                                    else
                                    {
                                        ret = property.CreateBeamPropertyFromTable(7, staadSec, 0, 0.0, 0.0);
                                    }
                                    result.Add(revitSectionName, ret);
                                    continue;

                                case "HEB":
                                    numeric = new String(revitSectionName.Where(Char.IsDigit).ToArray());
                                    staadSec = "HE" + numeric + "B";
                                    if (StaadExistingSecNames.Contains(staadSec))
                                    {

                                        ret = staadExistingSecNameDic.FirstOrDefault(x => x.Value == staadSec).Key;
                                    }

                                    else
                                    {
                                        ret = property.CreateBeamPropertyFromTable(7, staadSec, 0, 0.0, 0.0);
                                    }
                                    result.Add(revitSectionName, ret);
                                    continue;


                                case "HEM":
                                    numeric = new String(revitSectionName.Where(Char.IsDigit).ToArray());
                                    staadSec = "HE" + numeric + "M";
                                    if (StaadExistingSecNames.Contains(staadSec))
                                    {

                                        ret = staadExistingSecNameDic.FirstOrDefault(x => x.Value == staadSec).Key;
                                    }

                                    else
                                    {
                                        ret = property.CreateBeamPropertyFromTable(7, staadSec, 0, 0.0, 0.0);
                                    }
                                    result.Add(revitSectionName, ret);
                                    continue;



                                case "IPE":
                                    numeric = new String(revitSectionName.Where(Char.IsDigit).ToArray());
                                    staadSec = "IPE" + numeric;
                                    if (StaadExistingSecNames.Contains(staadSec))
                                    {

                                        ret = staadExistingSecNameDic.FirstOrDefault(x => x.Value == staadSec).Key;
                                    }

                                    else
                                    {
                                        ret = property.CreateBeamPropertyFromTable(7, staadSec, 0, 0.0, 0.0);
                                    }
                                    result.Add(revitSectionName, ret);
                                    continue;

                                case "IPEA":
                                    numeric = new String(revitSectionName.Where(Char.IsDigit).ToArray());
                                    staadSec = "IPEA" + numeric;
                                    if (StaadExistingSecNames.Contains(staadSec))
                                    {

                                        ret = staadExistingSecNameDic.FirstOrDefault(x => x.Value == staadSec).Key;
                                    }

                                    else
                                    {
                                        ret = property.CreateBeamPropertyFromTable(7, staadSec, 0, 0.0, 0.0);
                                    }
                                    result.Add(revitSectionName, ret);
                                    continue;

                                case "IPEO":
                                    numeric = new String(revitSectionName.Where(Char.IsDigit).ToArray());
                                    staadSec = "IPEO" + numeric;
                                    if (StaadExistingSecNames.Contains(staadSec))
                                    {

                                        ret = staadExistingSecNameDic.FirstOrDefault(x => x.Value == staadSec).Key;
                                    }

                                    else
                                    {
                                        ret = property.CreateBeamPropertyFromTable(7, staadSec, 0, 0.0, 0.0);
                                    }
                                    result.Add(revitSectionName, ret);
                                    continue;


                                default:

                                    result.Add(revitSectionName, -1);
                                    continue;
                            }

                        #endregion



                        #endregion

                        #region pipe
                        case SectType.COL_PIPE:
                            cc = revitSectionName.SkipWhile(c => !char.IsDigit(c)).ToList()[0];
                            lastSpacecharacterIndex = revitSectionName.IndexOf(cc);
                            dimentionswithX = revitSectionName.Substring(lastSpacecharacterIndex);
                            dimentions = dimentionswithX.Split(("X".ToCharArray()));

                            staadSec = dimentions[0] + "x" + dimentions[1] + "CHS";
                            string staadSec2 = "PIPE " + dimentions[0] + "X" + dimentions[1];

                            if (StaadExistingSecNames.Contains(staadSec))
                            {

                                ret = staadExistingSecNameDic.FirstOrDefault(x => x.Value == staadSec).Key;
                            }


                            else if (StaadExistingSecNames.Contains(staadSec2))
                            {
                                ret = staadExistingSecNameDic.FirstOrDefault(x => x.Value == staadSec2).Key;
                            }


                            else
                            {
                                ret = property.CreatePipePropertyFromTable(7, staadSec, 0, 0.0, 0.0);
                                if (ret == 0 || ret < 0)
                                {
                                    double outerDiameter = double.Parse(dimentions[0]) * lengthFactor;
                                    double InnerDiametr = outerDiameter - 2 * double.Parse(dimentions[1]) * lengthFactor;

                                    try
                                    {
                                        ret = property.CreatePipePropertyFromTable(7, "", -1, outerDiameter, InnerDiametr);
                                        if (ret == 0 || ret < 0)
                                        {
                                            ret = -1;
                                        }
                                    }

                                    catch (Exception ex2)
                                    {

                                        ret = -1;

                                    }


                                }

                            }

                            result.Add(revitSectionName, ret);
                            continue;



                        #endregion

                        #region BOX(TUBE)

                        case SectType.COL_TUBE:
                            cc = revitSectionName.SkipWhile(c => !char.IsDigit(c)).ToList()[0];
                            lastSpacecharacterIndex = revitSectionName.IndexOf(cc);
                            dimentionswithX = revitSectionName.Substring(lastSpacecharacterIndex);
                            dimentions = dimentionswithX.Split(("X".ToCharArray()));
                            int noOfX = revitSectionName.Count(a => a == 'X');
                            double thick = 0;


                            if (dimentions[0] == dimentions[1])
                            {
                                staadSec = dimentions[0] + "x" + dimentions[2] + "SHS";
                            }
                            else
                            {
                                staadSec = dimentions[0] + "x" + dimentions[1] + "x" + dimentions[2] + "RHS";
                            }


                            staadSec2 = "TUBE " + dimentions[0] + "X" + dimentions[1] + "X" + dimentions[2];


                            if (StaadExistingSecNames.Contains(staadSec))
                            {

                                ret = staadExistingSecNameDic.FirstOrDefault(x => x.Value == staadSec).Key;
                            }


                            else if (StaadExistingSecNames.Contains(staadSec2))
                            {
                                ret = staadExistingSecNameDic.FirstOrDefault(x => x.Value == staadSec2).Key;
                            }



                            else
                            {
                                ret = property.CreateTubePropertyFromTable(7, staadSec, 0, 0.0, 0.0, 0.0);

                                if (ret == 0 || ret < 0)
                                {

                                    thick = double.Parse(dimentions[2]) * lengthFactor;


                                    double width = double.Parse(dimentions[1]) * lengthFactor;
                                    double depth = double.Parse(dimentions[0]) * lengthFactor;

                                    try
                                    {
                                        ret = property.CreateTubePropertyFromTable(7, "", -1, thick, width, depth);
                                    }

                                    catch (Exception ex2)
                                    {

                                        ret = -1;


                                    }


                                }
                            }

                            result.Add(revitSectionName, ret);
                            continue;



                        #endregion

                        #region UPN
                        case SectType.COL_UPN:

                            staadSec = new string(revitSectionName.ToCharArray()
                            .Where(c => !Char.IsWhiteSpace(c))
                            .ToArray());



                            if (StaadExistingSecNames.Contains(staadSec))
                            {

                                ret = staadExistingSecNameDic.FirstOrDefault(x => x.Value == staadSec).Key;
                            }

                            else
                            {
                                ret = property.CreateChannelPropertyFromTable(7, staadSec, 0, 0.0);
                                if (ret == 0 || ret < 0)
                                {
                                    ret = property.CreateChannelPropertyFromTable(8, staadSec, 0, 0.0);

                                }
                                if (ret == 0 || ret < 0)
                                {
                                    ret = property.CreateChannelPropertyFromTable(9, staadSec, 0, 0.0);

                                }

                                if (ret == 0 || ret < 0)
                                {
                                    ret = -1;
                                }
                            }

                            result.Add(revitSectionName, ret);
                            continue;



                        #endregion

                        #region W
                        case SectType.COL_W:
                            cc = revitFrame.RvtName.SkipWhile(c => !char.IsDigit(c)).ToList()[0];
                            lastSpacecharacterIndex = revitFrame.RvtName.IndexOf(cc);
                            dimentionswithX = revitFrame.RvtName.Substring(lastSpacecharacterIndex);
                            staadSec = "W" + dimentionswithX;


                            if (StaadExistingSecNames.Contains(staadSec))
                            {

                                ret = staadExistingSecNameDic.FirstOrDefault(x => x.Value == staadSec).Key;
                            }

                            else
                            {
                                ret = property.CreateBeamPropertyFromTable(1, staadSec, 0, 0.0, 0.0);

                                if (ret == 0 || ret < 0)
                                {
                                    ret = -1;
                                }
                            }

                            result.Add(revitSectionName, ret);
                            continue;

                        #endregion

                        #region UB
                        case SectType.COL_UB:
                            cc = revitFrame.RvtName.SkipWhile(c => !char.IsDigit(c)).ToList()[0];
                            lastSpacecharacterIndex = revitFrame.RvtName.IndexOf(cc);
                            dimentionswithX = revitFrame.RvtName.Substring(lastSpacecharacterIndex);
                            staadSec = "UB" + dimentionswithX;



                            if (StaadExistingSecNames.Contains(staadSec))
                            {

                                ret = staadExistingSecNameDic.FirstOrDefault(x => x.Value == staadSec).Key;
                            }

                            else
                            {
                                ret = property.CreateBeamPropertyFromTable(3, staadSec, 0, 0.0, 0.0);
                                if (ret <= 0)
                                {
                                    ret = -1;
                                }
                            }


                            result.Add(revitSectionName, ret);
                            continue;
                        #endregion

                        #region UC
                        case SectType.COL_UC:
                            cc = revitFrame.RvtName.SkipWhile(c => !char.IsDigit(c)).ToList()[0];
                            lastSpacecharacterIndex = revitFrame.RvtName.IndexOf(cc);
                            dimentionswithX = revitFrame.RvtName.Substring(lastSpacecharacterIndex);
                            staadSec = "UC" + dimentionswithX;

                            if (StaadExistingSecNames.Contains(staadSec))
                            {

                                ret = staadExistingSecNameDic.FirstOrDefault(x => x.Value == staadSec).Key;
                            }

                            else
                            {
                                ret = property.CreateBeamPropertyFromTable(3, staadSec, 0, 0.0, 0.0);
                                if (ret <= 0)
                                {
                                    ret = -1;
                                }
                            }

                            result.Add(revitSectionName, ret);
                            continue;
                        #endregion

                        #region I_T
                        case SectType.COL_BUS_I_Tapered:

                            #region getDimentions 

                            seperatorIndex = revitSectionName.LastIndexOf("+WEB");

                            webString = revitSectionName.Substring(seperatorIndex + 4);
                            flgString = revitSectionName.Substring(0, seperatorIndex - 1);


                            webstring2 = webString.Split(("X".ToCharArray()));
                            webThick = webstring2[1];


                            flgString2 = flgString.Split(("X".ToCharArray()));
                            flgThick = flgString2[1];

                            string largeWebDepth = null;
                            string smallWebDepth = null;


                            consFlgDepth = null;


                            // get largeDepth ,smalldepth of the web and constWidth of the flange
                            string[] webstring3 = webstring2[0].Split(("/".ToCharArray()));


                            largeWebDepth = (Convert.ToDouble(new string((webstring3[0].SkipWhile(c => !char.IsDigit(c))
                                   .TakeWhile(c => char.IsDigit(c)).ToArray()
                                  ))) + 2 * Convert.ToDouble(flgThick)).ToString();



                            smallWebDepth = (Convert.ToDouble(new string((webstring3[1].SkipWhile(c => !char.IsDigit(c))
                                   .TakeWhile(c => char.IsDigit(c)).ToArray()
                                  ))) + 2 * Convert.ToDouble(flgThick)).ToString();




                            consFlgDepth = new string((flgString2[0].Substring(2).SkipWhile(c => !char.IsDigit(c))
                                   .TakeWhile(c => char.IsDigit(c)).ToArray()
                                   ));

                            #endregion


                            staadSec = "2FL. PL. " + consFlgDepth + "X" + flgThick + " +WEB PL. " + largeWebDepth + "/" + smallWebDepth + "X" + webThick;


                            if (StaadExistingSecNames.Contains(staadSec))
                            {

                                ret = staadExistingSecNameDic.FirstOrDefault(x => x.Value == staadSec).Key;
                            }

                            else
                            {
                                dimentionsArray = new double[7];
                                dimentionsArray[0] = double.Parse(largeWebDepth) * lengthFactor;
                                dimentionsArray[1] = double.Parse(webThick) * lengthFactor;
                                dimentionsArray[2] = double.Parse(smallWebDepth) * lengthFactor;
                                dimentionsArray[3] = double.Parse(consFlgDepth) * lengthFactor;
                                dimentionsArray[4] = double.Parse(flgThick) * lengthFactor;
                                dimentionsArray[5] = double.Parse(consFlgDepth) * lengthFactor;
                                dimentionsArray[6] = double.Parse(flgThick) * lengthFactor;

                                Odarray = dimentionsArray;
                                ret = property.CreateTaperedIProperty(dimentionsArray);
                                if (ret <= 0)
                                {
                                    ret = -1;
                                }
                            }


                            result.Add(revitSectionName, ret);
                            continue;
                        #endregion

                        #region Pipe_T
                        case SectType.COL_BUS_PIPE_Tapered:

                            int lastXIndex = revitSectionName.LastIndexOf("X");
                            string Diameters = revitSectionName.Substring(0, lastXIndex);

                            int lastseperetorIndex = Diameters.LastIndexOf("/");



                            string LargDiamterValue = new string((Diameters.SkipWhile(c => !char.IsDigit(c))
                                 .TakeWhile(c => char.IsDigit(c)).ToArray()
                                ));

                            string SmallDiamaterValue = Diameters.Substring(lastseperetorIndex + 1);
                            string PipeThicknessValue = revitSectionName.Substring(lastXIndex + 1);


                            staadSec = "PIPE " + LargDiamterValue + "/" + SmallDiamaterValue + "X" + PipeThicknessValue;


                            if (StaadExistingSecNames.Contains(staadSec))
                            {

                                ret = staadExistingSecNameDic.FirstOrDefault(x => x.Value == staadSec).Key;
                            }

                            else
                            {
                                double largeDiameter = double.Parse(LargDiamterValue) * lengthFactor;
                                double smallDiamater = double.Parse(SmallDiamaterValue) * lengthFactor;
                                thick = double.Parse(PipeThicknessValue) * lengthFactor;

                                ret = property.CreateTaperedTubeProperty(0, largeDiameter, smallDiamater, thick);
                                if (ret <= 0)
                                {
                                    ret = -1;
                                }
                            }

                            result.Add(revitSectionName, ret);
                            continue;



                        #endregion

                        #region TUBE_T
                        case SectType.COL_BUS_TUBE_Tapered:
                            #region GetDimentions 

                            largeWebDepth = null;
                            smallWebDepth = null;
                            consFlgDepth = null;
                            webThick = null;


                            string largeFlgWidth = null;
                            string smallFlgWidth = null;
                            string constFlgWidth = null;
                            flgThick = null;

                            bool isFlgTapered = false;



                            int no_of_x = Frame.CountCharsUsingIndex(revitSectionName, 'X');
                            int no_of_slash = Frame.CountCharsUsingIndex(revitSectionName, '/');


                            //  800/500X20
                            if (no_of_x == 1 && no_of_slash == 1)
                            {

                                largeWebDepth = largeFlgWidth = new string(revitSectionName.SkipWhile(c => !char.IsDigit(c)).TakeWhile(c => char.IsDigit(c)).ToArray());
                                smallWebDepth = smallFlgWidth = new string(revitSectionName.Substring(revitSectionName.IndexOf("/")).SkipWhile(c => !char.IsDigit(c)).TakeWhile(c => char.IsDigit(c)).ToArray());



                                webThick = flgThick = new string(revitSectionName.Substring(revitSectionName.IndexOf('X')).SkipWhile(c => !char.IsDigit(c)).TakeWhile(c => char.IsDigit(c)).ToArray());


                            }

                            //  800/500X300X20
                            else if (no_of_x == 2 && no_of_slash == 1)
                            {

                                largeWebDepth = new string(revitSectionName.SkipWhile(c => !char.IsDigit(c)).TakeWhile(c => char.IsDigit(c)).ToArray());
                                smallWebDepth = new string(revitSectionName.Substring(revitSectionName.IndexOf("/")).SkipWhile(c => !char.IsDigit(c)).TakeWhile(c => char.IsDigit(c)).ToArray());


                                largeFlgWidth = smallFlgWidth = new string(revitSectionName.Substring(revitSectionName.IndexOf('X')).SkipWhile(c => !char.IsDigit(c)).TakeWhile(c => char.IsDigit(c)).ToArray());
                                webThick = flgThick = new string(revitSectionName.Substring(revitSectionName.LastIndexOf('X')).SkipWhile(c => !char.IsDigit(c)).TakeWhile(c => char.IsDigit(c)).ToArray());

                            }

                            //  800/500X400/300X20
                            else if (no_of_x == 2 && no_of_slash == 2)
                            {


                                isFlgTapered = true;
                                int smallFlgWidthIndex = Frame.IndexOfNth(revitSectionName, "/", 1);



                                largeWebDepth = new string(revitSectionName.SkipWhile(c => !char.IsDigit(c)).TakeWhile(c => char.IsDigit(c)).ToArray());
                                smallWebDepth = new string(revitSectionName.Substring(revitSectionName.IndexOf("/")).SkipWhile(c => !char.IsDigit(c)).TakeWhile(c => char.IsDigit(c)).ToArray());



                                largeFlgWidth = new string(revitSectionName.Substring(revitSectionName.IndexOf('X')).SkipWhile(c => !char.IsDigit(c)).TakeWhile(c => char.IsDigit(c)).ToArray());
                                smallFlgWidth = new string(revitSectionName.Substring(smallFlgWidthIndex).SkipWhile(c => !char.IsDigit(c)).TakeWhile(c => char.IsDigit(c)).ToArray());


                                webThick = flgThick = new string(revitSectionName.Substring(revitSectionName.LastIndexOf("X")).SkipWhile(c => !char.IsDigit(c)).TakeWhile(c => char.IsDigit(c)).ToArray());



                            }

                            #endregion
                            double largeDim;
                            double smallDim;
                            if (double.Parse(largeWebDepth) < double.Parse(largeFlgWidth))
                            {
                                largeDim = double.Parse(largeWebDepth) * lengthFactor;
                            }
                            else
                            {
                                largeDim = double.Parse(largeFlgWidth) * lengthFactor;
                            }

                            if (double.Parse(smallWebDepth) < double.Parse(smallFlgWidth))
                            {
                                smallDim = double.Parse(smallWebDepth) * lengthFactor;
                            }

                            else
                            {
                                smallDim = double.Parse(smallFlgWidth) * lengthFactor;
                            }


                            thick = double.Parse(webThick) * lengthFactor;



                            staadSec = "TUBE " + largeDim / lengthFactor + "/" + smallDim / lengthFactor + "X" + thick / lengthFactor;
                            if (StaadExistingSecNames.Contains(staadSec))
                            {

                                ret = staadExistingSecNameDic.FirstOrDefault(x => x.Value == staadSec).Key;
                            }

                            else
                            {
                                ret = property.CreateTaperedTubeProperty(5, largeDim, smallDim, thick);
                                if (ret <= 0)
                                {
                                    ret = -1;
                                }
                            }


                            result.Add(revitSectionName, ret);
                            continue;
                        #endregion

                        #region equal single ANgle

                        case SectType.COL_L_EQUAL:

                            string firstDimention = null;
                            string secondDimention = null;

                            noOfX = Frame.CountCharsUsingIndex(revitSectionName, 'X');
                            dimentionswithX = new string(revitSectionName.SkipWhile(c => !char.IsDigit(c)).ToArray());
                            firstDimention = new string(dimentionswithX.TakeWhile(c => char.IsDigit(c)).ToArray());



                            if (noOfX == 2)
                            {
                                staadSec = "L" + dimentionswithX;

                                if (StaadExistingSecNames.Contains(staadSec))
                                {

                                    ret = staadExistingSecNameDic.FirstOrDefault(x => x.Value == staadSec).Key;
                                }

                                else
                                {
                                    ret = property.CreateAnglePropertyFromTable(7, staadSec, 0, 0.0);
                                    if (ret <= 0)
                                    {
                                        ret = -1;
                                    }
                                }



                            }

                            else
                            {
                                string dimentionswithXDublicated = dimentionswithX.Insert(firstDimention.IndexOf("X") + 1, firstDimention + "X"); ;
                                staadSec = "L" + dimentionswithXDublicated;

                                if (StaadExistingSecNames.Contains(staadSec))
                                {

                                    ret = staadExistingSecNameDic.FirstOrDefault(x => x.Value == staadSec).Key;
                                }

                                else
                                {
                                    ret = property.CreateAnglePropertyFromTable(7, staadSec, 0, 0.0);
                                    if (ret <= 0)
                                    {
                                        ret = -1;
                                    }
                                }


                            }
                            result.Add(revitSectionName, ret);
                            continue;

                        #endregion

                        #region UNEqual single ANgle

                        case SectType.COL_L_UNEQUAL:

                            firstDimention = null;
                            secondDimention = null;

                            noOfX = Frame.CountCharsUsingIndex(revitSectionName, 'X');
                            dimentionswithX = new string(revitSectionName.SkipWhile(c => !char.IsDigit(c)).ToArray());
                            firstDimention = new string(dimentionswithX.TakeWhile(c => char.IsDigit(c)).ToArray());


                            staadSec = "L" + dimentionswithX;


                            if (StaadExistingSecNames.Contains(staadSec))
                            {

                                ret = staadExistingSecNameDic.FirstOrDefault(x => x.Value == staadSec).Key;
                            }

                            else
                            {
                                ret = property.CreateAnglePropertyFromTable(7, staadSec, 0, 0.0);
                                if (ret <= 0)
                                {
                                    ret = -1;
                                }
                            }




                            result.Add(revitSectionName, ret);
                            continue;

                        #endregion

                        #region  dbl ANgle

                        case SectType.COL_2L_EQUALANGLE:
                            noOfX = Frame.CountCharsUsingIndex(revitFrame.RvtName, 'X');
                            string revitSectionModified = revitFrame.RvtName.Substring(2);
                            dimentionswithX = new string(revitSectionModified.SkipWhile(c => !char.IsDigit(c)).ToArray());

                            int indexOfSecondX = Frame.IndexOfNth(dimentionswithX, "X", 0);
                            int lastXindex = revitSectionModified.LastIndexOf("X");

                            string angleThick = revitSectionModified.Substring(lastXindex + 1);

                            firstDimention = new string(dimentionswithX.TakeWhile(c => char.IsDigit(c)).ToArray());

                            staadSec = "L" + firstDimention + "X" + firstDimention + "X" + angleThick;



                            if (StaadExistingSecNames.Contains(staadSec))
                            {

                                ret = staadExistingSecNameDic.FirstOrDefault(x => x.Value == staadSec).Key;
                            }

                            else
                            {
                                ret = property.CreateAnglePropertyFromTable(7, staadSec, 3, 0.0);

                                if (ret <= 0)
                                {
                                    ret = -1;
                                }
                            }


                            result.Add(revitSectionName, ret);
                            continue;


                        #endregion


                        default:
                            result.Add(revitSectionName, -1);
                            continue;





                    }
                }


                catch (Exception)
                {


                    result.Add(revitSectionName, -1);
                    continue;

                }


            }

            #region hadling of revit section not found in staad 
            int dummySection = 0;

            dummySection = staadExistingSecNameDic.FirstOrDefault(x => x.Value == "PIPE 40X10").Key;

            if (dummySection == 0)
            {

                double outerDiameter = double.Parse("40") * lengthFactor;
                double InnerDiametr = outerDiameter - 2 * double.Parse("10") * lengthFactor;

                ret = property.CreatePipePropertyFromTable(7, "", -1, outerDiameter, InnerDiametr);
                if (ret == 0 || ret < 0)
                {
                    ret = -1;
                }

                staadExistingSecNameDic.Add(ret, "PIPE 40X10");
            }


            List<string> keysToUpdate = new List<string>();
            foreach (var kvp in result)
            {
                if (kvp.Value == -1)
                {
                    keysToUpdate.Add(kvp.Key);
                }
            }

            foreach (var key in keysToUpdate)
            {
                result[key] = staadExistingSecNameDic.FirstOrDefault(x => x.Value == "PIPE 40X10").Key;
            }



            #endregion

            return result;



        }


        public static Dictionary<int, string> getModifiedNameOfAllDefinedSectionsInStaad()
        {
            #region defintions 

            double lengthFactor = 0.0;
            double forceFactor = 0.0;
            double momentFactor = 0.0;
            getUnitFactorsFactorFromStaad(ref forceFactor, ref momentFactor, ref lengthFactor);

            Dictionary<int, String> shapeDic = new Dictionary<int, string>();
            Dictionary<int, String> secNameDic = new Dictionary<int, string>();
            Dictionary<int, String> LibDic = new Dictionary<int, string>();
            Dictionary<int, bool> isBUSDIC = new Dictionary<int, bool>();

            #endregion

            List<int> sections = getSections();
            getDictionaries(lengthFactor, sections, ref shapeDic, ref secNameDic, ref LibDic, ref isBUSDIC);

            return secNameDic;


        }


        #endregion



    }
}
