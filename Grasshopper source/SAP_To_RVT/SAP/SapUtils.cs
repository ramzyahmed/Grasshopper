using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Microsoft.Office.Interop.Excel;
using SAP2000v1;
using System;
using System.Collections.Generic;
using System.Data;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Numerics;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Xml.Linq;
using Area = BIM_Integration.Area;
using Curve = Autodesk.Revit.DB.Curve;
using e = ETABSv1.eLoadPatternType;
using Excel = Microsoft.Office.Interop.Excel;
using Frame = BIM_Integration.Frame;
using Line = Autodesk.Revit.DB.Line;
using Parameter = Autodesk.Revit.DB.Parameter;
using Point = BIM_Integration.Point;


namespace BIM_Integration
{
    public class SapUtils

    {
        #region ForHydraMain
        public static string path2 = @"";
        public static cSapModel Get_sap_mode()
        {
            /// cOAPI mySapObject = (cOAPI)System.Runtime.InteropServices.Marshal.GetActiveObject("CSI.SAP2000.API.SapObject");
            // SAP2000v1.cSapModel mysapmodel = (cSapModel) System.Runtime.InteropServices.Marshal.GetActiveObject("CSI.SAP2000.API.SapModel");

            SAP2000v1.cOAPI mySapObject2 = ComActiveObject.Get("CSI.SAP2000.API.SapObject") as SAP2000v1.cOAPI;

            mySapObject2.SapModel.SetPresentUnits(eUnits.Ton_mm_C);



            #region getSteelMaterial


            int ret;
            string material = null;

            int numberOfNames = 0;
            string[] materialList = new string[0];
            ret = mySapObject2.SapModel.PropMaterial.GetNameList(ref numberOfNames, ref materialList);
            eMatType matType = eMatType.Steel;
            int a = 0;
            foreach (var mat in materialList)
            {

                mySapObject2.SapModel.PropMaterial.GetTypeOAPI(mat, ref matType, ref a);

                if (matType == eMatType.Steel)
                {
                    material = mat;
                    break;
                }


            }






            #endregion

            #region createDummySectionInSap

            int dummySectionId = mySapObject2.SapModel.PropFrame.SetCircle("dummySection", material, 10);


            #endregion  


            //aaa
            return mySapObject2.SapModel;


        }

        public class Sections
        {
            public static List<string> GetAllNamesAssignedToFramesInSap(cSapModel mySapModel, List<string> OptionalFrameIDS = null)
            {
                #region defintions 
                int ret;
                string PropName = null;
                string SAuto = null;
                List<string> All_sections = new List<string>();

                int numberOfITems = 0;
                string GUID = null;
                string notes = null;
                int color = 0;
                string[] StartSec = null;
                string[] Endsec = null;
                double[] MyLength = null;
                int[] MyType = null;
                int[] EI33 = null;
                int[] EI22 = null;

                #endregion

                List<string> frame_ids = getSapFrameIDS();

                if (OptionalFrameIDS == null)
                {
                    if (getSelectedFramesIDS().Count > 0)
                    {
                        frame_ids = getSelectedFramesIDS();
                    }

                }

                else
                {
                    frame_ids = OptionalFrameIDS;
                }

                eFramePropType e = eFramePropType.Angle;
                foreach (string Id in frame_ids)
                {
                    ret = mySapModel.FrameObj.GetSection(Id, ref PropName, ref SAuto);


                    mySapModel.PropFrame.GetTypeOAPI(PropName, ref e);
                    if (e == eFramePropType.Variable)
                    {
                        ret = mySapModel.PropFrame.GetNonPrismatic(PropName, ref numberOfITems, ref StartSec, ref Endsec, ref MyLength, ref MyType, ref EI33, ref EI22, ref color, ref notes, ref GUID);
                        All_sections.Add(StartSec[0]);
                        All_sections.Add(Endsec[0]);

                    }

                    All_sections.Add(PropName);

                }



                All_sections = All_sections.Distinct().ToList();
                return All_sections;
            }
            public static List<string> getSections(cSapModel mySapModel)
            {
                #region defintions 
                int ret;
                int numberOfSections = 0;
                string[] Sections = null;
                string PropName = null;
                string SAuto = null;
                List<string> All_sections = new List<string>();
                #endregion

                ret = mySapModel.PropFrame.GetNameList(ref numberOfSections, ref Sections);

                All_sections = Sections.ToList();

                return All_sections;
            }

            public static Dictionary<string, Section_type> getSectionTypes(List<string> sapSections)

            {
                cSapModel mySapModel = SapUtils.Get_sap_mode();

                Dictionary<string, string> libDic = new Dictionary<string, string>();

                #region definitions 
                List<Column> Columns2 = new List<Column>();
                List<Beam> Beams2 = new List<Beam>();
                List<Beam> all_members2 = new List<Beam>();
                int ret;

                List<string> all_profiles = new List<string>();
                Dictionary<string, Section_type> mydic = new Dictionary<string, Section_type>();


                #endregion

                foreach (string PropName2 in sapSections)
                {
                    #region defintions
                    string NameInFile = null;
                    string FileName = null;
                    string MatProp = null;
                    eFramePropType PropType = eFramePropType.Angle;
                    Section_type mysectiontype;
                    #endregion
                    // get type of each prop
                    ret = mySapModel.PropFrame.GetTypeOAPI(PropName2, ref PropType);


                    //  H for HotRolled and B for Built up 

                    switch (PropType)
                    {

                        //case eFramePropType.I:
                        //    ret = mySapModel.PropFrame.GetNameInPropFile(PropName2, ref NameInFile, ref FileName, ref MatProp, ref PropType);

                        //    if (NameInFile != "")
                        //    {
                        //        mysectiontype = Section_type.I_SECTION_H;
                        //        if (NameInFile.ToUpper().Contains("Euro"))
                        //        {
                        //            libDic.Add(PropName2, "euro");
                        //        }

                        //        else if ((NameInFile.ToUpper().Contains("Aisc") || NameInFile.ToUpper().Contains("Australia")))
                        //        {
                        //            libDic.Add(PropName2, "AISC");
                        //        }

                        //        else if ((NameInFile.ToUpper().Contains("BSShape") || NameInFile.ToUpper().Contains("ArcelorMittal_British")))
                        //        {
                        //            libDic.Add(PropName2, "BSSHAPES");
                        //        }
                        //        else
                        //        {
                        //            mysectiontype = Section_type.I_SECTION_B;
                        //        }
                        //    }

                        //    else
                        //    {
                        //        mysectiontype = Section_type.I_SECTION_B;
                        //    }
                        //    mydic.Add(PropName2, mysectiontype);
                        //    break;

                        case eFramePropType.I:
                            ret = mySapModel.PropFrame.GetNameInPropFile(PropName2, ref NameInFile, ref FileName, ref MatProp, ref PropType);

                            if (NameInFile != "")
                            {
                                mysectiontype = Section_type.I_SECTION_H;
                                if (FileName.ToUpper().Contains("EURO"))
                                {
                                    libDic.Add(PropName2, "euro");
                                }

                                else if ((FileName.ToUpper().Contains("AISC") || FileName.ToUpper().Contains("AUSTRALIA")))
                                {
                                    libDic.Add(PropName2, "AISC");
                                }

                                else if ((FileName.ToUpper().Contains("BSSHAPE") || FileName.ToUpper().Contains("ARCELORMITTAL_BRITISH")))
                                {
                                    libDic.Add(PropName2, "BSSHAPES");

                                }
                                else
                                {
                                    mysectiontype = Section_type.I_SECTION_B;
                                }
                            }

                            else
                            {
                                mysectiontype = Section_type.I_SECTION_B;
                            }
                            mydic.Add(PropName2, mysectiontype);
                            break;


                        case eFramePropType.Channel:
                            ret = mySapModel.PropFrame.GetNameInPropFile(PropName2, ref NameInFile, ref FileName, ref MatProp, ref PropType);

                            if (NameInFile != "")
                            {
                                mysectiontype = Section_type.Channel_H;
                            }
                            else
                            {
                                mysectiontype = Section_type.Channel_B;
                            }
                            mydic.Add(PropName2, mysectiontype);
                            break;




                        case eFramePropType.T:
                            ret = mySapModel.PropFrame.GetNameInPropFile(PropName2, ref NameInFile, ref FileName, ref MatProp, ref PropType);

                            if (NameInFile != "")
                            {
                                mysectiontype = Section_type.T_Section_H;
                            }
                            else
                            {
                                mysectiontype = Section_type.T_Section_B;
                            }
                            mydic.Add(PropName2, mysectiontype);
                            break;




                        case eFramePropType.Angle:
                            ret = mySapModel.PropFrame.GetNameInPropFile(PropName2, ref NameInFile, ref FileName, ref MatProp, ref PropType);

                            if (NameInFile != "")
                            {
                                mysectiontype = Section_type.S_Angle_H;
                            }
                            else
                            {
                                mysectiontype = Section_type.S_Angle_B;
                            }
                            mydic.Add(PropName2, mysectiontype);
                            break;




                        case eFramePropType.DblAngle:
                            ret = mySapModel.PropFrame.GetNameInPropFile(PropName2, ref NameInFile, ref FileName, ref MatProp, ref PropType);

                            if (NameInFile != "")
                            {
                                mysectiontype = Section_type.D_Angle_H;
                            }
                            else
                            {
                                mysectiontype = Section_type.D_Angle_B;
                            }
                            mydic.Add(PropName2, mysectiontype);
                            break;




                        case eFramePropType.Box:
                            ret = mySapModel.PropFrame.GetNameInPropFile(PropName2, ref NameInFile, ref FileName, ref MatProp, ref PropType);

                            if (NameInFile != "")
                            {
                                mysectiontype = Section_type.Hollow_box_H;
                            }
                            else
                            {
                                mysectiontype = Section_type.Hollow_box_B;
                            }
                            mydic.Add(PropName2, mysectiontype);
                            break;


                        case eFramePropType.Pipe:

                            ret = mySapModel.PropFrame.GetNameInPropFile(PropName2, ref NameInFile, ref FileName, ref MatProp, ref PropType);

                            if (NameInFile != "")
                            {
                                mysectiontype = Section_type.Hollow_pipe_H;
                            }
                            else
                            {
                                mysectiontype = Section_type.Hollow_pipe_B;
                            }
                            mydic.Add(PropName2, mysectiontype);
                            break;



                        case eFramePropType.Rectangular:
                            double t3 = 0;
                            double t2 = 0;
                            int Color2 = 0;
                            string Notes2 = null;
                            string GUID2 = null;

                            // filter between solid concrete and solid steel
                            ret = mySapModel.PropFrame.GetRectangle(PropName2, ref FileName, ref MatProp, ref t3, ref t2, ref Color2, ref Notes2, ref GUID2);
                            if (t3 > 150 && t2 > 150)
                            {
                                mysectiontype = Section_type.Concrete_rect;

                            }
                            else
                            {
                                mysectiontype = Section_type.Solid_box;

                            }
                            mydic.Add(PropName2, mysectiontype);
                            break;


                        case eFramePropType.Circle:

                            t3 = 0;
                            Color2 = 0;
                            Notes2 = null;
                            GUID2 = null;

                            // filter between solid concrete and solid steel
                            ret = mySapModel.PropFrame.GetCircle(PropName2, ref FileName, ref MatProp, ref t3, ref Color2, ref Notes2, ref GUID2);
                            if (t3 > 150)
                            {
                                mysectiontype = Section_type.Concrete_circ;

                            }
                            else
                            {
                                mysectiontype = Section_type.Solid_pipe;

                            }
                            mydic.Add(PropName2, mysectiontype);
                            break;


                        case eFramePropType.Cold_C:
                            mysectiontype = Section_type.Channel_cold;
                            mydic.Add(PropName2, mysectiontype);
                            break;

                        case eFramePropType.Cold_Z:
                            mysectiontype = Section_type.z_cold_Section;
                            mydic.Add(PropName2, mysectiontype);
                            break;


                        case eFramePropType.ConcreteBox:
                            mysectiontype = Section_type.Concrete_rect;
                            mydic.Add(PropName2, mysectiontype);
                            break;


                        case eFramePropType.ConcretePipe:
                            mysectiontype = Section_type.Concrete_circ;
                            mydic.Add(PropName2, mysectiontype);
                            break;
                            ;


                        case eFramePropType.SD:
                            mysectiontype = Section_type.Section_Designer;
                            mydic.Add(PropName2, mysectiontype);
                            break;


                        ////tapered
                        case eFramePropType.Variable:

                            #region defintions
                            int NumberItems = 0;
                            string[] StartSec = null;
                            string[] EndSec = null;
                            double[] MyLength = null;
                            int[] MyType = null;
                            int[] EI33 = null;
                            int[] EI22 = null;
                            int Color = 0;
                            string Notes = null;
                            string GUID = null;
                            #endregion

                            ret = mySapModel.PropFrame.GetNonPrismatic(PropName2, ref NumberItems, ref StartSec, ref EndSec, ref MyLength, ref MyType, ref EI33, ref EI22, ref Color, ref Notes, ref GUID);
                            ret = mySapModel.PropFrame.GetTypeOAPI(StartSec[0], ref PropType);

                            //tapered i or pipe or box
                            switch (PropType)
                            {
                                case eFramePropType.I:
                                    mysectiontype = Section_type.I_Tapered;
                                    mydic.Add(PropName2, mysectiontype);
                                    break;

                                case eFramePropType.Box:
                                    mysectiontype = Section_type.Box_Tapered;
                                    mydic.Add(PropName2, mysectiontype);
                                    break;

                                case eFramePropType.Pipe:
                                    mysectiontype = Section_type.Pipe_Tapered;
                                    mydic.Add(PropName2, mysectiontype);
                                    break;

                            }
                            break;


                    }


                }

                return mydic;
            }

            public static object Get_i_section_dimentions(cSapModel mySapModel, List<string> profiles)
            {

                #region defintions 
                string FileName = null;
                string MatProp = null;
                double t3 = 0;
                double t2 = 0;
                double tf = 0;
                double tw = 0;
                double t2b = 0;
                double tfb = 0;
                int Color = 0;
                string Notes = null;
                string GUID = null;
                profiles = profiles.Distinct().ToList();

                #endregion

                List<object> section_dimentions = new List<object>();

                foreach (string profile in profiles)
                {

                    int ret = mySapModel.PropFrame.GetISection(profile, ref FileName, ref MatProp, ref t3, ref t2, ref tf, ref tw, ref t2b, ref tfb, ref Color, ref Notes, ref GUID);
                    tw = System.Math.Round(tw, 2);
                    tf = System.Math.Round(tf, 2);
                    t3 = System.Math.Round(t3, 2);
                    t2 = System.Math.Round(t2, 2);


                    var obj2 = new { name = profile, depth = t3, width = t2, flg_thick = tf, web_thick = tw, rf = 0, mat = MatProp };
                    section_dimentions.Add(obj2);



                }

                return section_dimentions;
            }

            public static object Get_Box_section_dimentions(cSapModel mySapModel, List<string> profiles)
            {

                #region defintions 
                string FileName = null;
                string MatProp = null;
                double t3 = 0;
                double t2 = 0;
                double tf = 0;
                double tw = 0;
                int Color = 0;
                string Notes = null;
                string GUID = null;
                profiles = profiles.Distinct().ToList();
                #endregion

                List<object> section_dimentions = new List<object>();

                foreach (string profile in profiles)
                {
                    int ret = mySapModel.PropFrame.GetTube(profile, ref FileName, ref MatProp, ref t3, ref t2, ref tf, ref tw, ref Color, ref Notes, ref GUID);
                    tw = System.Math.Round(tw, 2);
                    tf = System.Math.Round(tf, 2);
                    t3 = System.Math.Round(t3, 2);
                    t2 = System.Math.Round(t2, 2);


                    var obj = new { name = profile, depth = t3, width = t2, flg_thick = tf, web_thick = tw, rf = 0, mat = MatProp };
                    section_dimentions.Add(obj);


                }

                return section_dimentions;
            }

            public static object Get_pipe_section_dimentions(cSapModel mySapModel, List<string> profiles)
            {

                #region defintions 
                string FileName = null;
                string MatProp = null;
                double t3 = 0;
                double tw = 0;

                int Color = 0;
                string Notes = null;
                string GUID = null;
                profiles = profiles.Distinct().ToList();
                #endregion

                List<object> section_dimentions = new List<object>();

                foreach (string profile in profiles)
                {
                    int ret = mySapModel.PropFrame.GetPipe(profile, ref FileName, ref MatProp, ref t3, ref tw, ref Color, ref Notes, ref GUID);
                    t3 = System.Math.Round(t3, 2);
                    tw = System.Math.Round(tw, 2);
                    var obj = new { name = profile, diameter = t3, thick = tw, mat = MatProp };
                    section_dimentions.Add(obj);




                    //handle if file name is not blank (improted section) you need to access xml files

                }

                return section_dimentions;
            }

            public static object Get_channel_section_dimentions(cSapModel mySapModel, List<string> profiles)
            {

                #region defintions 
                string FileName = null;
                string MatProp = null;
                double t3 = 0;
                double t2 = 0;
                double tf = 0;
                double tw = 0;
                int Color = 0;
                string Notes = null;
                string GUID = null;
                profiles = profiles.Distinct().ToList();
                #endregion

                List<object> section_dimentions = new List<object>();

                foreach (string profile in profiles)
                {
                    int ret = mySapModel.PropFrame.GetChannel(profile, ref FileName, ref MatProp, ref t3, ref t2, ref tf, ref tw, ref Color, ref Notes, ref GUID);
                    tw = System.Math.Round(tw, 2);
                    tf = System.Math.Round(tf, 2);
                    t3 = System.Math.Round(t3, 2);
                    t2 = System.Math.Round(t2, 2);

                    if (FileName != "")
                    {
                        #region get radius

                        FileName = Path.GetFileNameWithoutExtension(FileName);
                        FileName = FileName + ".xml";
                        string profile2 = profile.ToUpper();
                        XDocument doc = XDocument.Load(SapUtils.path2 + FileName);
                        var x = doc.Element("PROPERTY_FILE").Elements("STEEL_CHANNEL").ToList();

                        //
                        string units = doc.Element("PROPERTY_FILE").Element("CONTROL").Element("LENGTH_UNITS").Value;
                        double factor;
                        if (units == "in")
                        {
                            factor = 2.54 * 10;
                        }
                        if (units == "cm")

                        {
                            factor = 10;
                        }
                        else
                        {
                            factor = 1;
                        }


                        //
                        var y = x.Where(c => c.Element("LABEL").Value == profile2).ToList();
                        var t = y[0].Element("KDES");
                        string v;
                        double rf;
                        if (t != null)
                        {
                            v = t.Value;
                            rf = Convert.ToDouble(v) * factor - tf;
                        }

                        else
                        {
                            rf = 0;

                        }


                        #endregion

                        var obj = new { name = profile, depth = t3, width = t2, flg_thick = tf, web_thick = tw, rf = rf, mat = MatProp };
                        section_dimentions.Add(obj);
                    }

                    else
                    {
                        var obj = new { name = profile, depth = t3, width = t2, flg_thick = tf, web_thick = tw, rf = 0, mat = MatProp };
                        section_dimentions.Add(obj);
                    }



                }

                return section_dimentions;
            }

            public static object Get_dbl_angle_section_dimentions(cSapModel mySapModel, List<string> profiles)
            {

                #region defintions 
                string FileName = null;
                string MatProp = null;
                double t3 = 0;
                double t2 = 0;
                double tf = 0;
                double tw = 0;
                double dis = 0;
                int Color = 0;
                string Notes = null;
                string GUID = null;
                profiles = profiles.Distinct().ToList();
                #endregion

                List<object> section_dimentions = new List<object>();

                foreach (string profile in profiles)
                {

                    int ret = mySapModel.PropFrame.GetDblAngle(profile, ref FileName, ref MatProp, ref t3, ref t2, ref tf, ref tw, ref dis, ref Color, ref Notes, ref GUID);
                    tw = System.Math.Round(tw, 2);
                    tf = System.Math.Round(tf, 2);
                    t3 = System.Math.Round(t3, 2);
                    t2 = System.Math.Round(t2, 2);

                    //handle if file name is not blank (improted section) you need to access xml files (no radiuse)
                    var obj = new { name = profile, vr_leg_depth = t3, width = t2, HZ_leg_thick = tf, Vr_leg_thick = tw, dist = dis, mat = MatProp };
                    section_dimentions.Add(obj);
                }

                return section_dimentions;
            }

            public static object Get_Sig_angle_section_dimentions(cSapModel mySapModel, List<string> profiles)
            {

                #region defintions 
                string FileName = null;
                string MatProp = null;
                double t3 = 0;
                double t2 = 0;
                double tf = 0;
                double tw = 0;
                int Color = 0;
                string Notes = null;
                string GUID = null;
                profiles = profiles.Distinct().ToList();
                #endregion

                List<object> section_dimentions = new List<object>();

                foreach (string profile in profiles)
                {
                    int ret = mySapModel.PropFrame.GetAngle(profile, ref FileName, ref MatProp, ref t3, ref t2, ref tf, ref tw, ref Color, ref Notes, ref GUID);
                    tw = System.Math.Round(tw, 2);
                    tf = System.Math.Round(tf, 2);
                    t3 = System.Math.Round(t3, 2);
                    t2 = System.Math.Round(t2, 2);

                    if (FileName != "")
                    {
                        #region get radius

                        FileName = Path.GetFileNameWithoutExtension(FileName);
                        FileName = FileName + ".xml";
                        string profile2 = profile.ToUpper();
                        XDocument doc = XDocument.Load(SapUtils.path2 + FileName);
                        var x = doc.Element("PROPERTY_FILE").Elements("STEEL_ANGLE").ToList();
                        //
                        string units = doc.Element("PROPERTY_FILE").Element("CONTROL").Element("LENGTH_UNITS").Value;
                        double factor;
                        if (units == "in")
                        {
                            factor = 2.54 * 10;
                        }
                        if (units == "cm")

                        {
                            factor = 10;
                        }
                        else
                        {
                            factor = 1;
                        }
                        //
                        var y = x.Where(c => c.Element("LABEL").Value == profile2).ToList();
                        var t = y[0].Element("KDES");
                        string v;
                        double rf;
                        if (t != null)
                        {
                            v = t.Value;
                            rf = Convert.ToDouble(v) * factor - tf;
                        }

                        else
                        {
                            rf = 0;

                        }



                        #endregion

                        tw = System.Math.Round(tw, 2);
                        var obj = new { name = profile, vr_leg_depth = t3, width = t2, flg_thick = tf, web_thick = tw, rf = rf, mat = MatProp };
                        section_dimentions.Add(obj);
                    }

                    else
                    {
                        var obj = new { name = profile, vr_leg_depth = t3, width = t2, flg_thick = tf, web_thick = tw, rf = 0, mat = MatProp };
                        section_dimentions.Add(obj);
                    }


                }

                return section_dimentions;
            }

            public static object Get_Tee_section_dimentions(cSapModel mySapModel, List<string> profiles)
            {

                #region defintions 
                string FileName = null;
                string MatProp = null;
                double t3 = 0;
                double t2 = 0;
                double tf = 0;
                double tw = 0;
                int Color = 0;
                string Notes = null;
                string GUID = null;
                profiles = profiles.Distinct().ToList();
                #endregion

                List<object> section_dimentions = new List<object>();

                foreach (string profile in profiles)
                {
                    int ret = mySapModel.PropFrame.GetTee(profile, ref FileName, ref MatProp, ref t3, ref t2, ref tf, ref tw, ref Color, ref Notes, ref GUID);
                    tw = System.Math.Round(tw, 2);
                    tf = System.Math.Round(tf, 2);
                    t3 = System.Math.Round(t3, 2);
                    t2 = System.Math.Round(t2, 2);

                    if (FileName != "")
                    {
                        #region get radius

                        FileName = Path.GetFileNameWithoutExtension(FileName);
                        FileName = FileName + ".xml";
                        string profile2 = profile.ToUpper();
                        XDocument doc = XDocument.Load(SapUtils.path2 + FileName);
                        var x = doc.Element("PROPERTY_FILE").Elements("STEEL_TEE").ToList();
                        //
                        string units = doc.Element("PROPERTY_FILE").Element("CONTROL").Element("LENGTH_UNITS").Value;
                        double factor;
                        if (units == "in")
                        {
                            factor = 2.54 * 10;
                        }
                        if (units == "cm")

                        {
                            factor = 10;
                        }
                        else
                        {
                            factor = 1;
                        }
                        //
                        var y = x.Where(c => c.Element("LABEL").Value == profile2).ToList();
                        var t = y[0].Element("KDES");
                        string v;
                        double rf;
                        if (t != null)
                        {
                            v = t.Value;
                            rf = Convert.ToDouble(v) * factor - tf;
                        }

                        else
                        {
                            rf = 0;

                        }

                        #endregion

                        var obj = new { name = profile, depth = t3, width = t2, flg_thick = tf, web_thick = tw, rf = rf, mat = MatProp };
                        section_dimentions.Add(obj);
                    }

                    else
                    {
                        var obj = new { name = profile, depth = t3, width = t2, flg_thick = tf, web_thick = tw, rf = 0, mat = MatProp };
                        section_dimentions.Add(obj);
                    }


                }

                return section_dimentions;
            }

            public static object Get_solid_pipe_section_dimentions(cSapModel mySapModel, List<string> profiles)
            {

                #region defintions 
                string FileName = null;
                string MatProp = null;
                double t3 = 0;
                int Color = 0;
                string Notes = null;
                string GUID = null;
                profiles = profiles.Distinct().ToList();
                #endregion

                List<object> section_dimentions = new List<object>();

                foreach (string profile in profiles)
                {
                    int ret = mySapModel.PropFrame.GetCircle(profile, ref FileName, ref MatProp, ref t3, ref Color, ref Notes, ref GUID);

                    //handle if file name is not blank (improted section) you need to access xml files
                    var obj = new { name = profile, diameter = t3, mat = MatProp };
                    section_dimentions.Add(obj);
                }

                return section_dimentions;
            }

            public static object Get_solid_rect_section_dimentions(cSapModel mySapModel, List<string> profiles)
            {

                #region defintions 
                string FileName = null;
                string MatProp = null;
                double t3 = 0;
                double t2 = 0;
                int Color = 0;
                string Notes = null;
                string GUID = null;
                profiles = profiles.Distinct().ToList();
                #endregion

                List<object> section_dimentions = new List<object>();

                foreach (string profile in profiles)
                {
                    int ret = mySapModel.PropFrame.GetRectangle(profile, ref FileName, ref MatProp, ref t3, ref t2, ref Color, ref Notes, ref GUID);

                    //handle if file name is not blank (improted section) you need to access xml files
                    var obj = new { name = profile, depth = t3, width = t2, mat = MatProp };
                    section_dimentions.Add(obj);
                }

                return section_dimentions;
            }

            public static object Get_z_cold_section_dimentions(cSapModel mySapModel, List<string> profiles)
            {

                #region defintions 
                string FileName = null;
                string MatProp = null;
                double t3 = 0;
                double t2 = 0;
                double Thickness = 0;
                double LipDepth = 0;
                double Radius = 0;
                double LipAngle = 0;
                int Color = 0;
                string Notes = null;
                string GUID = null;
                profiles = profiles.Distinct().ToList();
                #endregion

                List<object> section_dimentions = new List<object>();

                foreach (string profile in profiles)
                {


                    int ret = mySapModel.PropFrame.GetColdZ(profile, ref FileName, ref MatProp, ref t3, ref t2, ref Thickness, ref Radius, ref LipDepth, ref LipAngle, ref Color, ref Notes, ref GUID);

                    //handle if file name is not blank (improted section) you need to access xml files
                    var obj = new { name = profile, depth = t3, width = t2, thickness = Thickness, radius = Radius, lipDepth = LipDepth, lipAngle = LipAngle, mat = MatProp };
                    section_dimentions.Add(obj);
                }

                return section_dimentions;
            }

            public static object Get_C_cold_section_dimentions(cSapModel mySapModel, List<string> profiles)
            {

                #region defintions 
                string FileName = null;
                string MatProp = null;
                double t3 = 0;
                double t2 = 0;
                double Thickness = 0;
                double LipDepth = 0;
                double Radius = 0;
                int Color = 0;
                string Notes = null;
                string GUID = null;
                profiles = profiles.Distinct().ToList();
                #endregion

                List<object> section_dimentions = new List<object>();

                foreach (string profile in profiles)
                {


                    int ret = mySapModel.PropFrame.GetColdC(profile, ref FileName, ref MatProp, ref t3, ref t2, ref Thickness, ref Radius, ref LipDepth, ref Color, ref Notes, ref GUID);

                    //handle if file name is not blank (improted section) you need to access xml files
                    var obj = new { name = profile, depth = t3, width = t2, thickness = Thickness, radius = Radius, lipDepth = LipDepth, mat = MatProp };
                    section_dimentions.Add(obj);
                }

                return section_dimentions;
            }

            public static object Get_Tapered_i_section_dimentions(cSapModel mySapModel, List<string> profiles)
            {

                #region defintions 
                string FileName = null;
                string MatProp = null;
                double t3 = 0;
                double t2 = 0;
                double tf = 0;
                double tw = 0;
                double t2b = 0;
                double tfb = 0;
                double t32 = 0;
                double t22 = 0;
                double tf2 = 0;
                double tw2 = 0;
                double t2b2 = 0;
                double tfb2 = 0;
                int ret = 0;
                int Color = 0;
                string Notes = null;
                string GUID = null;
                int NumberItems = 0;
                string[] StartSec = null;
                string[] EndSec = null;
                double[] MyLength = null;
                int[] MyType = null;
                int[] EI33 = null;
                int[] EI22 = null;

                profiles = profiles.Distinct().ToList();
                #endregion

                List<object> section_dimentions = new List<object>();



                foreach (string profile in profiles)
                {
                    ret = mySapModel.PropFrame.GetNonPrismatic(profile, ref NumberItems, ref StartSec, ref EndSec, ref MyLength, ref MyType, ref EI33, ref EI22, ref Color, ref Notes, ref GUID);


                    ret = mySapModel.PropFrame.GetISection(StartSec[0], ref FileName, ref MatProp, ref t3, ref t2, ref tf, ref tw, ref t2b, ref tfb, ref Color, ref Notes, ref GUID);
                    ret = mySapModel.PropFrame.GetISection(EndSec[0], ref FileName, ref MatProp, ref t32, ref t22, ref tf2, ref tw2, ref t2b2, ref tfb2, ref Color, ref Notes, ref GUID);

                    //handle if file name is not blank (improted section) you need to access xml files
                    var obj = new { name = profile, st_depth = t3, st_width = t2, flg_thick = tf, web_thick = tw, end_depth = t32, end_width = t22, mat = MatProp };
                    section_dimentions.Add(obj);
                }

                return section_dimentions;
            }

            public static object Get_Tapered_BOX_section_dimentions(cSapModel mySapModel, List<string> profiles)
            {

                #region defintions 
                string FileName = null;
                string MatProp = null;
                double t3 = 0;
                double t2 = 0;
                double tf = 0;
                double tw = 0;
                double t32 = 0;
                double t22 = 0;
                double tf2 = 0;
                double tw2 = 0;
                int ret = 0;
                int Color = 0;
                string Notes = null;
                string GUID = null;
                int NumberItems = 0;
                string[] StartSec = null;
                string[] EndSec = null;
                double[] MyLength = null;
                int[] MyType = null;
                int[] EI33 = null;
                int[] EI22 = null;

                profiles = profiles.Distinct().ToList();
                #endregion

                List<object> section_dimentions = new List<object>();



                foreach (string profile in profiles)
                {
                    ret = mySapModel.PropFrame.GetNonPrismatic(profile, ref NumberItems, ref StartSec, ref EndSec, ref MyLength, ref MyType, ref EI33, ref EI22, ref Color, ref Notes, ref GUID);

                    ret = mySapModel.PropFrame.GetTube(StartSec[0], ref FileName, ref MatProp, ref t3, ref t2, ref tf, ref tw, ref Color, ref Notes, ref GUID);
                    ret = mySapModel.PropFrame.GetTube(EndSec[0], ref FileName, ref MatProp, ref t32, ref t22, ref tf2, ref tw2, ref Color, ref Notes, ref GUID);

                    var obj = new { name = profile, st_depth = t3, st_width = t2, flg_thick = tf, web_thick = tw, end_depth = t32, end_width = t22, mat = MatProp };
                    section_dimentions.Add(obj);
                }

                return section_dimentions;
            }

            public static object Get_Tapered_PIPE_section_dimentions(cSapModel mySapModel, List<string> profiles)
            {

                #region defintions 
                string FileName = null;
                string MatProp = null;
                double t3 = 0;
                double tw = 0;
                double t32 = 0;
                double tw2 = 0;
                int ret = 0;
                int Color = 0;
                string Notes = null;
                string GUID = null;
                int NumberItems = 0;
                string[] StartSec = null;
                string[] EndSec = null;
                double[] MyLength = null;
                int[] MyType = null;
                int[] EI33 = null;
                int[] EI22 = null;

                profiles = profiles.Distinct().ToList();
                #endregion

                List<object> section_dimentions = new List<object>();



                foreach (string profile in profiles)
                {
                    ret = mySapModel.PropFrame.GetNonPrismatic(profile, ref NumberItems, ref StartSec, ref EndSec, ref MyLength, ref MyType, ref EI33, ref EI22, ref Color, ref Notes, ref GUID);

                    ret = mySapModel.PropFrame.GetPipe(StartSec[0], ref FileName, ref MatProp, ref t3, ref tw, ref Color, ref Notes, ref GUID);
                    ret = mySapModel.PropFrame.GetPipe(EndSec[0], ref FileName, ref MatProp, ref t32, ref tw2, ref Color, ref Notes, ref GUID);

                    //handle if file name is not blank (improted section) you need to access xml files
                    var obj = new { name = profile, st_diameter = t3, st_thick = tw, end_diameter = t32, end_thick = tw2, mat = MatProp };
                    section_dimentions.Add(obj);
                }

                return section_dimentions;
            }


            public static Dictionary<string, Section_type> getSectionTypesForHydraTest(List<string> sapSections, ref Dictionary<string, string> libDic2)

            {
                cSapModel mySapModel = SapUtils.Get_sap_mode();



                Dictionary<string, string> libDic = new Dictionary<string, string>();

                #region definitions 
                List<Column> Columns2 = new List<Column>();
                List<Beam> Beams2 = new List<Beam>();
                List<Beam> all_members2 = new List<Beam>();
                int ret;

                List<string> all_profiles = new List<string>();
                Dictionary<string, Section_type> mydic = new Dictionary<string, Section_type>();


                #endregion

                foreach (string PropName2 in sapSections)
                {


                    #region defintions
                    string NameInFile = null;
                    string FileName = null;
                    string MatProp = null;
                    eFramePropType PropType = eFramePropType.Angle;
                    Section_type mysectiontype;
                    #endregion
                    // get type of each prop
                    ret = mySapModel.PropFrame.GetTypeOAPI(PropName2, ref PropType);


                    //  H for HotRolled and B for Built up 

                    switch (PropType)
                    {

                        case eFramePropType.I:
                            ret = mySapModel.PropFrame.GetNameInPropFile(PropName2, ref NameInFile, ref FileName, ref MatProp, ref PropType);

                            if (FileName != "")
                            {
                                mysectiontype = Section_type.I_SECTION_H;
                                if (FileName.ToUpper().Contains("EURO"))
                                {
                                    libDic.Add(PropName2, "euro");
                                }

                                else if ((FileName.ToUpper().Contains("AISC") || FileName.ToUpper().Contains("AUSTRALIA")))
                                {
                                    libDic.Add(PropName2, "AISC");
                                }

                                else if ((FileName.ToUpper().Contains("BSSHAPE") || FileName.ToUpper().Contains("BRITISH")))
                                {
                                    libDic.Add(PropName2, "BSSHAPES");
                                }
                            }

                            else
                            {
                                mysectiontype = Section_type.I_SECTION_B;
                            }
                            mydic.Add(PropName2, mysectiontype);
                            break;




                        case eFramePropType.Channel:
                            ret = mySapModel.PropFrame.GetNameInPropFile(PropName2, ref NameInFile, ref FileName, ref MatProp, ref PropType);

                            if (NameInFile != "")
                            {
                                mysectiontype = Section_type.Channel_H;
                            }
                            else
                            {
                                mysectiontype = Section_type.Channel_B;
                            }
                            mydic.Add(PropName2, mysectiontype);
                            break;




                        case eFramePropType.T:
                            ret = mySapModel.PropFrame.GetNameInPropFile(PropName2, ref NameInFile, ref FileName, ref MatProp, ref PropType);

                            if (NameInFile != "")
                            {
                                mysectiontype = Section_type.T_Section_H;
                            }
                            else
                            {
                                mysectiontype = Section_type.T_Section_B;
                            }
                            mydic.Add(PropName2, mysectiontype);
                            break;




                        case eFramePropType.Angle:
                            ret = mySapModel.PropFrame.GetNameInPropFile(PropName2, ref NameInFile, ref FileName, ref MatProp, ref PropType);

                            if (NameInFile != "")
                            {
                                mysectiontype = Section_type.S_Angle_H;
                            }
                            else
                            {
                                mysectiontype = Section_type.S_Angle_B;
                            }
                            mydic.Add(PropName2, mysectiontype);
                            break;




                        case eFramePropType.DblAngle:
                            ret = mySapModel.PropFrame.GetNameInPropFile(PropName2, ref NameInFile, ref FileName, ref MatProp, ref PropType);

                            if (NameInFile != "")
                            {
                                mysectiontype = Section_type.D_Angle_H;
                            }
                            else
                            {
                                mysectiontype = Section_type.D_Angle_B;
                            }
                            mydic.Add(PropName2, mysectiontype);
                            break;




                        case eFramePropType.Box:
                            ret = mySapModel.PropFrame.GetNameInPropFile(PropName2, ref NameInFile, ref FileName, ref MatProp, ref PropType);

                            if (NameInFile != "")
                            {
                                mysectiontype = Section_type.Tube_H;
                            }
                            else
                            {
                                mysectiontype = Section_type.Tube_B;
                            }
                            mydic.Add(PropName2, mysectiontype);
                            break;


                        case eFramePropType.Pipe:

                            ret = mySapModel.PropFrame.GetNameInPropFile(PropName2, ref NameInFile, ref FileName, ref MatProp, ref PropType);

                            if (NameInFile != "")
                            {
                                mysectiontype = Section_type.Pipe_H;
                            }
                            else
                            {
                                mysectiontype = Section_type.Pipe_B;
                            }
                            mydic.Add(PropName2, mysectiontype);
                            break;



                        case eFramePropType.Rectangular:
                            double t3 = 0;
                            double t2 = 0;
                            int Color2 = 0;
                            string Notes2 = null;
                            string GUID2 = null;

                            // filter between solid concrete and solid steel
                            ret = mySapModel.PropFrame.GetRectangle(PropName2, ref FileName, ref MatProp, ref t3, ref t2, ref Color2, ref Notes2, ref GUID2);
                            if (t3 > 150 && t2 > 150)
                            {
                                mysectiontype = Section_type.Concrete_rect;

                            }
                            else
                            {
                                mysectiontype = Section_type.Solid_box;

                            }
                            mydic.Add(PropName2, mysectiontype);
                            break;


                        case eFramePropType.Circle:

                            t3 = 0;
                            Color2 = 0;
                            Notes2 = null;
                            GUID2 = null;

                            // filter between solid concrete and solid steel
                            ret = mySapModel.PropFrame.GetCircle(PropName2, ref FileName, ref MatProp, ref t3, ref Color2, ref Notes2, ref GUID2);
                            if (t3 > 150)
                            {
                                mysectiontype = Section_type.Concrete_circ;

                            }
                            else
                            {
                                mysectiontype = Section_type.Solid_pipe;

                            }
                            mydic.Add(PropName2, mysectiontype);
                            break;


                        case eFramePropType.Cold_C:
                            mysectiontype = Section_type.Channel_cold;
                            mydic.Add(PropName2, mysectiontype);
                            break;

                        case eFramePropType.Cold_Z:
                            mysectiontype = Section_type.z_cold_Section;
                            mydic.Add(PropName2, mysectiontype);
                            break;


                        case eFramePropType.ConcreteBox:
                            mysectiontype = Section_type.Concrete_rect;
                            mydic.Add(PropName2, mysectiontype);
                            break;


                        case eFramePropType.ConcretePipe:
                            mysectiontype = Section_type.Concrete_circ;
                            mydic.Add(PropName2, mysectiontype);
                            break;
                            ;


                        case eFramePropType.SD:
                            mysectiontype = Section_type.Section_Designer;
                            mydic.Add(PropName2, mysectiontype);
                            break;


                        ////tapered
                        case eFramePropType.Variable:
                            #region defintions
                            int NumberItems = 0;
                            string[] StartSec = null;
                            string[] EndSec = null;
                            double[] MyLength = null;
                            int[] MyType = null;
                            int[] EI33 = null;
                            int[] EI22 = null;
                            int Color = 0;
                            string Notes = null;
                            string GUID = null;
                            #endregion
                            ret = mySapModel.PropFrame.GetNonPrismatic(PropName2, ref NumberItems, ref StartSec, ref EndSec, ref MyLength, ref MyType, ref EI33, ref EI22, ref Color, ref Notes, ref GUID);
                            ret = mySapModel.PropFrame.GetTypeOAPI(StartSec[0], ref PropType);

                            //tapered i or pipe or box
                            switch (PropType)
                            {
                                case eFramePropType.I:
                                    mysectiontype = Section_type.I_Tapered;
                                    mydic.Add(PropName2, mysectiontype);
                                    break;

                                case eFramePropType.Box:
                                    mysectiontype = Section_type.Box_Tapered;
                                    mydic.Add(PropName2, mysectiontype);
                                    break;

                                case eFramePropType.Pipe:
                                    mysectiontype = Section_type.Pipe_Tapered;
                                    mydic.Add(PropName2, mysectiontype);
                                    break;

                                default:
                                    mydic.Add(PropName2, Section_type.Dummy);
                                    break;

                            }


                            break;

                        default:
                            mydic.Add(PropName2, Section_type.Dummy);
                            break;



                    }


                }

                libDic2 = libDic;
                return mydic;
            }

        }


        public static List<string> getSelectedFramesIDS()
        {


            cSapModel mySapModel = Get_sap_mode();
            List<string> result = new List<string>();


            #region defintions 
            double x1, y1, z1;
            x1 = y1 = z1 = 0;
            int NumberItems = 0;
            int[] ObjectType = null;
            string[] ObjectName = null;
            string st_id = "";
            string end_id = "";
            int ret = 0;
            #endregion

            ret = mySapModel.SelectObj.GetSelected(ref NumberItems, ref ObjectType, ref ObjectName);

            for (int i = 0; i < NumberItems; i++)
            {
                if (ObjectType[i] == 2)
                {
                    result.Add(ObjectName[i]);
                }
            }




            return result;
        }

        public static List<string> getSapFrameIDS()
        {

            cSapModel mySapModel = Get_sap_mode();

            #region defintions 
            string[] framesIDS = null;
            int noOfFrames = 0;
            int ret;

            #endregion

            ret = mySapModel.FrameObj.GetNameList(ref noOfFrames, ref framesIDS);

            if (noOfFrames == 0)
            {
                return new List<string>();
            }

            else
            {
                return framesIDS.ToList();
            }
        }

        public static List<Frame> getSapFrames(ref List<Beam> beams, ref List<Column> columns)
        {

            cSapModel mySapModel = Get_sap_mode();
            List<string> frame_ids = getSapFrameIDS();

            if (getSelectedFramesIDS().Count > 0)
            {
                frame_ids = getSelectedFramesIDS();
            }


            List<Frame> sapFrames = new List<Frame>();

            #region defintions 
            string st_point = null;
            string end_point = null;
            int ret;
            double x1, y1, z1, x2, y2, z2;
            x1 = y1 = z1 = x2 = y2 = z2 = 0;
            String sectionName = null;
            string auto = null;
            string shapeName = null;
            eFramePropType shape = eFramePropType.Angle;
            int numberOfGroup = 0;
            string[] assignedGroups = null;
            string materialProperty = null;

            double rotation = 0;
            bool isAdvancedAxes = false;

            string NameInFile = null;
            string FileName = null;

            #endregion


            foreach (string id in frame_ids)
            {
                Frame frame = new Frame();

                try
                {


                    #region getPoints and length
                    ret = mySapModel.FrameObj.GetPoints(id, ref st_point, ref end_point);
                    ret = mySapModel.PointObj.GetCoordCartesian(st_point, ref x1, ref y1, ref z1);
                    ret = mySapModel.PointObj.GetCoordCartesian(end_point, ref x2, ref y2, ref z2);

                    frame.ID = id;
                    frame.StartPoint = new Point(st_point, x1, y1, z1);
                    frame.EndPoint = new Point(end_point, x2, y2, z2);
                    frame.MidPoint = Point.getMidPoint(frame.StartPoint, frame.EndPoint);
                    frame.Length = System.Numerics.Vector3.Distance(Point.ToVector(frame.StartPoint), Point.ToVector(frame.EndPoint));
                    #endregion

                    #region get section
                    mySapModel.FrameObj.GetSection(id, ref sectionName, ref auto);
                    frame.Sec = sectionName;

                    mySapModel.PropFrame.GetTypeOAPI(sectionName, ref shape);

                    if (shape == eFramePropType.Variable)
                    {
                        #region defintions
                        int NumberItems = 0;
                        string[] StartSec = null;
                        string[] EndSec = null;
                        double[] MyLength = null;
                        int[] MyType = null;
                        int[] EI33 = null;
                        int[] EI22 = null;
                        int Color = 0;
                        string Notes = null;
                        string GUID = null;
                        eFramePropType shape2 = eFramePropType.Angle;
                        #endregion

                        ret = mySapModel.PropFrame.GetNonPrismatic(sectionName, ref NumberItems, ref StartSec, ref EndSec, ref MyLength, ref MyType, ref EI33, ref EI22, ref Color, ref Notes, ref GUID);
                        ret = mySapModel.PropFrame.GetTypeOAPI(StartSec[0], ref shape2);
                        sectionName = StartSec[0];

                        //tapered i or pipe or box
                        switch (shape2)
                        {
                            case eFramePropType.I:
                                shapeName = "VARI";

                                break;

                            case eFramePropType.Box:
                                shapeName = "VARBox";
                                break;

                            case eFramePropType.Pipe:
                                shapeName = "VARPipe";
                                break;

                        }



                    }

                    else
                    {
                        shapeName = shape.ToString();
                    }

                    frame.Shape = shapeName;
                    #endregion

                    #region get material and group and rotation angle
                    mySapModel.FrameObj.GetGroupAssign(id, ref numberOfGroup, ref assignedGroups);
                    frame.Group = assignedGroups[0];
                    frame.Groups = assignedGroups.ToList();
                    mySapModel.PropFrame.GetMaterial(sectionName, ref materialProperty);
                    frame.Material = materialProperty;

                    mySapModel.FrameObj.GetLocalAxes(id, ref rotation, ref isAdvancedAxes);

                    #endregion

                    #region getBusOption
                    ret = mySapModel.PropFrame.GetNameInPropFile(sectionName, ref NameInFile, ref FileName, ref materialProperty, ref shape);

                    if (NameInFile == "")
                    {
                        frame.IsBUS = true;
                    }
                    else
                    {

                        frame.IsBUS = false;
                        frame.Lib = FileName;
                    }

                    #endregion

                    #region  get Adanced Rotation 

                    int[] mypDir = new int[1];
                    string[] myplpt = new string[1];
                    double[] myplvector = new double[2];


                    string[] myPoints = new string[2];
                    double[] MyPlVect = new double[3];
                    int[] MyPlDir = new int[2];
                    string st_id = "";
                    string end_id = "";
                    Boolean aa = true;
                    bool Advanced = true;
                    int plane = 0;
                    int plVect = 0;
                    string cys = null;

                    double angleInDegrees = 0;

                    angleInDegrees = rotation;

                    if (isAdvancedAxes)
                    {


                        ret = mySapModel.FrameObj.GetLocalAxesAdvanced(id, ref Advanced, ref plane, ref plVect, ref cys, ref MyPlDir, ref myPoints, ref MyPlVect);

                        st_id = myPoints[0];
                        end_id = myPoints[1];

                        mySapModel.PointObj.GetCoordCartesian(st_id, ref x1, ref y1, ref z1);
                        mySapModel.PointObj.GetCoordCartesian(end_id, ref x2, ref y2, ref z2);

                        Vector3 v1 = new Vector3((float)(x1 - x2), (float)(y1 - y2), 0);


                        angleInDegrees = Math.Atan2(y1 - y2, x1 - x2) * 180 / Math.PI + rotation;


                    }




                    #endregion


                    #region getInsertionPoints

                    int cardinalPoint = 0;
                    bool m2 = false;
                    bool m3 = false;
                    bool stiffTransform = false;
                    double[] offset1 = null;
                    double[] offset2 = null;

                    mySapModel.FrameObj.GetInsertionPoint_1(id, ref cardinalPoint, ref m2, ref m3, ref stiffTransform, ref offset1, ref offset2, ref cys);

                    frame.InsPt = (InsertionPoint)cardinalPoint;
                    frame.StartZOffset = offset1[0];
                    frame.EndZOoffset = offset2[0];

                    #endregion


                    Point.getMemberType(frame);


                    if (frame.MemberType == MemberType.SlanttedColumn || frame.MemberType == MemberType.Beam || frame.MemberType == MemberType.VLBracing)
                    {
                        frame.Beta = -angleInDegrees;
                    }


                    else
                    {
                        frame.Beta = angleInDegrees + 90.00;
                    }


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

                    sapFrames.Add(frame);
                }


                catch (Exception ex2)
                {

                    System.Diagnostics.StackTrace trace = new System.Diagnostics.StackTrace(ex2, true);

                    RevitUtils.Log.Add("Error in getsapFrames Method ... Frame no" + frame.ID + ex2.Message + trace.GetFrame(trace.FrameCount - 1).GetFileLineNumber());

                    // TaskDialog.Show("ErrorInGetRevitFramesMethod", "ErrorInRevitBeamID= "+ instance.Id + ex2.Message + trace.GetFrame(trace.FrameCount - 1).GetFileLineNumber());
                    continue;
                }


            }

            #region Assignment of the supports IDS inside the Sap Frame
            
            bool[] Value = null;

            foreach (var item in sapFrames)
            {
                mySapModel.PointObj.GetRestraint(item.StartPoint.Id, ref Value);

                if (Value != null && Value.Any(v => v))
                {
                    item.BasePlateID = item.StartPoint.Id;
                }


                mySapModel.PointObj.GetRestraint(item.EndPoint.Id, ref Value);

                if (Value != null && Value.Any(v => v))
                {
                    item.BasePlateID = item.EndPoint.Id;
                }


            }

            #endregion


            return sapFrames;
        }
        public static List<Frame> getSapFrames()
        {

            cSapModel mySapModel = Get_sap_mode();
            List<string> frame_ids = getSapFrameIDS();

            if (getSelectedFramesIDS().Count > 0)
            {
                frame_ids = getSelectedFramesIDS();
            }


            List<Frame> sapFrames = new List<Frame>();

            #region defintions 
            string st_point = null;
            string end_point = null;
            int ret;
            double x1, y1, z1, x2, y2, z2;
            x1 = y1 = z1 = x2 = y2 = z2 = 0;
            String sectionName = null;
            string auto = null;
            string shapeName = null;
            eFramePropType shape = eFramePropType.Angle;
            int numberOfGroup = 0;
            string[] assignedGroups = null;
            string materialProperty = null;

            double rotation = 0;
            bool isAdvancedAxes = false;

            string NameInFile = null;
            string FileName = null;

            #endregion


            foreach (string id in frame_ids)
            {
                Frame frame = new Frame();

                try
                {


                    #region getPoints and length
                    ret = mySapModel.FrameObj.GetPoints(id, ref st_point, ref end_point);
                    ret = mySapModel.PointObj.GetCoordCartesian(st_point, ref x1, ref y1, ref z1);
                    ret = mySapModel.PointObj.GetCoordCartesian(end_point, ref x2, ref y2, ref z2);

                    frame.ID = id;
                    frame.StartPoint = new Point(st_point, x1, y1, z1);
                    frame.EndPoint = new Point(end_point, x2, y2, z2);
                    frame.MidPoint = Point.getMidPoint(frame.StartPoint, frame.EndPoint);
                    frame.Length = System.Numerics.Vector3.Distance(Point.ToVector(frame.StartPoint), Point.ToVector(frame.EndPoint));
                    #endregion

                    #region get section
                    mySapModel.FrameObj.GetSection(id, ref sectionName, ref auto);

                    #region get Dimensions for GH
                    // Define and initialize output variables
                    string fileName = "";
                    string matProp = "";
                    double t3 = 0;
                    double t2 = 0;
                    double area = 0;
                    double as2 = 0;
                    double as3 = 0;
                    double torsion = 0;
                    double I22 = 0;
                    double I33 = 0;
                    double S22 = 0;
                    double S33 = 0;
                    double Z22 = 0;
                    double Z33 = 0;
                    double R22 = 0;
                    double R33 = 0;
                    int color = 0;
                    string notes = "";
                    string guid = "";

                    mySapModel.PropFrame.GetGeneral(
                                                    sectionName,
                                                    ref fileName,
                                                    ref matProp,
                                                    ref t3,
                                                    ref t2,
                                                    ref area,
                                                    ref as2,
                                                    ref as3,
                                                    ref torsion,
                                                    ref I22,
                                                    ref I33,
                                                    ref S22,
                                                    ref S33,
                                                    ref Z22,
                                                    ref Z33,
                                                    ref R22,
                                                    ref R33,
                                                    ref color,
                                                    ref notes,
                                                    ref guid
                                                );

                    frame.SecDepth = t3;
                    frame.SecWidth = t2;

                    #endregion

                    frame.Sec = sectionName;
                    frame.ExistSec = sectionName;

                    mySapModel.PropFrame.GetTypeOAPI(sectionName, ref shape);

                    if (shape == eFramePropType.Variable)
                    {
                        #region defintions
                        int NumberItems = 0;
                        string[] StartSec = null;
                        string[] EndSec = null;
                        double[] MyLength = null;
                        int[] MyType = null;
                        int[] EI33 = null;
                        int[] EI22 = null;
                        int Color = 0;
                        string Notes = null;
                        string GUID = null;
                        eFramePropType shape2 = eFramePropType.Angle;
                        #endregion
                        frame.IsNonPresmatic = true;

                        ret = mySapModel.PropFrame.GetNonPrismatic(sectionName, ref NumberItems, ref StartSec, ref EndSec, ref MyLength, ref MyType, ref EI33, ref EI22, ref Color, ref Notes, ref GUID);
                        ret = mySapModel.PropFrame.GetTypeOAPI(StartSec[0], ref shape2);
                        sectionName = StartSec[0];

                        //tapered i or pipe or box
                        switch (shape2)
                        {
                            case eFramePropType.I:
                                shapeName = "VARI";

                                break;

                            case eFramePropType.Box:
                                shapeName = "VARBox";
                                break;

                            case eFramePropType.Pipe:
                                shapeName = "VARPipe";
                                break;

                        }



                    }

                    else
                    {
                        shapeName = shape.ToString();
                        frame.IsNonPresmatic = false;

                    }

                    frame.Shape = shapeName;
                    #endregion

                    #region get material and group and rotation angle
                    mySapModel.FrameObj.GetGroupAssign(id, ref numberOfGroup, ref assignedGroups);
                    frame.Group = assignedGroups[0];
                    frame.Groups = assignedGroups.ToList();
                    mySapModel.PropFrame.GetMaterial(sectionName, ref materialProperty);
                    frame.Material = materialProperty;

                    mySapModel.FrameObj.GetLocalAxes(id, ref rotation, ref isAdvancedAxes);

                    #endregion

                    #region getBusOption
                    ret = mySapModel.PropFrame.GetNameInPropFile(sectionName, ref NameInFile, ref FileName, ref materialProperty, ref shape);

                    if (NameInFile == "")
                    {
                        frame.IsBUS = true;
                    }
                    else
                    {

                        frame.IsBUS = false;
                        frame.Lib = FileName;
                    }

                    #endregion

                    #region  get Adanced Rotation 

                    int[] mypDir = new int[1];
                    string[] myplpt = new string[1];
                    double[] myplvector = new double[2];


                    string[] myPoints = new string[2];
                    double[] MyPlVect = new double[3];
                    int[] MyPlDir = new int[2];
                    string st_id = "";
                    string end_id = "";
                    Boolean aa = true;
                    bool Advanced = true;
                    int plane = 0;
                    int plVect = 0;
                    string cys = null;

                    double angleInDegrees = 0;

                    angleInDegrees = rotation;

                    if (isAdvancedAxes)
                    {


                        ret = mySapModel.FrameObj.GetLocalAxesAdvanced(id, ref Advanced, ref plane, ref plVect, ref cys, ref MyPlDir, ref myPoints, ref MyPlVect);

                        st_id = myPoints[0];
                        end_id = myPoints[1];

                        mySapModel.PointObj.GetCoordCartesian(st_id, ref x1, ref y1, ref z1);
                        mySapModel.PointObj.GetCoordCartesian(end_id, ref x2, ref y2, ref z2);

                        Vector3 v1 = new Vector3((float)(x1 - x2), (float)(y1 - y2), 0);


                        angleInDegrees = Math.Atan2(y1 - y2, x1 - x2) * 180 / Math.PI + rotation;


                    }




                    #endregion

                    #region getInsertionPoints

                    int cardinalPoint = 0;
                    bool m2 = false;
                    bool m3 = false;
                    bool stiffTransform = false;
                    double[] offset1 = null;
                    double[] offset2 = null;

                    mySapModel.FrameObj.GetInsertionPoint_1(id, ref cardinalPoint, ref m2, ref m3, ref stiffTransform, ref offset1, ref offset2, ref cys);

                    frame.InsPt = (InsertionPoint)cardinalPoint;

                    #endregion

                    Point.getMemberType(frame);


                    if (frame.MemberType == MemberType.SlanttedColumn || frame.MemberType == MemberType.Beam || frame.MemberType == MemberType.VLBracing)
                    {
                        frame.Beta = -angleInDegrees;
                    }


                    else
                    {
                        frame.Beta = angleInDegrees + 90.00;
                    }


                    sapFrames.Add(frame);
                }


                catch (Exception ex2)
                {

                    System.Diagnostics.StackTrace trace = new System.Diagnostics.StackTrace(ex2, true);

                    RevitUtils.Log.Add("Error in getsapFrames Method ... Frame no" + frame.ID + ex2.Message + trace.GetFrame(trace.FrameCount - 1).GetFileLineNumber());

                    // TaskDialog.Show("ErrorInGetRevitFramesMethod", "ErrorInRevitBeamID= "+ instance.Id + ex2.Message + trace.GetFrame(trace.FrameCount - 1).GetFileLineNumber());
                    continue;
                }


            }


            return sapFrames;
        }

        public static List<Frame> getSapFramesForRhino(ref List<Beam> beams, ref List<Column> columns)
        {

            cSapModel mySapModel = Get_sap_mode();
            List<string> frame_ids = getSapFrameIDS();

            if (getSelectedFramesIDS().Count > 0)
            {
                frame_ids = getSelectedFramesIDS();
            }


            List<Frame> sapFrames = new List<Frame>();

            #region defintions 
            string st_point = null;
            string end_point = null;
            int ret;
            double x1, y1, z1, x2, y2, z2;
            x1 = y1 = z1 = x2 = y2 = z2 = 0;
            String sectionName = null;
            string auto = null;
            string shapeName = null;
            eFramePropType shape = eFramePropType.Angle;
            int numberOfGroup = 0;
            string[] assignedGroups = null;
            string materialProperty = null;

            double rotation = 0;
            bool isAdvancedAxes = false;

            string NameInFile = null;
            string FileName = null;

            #endregion


            foreach (string id in frame_ids)
            {
                Frame frame = new Frame();

                try
                {


                    #region getPoints and length
                    ret = mySapModel.FrameObj.GetPoints(id, ref st_point, ref end_point);
                    ret = mySapModel.PointObj.GetCoordCartesian(st_point, ref x1, ref y1, ref z1);
                    ret = mySapModel.PointObj.GetCoordCartesian(end_point, ref x2, ref y2, ref z2);

                    frame.ID = id;
                    frame.StartPoint = new Point(st_point, x1, y1, z1);
                    frame.EndPoint = new Point(end_point, x2, y2, z2);
                    frame.MidPoint = Point.getMidPoint(frame.StartPoint, frame.EndPoint);
                    frame.Length = System.Numerics.Vector3.Distance(Point.ToVector(frame.StartPoint), Point.ToVector(frame.EndPoint));
                    #endregion

                    #region get section
                    mySapModel.FrameObj.GetSection(id, ref sectionName, ref auto);
                    frame.Sec = sectionName;

                    mySapModel.PropFrame.GetTypeOAPI(sectionName, ref shape);

                    if (shape == eFramePropType.Variable)
                    {
                        #region defintions
                        int NumberItems = 0;
                        string[] StartSec = null;
                        string[] EndSec = null;
                        double[] MyLength = null;
                        int[] MyType = null;
                        int[] EI33 = null;
                        int[] EI22 = null;
                        int Color = 0;
                        string Notes = null;
                        string GUID = null;
                        eFramePropType shape2 = eFramePropType.Angle;
                        #endregion

                        ret = mySapModel.PropFrame.GetNonPrismatic(sectionName, ref NumberItems, ref StartSec, ref EndSec, ref MyLength, ref MyType, ref EI33, ref EI22, ref Color, ref Notes, ref GUID);
                        ret = mySapModel.PropFrame.GetTypeOAPI(StartSec[0], ref shape2);
                        sectionName = StartSec[0];

                        //tapered i or pipe or box
                        switch (shape2)
                        {
                            case eFramePropType.I:
                                shapeName = "VARI";

                                break;

                            case eFramePropType.Box:
                                shapeName = "VARBox";
                                break;

                            case eFramePropType.Pipe:
                                shapeName = "VARPipe";
                                break;

                        }



                    }

                    else
                    {
                        shapeName = shape.ToString();
                    }

                    frame.Shape = shapeName;
                    #endregion

                    #region get material and group and rotation angle
                    mySapModel.FrameObj.GetGroupAssign(id, ref numberOfGroup, ref assignedGroups);
                    frame.Group = assignedGroups[0];
                    frame.Groups = assignedGroups.ToList();
                    mySapModel.PropFrame.GetMaterial(sectionName, ref materialProperty);
                    frame.Material = materialProperty;

                    mySapModel.FrameObj.GetLocalAxes(id, ref rotation, ref isAdvancedAxes);

                    #endregion

                    #region getBusOption
                    ret = mySapModel.PropFrame.GetNameInPropFile(sectionName, ref NameInFile, ref FileName, ref materialProperty, ref shape);

                    if (NameInFile == "")
                    {
                        frame.IsBUS = true;
                    }
                    else
                    {

                        frame.IsBUS = false;
                        frame.Lib = FileName;
                    }

                    #endregion

                    #region  get Adanced Rotation 

                    int[] mypDir = new int[1];
                    string[] myplpt = new string[1];
                    double[] myplvector = new double[2];


                    string[] myPoints = new string[2];
                    double[] MyPlVect = new double[3];
                    int[] MyPlDir = new int[2];
                    string st_id = "";
                    string end_id = "";
                    Boolean aa = true;
                    bool Advanced = true;
                    int plane = 0;
                    int plVect = 0;
                    string cys = null;

                    double angleInDegrees = 0;

                    angleInDegrees = rotation;

                    if (isAdvancedAxes)
                    {


                        ret = mySapModel.FrameObj.GetLocalAxesAdvanced(id, ref Advanced, ref plane, ref plVect, ref cys, ref MyPlDir, ref myPoints, ref MyPlVect);

                        st_id = myPoints[0];
                        end_id = myPoints[1];

                        mySapModel.PointObj.GetCoordCartesian(st_id, ref x1, ref y1, ref z1);
                        mySapModel.PointObj.GetCoordCartesian(end_id, ref x2, ref y2, ref z2);

                        Vector3 v1 = new Vector3((float)(x1 - x2), (float)(y1 - y2), 0);


                        angleInDegrees = Math.Atan2(y1 - y2, x1 - x2) * 180 / Math.PI + rotation;


                    }




                    #endregion


                    #region GetLoads
                     ret = 0;
                    int NumberItems2 = 0;
                    string[] FrameName = null;
                    string[] LoadPat = null;
                    int[] MyType2 = null;
                    string[] CSys = null;
                    int[] Dir = null;
                    double[] RD1 = null;
                    double[] RD2 = null;
                    double[] Dist1 = null;
                    double[] Dist2 = null;
                    double[] Val1 = null;
                    double[] Val2 = null;


                    ret = mySapModel.FrameObj.GetLoadDistributed(id, ref NumberItems2, ref FrameName, ref LoadPat, ref MyType2, ref CSys, ref Dir, ref RD1, ref RD2, ref Dist1, ref Dist2, ref Val1, ref Val2);


                    string DistrLoadsAssignments = $"{ret},{NumberItems2},{string.Join(",", FrameName)},{string.Join(",", LoadPat)},{string.Join(",", MyType2)},{string.Join(",", CSys)},{string.Join(",", Dir)},{string.Join(",", RD1)},{string.Join(",", RD2)},{string.Join(",", Dist1)},{string.Join(",", Dist2)},{string.Join(",", Val1)},{string.Join(",", Val2)}";

                    frame.distLoadAssignments = DistrLoadsAssignments;


                    bool[] ii = null;
                    bool[] jj = null;
                    double[] StartValue = null;
                    double[] EndValue = null;

                    ret = mySapModel.FrameObj.GetReleases(id, ref ii, ref jj, ref StartValue, ref EndValue);
                    string releasesAssignmetns = $"{string.Join(",", ii.Select(b => b.ToString()))};{string.Join(",", jj.Select(b => b.ToString()))};{string.Join(",", StartValue)};{string.Join(",", EndValue)}";



                    #endregion

                    Point.getMemberType(frame);


                    if (frame.MemberType == MemberType.SlanttedColumn || frame.MemberType == MemberType.Beam || frame.MemberType == MemberType.VLBracing)
                    {
                        frame.Beta = -angleInDegrees;
                    }


                    else
                    {
                        frame.Beta = angleInDegrees + 90.00;
                    }


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

                    sapFrames.Add(frame);
                }


                catch (Exception ex2)
                {

                    System.Diagnostics.StackTrace trace = new System.Diagnostics.StackTrace(ex2, true);

                    RevitUtils.Log.Add("Error in getsapFrames Method ... Frame no" + frame.ID + ex2.Message + trace.GetFrame(trace.FrameCount - 1).GetFileLineNumber());

                    // TaskDialog.Show("ErrorInGetRevitFramesMethod", "ErrorInRevitBeamID= "+ instance.Id + ex2.Message + trace.GetFrame(trace.FrameCount - 1).GetFileLineNumber());
                    continue;
                }


            }


            return sapFrames;
        }

        public static Frame getSapSelectedFrame()
        {

            cSapModel mySapModel = Get_sap_mode();
            Frame result = new Frame();

            #region defintions 
            double x1, y1, z1;
            x1 = y1 = z1 = 0;
            int NumberItems = 0;
            int[] ObjectType = null;
            string[] ObjectName = null;
            string st_id = "";
            string end_id = "";
            int ret = 0;
            #endregion


            ret = mySapModel.SelectObj.GetSelected(ref NumberItems, ref ObjectType, ref ObjectName);
            result.ID = ObjectName[0];
            mySapModel.FrameObj.GetPoints(ObjectName[0], ref st_id, ref end_id);


            ret = mySapModel.PointObj.GetCoordCartesian(st_id, ref x1, ref y1, ref z1);
            result.StartPoint.X = Math.Round(x1, 0);
            result.StartPoint.Y = Math.Round(y1, 0);
            result.StartPoint.Z = Math.Round(z1, 0);


            ret = mySapModel.PointObj.GetCoordCartesian(end_id, ref x1, ref y1, ref z1);
            result.EndPoint.X = Math.Round(x1, 0);
            result.EndPoint.Y = Math.Round(y1, 0);
            result.EndPoint.Z = Math.Round(z1, 0);

            result.MidPoint = Point.getMidPoint(result.StartPoint, result.EndPoint);

            return result;


        }

        public static string createSapBeam(Frame revitFrame, cSapModel mysapmodel, Document doc, string type, Dictionary<string, string> secDicForBeams, Dictionary<string, string> secDicForColumns)
        {
            #region defintions

            int ret;
            string id = revitFrame.RvtID;
            string sapSection = null;
            string materialName = null;

            #endregion


            ret = mysapmodel.FrameObj.AddByCoord(revitFrame.StartPoint.X, revitFrame.StartPoint.Y, revitFrame.StartPoint.Z, revitFrame.EndPoint.X, revitFrame.EndPoint.Y, revitFrame.EndPoint.Z, ref id);

            if (type == "BM")
            {
                sapSection = secDicForBeams[revitFrame.RvtName];
            }
            else
            {
                sapSection = secDicForColumns[revitFrame.RvtName];
            }

            ret = mysapmodel.PropFrame.GetMaterial(sapSection, ref materialName);

            ret = mysapmodel.FrameObj.SetSection(id, sapSection);
            if (ret == 1)
            {
                mysapmodel.FrameObj.SetSection(id, "dummySection");
            }

            ret = mysapmodel.FrameObj.SetLocalAxes(id, revitFrame.Beta);





            // set mark in revit 
            FamilyInstance instance = (FamilyInstance)doc.GetElement(revitFrame.revitEleId);
            using (Transaction trns = new Transaction(doc, "aa"))
            {
                trns.Start();
                FailureHandlingOptions failOpt = trns.GetFailureHandlingOptions();
                Frame newBeam = new Frame();
                newBeam.ID = id;
                newBeam.Material = materialName;
                RevitUtils.setParametersSAP(instance, newBeam);

                trns.Commit();
            }


            return id;
        }

        public static string createSapBeam2(ref Frame revitFrame, cSapModel mysapmodel, Document doc, string type, Dictionary<string, string> secDicForBeams, Dictionary<string, string> secDicForColumns)
        {
            #region defintions

            int ret;
            string id = revitFrame.RvtID;
            string sapSection = null;
            string materialName = null;

            #endregion


            ret = mysapmodel.FrameObj.AddByCoord(revitFrame.StartPoint.X, revitFrame.StartPoint.Y, revitFrame.StartPoint.Z, revitFrame.EndPoint.X, revitFrame.EndPoint.Y, revitFrame.EndPoint.Z, ref id);

            if (type == "BM")
            {
                sapSection = secDicForBeams[revitFrame.RvtName];
            }
            else
            {
                sapSection = secDicForColumns[revitFrame.RvtName];
            }

            ret = mysapmodel.PropFrame.GetMaterial(sapSection, ref materialName);

            ret = mysapmodel.FrameObj.SetSection(id, sapSection);
            if (ret == 1)
            {
                mysapmodel.FrameObj.SetSection(id, "dummySection");
            }

            ret = mysapmodel.FrameObj.SetLocalAxes(id, revitFrame.Beta);





            // set mark in revit 
            FamilyInstance instance = (FamilyInstance)doc.GetElement(revitFrame.revitEleId);
            using (Transaction trns = new Transaction(doc, "aa"))
            {
                trns.Start();
                FailureHandlingOptions failOpt = trns.GetFailureHandlingOptions();
                Frame newBeam = new Frame();
                newBeam.ID = id;
                newBeam.Material = materialName;
                RevitUtils.setParametersSAP(instance, newBeam);

                trns.Commit();
            }

            #region set new data in Revit Element
            string stPt = null;
            string endPt = null;
            eFramePropType PropType = eFramePropType.SD;




            mysapmodel.FrameObj.GetPoints(id, ref stPt, ref endPt);
            mysapmodel.PropFrame.GetTypeOAPI(sapSection, ref PropType);

            revitFrame.Sec = sapSection;
            revitFrame.Material = materialName;
            revitFrame.ID = id;
            revitFrame.Shape = PropType.ToString();
            revitFrame.StartPoint.Id = stPt;
            revitFrame.EndPoint.Id = endPt;


            #endregion

            return id;
        }

        public static string createSapColumn(Frame revitFrame, cSapModel mysapmodel, Document doc, string type, Dictionary<string, string> secDicForBeams, Dictionary<string, string> secDicForColumns)
        {
            #region defintions

            int ret;
            string id = revitFrame.RvtID;
            string sapSection = null;
            string materialName = null;

            #endregion


            ret = mysapmodel.FrameObj.AddByCoord(revitFrame.StartPoint.X, revitFrame.StartPoint.Y, revitFrame.StartPoint.Z, revitFrame.EndPoint.X, revitFrame.EndPoint.Y, revitFrame.EndPoint.Z, ref id);

            if (type == "BM")
            {
                sapSection = secDicForBeams[revitFrame.RvtName];
            }
            else
            {
                sapSection = secDicForColumns[revitFrame.RvtName];
            }

            ret = mysapmodel.FrameObj.SetSection(id, sapSection);
            if (ret == 1)
            {
                mysapmodel.FrameObj.SetSection(id, "dummySection");
            }

            ret = mysapmodel.FrameObj.SetLocalAxes(id, revitFrame.Beta);
            ret = mysapmodel.PropFrame.GetMaterial(sapSection, ref materialName);

            // set mark in revit 
            FamilyInstance instance = (FamilyInstance)doc.GetElement(revitFrame.revitEleId);
            using (Transaction trns = new Transaction(doc, "aa"))
            {
                trns.Start();
                FailureHandlingOptions failOpt = trns.GetFailureHandlingOptions();

                Frame newColumn = new Frame();
                newColumn.ID = id;
                newColumn.Material = materialName;
                RevitUtils.setParametersSAP(instance, newColumn);

                trns.Commit();
            }

            return id;
        }
        public static string createSapColumn2(ref Frame revitFrame, cSapModel mysapmodel, Document doc, string type, Dictionary<string, string> secDicForBeams, Dictionary<string, string> secDicForColumns)
        {
            #region defintions

            int ret;
            string id = revitFrame.RvtID;
            string sapSection = null;
            string materialName = null;

            #endregion


            ret = mysapmodel.FrameObj.AddByCoord(revitFrame.StartPoint.X, revitFrame.StartPoint.Y, revitFrame.StartPoint.Z, revitFrame.EndPoint.X, revitFrame.EndPoint.Y, revitFrame.EndPoint.Z, ref id);

            if (type == "BM")
            {
                sapSection = secDicForBeams[revitFrame.RvtName];
            }
            else
            {
                sapSection = secDicForColumns[revitFrame.RvtName];
            }

            ret = mysapmodel.FrameObj.SetSection(id, sapSection);
            if (ret == 1)
            {
                mysapmodel.FrameObj.SetSection(id, "dummySection");
            }

            ret = mysapmodel.FrameObj.SetLocalAxes(id, revitFrame.Beta);
            ret = mysapmodel.PropFrame.GetMaterial(sapSection, ref materialName);

            // set mark in revit 
            FamilyInstance instance = (FamilyInstance)doc.GetElement(revitFrame.revitEleId);
            using (Transaction trns = new Transaction(doc, "aa"))
            {
                trns.Start();
                FailureHandlingOptions failOpt = trns.GetFailureHandlingOptions();

                Frame newColumn = new Frame();
                newColumn.ID = id;
                newColumn.Material = materialName;
                RevitUtils.setParametersSAP(instance, newColumn);

                trns.Commit();
            }

            #region set new data in Revit Element
            string stPt = null;
            string endPt = null;
            eFramePropType PropType = eFramePropType.SD;




            mysapmodel.FrameObj.GetPoints(id, ref stPt, ref endPt);
            mysapmodel.PropFrame.GetTypeOAPI(sapSection, ref PropType);

            revitFrame.Sec = sapSection;
            revitFrame.Material = materialName;
            revitFrame.ID = id;
            revitFrame.Shape = PropType.ToString();
            revitFrame.StartPoint.Id = stPt;
            revitFrame.EndPoint.Id = endPt;


            #endregion


            return id;
        }



        #region RelatedToSections
        public static Dictionary<string, string> createSapSectionForBeams2(List<Frame> revitFrmaes, cSapModel mysapmodel)
        {
            #region definitions 
            string sec;
            string sapSection = null;
            #endregion
            List<object> sapSecNameModifiedSapSecName = ChangeSapSectionNames.GetSapSectionsDictionary(mysapmodel, ref revitFrmaes);
            List<string> updatedProfilesName = sapSecNameModifiedSapSecName.Cast<dynamic>().Select(x => x.newSec).Cast<string>().ToList();
            List<Frame> filteredRevitFrames = revitFrmaes.GroupBy(x => x.RvtName).Select(g => g.First()).ToList();

            Dictionary<string, string> dic = new Dictionary<string, string>();

            foreach (Frame revitFrame in filteredRevitFrames)
            {
                try
                {
                    sec = updatedProfilesName.AsParallel().FirstOrDefault(x => x == revitFrame.RvtName);
                    if (sec != null)
                    {
                        revitFrame.Sec = sec;
                        sapSection = sec;
                        dic.Add(revitFrame.RvtName, sapSecNameModifiedSapSecName.Cast<dynamic>().First(x => x.newSec == sec).SapSec);
                        continue;
                    }

                    else
                    {
                        sec = createSapSectionForBeam(revitFrame, mysapmodel);
                        revitFrame.Sec = sec;
                        sapSection = sec;
                        dic.Add(revitFrame.RvtName, sec);
                        continue;
                    }
                }

                catch (Exception)
                {

                    continue;
                }



            }



            return dic;
        }
        public static Dictionary<string, string> createSapSectionForColumns2(List<Frame> revitFrmaes, cSapModel mysapmodel)
        {
            #region definitions 
            string sec;
            string sapSection = null;
            #endregion


            List<object> sapSecNameModifiedSapSecName = ChangeSapSectionNames.GetSapSectionsDictionary(mysapmodel, ref revitFrmaes);
            List<string> updatedProfilesName = sapSecNameModifiedSapSecName.Cast<dynamic>().Select(x => x.newSec).Cast<string>().ToList();
            List<Frame> filteredRevitFrames = revitFrmaes.GroupBy(x => x.RvtName).Select(g => g.First()).ToList();

            Dictionary<string, string> dic = new Dictionary<string, string>();


            foreach (Frame revitFrame in filteredRevitFrames)
            {
                try
                {
                    sec = updatedProfilesName.AsParallel().FirstOrDefault(x => x == revitFrame.RvtName);
                    if (sec != null)
                    {
                        revitFrame.Sec = sec;
                        sapSection = sec;
                        dic.Add(revitFrame.RvtName, sapSecNameModifiedSapSecName.Cast<dynamic>().First(x => x.newSec == sec).SapSec);
                        continue;
                    }

                    else
                    {
                        sec = createSapSectionForColumn(revitFrame, mysapmodel);
                        revitFrame.Sec = sec;
                        sapSection = sec;
                        dic.Add(revitFrame.RvtName, sec);
                        continue;
                    }
                }


                catch (Exception)
                {

                    continue;
                }



            }



            return dic;
        }

        public static string createSapSectionForBeam(Frame revitFrame, cSapModel mysapmodel)
        {
            #region defintions
            int ret;
            string id = revitFrame.RvtID;
            char cc;
            int lastSpacecharacterIndex;
            string dimentionswithX;
            string sec;
            string[] dimentions;
            string material = revitFrame.Material;
            List<string> profiles = Sections.GetAllNamesAssignedToFramesInSap(mysapmodel);

            if (material == null || material == "" || material == "STEEL")
            {
                int numberOfNames = 0;
                string[] materialList = new string[0];
                ret = mysapmodel.PropMaterial.GetNameList(ref numberOfNames, ref materialList);
                eMatType matType = eMatType.Steel;
                int a = 0;
                foreach (var mat in materialList)
                {

                    mysapmodel.PropMaterial.GetTypeOAPI(mat, ref matType, ref a);

                    if (matType == eMatType.Steel)
                    {
                        material = mat;
                        break;
                    }
                }

            }



            #endregion

            sec = profiles.AsParallel().FirstOrDefault(x => x == revitFrame.RvtName);

            if (sec != null)
            {
                revitFrame.Sec = sec;
                return sec;
            }

            try
            {
                switch (revitFrame.FamilyType)
                {


                    #region I_B
                    case SectType.BM_BUS_I:

                        #region getDimentions 
                        int seperatorIndex = revitFrame.RvtName.LastIndexOf("+WEB");

                        string webString = revitFrame.RvtName.Substring(seperatorIndex + 4);
                        string flgString = revitFrame.RvtName.Substring(0, seperatorIndex - 1);


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
                        ret = mysapmodel.PropFrame.SetISection(revitFrame.RvtName, material, double.Parse(consWebDepth), double.Parse(consFlgDepth), double.Parse(flgThick), double.Parse(webThick), double.Parse(consFlgDepth), double.Parse(flgThick));
                        revitFrame.Sec = revitFrame.RvtName;
                        return revitFrame.Sec;
                        #endregion

                        #region I_H
                        string numeric = null;

                    case SectType.BM_HEA:
                        numeric = new String(revitFrame.RvtName.Where(Char.IsDigit).ToArray());
                        revitFrame.Sec = "HE" + numeric + "A";
                        ret = mysapmodel.PropFrame.ImportProp(revitFrame.Sec, material, "ArcelorMittal_Europe.xml", revitFrame.Sec);
                        return revitFrame.Sec;


                    case SectType.BM_HEAA:
                        numeric = new String(revitFrame.RvtName.Where(Char.IsDigit).ToArray());
                        revitFrame.Sec = "HE" + numeric + "AA";
                        ret = mysapmodel.PropFrame.ImportProp(revitFrame.Sec, material, "ArcelorMittal_Europe.xml", revitFrame.Sec);
                        return revitFrame.Sec;



                    case SectType.BM_HEB:
                        numeric = new String(revitFrame.RvtName.Where(Char.IsDigit).ToArray());
                        revitFrame.Sec = "HE" + numeric + "B";
                        ret = mysapmodel.PropFrame.ImportProp(revitFrame.Sec, material, "ArcelorMittal_Europe.xml", revitFrame.Sec);
                        return revitFrame.Sec;


                    case SectType.BM_HEM:
                        numeric = new String(revitFrame.RvtName.Where(Char.IsDigit).ToArray());
                        revitFrame.Sec = "HE" + numeric + "M";
                        ret = mysapmodel.PropFrame.ImportProp(revitFrame.Sec, material, "ArcelorMittal_Europe.xml", revitFrame.Sec);
                        return revitFrame.Sec;


                    case SectType.BM_IPE:
                        numeric = new String(revitFrame.RvtName.Where(Char.IsDigit).ToArray());
                        revitFrame.Sec = "IPE" + numeric;
                        ret = mysapmodel.PropFrame.ImportProp(revitFrame.Sec, material, "ArcelorMittal_Europe.xml", revitFrame.Sec);
                        return revitFrame.Sec;


                    case SectType.BM_IPEA:
                        numeric = new String(revitFrame.RvtName.Where(Char.IsDigit).ToArray());
                        revitFrame.Sec = "IPEA" + numeric;
                        ret = mysapmodel.PropFrame.ImportProp(revitFrame.Sec, material, "ArcelorMittal_Europe.xml", revitFrame.Sec);
                        return revitFrame.Sec;


                    case SectType.BM_IPEO:
                        numeric = new String(revitFrame.RvtName.Where(Char.IsDigit).ToArray());
                        revitFrame.Sec = "IPEO" + numeric;
                        ret = mysapmodel.PropFrame.ImportProp(revitFrame.Sec, material, "ArcelorMittal_Europe.xml", revitFrame.Sec);
                        return revitFrame.Sec;

                    #endregion

                    #region pipe
                    //case SectType.BM_PIPE:
                    //    cc = revitFrame.RvtName.SkipWhile(c => !char.IsDigit(c)).ToList()[0];
                    //    lastSpacecharacterIndex = revitFrame.RvtName.IndexOf(cc);
                    //    dimentionswithX = revitFrame.RvtName.Substring(lastSpacecharacterIndex);
                    //    sec = profiles.AsParallel().FirstOrDefault(x => x.Contains(dimentionswithX));

                    //    if (sec != null)
                    //    {
                    //        revitFrame.Sec = sec;
                    //        return sec;
                    //    }

                    //    else
                    //    {
                    //        dimentions = dimentionswithX.Split(("X".ToCharArray()));
                    //        ret = mysapmodel.PropFrame.SetPipe(revitFrame.RvtName, material, double.Parse(dimentions[0]), double.Parse(dimentions[1]));
                    //        revitFrame.Sec = revitFrame.RvtName;
                    //        return revitFrame.Sec;
                    //    }

                    case SectType.BM_PIPE:
                        cc = revitFrame.RvtName.SkipWhile(c => !char.IsDigit(c)).ToList()[0];
                        lastSpacecharacterIndex = revitFrame.RvtName.IndexOf(cc);
                        dimentionswithX = revitFrame.RvtName.Substring(lastSpacecharacterIndex);
                        sec = profiles.AsParallel().FirstOrDefault(x => x.Contains(dimentionswithX));


                        dimentions = dimentionswithX.Split(("X".ToCharArray()));
                        ret = mysapmodel.PropFrame.SetPipe(revitFrame.RvtName, material, double.Parse(dimentions[0]), double.Parse(dimentions[1]));
                        revitFrame.Sec = revitFrame.RvtName;
                        return revitFrame.Sec;


                    #endregion

                    #region BOX(TUBE)

                    //case SectType.BM_TUBE:
                    //    cc = revitFrame.RvtName.SkipWhile(c => !char.IsDigit(c)).ToList()[0];
                    //    lastSpacecharacterIndex = revitFrame.RvtName.IndexOf(cc);
                    //    dimentionswithX = revitFrame.RvtName.Substring(lastSpacecharacterIndex);
                    //    int noOfX2 = revitFrame.RvtName.Count(x => x == 'X');
                    //    sec = profiles.AsParallel().FirstOrDefault(x => x.Contains(dimentionswithX) && x.Count(a => a == 'X') == noOfX2);

                    //    if (sec != null)
                    //    {
                    //        revitFrame.Sec = revitFrame.RvtName;
                    //        return sec;
                    //    }

                    //    else
                    //    {
                    //        dimentions = dimentionswithX.Split(("X".ToCharArray()));
                    //        if (dimentions.Length == 4)
                    //        {
                    //            ret = mysapmodel.PropFrame.SetTube(revitFrame.RvtName, material, double.Parse(dimentions[0]), double.Parse(dimentions[1]), double.Parse(dimentions[3]), double.Parse(dimentions[2]));
                    //        }


                    //        else
                    //        {
                    //            ret = mysapmodel.PropFrame.SetTube(revitFrame.RvtName, material, double.Parse(dimentions[0]), double.Parse(dimentions[1]), double.Parse(dimentions[2]), double.Parse(dimentions[2]));

                    //        }

                    //        revitFrame.Sec = revitFrame.RvtName;
                    //        return revitFrame.RvtName;
                    //    }


                    case SectType.BM_TUBE:
                        cc = revitFrame.RvtName.SkipWhile(c => !char.IsDigit(c)).ToList()[0];
                        lastSpacecharacterIndex = revitFrame.RvtName.IndexOf(cc);
                        dimentionswithX = revitFrame.RvtName.Substring(lastSpacecharacterIndex);
                        int noOfX2 = revitFrame.RvtName.Count(x => x == 'X');

                        dimentions = dimentionswithX.Split(("X".ToCharArray()));
                        if (dimentions.Length == 4)
                        {
                            ret = mysapmodel.PropFrame.SetTube(revitFrame.RvtName, material, double.Parse(dimentions[0]), double.Parse(dimentions[1]), double.Parse(dimentions[3]), double.Parse(dimentions[2]));
                        }


                        else
                        {
                            ret = mysapmodel.PropFrame.SetTube(revitFrame.RvtName, material, double.Parse(dimentions[0]), double.Parse(dimentions[1]), double.Parse(dimentions[2]), double.Parse(dimentions[2]));

                        }

                        revitFrame.Sec = revitFrame.RvtName;
                        return revitFrame.RvtName;

                    #endregion

                    #region UPN
                    case SectType.BM_UPN:
                        numeric = new String(revitFrame.RvtName.Where(Char.IsDigit).ToArray());
                        revitFrame.Sec = "UPN" + numeric;
                        ret = mysapmodel.PropFrame.ImportProp(revitFrame.Sec, material, "Euro.xml", revitFrame.Sec);
                        return revitFrame.Sec;

                    #endregion

                    #region W

                    case SectType.BM_W:
                        int noOfX = Frame.CountCharsUsingIndex(revitFrame.RvtName, 'X');
                        dimentionswithX = new string(revitFrame.RvtName.SkipWhile(c => !char.IsDigit(c)).ToArray());
                        int indexOfSecondX = Frame.IndexOfNth(dimentionswithX, "X", 0);
                        int lastXindex = revitFrame.RvtName.LastIndexOf("X");

                        string firsapimention = null;
                        string secondDimention = null;

                        string lastDimension = revitFrame.RvtName.Substring(lastXindex + 1);

                        if (noOfX == 2)
                        {
                            firsapimention = new string(dimentionswithX.TakeWhile(c => char.IsDigit(c)).ToArray());
                            secondDimention = new string(dimentionswithX.Substring(indexOfSecondX).SkipWhile(c => !Char.IsDigit(c)).TakeWhile(c => char.IsDigit(c)).ToArray());

                            if (firsapimention == secondDimention)
                            {
                                revitFrame.Sec = "W" + firsapimention + "X" + lastDimension;
                            }
                            else
                            {
                                revitFrame.Sec = "W" + dimentionswithX;
                            }
                        }

                        else
                        {
                            revitFrame.Sec = "W" + dimentionswithX;
                        }


                        ret = mysapmodel.PropFrame.ImportProp(revitFrame.Sec, material, "AISC14.xml", revitFrame.Sec);
                        return revitFrame.Sec;



                    //cc = revitFrame.RvtName.SkipWhile(c => !char.IsDigit(c)).ToList()[0];
                    //lastSpacecharacterIndex = revitFrame.RvtName.IndexOf(cc);
                    //dimentionswithX = revitFrame.RvtName.Substring(lastSpacecharacterIndex);
                    //revitFrame.Sec = "W" + dimentionswithX;
                    //ret = mysapmodel.PropFrame.ImportProp(revitFrame.Sec, material, "AISC15.xml", revitFrame.Sec);
                    //return revitFrame.Sec;

                    #endregion

                    #region UB
                    case SectType.BM_UB:
                        cc = revitFrame.RvtName.SkipWhile(c => !char.IsDigit(c)).ToList()[0];
                        lastSpacecharacterIndex = revitFrame.RvtName.IndexOf(cc);
                        dimentionswithX = revitFrame.RvtName.Substring(lastSpacecharacterIndex);
                        revitFrame.Sec = "UKB" + dimentionswithX;
                        ret = mysapmodel.PropFrame.ImportProp(revitFrame.Sec, material, "BSShapes2006.xml", revitFrame.Sec);
                        return revitFrame.Sec;
                    #endregion

                    #region UC
                    case SectType.BM_UC:
                        cc = revitFrame.RvtName.SkipWhile(c => !char.IsDigit(c)).ToList()[0];
                        lastSpacecharacterIndex = revitFrame.RvtName.IndexOf(cc);
                        dimentionswithX = revitFrame.RvtName.Substring(lastSpacecharacterIndex);
                        revitFrame.Sec = "UKC" + dimentionswithX;
                        ret = mysapmodel.PropFrame.ImportProp(revitFrame.Sec, material, "BSShapes2006.xml", revitFrame.Sec);
                        return revitFrame.Sec;
                    #endregion

                    #region I_T
                    case SectType.BM_BUS_I_Tapered:

                        string[] StartSec = new string[1];
                        string[] EndSec = new string[1];
                        double[] MyLength = new double[1];
                        int[] MyType = new int[1];
                        int[] EI33 = new int[1];
                        int[] EI22 = new int[1];

                        #region getDimentions 

                        seperatorIndex = revitFrame.RvtName.LastIndexOf("+WEB");

                        webString = revitFrame.RvtName.Substring(seperatorIndex + 4);
                        flgString = revitFrame.RvtName.Substring(0, seperatorIndex - 1);


                        webstring2 = webString.Split(("X".ToCharArray()));
                        webThick = webstring2[1];


                        flgString2 = flgString.Split(("X".ToCharArray()));
                        flgThick = flgString2[1];

                        string startWebDepth = null;
                        string endWebDepth = null;
                        string largeWebDepthWithoutTF = null;
                        string smallWebDepthWithoutTF = null;



                        consFlgDepth = null;


                        // get largeDepth ,smalldepth of the web and constWidth of the flange
                        string[] webstring3 = webstring2[0].Split(("/".ToCharArray()));


                        largeWebDepthWithoutTF = (Convert.ToDouble(new string((webstring3[0].SkipWhile(c => !char.IsDigit(c))
                                               .TakeWhile(c => char.IsDigit(c)).ToArray()
                                              ))).ToString());




                        startWebDepth = (Convert.ToDouble(new string((webstring3[0].SkipWhile(c => !char.IsDigit(c))
                                           .TakeWhile(c => char.IsDigit(c)).ToArray()
                                          ))) + 2 * Convert.ToDouble(flgThick)).ToString();


                        smallWebDepthWithoutTF = (Convert.ToDouble(new string((webstring3[1].SkipWhile(c => !char.IsDigit(c))
                                                .TakeWhile(c => char.IsDigit(c)).ToArray()
                                                ))).ToString());


                        endWebDepth = (Convert.ToDouble(new string((webstring3[1].SkipWhile(c => !char.IsDigit(c))
                                           .TakeWhile(c => char.IsDigit(c)).ToArray()
                                          ))) + 2 * Convert.ToDouble(flgThick)).ToString();


                        consFlgDepth = new string((flgString2[0].Substring(2).SkipWhile(c => !char.IsDigit(c))
                                           .TakeWhile(c => char.IsDigit(c)).ToArray()
                                           ));

                        #endregion

                        string smallSection = "2FL. PL. " + consFlgDepth + "X" + flgThick + "  +" + "WEB PL. " + largeWebDepthWithoutTF + "X" + webThick;
                        string largeSection = "2FL. PL. " + consFlgDepth + "X" + flgThick + "  +" + "WEB PL. " + smallWebDepthWithoutTF + "X" + webThick;
                        // smallSection
                        ret = mysapmodel.PropFrame.SetISection(smallSection, material, double.Parse(endWebDepth), double.Parse(consFlgDepth), double.Parse(flgThick), double.Parse(webThick), double.Parse(consFlgDepth), double.Parse(flgThick));

                        // largeSection
                        ret = mysapmodel.PropFrame.SetISection(largeSection, material, double.Parse(startWebDepth), double.Parse(consFlgDepth), double.Parse(flgThick), double.Parse(webThick), double.Parse(consFlgDepth), double.Parse(flgThick));


                        StartSec[0] = smallSection;
                        EndSec[0] = largeSection;

                        MyType[0] = 1;

                        MyLength[0] = 1;
                        EI22[0] = EI33[0] = 1;

                        ret = mysapmodel.PropFrame.SetNonPrismatic(revitFrame.RvtName, 1, ref StartSec, ref EndSec, ref MyLength, ref MyType, ref EI33, ref EI22);

                        revitFrame.Sec = revitFrame.RvtName;
                        return revitFrame.Sec;
                    #endregion

                    #region Pipe_T
                    case SectType.BM_BUS_PIPE_Tapered:

                        #region Dimensions

                        StartSec = new string[1];
                        EndSec = new string[1];
                        MyLength = new double[1];
                        MyType = new int[1];
                        EI33 = new int[1];
                        EI22 = new int[1];



                        int lastXIndex = revitFrame.RvtName.LastIndexOf("X");
                        string Diameters = revitFrame.RvtName.Substring(0, lastXIndex);

                        int lastseperetorIndex = Diameters.LastIndexOf("/");

                        string startDiamaterValue = new string((Diameters.SkipWhile(c => !char.IsDigit(c))
                             .TakeWhile(c => char.IsDigit(c)).ToArray()
                            ));

                        string endDiameterValue = Diameters.Substring(lastseperetorIndex + 1);
                        string PipeThicknessValue = revitFrame.RvtName.Substring(lastXIndex + 1);
                        #endregion

                        string startSection = "PIPE " + startDiamaterValue + "X" + PipeThicknessValue;
                        string endSection = "PIPE " + endDiameterValue + "X" + PipeThicknessValue;

                        // smallSection
                        ret = mysapmodel.PropFrame.SetPipe(startSection, material, double.Parse(startDiamaterValue), double.Parse(PipeThicknessValue));

                        // largeSection
                        ret = mysapmodel.PropFrame.SetPipe(endSection, material, double.Parse(endDiameterValue), double.Parse(PipeThicknessValue));


                        StartSec[0] = startSection;
                        EndSec[0] = endSection;

                        MyType[0] = 1;

                        MyLength[0] = 1;
                        EI22[0] = EI33[0] = 1;

                        ret = mysapmodel.PropFrame.SetNonPrismatic(revitFrame.RvtName, 1, ref StartSec, ref EndSec, ref MyLength, ref MyType, ref EI33, ref EI22);

                        revitFrame.Sec = revitFrame.RvtName;
                        return revitFrame.Sec;
                        //TaperedSection



                        break;
                    #endregion

                    #region TUBE_T
                    case SectType.BM_BUS_TUBE_Tapered:

                        #region GetDimentions 
                        string largeFlgWidth = null;
                        string smallFlgWidth = null;
                        string constFlgWidth = null;

                        StartSec = new string[1];
                        EndSec = new string[1];
                        MyLength = new double[1];
                        MyType = new int[1];
                        EI33 = new int[1];
                        EI22 = new int[1];

                        bool isFlgTapered = false;

                        startSection = null;
                        endSection = null;
                        endWebDepth = null;
                        flgThick = null;
                        webThick = null;
                        startWebDepth = null;



                        int no_of_x = Frame.CountCharsUsingIndex(revitFrame.RvtName, 'X');
                        int no_of_slash = Frame.CountCharsUsingIndex(revitFrame.RvtName, '/');


                        //  800/500X20
                        if (no_of_x == 1 && no_of_slash == 1)
                        {

                            startWebDepth = largeFlgWidth = new string(revitFrame.RvtName.SkipWhile(c => !char.IsDigit(c)).TakeWhile(c => char.IsDigit(c)).ToArray());
                            endWebDepth = smallFlgWidth = new string(revitFrame.RvtName.Substring(revitFrame.RvtName.IndexOf("/")).SkipWhile(c => !char.IsDigit(c)).TakeWhile(c => char.IsDigit(c)).ToArray());



                            webThick = flgThick = new string(revitFrame.RvtName.Substring(revitFrame.RvtName.IndexOf('X')).SkipWhile(c => !char.IsDigit(c)).TakeWhile(c => char.IsDigit(c)).ToArray());

                            startSection = "TUBE " + startWebDepth + "X" + largeFlgWidth + "X" + webThick;
                            endSection = "TUBE " + endWebDepth + "X" + smallFlgWidth + "X" + flgThick;
                        }

                        //  800/500X300X20
                        else if (no_of_x == 2 && no_of_slash == 1)
                        {

                            startWebDepth = new string(revitFrame.RvtName.SkipWhile(c => !char.IsDigit(c)).TakeWhile(c => char.IsDigit(c)).ToArray());
                            endWebDepth = new string(revitFrame.RvtName.Substring(revitFrame.RvtName.IndexOf("/")).SkipWhile(c => !char.IsDigit(c)).TakeWhile(c => char.IsDigit(c)).ToArray());


                            largeFlgWidth = smallFlgWidth = new string(revitFrame.RvtName.Substring(revitFrame.RvtName.IndexOf('X')).SkipWhile(c => !char.IsDigit(c)).TakeWhile(c => char.IsDigit(c)).ToArray());
                            webThick = flgThick = new string(revitFrame.RvtName.Substring(revitFrame.RvtName.LastIndexOf('X')).SkipWhile(c => !char.IsDigit(c)).TakeWhile(c => char.IsDigit(c)).ToArray());

                            startSection = "TUBE " + startWebDepth + "X" + largeFlgWidth + "X" + webThick;
                            endSection = "TUBE " + endWebDepth + "X" + smallFlgWidth + "X" + flgThick;


                        }

                        //  800/500X400/300X20
                        else if (no_of_x == 2 && no_of_slash == 2)
                        {

                            isFlgTapered = true;
                            int smallFlgWidthIndex = Frame.IndexOfNth(revitFrame.RvtName, "/", 1);



                            startWebDepth = new string(revitFrame.RvtName.SkipWhile(c => !char.IsDigit(c)).TakeWhile(c => char.IsDigit(c)).ToArray());
                            endWebDepth = new string(revitFrame.RvtName.Substring(revitFrame.RvtName.IndexOf("/")).SkipWhile(c => !char.IsDigit(c)).TakeWhile(c => char.IsDigit(c)).ToArray());



                            largeFlgWidth = new string(revitFrame.RvtName.Substring(revitFrame.RvtName.IndexOf('X')).SkipWhile(c => !char.IsDigit(c)).TakeWhile(c => char.IsDigit(c)).ToArray());
                            smallFlgWidth = new string(revitFrame.RvtName.Substring(smallFlgWidthIndex).SkipWhile(c => !char.IsDigit(c)).TakeWhile(c => char.IsDigit(c)).ToArray());


                            webThick = flgThick = new string(revitFrame.RvtName.Substring(revitFrame.RvtName.LastIndexOf("X")).SkipWhile(c => !char.IsDigit(c)).TakeWhile(c => char.IsDigit(c)).ToArray());


                            startSection = "TUBE " + startWebDepth + "X" + largeFlgWidth + "X" + webThick;
                            endSection = "TUBE " + endWebDepth + "X" + smallFlgWidth + "X" + flgThick;

                        }


                        // createSmallSection
                        ret = mysapmodel.PropFrame.SetTube(startSection, material, double.Parse(endWebDepth), double.Parse(smallFlgWidth), double.Parse(flgThick), double.Parse(webThick));

                        // createLaregeSection
                        ret = mysapmodel.PropFrame.SetTube(endSection, material, double.Parse(startWebDepth), double.Parse(largeFlgWidth), double.Parse(flgThick), double.Parse(webThick));

                        // createTaperedSection
                        StartSec[0] = startSection;
                        EndSec[0] = endSection;

                        MyType[0] = 1;

                        MyLength[0] = 1;
                        EI22[0] = EI33[0] = 1;

                        ret = mysapmodel.PropFrame.SetNonPrismatic(revitFrame.RvtName, 1, ref StartSec, ref EndSec, ref MyLength, ref MyType, ref EI33, ref EI22);

                        revitFrame.Sec = revitFrame.RvtName;
                        return revitFrame.Sec;
                        #endregion

                        break;
                    #endregion

                    #region equal single ANgle

                    case SectType.BM_L_EQUAL:
                        noOfX = Frame.CountCharsUsingIndex(revitFrame.RvtName, 'X');
                        dimentionswithX = new string(revitFrame.RvtName.SkipWhile(c => !char.IsDigit(c)).ToArray());
                        indexOfSecondX = Frame.IndexOfNth(dimentionswithX, "X", 0);
                        lastXindex = revitFrame.RvtName.LastIndexOf("X");

                        string angleThick = revitFrame.RvtName.Substring(lastXindex + 1);

                        firsapimention = null;
                        secondDimention = null;



                        if (noOfX == 2)
                        {
                            dimentionswithX = new string(revitFrame.RvtName.SkipWhile(c => !char.IsDigit(c)).ToArray());
                            firsapimention = new string(dimentionswithX.TakeWhile(c => char.IsDigit(c)).ToArray());
                            secondDimention = new string(dimentionswithX.Substring(indexOfSecondX).SkipWhile(c => !Char.IsDigit(c)).TakeWhile(c => char.IsDigit(c)).ToArray());

                            revitFrame.Sec = "L" + firsapimention + "X" + angleThick;
                            ret = mysapmodel.PropFrame.ImportProp(revitFrame.Sec, material, "Euro.xml", revitFrame.Sec);

                            return revitFrame.Sec;
                            //revitFrame.Sec = "L" + firsapimention + "X" + secondDimention + "X" + angleThick;

                        }



                        else if (noOfX == 1)
                        {

                            dimentionswithX = new string(revitFrame.RvtName.SkipWhile(c => !char.IsDigit(c)).ToArray());
                            firsapimention = new string(dimentionswithX.TakeWhile(c => char.IsDigit(c)).ToArray());
                            revitFrame.Sec = "L" + firsapimention + "X" + angleThick;
                            ret = mysapmodel.PropFrame.ImportProp(revitFrame.Sec, material, "Euro.xml", revitFrame.Sec);

                            return revitFrame.Sec;

                        }

                        return revitFrame.Sec;
                        break;

                    #endregion

                    #region UNEqual single ANgle

                    case SectType.BM_L_UNEQUAL:
                        noOfX = Frame.CountCharsUsingIndex(revitFrame.RvtName, 'X');
                        dimentionswithX = new string(revitFrame.RvtName.SkipWhile(c => !char.IsDigit(c)).ToArray());
                        indexOfSecondX = Frame.IndexOfNth(dimentionswithX, "X", 0);
                        lastXindex = revitFrame.RvtName.LastIndexOf("X");

                        angleThick = revitFrame.RvtName.Substring(lastXindex + 1);

                        firsapimention = null;
                        secondDimention = null;



                        dimentionswithX = new string(revitFrame.RvtName.SkipWhile(c => !char.IsDigit(c)).ToArray());
                        firsapimention = new string(dimentionswithX.TakeWhile(c => char.IsDigit(c)).ToArray());
                        secondDimention = new string(dimentionswithX.Substring(indexOfSecondX).SkipWhile(c => !Char.IsDigit(c)).TakeWhile(c => char.IsDigit(c)).ToArray());

                        revitFrame.Sec = "L" + firsapimention + "X" + secondDimention + "X" + angleThick;

                        ret = mysapmodel.PropFrame.ImportProp(revitFrame.Sec, material, "Euro.xml", revitFrame.Sec);

                        return revitFrame.Sec;


                    #endregion

                    #region  dbl ANgle

                    case SectType.BM_2L_EQUALANGLE:
                        noOfX = Frame.CountCharsUsingIndex(revitFrame.RvtName, 'X');
                        revitFrame.RvtName = revitFrame.RvtName.Substring(2);
                        dimentionswithX = new string(revitFrame.RvtName.SkipWhile(c => !char.IsDigit(c)).ToArray());
                        indexOfSecondX = Frame.IndexOfNth(dimentionswithX, "X", 0);
                        lastXindex = revitFrame.RvtName.LastIndexOf("X");

                        angleThick = revitFrame.RvtName.Substring(lastXindex + 1);

                        dimentionswithX = new string(revitFrame.RvtName.SkipWhile(c => !char.IsDigit(c)).ToArray());
                        firsapimention = new string(dimentionswithX.TakeWhile(c => char.IsDigit(c)).ToArray());
                        secondDimention = new string(dimentionswithX.Substring(indexOfSecondX).SkipWhile(c => !Char.IsDigit(c)).TakeWhile(c => char.IsDigit(c)).ToArray());

                        revitFrame.Sec = "2L" + firsapimention + "X" + angleThick + "/" + "10" + "/";

                        ret = mysapmodel.PropFrame.ImportProp(revitFrame.Sec, material, "Euro.xml", revitFrame.Sec);

                        return revitFrame.Sec;


                    #endregion

                    default:
                        revitFrame.Sec = "dummySection";
                        return revitFrame.Sec;


                }
            }
            catch (Exception)
            {

                return "dummySection";
            }



        }
        public static string createSapSectionForColumn(Frame revitFrame, cSapModel mysapmodel)
        {
            #region defintions

            int ret;
            string id = revitFrame.RvtID;
            char cc;
            int lastSpacecharacterIndex;
            string dimentionswithX;
            string sec;
            string[] dimentions;
            string material = revitFrame.Material;
            List<string> profiles = Sections.GetAllNamesAssignedToFramesInSap(mysapmodel);

            if (material == null || material == "" || material == "STEEL")
            {
                int numberOfNames = 0;
                string[] materialList = new string[0];
                ret = mysapmodel.PropMaterial.GetNameList(ref numberOfNames, ref materialList);
                eMatType matType = eMatType.Steel;
                int a = 0;
                foreach (var mat in materialList)
                {

                    mysapmodel.PropMaterial.GetTypeOAPI(mat, ref matType, ref a);

                    if (matType == eMatType.Steel)
                    {
                        material = mat;
                        break;
                    }
                }

            }




            #endregion

            #region createDummySectionInSap


            int dummySectionId = mysapmodel.PropFrame.SetCircle("dummySection", material, 10);


            #endregion  

            sec = profiles.AsParallel().FirstOrDefault(x => x == revitFrame.RvtName);

            if (sec != null)
            {
                revitFrame.Sec = sec;
                return sec;
            }

            try
            {
                switch (revitFrame.FamilyType)
                {

                    #region I_B
                    case SectType.COL_BUS_I:
                        #region getDimentions 
                        int seperatorIndex = revitFrame.RvtName.LastIndexOf("+WEB");

                        string webString = revitFrame.RvtName.Substring(seperatorIndex + 4);
                        string flgString = revitFrame.RvtName.Substring(0, seperatorIndex - 1);


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
                        ret = mysapmodel.PropFrame.SetISection(revitFrame.RvtName, material, double.Parse(consWebDepth), double.Parse(consFlgDepth), double.Parse(flgThick), double.Parse(webThick), double.Parse(consFlgDepth), double.Parse(flgThick));



                        revitFrame.Sec = revitFrame.RvtName;
                        return revitFrame.Sec;
                        #endregion

                        #region I_H
                        string numeric = null;

                    case SectType.COL_HEA:
                        numeric = new String(revitFrame.RvtName.Where(Char.IsDigit).ToArray());
                        revitFrame.Sec = "HE" + numeric + "A";
                        ret = mysapmodel.PropFrame.ImportProp(revitFrame.Sec, material, "ArcelorMittal_Europe.xml", revitFrame.Sec);
                        return revitFrame.Sec;


                    case SectType.COL_HEAA:
                        numeric = new String(revitFrame.RvtName.Where(Char.IsDigit).ToArray());
                        revitFrame.Sec = "HE" + numeric + "AA";
                        ret = mysapmodel.PropFrame.ImportProp(revitFrame.Sec, material, "ArcelorMittal_Europe.xml", revitFrame.Sec);
                        return revitFrame.Sec;



                    case SectType.COL_HEB:
                        numeric = new String(revitFrame.RvtName.Where(Char.IsDigit).ToArray());
                        revitFrame.Sec = "HE" + numeric + "B";
                        ret = mysapmodel.PropFrame.ImportProp(revitFrame.Sec, material, "ArcelorMittal_Europe.xml", revitFrame.Sec);
                        return revitFrame.Sec;


                    case SectType.COL_HEM:
                        numeric = new String(revitFrame.RvtName.Where(Char.IsDigit).ToArray());
                        revitFrame.Sec = "HE" + numeric + "M";
                        ret = mysapmodel.PropFrame.ImportProp(revitFrame.Sec, material, "ArcelorMittal_Europe.xml", revitFrame.Sec);
                        return revitFrame.Sec;


                    case SectType.COL_IPE:
                        numeric = new String(revitFrame.RvtName.Where(Char.IsDigit).ToArray());
                        revitFrame.Sec = "IPE" + numeric;
                        ret = mysapmodel.PropFrame.ImportProp(revitFrame.Sec, material, "ArcelorMittal_Europe.xml", revitFrame.Sec);
                        return revitFrame.Sec;


                    case SectType.COL_IPEA:
                        numeric = new String(revitFrame.RvtName.Where(Char.IsDigit).ToArray());
                        revitFrame.Sec = "IPEA" + numeric;
                        ret = mysapmodel.PropFrame.ImportProp(revitFrame.Sec, material, "ArcelorMittal_Europe.xml", revitFrame.Sec);
                        return revitFrame.Sec;


                    case SectType.COL_IPEO:
                        numeric = new String(revitFrame.RvtName.Where(Char.IsDigit).ToArray());
                        revitFrame.Sec = "IPEO" + numeric;
                        ret = mysapmodel.PropFrame.ImportProp(revitFrame.Sec, material, "ArcelorMittal_Europe.xml", revitFrame.Sec);
                        return revitFrame.Sec;

                    #endregion

                    #region pipe
                    //case SectType.COL_PIPE:
                    //    cc = revitFrame.RvtName.SkipWhile(c => !char.IsDigit(c)).ToList()[0];
                    //    lastSpacecharacterIndex = revitFrame.RvtName.IndexOf(cc);
                    //    dimentionswithX = revitFrame.RvtName.Substring(lastSpacecharacterIndex);
                    //    sec = profiles.AsParallel().FirstOrDefault(x => x.Contains(dimentionswithX));

                    //    if (sec != null)
                    //    {
                    //        revitFrame.Sec = sec;
                    //        return sec;
                    //    }

                    //    else
                    //    {
                    //        dimentions = dimentionswithX.Split(("X".ToCharArray()));
                    //        ret = mysapmodel.PropFrame.SetPipe(revitFrame.RvtName, material, double.Parse(dimentions[0]), double.Parse(dimentions[1]));
                    //        revitFrame.Sec = revitFrame.RvtName;
                    //        return revitFrame.Sec;
                    //    }

                    case SectType.COL_PIPE:
                        cc = revitFrame.RvtName.SkipWhile(c => !char.IsDigit(c)).ToList()[0];
                        lastSpacecharacterIndex = revitFrame.RvtName.IndexOf(cc);
                        dimentionswithX = revitFrame.RvtName.Substring(lastSpacecharacterIndex);
                        sec = profiles.AsParallel().FirstOrDefault(x => x.Contains(dimentionswithX));


                        dimentions = dimentionswithX.Split(("X".ToCharArray()));
                        ret = mysapmodel.PropFrame.SetPipe(revitFrame.RvtName, material, double.Parse(dimentions[0]), double.Parse(dimentions[1]));
                        revitFrame.Sec = revitFrame.RvtName;
                        return revitFrame.Sec;


                    #endregion

                    #region BOX(TUBE)

                    //case SectType.COL_TUBE:
                    //    cc = revitFrame.RvtName.SkipWhile(c => !char.IsDigit(c)).ToList()[0];
                    //    lastSpacecharacterIndex = revitFrame.RvtName.IndexOf(cc);
                    //    dimentionswithX = revitFrame.RvtName.Substring(lastSpacecharacterIndex);
                    //    int noOfX2 = revitFrame.RvtName.Count(x => x == 'X');
                    //    sec = profiles.AsParallel().FirstOrDefault(x => x.Contains(dimentionswithX) && x.Count(a => a == 'X') == noOfX2);

                    //    if (sec != null)
                    //    {
                    //        revitFrame.Sec = revitFrame.RvtName;
                    //        return sec;
                    //    }

                    //    else
                    //    {
                    //        dimentions = dimentionswithX.Split(("X".ToCharArray()));
                    //        if (dimentions.Length == 4)
                    //        {
                    //            ret = mysapmodel.PropFrame.SetTube(revitFrame.RvtName, material, double.Parse(dimentions[0]), double.Parse(dimentions[1]), double.Parse(dimentions[3]), double.Parse(dimentions[2]));
                    //        }


                    //        else
                    //        {
                    //            ret = mysapmodel.PropFrame.SetTube(revitFrame.RvtName, material, double.Parse(dimentions[0]), double.Parse(dimentions[1]), double.Parse(dimentions[2]), double.Parse(dimentions[2]));

                    //        }

                    //        revitFrame.Sec = revitFrame.RvtName;
                    //        return revitFrame.RvtName;
                    //    }


                    case SectType.COL_TUBE:
                        cc = revitFrame.RvtName.SkipWhile(c => !char.IsDigit(c)).ToList()[0];
                        lastSpacecharacterIndex = revitFrame.RvtName.IndexOf(cc);
                        dimentionswithX = revitFrame.RvtName.Substring(lastSpacecharacterIndex);
                        int noOfX2 = revitFrame.RvtName.Count(x => x == 'X');

                        dimentions = dimentionswithX.Split(("X".ToCharArray()));
                        if (dimentions.Length == 4)
                        {
                            ret = mysapmodel.PropFrame.SetTube(revitFrame.RvtName, material, double.Parse(dimentions[0]), double.Parse(dimentions[1]), double.Parse(dimentions[3]), double.Parse(dimentions[2]));
                        }


                        else
                        {
                            ret = mysapmodel.PropFrame.SetTube(revitFrame.RvtName, material, double.Parse(dimentions[0]), double.Parse(dimentions[1]), double.Parse(dimentions[2]), double.Parse(dimentions[2]));

                        }

                        revitFrame.Sec = revitFrame.RvtName;
                        return revitFrame.RvtName;

                    #endregion

                    #region UPN
                    case SectType.COL_UPN:
                        numeric = new String(revitFrame.RvtName.Where(Char.IsDigit).ToArray());
                        revitFrame.Sec = "UPN" + numeric;
                        ret = mysapmodel.PropFrame.ImportProp(revitFrame.Sec, material, "Euro.xml", revitFrame.Sec);
                        return revitFrame.Sec;

                    #endregion

                    #region W


                    case SectType.BM_W:
                        int noOfX = Frame.CountCharsUsingIndex(revitFrame.RvtName, 'X');
                        dimentionswithX = new string(revitFrame.RvtName.SkipWhile(c => !char.IsDigit(c)).ToArray());
                        int indexOfSecondX = Frame.IndexOfNth(dimentionswithX, "X", 0);
                        int lastXindex = revitFrame.RvtName.LastIndexOf("X");

                        string firsapimention = null;
                        string secondDimention = null;

                        string lastDimension = revitFrame.RvtName.Substring(lastXindex + 1);

                        if (noOfX == 2)
                        {
                            firsapimention = new string(dimentionswithX.TakeWhile(c => char.IsDigit(c)).ToArray());
                            secondDimention = new string(dimentionswithX.Substring(indexOfSecondX).SkipWhile(c => !Char.IsDigit(c)).TakeWhile(c => char.IsDigit(c)).ToArray());

                            if (firsapimention == secondDimention)
                            {
                                revitFrame.Sec = "W" + firsapimention + "X" + lastDimension;
                            }
                            else
                            {
                                revitFrame.Sec = "W" + dimentionswithX;
                            }
                        }

                        else
                        {
                            revitFrame.Sec = "W" + dimentionswithX;
                        }


                        ret = mysapmodel.PropFrame.ImportProp(revitFrame.Sec, material, "AISC14.xml", revitFrame.Sec);
                        return revitFrame.Sec;


                    //    cc = revitFrame.RvtName.SkipWhile(c => !char.IsDigit(c)).ToList()[0];
                    //    lastSpacecharacterIndex = revitFrame.RvtName.IndexOf(cc);
                    //    dimentionswithX = revitFrame.RvtName.Substring(lastSpacecharacterIndex);
                    //    revitFrame.Sec = "W" + dimentionswithX;
                    //    ret = mysapmodel.PropFrame.ImportProp(revitFrame.Sec, material, "AISC15.xml", revitFrame.Sec);
                    //    return revitFrame.Sec;

                    #endregion

                    #region UB
                    case SectType.COL_UB:
                        cc = revitFrame.RvtName.SkipWhile(c => !char.IsDigit(c)).ToList()[0];
                        lastSpacecharacterIndex = revitFrame.RvtName.IndexOf(cc);
                        dimentionswithX = revitFrame.RvtName.Substring(lastSpacecharacterIndex);
                        revitFrame.Sec = "UKB" + dimentionswithX;
                        ret = mysapmodel.PropFrame.ImportProp(revitFrame.Sec, material, "BSShapes2006.xml", revitFrame.Sec);
                        return revitFrame.Sec;
                    #endregion

                    #region UC
                    case SectType.COL_UC:
                        cc = revitFrame.RvtName.SkipWhile(c => !char.IsDigit(c)).ToList()[0];
                        lastSpacecharacterIndex = revitFrame.RvtName.IndexOf(cc);
                        dimentionswithX = revitFrame.RvtName.Substring(lastSpacecharacterIndex);
                        revitFrame.Sec = "UKC" + dimentionswithX;
                        ret = mysapmodel.PropFrame.ImportProp(revitFrame.Sec, material, "BSShapes2006.xml", revitFrame.Sec);
                        return revitFrame.Sec;
                    #endregion

                    #region I_T
                    case SectType.BM_BUS_I_Tapered:

                        string[] StartSec = new string[1];
                        string[] EndSec = new string[1];
                        double[] MyLength = new double[1];
                        int[] MyType = new int[1];
                        int[] EI33 = new int[1];
                        int[] EI22 = new int[1];

                        #region getDimentions 

                        seperatorIndex = revitFrame.RvtName.LastIndexOf("+WEB");

                        webString = revitFrame.RvtName.Substring(seperatorIndex + 4);
                        flgString = revitFrame.RvtName.Substring(0, seperatorIndex - 1);


                        webstring2 = webString.Split(("X".ToCharArray()));
                        webThick = webstring2[1];


                        flgString2 = flgString.Split(("X".ToCharArray()));
                        flgThick = flgString2[1];

                        string startWebDepth = null;
                        string endWebDepth = null;
                        string largeWebDepthWithoutTF = null;
                        string smallWebDepthWithoutTF = null;



                        consFlgDepth = null;


                        // get largeDepth ,smalldepth of the web and constWidth of the flange
                        string[] webstring3 = webstring2[0].Split(("/".ToCharArray()));


                        largeWebDepthWithoutTF = (Convert.ToDouble(new string((webstring3[0].SkipWhile(c => !char.IsDigit(c))
                                               .TakeWhile(c => char.IsDigit(c)).ToArray()
                                              ))).ToString());




                        startWebDepth = (Convert.ToDouble(new string((webstring3[0].SkipWhile(c => !char.IsDigit(c))
                                           .TakeWhile(c => char.IsDigit(c)).ToArray()
                                          ))) + 2 * Convert.ToDouble(flgThick)).ToString();


                        smallWebDepthWithoutTF = (Convert.ToDouble(new string((webstring3[1].SkipWhile(c => !char.IsDigit(c))
                                                .TakeWhile(c => char.IsDigit(c)).ToArray()
                                                ))).ToString());


                        endWebDepth = (Convert.ToDouble(new string((webstring3[1].SkipWhile(c => !char.IsDigit(c))
                                           .TakeWhile(c => char.IsDigit(c)).ToArray()
                                          ))) + 2 * Convert.ToDouble(flgThick)).ToString();


                        consFlgDepth = new string((flgString2[0].Substring(2).SkipWhile(c => !char.IsDigit(c))
                                           .TakeWhile(c => char.IsDigit(c)).ToArray()
                                           ));

                        #endregion

                        string smallSection = "2FL. PL. " + consFlgDepth + "X" + flgThick + "  +" + "WEB PL. " + largeWebDepthWithoutTF + "X" + webThick;
                        string largeSection = "2FL. PL. " + consFlgDepth + "X" + flgThick + "  +" + "WEB PL. " + smallWebDepthWithoutTF + "X" + webThick;
                        // smallSection
                        ret = mysapmodel.PropFrame.SetISection(smallSection, material, double.Parse(endWebDepth), double.Parse(consFlgDepth), double.Parse(flgThick), double.Parse(webThick), double.Parse(consFlgDepth), double.Parse(flgThick));

                        // largeSection
                        ret = mysapmodel.PropFrame.SetISection(largeSection, material, double.Parse(startWebDepth), double.Parse(consFlgDepth), double.Parse(flgThick), double.Parse(webThick), double.Parse(consFlgDepth), double.Parse(flgThick));


                        StartSec[0] = smallSection;
                        EndSec[0] = largeSection;

                        MyType[0] = 1;

                        MyLength[0] = 1;
                        EI22[0] = EI33[0] = 1;

                        ret = mysapmodel.PropFrame.SetNonPrismatic(revitFrame.RvtName, 1, ref StartSec, ref EndSec, ref MyLength, ref MyType, ref EI33, ref EI22);

                        revitFrame.Sec = revitFrame.RvtName;
                        return revitFrame.Sec;
                    #endregion

                    #region Pipe_T
                    case SectType.BM_BUS_PIPE_Tapered:

                        #region Dimensions

                        StartSec = new string[1];
                        EndSec = new string[1];
                        MyLength = new double[1];
                        MyType = new int[1];
                        EI33 = new int[1];
                        EI22 = new int[1];



                        int lastXIndex = revitFrame.RvtName.LastIndexOf("X");
                        string Diameters = revitFrame.RvtName.Substring(0, lastXIndex);

                        int lastseperetorIndex = Diameters.LastIndexOf("/");

                        string startDiamaterValue = new string((Diameters.SkipWhile(c => !char.IsDigit(c))
                             .TakeWhile(c => char.IsDigit(c)).ToArray()
                            ));

                        string endDiameterValue = Diameters.Substring(lastseperetorIndex + 1);
                        string PipeThicknessValue = revitFrame.RvtName.Substring(lastXIndex + 1);
                        #endregion

                        string startSection = "PIPE " + startDiamaterValue + "X" + PipeThicknessValue;
                        string endSection = "PIPE " + endDiameterValue + "X" + PipeThicknessValue;

                        // smallSection
                        ret = mysapmodel.PropFrame.SetPipe(startSection, material, double.Parse(startDiamaterValue), double.Parse(PipeThicknessValue));

                        // largeSection
                        ret = mysapmodel.PropFrame.SetPipe(endSection, material, double.Parse(endDiameterValue), double.Parse(PipeThicknessValue));


                        StartSec[0] = startSection;
                        EndSec[0] = endSection;

                        MyType[0] = 1;

                        MyLength[0] = 1;
                        EI22[0] = EI33[0] = 1;

                        ret = mysapmodel.PropFrame.SetNonPrismatic(revitFrame.RvtName, 1, ref StartSec, ref EndSec, ref MyLength, ref MyType, ref EI33, ref EI22);

                        revitFrame.Sec = revitFrame.RvtName;
                        return revitFrame.Sec;
                        //TaperedSection



                        break;
                    #endregion

                    #region TUBE_T
                    case SectType.BM_BUS_TUBE_Tapered:

                        #region GetDimentions 
                        string largeFlgWidth = null;
                        string smallFlgWidth = null;
                        string constFlgWidth = null;

                        StartSec = new string[1];
                        EndSec = new string[1];
                        MyLength = new double[1];
                        MyType = new int[1];
                        EI33 = new int[1];
                        EI22 = new int[1];

                        bool isFlgTapered = false;

                        startSection = null;
                        endSection = null;
                        endWebDepth = null;
                        flgThick = null;
                        webThick = null;
                        startWebDepth = null;



                        int no_of_x = Frame.CountCharsUsingIndex(revitFrame.RvtName, 'X');
                        int no_of_slash = Frame.CountCharsUsingIndex(revitFrame.RvtName, '/');


                        //  800/500X20
                        if (no_of_x == 1 && no_of_slash == 1)
                        {

                            startWebDepth = largeFlgWidth = new string(revitFrame.RvtName.SkipWhile(c => !char.IsDigit(c)).TakeWhile(c => char.IsDigit(c)).ToArray());
                            endWebDepth = smallFlgWidth = new string(revitFrame.RvtName.Substring(revitFrame.RvtName.IndexOf("/")).SkipWhile(c => !char.IsDigit(c)).TakeWhile(c => char.IsDigit(c)).ToArray());



                            webThick = flgThick = new string(revitFrame.RvtName.Substring(revitFrame.RvtName.IndexOf('X')).SkipWhile(c => !char.IsDigit(c)).TakeWhile(c => char.IsDigit(c)).ToArray());

                            startSection = "TUBE " + startWebDepth + "X" + largeFlgWidth + "X" + webThick;
                            endSection = "TUBE " + endWebDepth + "X" + smallFlgWidth + "X" + flgThick;
                        }

                        //  800/500X300X20
                        else if (no_of_x == 2 && no_of_slash == 1)
                        {

                            startWebDepth = new string(revitFrame.RvtName.SkipWhile(c => !char.IsDigit(c)).TakeWhile(c => char.IsDigit(c)).ToArray());
                            endWebDepth = new string(revitFrame.RvtName.Substring(revitFrame.RvtName.IndexOf("/")).SkipWhile(c => !char.IsDigit(c)).TakeWhile(c => char.IsDigit(c)).ToArray());


                            largeFlgWidth = smallFlgWidth = new string(revitFrame.RvtName.Substring(revitFrame.RvtName.IndexOf('X')).SkipWhile(c => !char.IsDigit(c)).TakeWhile(c => char.IsDigit(c)).ToArray());
                            webThick = flgThick = new string(revitFrame.RvtName.Substring(revitFrame.RvtName.LastIndexOf('X')).SkipWhile(c => !char.IsDigit(c)).TakeWhile(c => char.IsDigit(c)).ToArray());

                            startSection = "TUBE " + startWebDepth + "X" + largeFlgWidth + "X" + webThick;
                            endSection = "TUBE " + endWebDepth + "X" + smallFlgWidth + "X" + flgThick;


                        }

                        //  800/500X400/300X20
                        else if (no_of_x == 2 && no_of_slash == 2)
                        {

                            isFlgTapered = true;
                            int smallFlgWidthIndex = Frame.IndexOfNth(revitFrame.RvtName, "/", 1);



                            startWebDepth = new string(revitFrame.RvtName.SkipWhile(c => !char.IsDigit(c)).TakeWhile(c => char.IsDigit(c)).ToArray());
                            endWebDepth = new string(revitFrame.RvtName.Substring(revitFrame.RvtName.IndexOf("/")).SkipWhile(c => !char.IsDigit(c)).TakeWhile(c => char.IsDigit(c)).ToArray());



                            largeFlgWidth = new string(revitFrame.RvtName.Substring(revitFrame.RvtName.IndexOf('X')).SkipWhile(c => !char.IsDigit(c)).TakeWhile(c => char.IsDigit(c)).ToArray());
                            smallFlgWidth = new string(revitFrame.RvtName.Substring(smallFlgWidthIndex).SkipWhile(c => !char.IsDigit(c)).TakeWhile(c => char.IsDigit(c)).ToArray());


                            webThick = flgThick = new string(revitFrame.RvtName.Substring(revitFrame.RvtName.LastIndexOf("X")).SkipWhile(c => !char.IsDigit(c)).TakeWhile(c => char.IsDigit(c)).ToArray());


                            startSection = "TUBE " + startWebDepth + "X" + largeFlgWidth + "X" + webThick;
                            endSection = "TUBE " + endWebDepth + "X" + smallFlgWidth + "X" + flgThick;

                        }


                        // createSmallSection
                        ret = mysapmodel.PropFrame.SetTube(startSection, material, double.Parse(endWebDepth), double.Parse(smallFlgWidth), double.Parse(flgThick), double.Parse(webThick));

                        // createLaregeSection
                        ret = mysapmodel.PropFrame.SetTube(endSection, material, double.Parse(startWebDepth), double.Parse(largeFlgWidth), double.Parse(flgThick), double.Parse(webThick));

                        // createTaperedSection
                        StartSec[0] = startSection;
                        EndSec[0] = endSection;

                        MyType[0] = 1;

                        MyLength[0] = 1;
                        EI22[0] = EI33[0] = 1;

                        ret = mysapmodel.PropFrame.SetNonPrismatic(revitFrame.RvtName, 1, ref StartSec, ref EndSec, ref MyLength, ref MyType, ref EI33, ref EI22);

                        revitFrame.Sec = revitFrame.RvtName;
                        return revitFrame.Sec;
                        #endregion

                        break;
                    #endregion

                    #region equal single ANgle

                    case SectType.COL_L_EQUAL:
                        noOfX = Frame.CountCharsUsingIndex(revitFrame.RvtName, 'X');
                        dimentionswithX = new string(revitFrame.RvtName.SkipWhile(c => !char.IsDigit(c)).ToArray());
                        indexOfSecondX = Frame.IndexOfNth(dimentionswithX, "X", 0);
                        lastXindex = revitFrame.RvtName.LastIndexOf("X");

                        string angleThick = revitFrame.RvtName.Substring(lastXindex + 1);

                        firsapimention = null;
                        secondDimention = null;



                        if (noOfX == 2)
                        {
                            dimentionswithX = new string(revitFrame.RvtName.SkipWhile(c => !char.IsDigit(c)).ToArray());
                            firsapimention = new string(dimentionswithX.TakeWhile(c => char.IsDigit(c)).ToArray());
                            secondDimention = new string(dimentionswithX.Substring(indexOfSecondX).SkipWhile(c => !Char.IsDigit(c)).TakeWhile(c => char.IsDigit(c)).ToArray());

                            revitFrame.Sec = "L" + firsapimention + "X" + angleThick;
                            ret = mysapmodel.PropFrame.ImportProp(revitFrame.Sec, material, "Euro.xml", revitFrame.Sec);

                            return revitFrame.Sec;
                            //revitFrame.Sec = "L" + firsapimention + "X" + secondDimention + "X" + angleThick;

                        }



                        else if (noOfX == 1)
                        {

                            dimentionswithX = new string(revitFrame.RvtName.SkipWhile(c => !char.IsDigit(c)).ToArray());
                            firsapimention = new string(dimentionswithX.TakeWhile(c => char.IsDigit(c)).ToArray());
                            revitFrame.Sec = "L" + firsapimention + "X" + angleThick;
                            ret = mysapmodel.PropFrame.ImportProp(revitFrame.Sec, material, "Euro.xml", revitFrame.Sec);

                            return revitFrame.Sec;

                        }

                        return revitFrame.Sec;
                        break;

                    #endregion

                    #region UNEqual single ANgle

                    case SectType.COL_L_UNEQUAL:
                        noOfX = Frame.CountCharsUsingIndex(revitFrame.RvtName, 'X');
                        dimentionswithX = new string(revitFrame.RvtName.SkipWhile(c => !char.IsDigit(c)).ToArray());
                        indexOfSecondX = Frame.IndexOfNth(dimentionswithX, "X", 0);
                        lastXindex = revitFrame.RvtName.LastIndexOf("X");

                        angleThick = revitFrame.RvtName.Substring(lastXindex + 1);

                        firsapimention = null;
                        secondDimention = null;



                        dimentionswithX = new string(revitFrame.RvtName.SkipWhile(c => !char.IsDigit(c)).ToArray());
                        firsapimention = new string(dimentionswithX.TakeWhile(c => char.IsDigit(c)).ToArray());
                        secondDimention = new string(dimentionswithX.Substring(indexOfSecondX).SkipWhile(c => !Char.IsDigit(c)).TakeWhile(c => char.IsDigit(c)).ToArray());

                        revitFrame.Sec = "L" + firsapimention + "X" + secondDimention + "X" + angleThick;

                        ret = mysapmodel.PropFrame.ImportProp(revitFrame.Sec, material, "Euro.xml", revitFrame.Sec);

                        return revitFrame.Sec;


                    #endregion

                    #region  dbl ANgle

                    case SectType.COL_2L_EQUALANGLE:
                        noOfX = Frame.CountCharsUsingIndex(revitFrame.RvtName, 'X');
                        revitFrame.RvtName = revitFrame.RvtName.Substring(2);
                        dimentionswithX = new string(revitFrame.RvtName.SkipWhile(c => !char.IsDigit(c)).ToArray());
                        indexOfSecondX = Frame.IndexOfNth(dimentionswithX, "X", 0);
                        lastXindex = revitFrame.RvtName.LastIndexOf("X");

                        angleThick = revitFrame.RvtName.Substring(lastXindex + 1);

                        dimentionswithX = new string(revitFrame.RvtName.SkipWhile(c => !char.IsDigit(c)).ToArray());
                        firsapimention = new string(dimentionswithX.TakeWhile(c => char.IsDigit(c)).ToArray());
                        secondDimention = new string(dimentionswithX.Substring(indexOfSecondX).SkipWhile(c => !Char.IsDigit(c)).TakeWhile(c => char.IsDigit(c)).ToArray());

                        revitFrame.Sec = "2L" + firsapimention + "X" + angleThick + "/" + "10" + "/";

                        ret = mysapmodel.PropFrame.ImportProp(revitFrame.Sec, material, "Euro.xml", revitFrame.Sec);

                        return revitFrame.Sec;


                    #endregion



                    default:
                        return "dummySection";
                }
            }
            catch (Exception)
            {

                return "dummySection";
            }



        }

        public static void UpdateBeamDicWithNames(cSapModel mysapmodel, Dictionary<string, string> sapSectionsDicForBeams)
        {
            int count = 0;
            int countMaterials = 0;
            string[] materList = null;
            string[] PropsList = null;
            mysapmodel.PropFrame.GetNameList(ref count, ref PropsList);
            List<string> SapPropListFromDic = sapSectionsDicForBeams.Values.ToList();
            List<string> PropsNotCreated = SapPropListFromDic.Except(PropsList).ToList();

            mysapmodel.PropMaterial.GetNameList(ref countMaterials, ref materList);
            foreach (var item in PropsNotCreated)
            {
                string RvtCrosspondName = sapSectionsDicForBeams.First(x => x.Value == item).Key;
                string newName = "dummy section" + "(" + RvtCrosspondName + ")";
                mysapmodel.PropFrame.SetCircle(newName, materList[0], 10);
                sapSectionsDicForBeams[RvtCrosspondName] = newName;

            }
        }


        #endregion

        #region  SapGroupUtils
        public static void createGroupinSap(List<string> frameIds, cSapModel mySapModel, string groupName)

        {
            mySapModel.GroupDef.SetGroup(groupName);
            mySapModel.GroupDef.Clear(groupName);


            foreach (var id in frameIds.AsParallel())
            {
                mySapModel.FrameObj.SetGroupAssign(id, groupName);

            }



        }
        public static List<string> getGroupsList()
        {

            int NumberOFGroups = 0;
            string[] groups = null;
            Get_sap_mode().GroupDef.GetNameList(ref NumberOFGroups, ref groups);
            return groups.ToList();
        }

        public static List<string> getIdsOfElementsInGroup(List<String> groupNames, int targetGroupType)
        {
            //1 = Point object // 2 = Frame object // 3 = Cable object //4 = Tendon object //5 = Area object //6 = Solid object //7 = Link object

            List<string> TargetObjectIds = new List<string>();
            int NumberOfItems = 0;

            string[] allReturnedIds = null;
            int[] returnedObjectTypes = null;

            foreach (var g in groupNames)
            {
                Get_sap_mode().GroupDef.GetAssignments(g, ref NumberOfItems, ref returnedObjectTypes, ref allReturnedIds);

                for (int i = 0; i < returnedObjectTypes.Length; i++)
                {
                    if (returnedObjectTypes[i] == targetGroupType)
                    {
                        TargetObjectIds.Add(allReturnedIds[i]);
                    }
                }
            }



            return TargetObjectIds.Distinct().ToList();
        }


        #endregion

        #region  SapPointsUtils

        public static List<string> getSapPointIDS()
        {

            cSapModel mySapModel = Get_sap_mode();

            #region defintions 
            string[] framesIDS = null;
            int noOfFrames = 0;
            int ret;

            #endregion

            ret = mySapModel.PointObj.GetNameList(ref noOfFrames, ref framesIDS);

            if (noOfFrames == 0)
            {
                return new List<string>();
            }

            else
            {
                return framesIDS.ToList();
            }
        }

        public static List<Point> getSapPoints(cSapModel mySapModel)
        {
            List<Point> pts = new List<Point>();


            #region defintions 
            string[] PointIds = null;
            int noOfFrames = 0;
            double x, y, z;
            x = y = z = 0;

            int ret;

            #endregion

            ret = mySapModel.PointObj.GetNameList(ref noOfFrames, ref PointIds);



            for (int i = 0; i < noOfFrames; i++)
            {
                ret = mySapModel.PointObj.GetCoordCartesian(PointIds[i], ref x, ref y, ref z);
                Point pt = new Point();
                pt.X = x;
                pt.Y = y;
                pt.Z = z;
                pt.Id = PointIds[i];
                pts.Add(pt);
            }

            return pts;

        }

        // not used for rhino
        public static List<string> addNewPoints(List<Point> points)
        {
            cSapModel mySapModel = Get_sap_mode();
            List<string> Ids = new List<string>();
            string id = null;

            foreach (var item in points)
            {
                mySapModel.PointObj.AddCartesian(item.X, item.Y, item.Z, ref id);
                Ids.Add(id);
            }

            return Ids;
        }

        // used for rhino
        public static List<string> addNewPointsNew(ref List<Point> points)
        {
            cSapModel mySapModel = Get_sap_mode();
            List<string> Ids = new List<string>();

 
            foreach (var item in points)
            {
                string id = null;

                mySapModel.PointObj.AddCartesian(item.X, item.Y, item.Z, ref id, "");
                if (item.Id == "3790")
                {
                    int aaa = 22;
                }
                Ids.Add(id);
                item.Id = id;
                bool[] arr = item.Restrains;
                mySapModel.PointObj.SetRestraint(id,ref arr);
            }

            return Ids;
        }

        #endregion

        #region convertToSegments
        public static List<List<Frame>> GroupFrames(List<Frame> frames)
        {
            List<List<Frame>> groups = new List<List<Frame>>();
            List<Frame> remainingFrames = new List<Frame>(frames);

            while (remainingFrames.Count > 0)
            {
                // initialize a new group with the first frame in the list
                List<Frame> group = new List<Frame> { remainingFrames[0] };
                List<Frame> group2 = new List<Frame> { remainingFrames[0] };

                // remove the first frame from the list of remaining frames
                remainingFrames.RemoveAt(0);

                bool framesAdded = true;

                while (framesAdded)
                {
                    framesAdded = false;

                    // loop through the remaining frames to find any that connect to the end of the current group
                    for (int i = 0; i < remainingFrames.Count; i++)
                    {

                        Frame frame = remainingFrames[i];

                        if (frame.ID == "22")
                        {
                            if (frame.StartPoint.Id == group[group.Count - 1].EndPoint.Id || frame.StartPoint.Id == group[group.Count - 1].StartPoint.Id)
                            {

                            }

                            if (isSameVector(frame, group[group.Count - 1]))
                            {

                            }
                        }



                        if ((frame.StartPoint.Id == group[group.Count - 1].EndPoint.Id || frame.StartPoint.Id == group[group.Count - 1].StartPoint.Id) && isSameVector(frame, group[group.Count - 1]))
                        {
                            // add the frame to the group and remove it from the list of remaining frames
                            group.Add(frame);
                            remainingFrames.RemoveAt(i);

                            framesAdded = true;
                            break;
                        }

                        else if ((frame.EndPoint.Id == group[group.Count - 1].EndPoint.Id || frame.EndPoint.Id == group[group.Count - 1].StartPoint.Id) && isSameVector(frame, group[group.Count - 1]))
                        {
                            // reverse the frame and add it to the group, then remove it from the list of remaining frames
                            group.Add(new Frame { ID = frame.ID, StartPoint = frame.EndPoint, EndPoint = frame.StartPoint, MidPoint = frame.MidPoint });
                            remainingFrames.RemoveAt(i);

                            framesAdded = true;
                            break;
                        }


                    }



                }


                framesAdded = true;

                while (framesAdded)
                {
                    framesAdded = false;

                    // loop through the remaining frames to find any that connect to the end of the current group2
                    for (int i = 0; i < remainingFrames.Count; i++)
                    {

                        Frame frame = remainingFrames[i];

                        if ((frame.StartPoint.Id == group2[group2.Count - 1].EndPoint.Id || frame.StartPoint.Id == group2[group2.Count - 1].StartPoint.Id) && isSameVector(frame, group2[group2.Count - 1]))
                        {
                            // add the frame to the group2 and remove it from the list of remaining frames
                            group2.Add(frame);
                            remainingFrames.RemoveAt(i);

                            framesAdded = true;
                            break;
                        }

                        else if ((frame.EndPoint.Id == group2[group2.Count - 1].EndPoint.Id || frame.EndPoint.Id == group2[group2.Count - 1].StartPoint.Id) && isSameVector(frame, group2[group2.Count - 1]))
                        {
                            // reverse the frame and add it to the group2, then remove it from the list of remaining frames
                            group2.Add(new Frame { ID = frame.ID, StartPoint = frame.EndPoint, EndPoint = frame.StartPoint, MidPoint = frame.MidPoint });
                            remainingFrames.RemoveAt(i);

                            framesAdded = true;
                            break;
                        }


                    }



                }


                group = group.Union(group2).Distinct().ToList();
                groups.Add(group);
            }

            return groups;
        }

        public static bool isSameVector2(Frame x, Frame y)
        {
            bool iSameVector = false;

            Vector3 v1 = new Vector3((float)(x.EndPoint.X - x.StartPoint.X), (float)(x.EndPoint.Y - x.StartPoint.Y), (float)(x.EndPoint.Z - x.StartPoint.Z));
            Vector3 v2 = new Vector3((float)(y.EndPoint.X - y.StartPoint.X), (float)(y.EndPoint.Y - y.StartPoint.Y), (float)(y.EndPoint.Z - y.StartPoint.Z));

            v1 = Vector3.Normalize(v1);
            v2 = Vector3.Normalize(v2);

            float dotProduct = Vector3.Dot(v1, v2);

            float magV1 = v1.Length();
            float magV2 = v2.Length();

            float angle = (float)Math.Acos(dotProduct / (magV1 * magV2));
            if (float.IsNaN(angle))
            {
                iSameVector = true;
            }
            else
            {
                // Convert the angle from radians to degrees
                float angleDegrees = angle * (float)(180 / Math.PI);

                // Check if the angle is close to 0 or 180 degrees (within a tolerance of 0.1 degrees)
                if (Math.Abs(angleDegrees) < 0.1f || Math.Abs(angleDegrees - 180) < 0.1f)
                {
                    iSameVector = true;

                }

            }







            //Vector3 crossProduct = Vector3.Cross(v1, v2);

            //float aa = crossProduct.Length();

            //if (Math.Abs( crossProduct.Length() )< 0.001)
            //{
            //    iSameVector = true;
            //}





            //float Xangle1 = (float)Math.Round( Math.Abs( Vector3.Dot(Vector3.UnitX, v1)),2);
            //float Yangle1 = (float)Math.Round(Math.Abs( Vector3.Dot(Vector3.UnitY, v1)),2);
            //float Zangle1 = (float)Math.Round(Math.Abs( Vector3.Dot(Vector3.UnitZ, v1)),2);


            //float Xangle2 = (float)Math.Round(Math.Abs( Vector3.Dot(Vector3.UnitX, v2)),2);
            //float Yangle2 = (float)Math.Round(Math.Abs( Vector3.Dot(Vector3.UnitY, v2)),2);
            //float Zangl2 = (float)Math.Round(Math.Abs( Vector3.Dot(Vector3.UnitZ, v2)),2);

            //if (Math.Abs( Xangle1- Xangle2 ) < 0.05 && Math.Abs( Yangle1 - Yangle2 ) < 0.05&& Math.Abs( Zangle1 -Zangl2) < 0.05)
            //{
            //    iSameVector = true;
            //}

            return iSameVector;
        }

        public static bool isSameVector(Frame frame1, Frame frame2, double tolerance = 1e-10)
        {
            double slopeXY1, slopeXY2, slopeYZ1, slopeYZ2, slopeXZ1, slopeXZ2;

            bool isSameVectorr = false;

            tolerance = 0.01;
            // calculate slopes in XY plane
            if (Math.Abs(frame1.EndPoint.X - frame1.StartPoint.X) < tolerance)
            {
                // line 1 is vertical in XY plane
                slopeXY1 = 0;
            }
            else
            {
                slopeXY1 = (frame1.EndPoint.Y - frame1.StartPoint.Y) / (frame1.EndPoint.X - frame1.StartPoint.X);
            }

            if (Math.Abs(frame2.EndPoint.X - frame2.StartPoint.X) < tolerance)
            {
                // line 2 is vertical in XY plane
                slopeXY2 = 0;
            }
            else
            {
                slopeXY2 = (frame2.EndPoint.Y - frame2.StartPoint.Y) / (frame2.EndPoint.X - frame2.StartPoint.X);
            }

            // calculate slopes in YZ plane
            if (Math.Abs(frame1.EndPoint.Y - frame1.StartPoint.Y) < tolerance)
            {
                // line 1 is vertical in YZ plane
                slopeYZ1 = 0;
            }
            else
            {
                slopeYZ1 = (frame1.EndPoint.Z - frame1.StartPoint.Z) / (frame1.EndPoint.Y - frame1.StartPoint.Y);
            }

            if (Math.Abs(frame2.EndPoint.Y - frame2.StartPoint.Y) < tolerance)
            {
                // line 2 is vertical in YZ plane
                slopeYZ2 = 0;
            }
            else
            {
                slopeYZ2 = (frame2.EndPoint.Z - frame2.StartPoint.Z) / (frame2.EndPoint.Y - frame2.StartPoint.Y);
            }

            // calculate slopes in XZ plane
            if (Math.Abs(frame1.EndPoint.X - frame1.StartPoint.X) < tolerance)
            {
                // line 1 is vertical in XZ plane
                slopeXZ1 = 0;
            }
            else
            {
                slopeXZ1 = (frame1.EndPoint.Z - frame1.StartPoint.Z) / (frame1.EndPoint.X - frame1.StartPoint.X);
            }

            if (Math.Abs(frame2.EndPoint.X - frame2.StartPoint.X) < tolerance)
            {
                // line 2 is vertical in XZ plane
                slopeXZ2 = 0;
            }
            else
            {
                slopeXZ2 = (frame2.EndPoint.Z - frame2.StartPoint.Z) / (frame2.EndPoint.X - frame2.StartPoint.X);
            }

            // compare slopes with a tolerance, allowing for slopes to be equal
            isSameVectorr = (Math.Abs(slopeXY1 - slopeXY2) < tolerance) && (Math.Abs(slopeYZ1 - slopeYZ2) < tolerance) && (Math.Abs(slopeXZ1 - slopeXZ2) < tolerance);

            isSameVectorr = Math.Abs(slopeXY1 - slopeXY2) < tolerance;
            isSameVectorr = Math.Abs(slopeYZ1 - slopeYZ2) < tolerance;
            isSameVectorr = Math.Abs(slopeXZ1 - slopeXZ2) < tolerance;


            return isSameVectorr;
        }

        public static bool isSameVector1(Frame frame1, Frame frame2)
        {
            double slope1, slope2;

            // handle vertical line 1
            if (Math.Abs(frame1.EndPoint.X - frame1.StartPoint.X) < 0.00001)
            {
                // both lines are vertical
                if (Math.Abs(frame2.EndPoint.X - frame2.StartPoint.X) < 0.00001)
                {
                    return true;
                }
                // only line 1 is vertical
                else
                {
                    return false;
                }
            }
            // line 1 is not vertical
            else
            {
                slope1 = (frame1.EndPoint.Y - frame1.StartPoint.Y) / (frame1.EndPoint.X - frame1.StartPoint.X);
            }

            // handle vertical line 2
            if (Math.Abs(frame2.EndPoint.X - frame2.StartPoint.X) < 0.00001)
            {
                // only line 2 is vertical
                return false;
            }
            // both lines are not vertical
            else
            {
                slope2 = (frame2.EndPoint.Y - frame2.StartPoint.Y) / (frame2.EndPoint.X - frame2.StartPoint.X);
            }

            // compare slopes with a tolerance, allowing for slopes to be equal
            if (Math.Abs(slope1 - slope2) < 0.00001)
            {
                // check if both slopes are zero
                if (Math.Abs(slope1) < 0.00001)
                {
                    // check if y-coordinates of start and end points are equal
                    return Math.Abs(frame1.StartPoint.Y - frame1.EndPoint.Y - (frame2.StartPoint.Y - frame2.EndPoint.Y)) < 0.0001;
                }
                else
                {
                    return true;
                }
            }
            else
            {
                return false;
            }
        }

        public static bool isSameVector4(Frame frame1, Frame frame2)
        {
            double slope1, slope2;
            if (Math.Abs(frame1.EndPoint.X - frame1.StartPoint.X) < double.Epsilon) // handle vertical line 1
            {
                if (Math.Abs(frame2.EndPoint.X - frame2.StartPoint.X) < double.Epsilon) // both lines are vertical
                {
                    return true;
                }
                else // only line 1 is vertical
                {
                    return false;
                }
            }

            else // line 1 is not vertical
            {
                slope1 = (frame1.EndPoint.Y - frame1.StartPoint.Y) / (frame1.EndPoint.X - frame1.StartPoint.X);
            }

            if (Math.Abs(frame2.EndPoint.X - frame2.StartPoint.X) < double.Epsilon) // handle vertical line 2
            {
                // only line 2 is vertical, since line 1 is not vertical
                return false;
            }
            else // both lines are not vertical
            {
                slope2 = (frame2.EndPoint.Y - frame2.StartPoint.Y) / (frame2.EndPoint.X - frame2.StartPoint.X);
            }

            // compare slopes with a tolerance, allowing for slopes to be equal or equal with negative sign
            return slope1 == slope2;
        }
        #endregion

        // not used
        public static void setDataInExcel(List<Frame> sapFrames)
        {

            Microsoft.Office.Interop.Excel.Application Excel = null;
            Microsoft.Office.Interop.Excel._Workbook wb = null;
            Microsoft.Office.Interop.Excel._Worksheet ws = null;
            ExcelUtils.CreateExcelForSAP(ref Excel, ref wb, ref ws, "Sap");

            Excel.Range sapID = ws.Range["A7"];
            Excel.Range sapStartPointId = ws.Range["B7"];
            Excel.Range sapEndPointId = ws.Range["C7"];
            Excel.Range sapStartPtX = ws.Range["D7"];
            Excel.Range sapSec = ws.Range["J7"];
            Excel.Range sapLibrary = ws.Range["K7"];
            Excel.Range sapShape = ws.Range["L7"];
            Excel.Range sapISBUs = ws.Range["M7"];
            Excel.Range sapGroup = ws.Range["Q7"];
            Excel.Range sapMaterial = ws.Range["R7"];
            Excel.Range sapLength = ws.Range["S7"];
            Excel.Range sapBeta = ws.Range["T7"];

            Excel.Range["F5"].Value = Convert.ToInt32(sapFrames.Count);

            for (int i = 0; i < sapFrames.Count; i++)
            {
                sapID.Offset[i, 0].Value = sapFrames[i].ID;
                sapStartPointId.Offset[i, 0].Value = sapFrames[i].StartPoint.Id;
                sapEndPointId.Offset[i, 0].Value = sapFrames[i].EndPoint.Id;
                sapStartPtX.Offset[i, 0].Value = Convert.ToString(sapFrames[i].StartPoint.X);
                sapStartPtX.Offset[i, 1].Value = Convert.ToString(sapFrames[i].StartPoint.Y);
                sapStartPtX.Offset[i, 2].Value = Convert.ToString(sapFrames[i].StartPoint.Z);
                sapStartPtX.Offset[i, 3].Value = Convert.ToString(sapFrames[i].EndPoint.X);
                sapStartPtX.Offset[i, 4].Value = Convert.ToString(sapFrames[i].EndPoint.Y);
                sapStartPtX.Offset[i, 5].Value = Convert.ToString(sapFrames[i].EndPoint.Z);
                sapSec.Offset[i, 0].Value = sapFrames[i].Sec;
                sapShape.Offset[i, 0].Value = sapFrames[i].Shape;
                sapLibrary.Offset[i, 0].Value = sapFrames[i].Lib;
                sapISBUs.Offset[i, 0].Value = Convert.ToString(sapFrames[i].IsBUS);
                sapGroup.Offset[i, 0].Value = sapFrames[i].Group;
                sapMaterial.Offset[i, 0].Value = sapFrames[i].Material;

                sapLength.Offset[i, 0].Value = Convert.ToString(sapFrames[i].Length);


                if (sapFrames[i].MemberType == MemberType.SlanttedColumn || sapFrames[i].MemberType == MemberType.Beam || sapFrames[i].MemberType == MemberType.VLBracing)
                {
                    sapBeta.Offset[i, 0].Value = Convert.ToDouble(sapFrames[i].Beta * -1);
                }


                else
                {
                    sapBeta.Offset[i, 0].Value = Convert.ToDouble(sapFrames[i].Beta - 90.00);
                }


            }

            wb.Save();
            wb.Close();
            Excel.DisplayAlerts = true;
            Excel.Quit();
        }


        #endregion



        #region RelatedToDataBase

        public static int getDataBaseFromSap(string typeOrTableName, bool isMultibleTables)
        {
            string tableKey;
            int tableVersion = 1;
            string fullPath;
            string seprator = ",";
            string groupName = "All";
            string[] filedKeyList = null;
            int ret = 1;



            cSapModel mySapModel = SapUtils.Get_sap_mode();

            int NumberTables = 0;
            string[] tableKeys = null;
            string[] availableTablesNames = null;
            int[] importType = null;
            mySapModel.DatabaseTables.GetAvailableTables(ref NumberTables, ref tableKeys, ref availableTablesNames, ref importType);


            if (!isMultibleTables)
            {
                // singleTble
                if (availableTablesNames.Contains(typeOrTableName))
                {
                    string[] FieldKeyList = null;
                    string GroupName = "All";
                    int TableVersion = 0;


                    string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                    string dataBaseFolderName = "SapDataBAse";
                    string dataBaseFolderPath = Path.Combine(desktopPath, dataBaseFolderName);
                    Directory.CreateDirectory(dataBaseFolderPath);

                    string tablesFolderName = "SinglTables";
                    string tableFolderPath = Path.Combine(dataBaseFolderPath, tablesFolderName);
                    Directory.CreateDirectory(tableFolderPath);

                    fullPath = Path.Combine(tableFolderPath, typeOrTableName + ".csv");

                    ret = mySapModel.DatabaseTables.GetTableForDisplayCSVFile(typeOrTableName, ref FieldKeyList, GroupName, ref TableVersion, fullPath, ",");
                }



            }


            else
            {
                List<string> AllTables = new List<string>();
                List<string> Areas = new List<string>();
                List<string> Frames = new List<string>();
                List<string> Joints = new List<string>();
                List<string> LoadDefinitions = new List<string>();
                List<string> OverwritesAndPrefernces = new List<string>();
                List<string> Materials = new List<string>();
                List<string> Miscellaneous = new List<string>();



                #region All Tables

                AllTables.Add("Connectivity - Area");
                AllTables.Add("Area Loads - Uniform");
                AllTables.Add("Area Loads - Uniform To Frame");
                AllTables.Add("Area Loads - Gravity");
                AllTables.Add("Area Loads - Temperature");
                AllTables.Add("Area Local Axes Assignments 1 - Typical");
                AllTables.Add("Area Local Axes Assignments 2 - Advanced");
                AllTables.Add("Area Section Assignments");
                AllTables.Add("Area Section Properties");
                AllTables.Add("Area Section Property Design Parameters");
                AllTables.Add("Area Section Property - Time Dependent");





                AllTables.Add("Connectivity - Frame");
                AllTables.Add("Frame Loads - Distributed");
                AllTables.Add("Frame Loads - Temperature");
                AllTables.Add("Frame Loads - Gravity");
                AllTables.Add("Frame Loads - Point");
                AllTables.Add("Frame Loads - Deformation");
                AllTables.Add("Frame Load Transfer Options");
                AllTables.Add("Frame Local Axes Assignments 1 - Typical");
                AllTables.Add("Frame Local Axes Assignments 2 - Advanced");
                AllTables.Add("Frame Auto Mesh Assignments");
                AllTables.Add("Frame Design Procedures");
                AllTables.Add("Frame Output Station Assignments");
                AllTables.Add("Frame Release Assignments 1 - General");
                AllTables.Add("Frame Section Assignments");
                AllTables.Add("Frame Section Properties 01 - General");
                AllTables.Add("Frame Section Properties 05 - Nonprismatic");
                AllTables.Add("Frame Section Properties 13 - Time Dependent");




                AllTables.Add("Joint Coordinates");
                AllTables.Add("Joint Pattern Definitions");
                AllTables.Add("Joint Restraint Assignments");
                AllTables.Add("Joint Loads - Force");
                AllTables.Add("Joint Loads - Ground Displacement");


                AllTables.Add("Solid Property Definitions");



                AllTables.Add("Auto Combination Option Data 01 - General");
                AllTables.Add("Case - Modal 1 - General");
                AllTables.Add("Case - Response Spectrum 1 - General");
                AllTables.Add("Case - Static 1 - Load Assignments");
                AllTables.Add("Combination Definitions");
                AllTables.Add("Load Case Definitions");
                AllTables.Add("Load Pattern Definitions");



                AllTables.Add("Material Properties 01 - General");
                AllTables.Add("Material Properties 02 - Basic Mechanical Properties");
                AllTables.Add("Material Properties 03a - Steel Data");
                AllTables.Add("Material Properties 03b - Concrete Data");
                AllTables.Add("Material Properties 03e - Rebar Data");
                AllTables.Add("Material Properties 03f - Tendon Data");
                AllTables.Add("Material Properties 03j - Coupled Nonlinear Von Mises Data");
                AllTables.Add("Material Properties 04 - User Stress-Strain Curves");
                AllTables.Add("Material Properties 06 - Damping Parameters");
                AllTables.Add("Material Properties 09 - Acceptance Criteria");



                AllTables.Add("Grid Lines");
                AllTables.Add("Groups 1 - Definitions");
                AllTables.Add("Groups 2 - Assignments");
                AllTables.Add("Mass Source");
                AllTables.Add("Coordinate Systems");
                AllTables.Add("Analysis Options");



                AllTables.Add("Overwrites - Steel Design - AISC 360-16");
                AllTables.Add("Preferences - Aluminum Design - AA 2015");
                AllTables.Add("Preferences - Cold Formed Design - AISI-16");
                AllTables.Add("Preferences - Concrete Shell Design - Eurocode 2-2004");
                AllTables.Add("Preferences - Steel Design - AISC 360-16");
                AllTables.Add("Preferences - Dimensional");

                #endregion

                #region Area Tables

                Areas.Add("Connectivity - Area");
                Areas.Add("Area Loads - Uniform");
                Areas.Add("Area Loads - Uniform To Frame");
                Areas.Add("Area Loads - Gravity");
                Areas.Add("Area Loads - Temperature");
                Areas.Add("Area Local Axes Assignments 1 - Typical");
                Areas.Add("Area Local Axes Assignments 2 - Advanced");
                Areas.Add("Area Section Assignments");
                Areas.Add("Area Section Properties");
                Areas.Add("Area Section Property Design Parameters");
                Areas.Add("Area Section Property - Time Dependent");

                #endregion

                #region frame Tables


                Frames.Add("Connectivity - Frame");
                Frames.Add("Frame Loads - Distributed");
                Frames.Add("Frame Loads - Temperature");
                Frames.Add("Frame Loads - Gravity");
                Frames.Add("Frame Loads - Point");
                Frames.Add("Frame Loads - Deformation");
                Frames.Add("Frame Load Transfer Options");
                Frames.Add("Frame Local Axes Assignments 1 - Typical");
                Frames.Add("Frame Local Axes Assignments 2 - Advanced");
                Frames.Add("Frame Auto Mesh Assignments");

                Frames.Add("Frame Section Assignments");
                Frames.Add("Frame Output Station Assignments");
                Frames.Add("Frame Release Assignments 1 - General");
                Frames.Add("Frame Section Assignments");


                Frames.Add("Preferences - Steel Design - AISC 360-16");
                Frames.Add("Overwrites - Steel Design - AISC 360-16");



                ////
                Frames.Add("Frame Section Properties 01 - General");
                Frames.Add("Frame Section Properties 05 - Nonprismatic");
                Frames.Add("Frame Section Properties 13 - Time Dependent");

                #endregion

                #region joint Tables

                Joints.Add("Joint Coordinates");
                Joints.Add("Joint Pattern Definitions");
                Joints.Add("Joint Restraint Assignments");
                Joints.Add("Joint Loads - Force");
                Joints.Add("Joint Loads - Ground Displacement");

                #endregion

                #region Load Definitions

                LoadDefinitions.Add("Auto Combination Option Data 01 - General");
                LoadDefinitions.Add("Case - Modal 1 - General");
                LoadDefinitions.Add("Case - Response Spectrum 1 - General");
                LoadDefinitions.Add("Case - Static 1 - Load Assignments");
                LoadDefinitions.Add("Combination Definitions");
                LoadDefinitions.Add("Load Case Definitions");
                LoadDefinitions.Add("Load Pattern Definitions");

                #endregion

                #region Overwrites And Prefernces

                OverwritesAndPrefernces.Add("Overwrites - Steel Design - AISC 360-16");
                OverwritesAndPrefernces.Add("Preferences - Aluminum Design - AA 2015");
                OverwritesAndPrefernces.Add("Preferences - Cold Formed Design - AISI-16");
                OverwritesAndPrefernces.Add("Preferences - Concrete Shell Design - Eurocode 2-2004");
                OverwritesAndPrefernces.Add("Preferences - Steel Design - AISC 360-16");
                OverwritesAndPrefernces.Add("Preferences - Dimensional");

                #endregion

                #region Materials

                Materials.Add("Material Properties 01 - General");
                Materials.Add("Material Properties 02 - Basic Mechanical Properties");
                Materials.Add("Material Properties 03a - Steel Data");
                Materials.Add("Material Properties 03b - Concrete Data");
                Materials.Add("Material Properties 03e - Rebar Data");
                Materials.Add("Material Properties 03f - Tendon Data");
                Materials.Add("Material Properties 03j - Coupled Nonlinear Von Mises Data");
                Materials.Add("Material Properties 04 - User Stress-Strain Curves");
                Materials.Add("Material Properties 06 - Damping Parameters");
                Materials.Add("Material Properties 09 - Acceptance Criteria");

                #endregion

                #region Miscellaneous

                Miscellaneous.Add("Grid Lines");
                Miscellaneous.Add("Groups 1 - Definitions");
                Miscellaneous.Add("Groups 2 - Assignments");
                Miscellaneous.Add("Mass Source");
                Miscellaneous.Add("Coordinate Systems");
                Miscellaneous.Add("Analysis Options");

                #endregion



                if (typeOrTableName == "All")
                {
                    string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                    string dataBaseFolderName = "SapDataBAse";
                    string dataBaseFolderPath = Path.Combine(desktopPath, dataBaseFolderName);
                    Directory.CreateDirectory(dataBaseFolderPath);

                    string tablesFolderName = "All";
                    string tableFolderPath = Path.Combine(dataBaseFolderPath, tablesFolderName);
                    Directory.CreateDirectory(tableFolderPath);




                    foreach (var item in AllTables)
                    {
                        if (availableTablesNames.Contains(item))
                        {
                            fullPath = Path.Combine(tableFolderPath, item + ".csv");


                            tableKey = item;

                            ret = mySapModel.DatabaseTables.GetTableForDisplayCSVFile(tableKey, ref filedKeyList, groupName, ref tableVersion, fullPath, seprator);

                        }


                    }
                }

                if (typeOrTableName == "Frames")
                {
                    string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                    string dataBaseFolderName = "SapDataBAse";
                    string dataBaseFolderPath = Path.Combine(desktopPath, dataBaseFolderName);
                    Directory.CreateDirectory(dataBaseFolderPath);

                    string tablesFolderName = "Frames";
                    string tableFolderPath = Path.Combine(dataBaseFolderPath, tablesFolderName);
                    Directory.CreateDirectory(tableFolderPath);




                    foreach (var item in Frames)
                    {
                        if (availableTablesNames.Contains(item))
                        {

                            fullPath = Path.Combine(tableFolderPath, item + ".csv");


                            tableKey = item;

                            ret = mySapModel.DatabaseTables.GetTableForDisplayCSVFile(tableKey, ref filedKeyList, groupName, ref tableVersion, fullPath, seprator);

                        }


                    }
                }

                if (typeOrTableName == "Areas")
                {
                    string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                    string dataBaseFolderName = "SapDataBAse";
                    string dataBaseFolderPath = Path.Combine(desktopPath, dataBaseFolderName);
                    Directory.CreateDirectory(dataBaseFolderPath);

                    string tablesFolderName = "Areas";
                    string tableFolderPath = Path.Combine(dataBaseFolderPath, tablesFolderName);
                    Directory.CreateDirectory(tableFolderPath);




                    foreach (var item in Areas)
                    {
                        if (availableTablesNames.Contains(item))
                        {
                            fullPath = Path.Combine(tableFolderPath, item + ".csv");


                            tableKey = item;

                            ret = mySapModel.DatabaseTables.GetTableForDisplayCSVFile(tableKey, ref filedKeyList, groupName, ref tableVersion, fullPath, seprator);

                        }


                    }
                }

                if (typeOrTableName == "Joints")
                {
                    string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                    string dataBaseFolderName = "SapDataBAse";
                    string dataBaseFolderPath = Path.Combine(desktopPath, dataBaseFolderName);
                    Directory.CreateDirectory(dataBaseFolderPath);

                    string tablesFolderName = "Joints";
                    string tableFolderPath = Path.Combine(dataBaseFolderPath, tablesFolderName);
                    Directory.CreateDirectory(tableFolderPath);




                    foreach (var item in Joints)
                    {
                        if (availableTablesNames.Contains(item))
                        {
                            fullPath = Path.Combine(tableFolderPath, item + ".csv");


                            tableKey = item;

                            ret = mySapModel.DatabaseTables.GetTableForDisplayCSVFile(tableKey, ref filedKeyList, groupName, ref tableVersion, fullPath, seprator);

                        }



                    }
                }

                if (typeOrTableName == "Loads")
                {
                    string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                    string dataBaseFolderName = "SapDataBAse";
                    string dataBaseFolderPath = Path.Combine(desktopPath, dataBaseFolderName);
                    Directory.CreateDirectory(dataBaseFolderPath);

                    string tablesFolderName = "Joints";
                    string tableFolderPath = Path.Combine(dataBaseFolderPath, tablesFolderName);
                    Directory.CreateDirectory(tableFolderPath);




                    foreach (var item in LoadDefinitions)
                    {
                        if (availableTablesNames.Contains(item))
                        {
                            fullPath = Path.Combine(tableFolderPath, item + ".csv");


                            tableKey = item;

                            ret = mySapModel.DatabaseTables.GetTableForDisplayCSVFile(tableKey, ref filedKeyList, groupName, ref tableVersion, fullPath, seprator);

                        }


                    }
                }

                if (typeOrTableName == "Over")
                {
                    string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                    string dataBaseFolderName = "SapDataBAse";
                    string dataBaseFolderPath = Path.Combine(desktopPath, dataBaseFolderName);
                    Directory.CreateDirectory(dataBaseFolderPath);

                    string tablesFolderName = "OverwritesAndPrefernces";
                    string tableFolderPath = Path.Combine(dataBaseFolderPath, tablesFolderName);
                    Directory.CreateDirectory(tableFolderPath);




                    foreach (var item in OverwritesAndPrefernces)
                    {
                        if (availableTablesNames.Contains(item))
                        {
                            fullPath = Path.Combine(tableFolderPath, item + ".csv");


                            tableKey = item;

                            ret = mySapModel.DatabaseTables.GetTableForDisplayCSVFile(tableKey, ref filedKeyList, groupName, ref tableVersion, fullPath, seprator);


                        }

                    }
                }

                if (typeOrTableName == "Materials")
                {
                    string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                    string dataBaseFolderName = "SapDataBAse";
                    string dataBaseFolderPath = Path.Combine(desktopPath, dataBaseFolderName);
                    Directory.CreateDirectory(dataBaseFolderPath);

                    string tablesFolderName = "Materials";
                    string tableFolderPath = Path.Combine(dataBaseFolderPath, tablesFolderName);
                    Directory.CreateDirectory(tableFolderPath);




                    foreach (var item in Materials)
                    {
                        if (availableTablesNames.Contains(item))
                        {
                            fullPath = Path.Combine(tableFolderPath, item + ".csv");


                            tableKey = item;

                            ret = mySapModel.DatabaseTables.GetTableForDisplayCSVFile(tableKey, ref filedKeyList, groupName, ref tableVersion, fullPath, seprator);

                        }


                    }
                }

                if (typeOrTableName == "Mis")
                {
                    string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                    string dataBaseFolderName = "SapDataBAse";
                    string dataBaseFolderPath = Path.Combine(desktopPath, dataBaseFolderName);
                    Directory.CreateDirectory(dataBaseFolderPath);

                    string tablesFolderName = "Miscellaneous";
                    string tableFolderPath = Path.Combine(dataBaseFolderPath, tablesFolderName);
                    Directory.CreateDirectory(tableFolderPath);




                    foreach (var item in Miscellaneous)
                    {
                        if (availableTablesNames.Contains(item))
                        {
                            fullPath = Path.Combine(tableFolderPath, item + ".csv");


                            tableKey = item;

                            ret = mySapModel.DatabaseTables.GetTableForDisplayCSVFile(tableKey, ref filedKeyList, groupName, ref tableVersion, fullPath, seprator);

                        }


                    }
                }
            }



            return ret;
        }


        public static string[] getDataTableFromSapToEtabs(string TableName, List<string> filedsNeeded, ref List<string> FiledsReturned, ref int numberOfRecords, cSapModel mysapmodel)
        {
            string[] result = null;

            string[] filedsNeededArr = filedsNeeded.ToArray();
            int tableVersion = 0;
            string[] filedsKeysIncluded = null;
            mysapmodel.DatabaseTables.GetTableForDisplayArray(TableName, ref filedsNeededArr, "", ref tableVersion, ref filedsKeysIncluded, ref numberOfRecords, ref result);

            FiledsReturned = filedsKeysIncluded.ToList();
            return result;
        }


        #endregion

        #region rhino
        public static void createGroupinSap(List<string> frameIds, string groupName)

        {
            cSapModel mySapModel = Get_sap_mode();
            mySapModel.GroupDef.SetGroup(groupName);
            mySapModel.GroupDef.Clear(groupName);


            foreach (var id in frameIds.AsParallel())
            {
                mySapModel.FrameObj.SetGroupAssign(id, groupName);

            }


        }


        public static void createGroupOfPointsinSap(List<string> frameIds, string groupName)

        {
            cSapModel mySapModel = Get_sap_mode();
            mySapModel.GroupDef.SetGroup(groupName);
            mySapModel.GroupDef.Clear(groupName);


            foreach (var id in frameIds.AsParallel())
            {
                mySapModel.PointObj.SetGroupAssign(id, groupName);

            }


        }

        public static void createGroupOfAreasinSap(List<string> frameIds, string groupName)

        {
            cSapModel mySapModel = Get_sap_mode();
            mySapModel.GroupDef.SetGroup(groupName);
            mySapModel.GroupDef.Clear(groupName);


            foreach (var id in frameIds.AsParallel())
            {
                mySapModel.AreaObj.SetGroupAssign(id, groupName);

            }


        }


        public static void clearGroupAssingments()
        {
            cSapModel mySapModel = Get_sap_mode();

            List<string> groupList = getGroupsList();

            foreach (var group in groupList)
            {
                mySapModel.GroupDef.Clear(group);
            }


        }

        public static void getGroupFramesAssignments(ref Dictionary<string, List<string>> dic)
        {
            cSapModel mySapModel = Get_sap_mode();
            int NumberOfGroups = 0;
            string[] groupNames = null;
            mySapModel.GroupDef.GetNameList(ref NumberOfGroups, ref groupNames);

            foreach (var group in groupNames)
            {

                int NumberOfElements = 0;
                int[] objectType = null;
                string[] elementsIds = null;
                mySapModel.GroupDef.GetAssignments(group, ref NumberOfElements, ref objectType, ref elementsIds);

                List<string> frameIds = new List<string>();

                for (int i = 0; i < NumberOfElements; i++)
                {
                    if (objectType[i] == 2)
                    {
                        frameIds.Add(elementsIds[i]);
                    }

                }

                dic.Add(group, frameIds);


            }







        }

        public static void getGroupAreasAssignments(ref Dictionary<string, List<string>> dic)
        {

            cSapModel mySapModel = Get_sap_mode();
            int NumberOfGroups = 0;
            string[] groupNames = null;
            mySapModel.GroupDef.GetNameList(ref NumberOfGroups, ref groupNames);

            foreach (var group in groupNames)
            {

                int NumberOfElements = 0;
                int[] objectType = null;
                string[] elementsIds = null;
                mySapModel.GroupDef.GetAssignments(group, ref NumberOfElements, ref objectType, ref elementsIds);

                List<string> AreaIds = new List<string>();

                for (int i = 0; i < NumberOfElements; i++)
                {
                    if (objectType[i] == 5)
                    {
                        AreaIds.Add(elementsIds[i]);
                    }

                }

                dic.Add(group, AreaIds);


            }







        }


        public static List<string> createSapFrames(List<Frame> frames)
        {
            cSapModel mySapModel = Get_sap_mode();

            List<string> sapIds = new List<string>();

            List<string> Allgroups = getGroupsList();

            var grouppedFrames = frames.GroupBy(x => x.LayerName);

            foreach (var item in grouppedFrames.Select(x=> x.Key))
            {
                mySapModel.GroupDef.SetGroup("RH_"+item);
            }

            foreach (var f in frames)
            {
                #region defintions

                int ret;
                string ID = f.ID;
                string sapSection = null;
                string materialName = null;

                #endregion


                ret = mySapModel.FrameObj.AddByCoord(f.StartPoint.X, f.StartPoint.Y, f.StartPoint.Z, f.EndPoint.X, f.EndPoint.Y, f.EndPoint.Z, ref ID);
                f.ID = ID;
                sapIds.Add(ID);

                #region set Attributes

                //// remove from all groups
                //foreach (var group in Allgroups)
                //{
                //    mySapModel.FrameObj.SetGroupAssign(f.ID, group, true);
                //}


                foreach (var group in f.Groups)
                {

                    if (!Allgroups.Contains(group))
                    {
                        mySapModel.GroupDef.SetGroup(group);
                    }

                    mySapModel.FrameObj.SetGroupAssign(f.ID, group);


                }

                mySapModel.FrameObj.SetSection(f.ID, f.Sec);


                // newlyAddedForLayers
                mySapModel.FrameObj.SetGroupAssign(f.ID, "RH_"+f.LayerName);

                #endregion
            }




            return sapIds;
        }


        public static List<string> createSapArea(List<Area> Areas)
        {
            cSapModel mySapModel = Get_sap_mode();

            List<string> sapIds = new List<string>();


            int noOfPoints = 0;
            double[] Xcoor = null;
            double[] Ycoor = null;
            double[] Zcoor = null;

            int ret = 0;
            String ID = null;

            List<string> Allgroups = getGroupsList();

            foreach (var A in Areas)
            {
                #region defintions
                noOfPoints = A.vertices.Count;
                Xcoor = A.vertices.Select(x => x.X).ToArray();
                Ycoor = A.vertices.Select(x => x.Y).ToArray();
                Zcoor = A.vertices.Select(x => x.Z).ToArray();

                #endregion


                ret = mySapModel.AreaObj.AddByCoord(noOfPoints, ref Xcoor, ref Ycoor, ref Zcoor, ref ID);
                A.ID = ID;
                sapIds.Add(ID);

                #region set Attributes

                //foreach (var group in Allgroups)
                //{
                //    mySapModel.AreaObj.SetGroupAssign(A.ID, group, true);
                //}




                foreach (var group in A.Groups)
                {

                    if (!Allgroups.Contains(group))
                    {
                        mySapModel.GroupDef.SetGroup(group);
                    }

                    mySapModel.AreaObj.SetGroupAssign(A.ID, group);


                }


                mySapModel.AreaObj.SetProperty(A.ID, A.Sec);


                #endregion
            }




            return sapIds;
        }


        public static void updataAssignmentsFromRhinoToSap(List<Frame> frames)
        {
            cSapModel mySapModel = Get_sap_mode();
            List<string> Allgroups = getGroupsList();


            foreach (var f in frames)
            {

                foreach (var group in f.Groups)
                {

                    if (!Allgroups.Contains(group))
                    {
                        mySapModel.GroupDef.SetGroup(group);
                    }

                    mySapModel.FrameObj.SetGroupAssign(f.ID, group);


                }

                // set section
                mySapModel.FrameObj.SetSection(f.ID, f.Sec);

                mySapModel.EditFrame.ChangeConnectivity(f.ID, f.StartPoint.Id, f.EndPoint.Id);


               // mySapModel.FrameObj.SetLocalAxes(f.ID, f.Beta);

            }





        }

        public static void updataAssignmentsFromRhinoToSap(List<Area> Areas)
        {
            cSapModel mySapModel = Get_sap_mode();
            List<string> Allgroups = getGroupsList();


            foreach (var A in Areas)
            {

                #region set Attributes

                foreach (var group in A.Groups)
                {

                    if (!Allgroups.Contains(group))
                    {
                        mySapModel.GroupDef.SetGroup(group);
                    }

                    mySapModel.AreaObj.SetGroupAssign(A.ID, group);


                }

                // set sec

                mySapModel.AreaObj.SetProperty(A.ID, A.Sec);

                // here you need to update the geomtery based on x.y.z  in start and end points not by id 

                string[] ptsIds = A.vertices.Select(x => x.Id).ToArray();
                mySapModel.EditArea.ChangeConnectivity(A.ID, A.vertices.Count, ref ptsIds);


                #endregion
            }


        }


        public static List<Rhino.Geometry.Line> getFramesRhino(List<string> frameIDs, string unit)
        {
            #region defintions 
            string st_point = null;
            string end_point = null;
            int ret;
            double x1, y1, z1, x2, y2, z2;
            x1 = y1 = z1 = x2 = y2 = z2 = 0;

            #endregion
            double factor = 1 / 1000;


            List<Rhino.Geometry.Line> Lines = new List<Rhino.Geometry.Line>();
            cSapModel mySapModel = Get_sap_mode();
            foreach (var id in frameIDs)
            {
                Rhino.Geometry.Line l = new Rhino.Geometry.Line();

                ret = mySapModel.FrameObj.GetPoints(id, ref st_point, ref end_point);
                ret = mySapModel.PointObj.GetCoordCartesian(st_point, ref x1, ref y1, ref z1);
                ret = mySapModel.PointObj.GetCoordCartesian(end_point, ref x2, ref y2, ref z2);

                if (unit.ToUpper() == "M")
                {
                    l = new Rhino.Geometry.Line(x1 * factor, y1 * factor, z1 * factor, x2 * factor, y2 * factor, z2 * factor);
                }

                else
                {
                    l = new Rhino.Geometry.Line(x1, y1, z1, x2, y2, z2);
                }

                Lines.Add(l);

            }


            return Lines;


        }


        public static List<Rhino.Geometry.Line> getFramesRhinoBasic(List<string> frameIDs, string unit)
        {
            #region defintions 
            string st_point = null;
            string end_point = null;
            int ret;
            double x1, y1, z1, x2, y2, z2;
            x1 = y1 = z1 = x2 = y2 = z2 = 0;

            #endregion
            double factor = 1 / 1000;


            List<Rhino.Geometry.Line> Lines = new List<Rhino.Geometry.Line>();
            cSapModel mySapModel = Get_sap_mode();
            foreach (var id in frameIDs)
            {
                Rhino.Geometry.Line l = new Rhino.Geometry.Line();

                ret = mySapModel.FrameObj.GetPoints(id, ref st_point, ref end_point);
                ret = mySapModel.PointObj.GetCoordCartesian(st_point, ref x1, ref y1, ref z1);
                ret = mySapModel.PointObj.GetCoordCartesian(end_point, ref x2, ref y2, ref z2);

                if (unit.ToUpper() == "M")
                {
                    l = new Rhino.Geometry.Line(x1 * factor, y1 * factor, z1 * factor, x2 * factor, y2 * factor, z2 * factor);
                }

                else
                {
                    l = new Rhino.Geometry.Line(x1, y1, z1, x2, y2, z2);
                }

                Lines.Add(l);

            }


            return Lines;


        }


        public static List<Rhino.Geometry.Line> getFramesRhinoParrallel(List<string> frameIDs, string unit)
        {
            #region defintions 
            string st_point = null;
            string end_point = null;
            int ret;
            double x1, y1, z1, x2, y2, z2;
            x1 = y1 = z1 = x2 = y2 = z2 = 0;

            #endregion
            double factor = 1 / 1000;


            List<Rhino.Geometry.Line> Lines = new List<Rhino.Geometry.Line>();
            cSapModel mySapModel = Get_sap_mode();

            Parallel.ForEach(frameIDs, id =>
            {
                Rhino.Geometry.Line l = new Rhino.Geometry.Line();

                ret = mySapModel.FrameObj.GetPoints(id, ref st_point, ref end_point);
                ret = mySapModel.PointObj.GetCoordCartesian(st_point, ref x1, ref y1, ref z1);
                ret = mySapModel.PointObj.GetCoordCartesian(end_point, ref x2, ref y2, ref z2);

                if (unit.ToUpper() == "M")
                {
                    l = new Rhino.Geometry.Line(x1 * factor, y1 * factor, z1 * factor, x2 * factor, y2 * factor, z2 * factor);
                }

                else
                {
                    l = new Rhino.Geometry.Line(x1, y1, z1, x2, y2, z2);
                }

                Lines.Add(l);
            });


            return Lines;


        }


        public static string[] CreateTableData(string tableKey, string[] fieldKeysIncluded, int numberRecords, List<List<string>> fieldValues)
        {
            // Initialize the tableData array
            string[] tableData = new string[fieldKeysIncluded.Length * numberRecords];

            // Populate the tableData based on the provided fieldValues
            for (int i = 0; i < numberRecords; i++)
            {
                for (int j = 0; j < fieldKeysIncluded.Length; j++)
                {
                    // Get the value for the current field from the corresponding list
                    string value = fieldValues[j][i];

                    // Store the value in the tableData array
                    tableData[i * fieldKeysIncluded.Length + j] = value;
                }
            }

            // Return the generated tableData
            return tableData;
        }


        public static void rhinoToStaad(List<Frame> rhinoFrames)
        {

            #region Create_Excel_object   

            // creating Excel Application  
            Microsoft.Office.Interop.Excel.Application app = null;
            // creating new WorkBook within Excel application  
            Microsoft.Office.Interop.Excel._Workbook workbook = null;
            // creating new Excelsheet in workbook  
            Microsoft.Office.Interop.Excel._Worksheet worksheet1 = null;
            ExcelUtils.CreateExcel(ref app, ref workbook, ref worksheet1);


            string fullPath2 = workbook.FullName;
            #endregion


            RevitUtils.setRevitDataInExcelBulk(rhinoFrames, (Microsoft.Office.Interop.Excel.Worksheet)worksheet1);

            #region saveExcel
            workbook.Save();
            workbook.Close();
            app.DisplayAlerts = true;
            app.Quit();
            #endregion


            #region write RvtToStd_1
            string username = Environment.UserName;
            string folder = @"C:\\Users\\" + username + "\\Desktop\\H Y D R A _temp\\ExcDataWithRvt\\";
            string fileName = "HYDRA Integration.txt";
            // Fullpath. You can direct hardcode it if you like.  
            string fullPath = folder + fileName;
            // An array of strings  
            string[] write = { "RvtToStd_ALL" };
            // If the file does not exists, it will create a new file.  
            // This method automatically opens the file, writes to it, and closes file  
            File.WriteAllLines(fullPath, write);



            #endregion

            ////////////////////////////////////////////////////////// new 

        }


        public static void RhinoToRevitBeams(List<Frame> beams, object doc2)
        {

            Document doc = doc2 as Document;
            ForgeTypeId dis = UnitTypeId.Millimeters;

            #region defintions for creation of beams and columns 
            FamilySymbol revitSection = null;
            double angle = 0;
            Level lvl0 = new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().ToList().FirstOrDefault();

            if (lvl0 == null)
            {
                using (Transaction t = new Transaction(doc, "Create Levels"))
                {
                    t.Start();
                    lvl0 = Level.Create(doc, 0.0);
                    lvl0.Name = "Hydra-level-1";

                    t.Commit();

                }
            }
        #endregion

        Beammodeling:


            RevitServices.Transactions.TransactionManager.Instance.ForceCloseTransaction();

            using (TransactionGroup tg = new TransactionGroup(doc, "aaa"))
            {

                tg.Start();

                var chopSize = Frame.chopSize;
                var beamChunks = beams.Chunk(chopSize).ToList();

                foreach (var beamChunk in beamChunks)
                {

                    using (Transaction trns = new Transaction(doc, "trns"))
                    {


                        trns.Start();
                        FailureHandlingOptions failOpt = trns.GetFailureHandlingOptions();
                        failOpt.SetFailuresPreprocessor(new RoomWarningSwallower());
                        trns.SetFailureHandlingOptions(failOpt);
                        foreach (var bm in beamChunk)
                        {

                            var x = UnitUtils.ConvertToInternalUnits(bm.StartPoint.X, dis);
                            var y = UnitUtils.ConvertToInternalUnits(bm.StartPoint.Y, dis);
                            var z = UnitUtils.ConvertToInternalUnits(bm.StartPoint.Z, dis);
                            var StartPoint = new XYZ(x, y, z);

                            var x1 = UnitUtils.ConvertToInternalUnits(bm.EndPoint.X, dis);
                            var y1 = UnitUtils.ConvertToInternalUnits(bm.EndPoint.Y, dis);
                            var z1 = UnitUtils.ConvertToInternalUnits(bm.EndPoint.Z, dis);
                            var EndPoint = new XYZ(x1, y1, z1);



                            try
                            {
                                Curve line = Line.CreateBound(StartPoint, EndPoint);
                                angle = (bm.Beta) * Math.PI / 180;

                                revitSection = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralFraming).WhereElementIsElementType().First() as FamilySymbol;




                                if (!revitSection.IsActive)
                                    revitSection.Activate();

                                FamilyInstance instance = doc.Create.NewFamilyInstance(line, revitSection, lvl0, StructuralType.Beam);
                                instance.LookupParameter("Cross-Section Rotation").Set(angle);
                                Parameter sss = instance.LookupParameter("z Justification");
                                int zjustf = (int)ZJustification.Center;
                                sss.Set(zjustf);
                                Autodesk.Revit.DB.Structure.StructuralFramingUtils.DisallowJoinAtEnd(instance, 0);
                                Autodesk.Revit.DB.Structure.StructuralFramingUtils.DisallowJoinAtEnd(instance, 1);

                            }


                            catch (Exception ex2)

                            {
                                System.Diagnostics.StackTrace trace = new System.Diagnostics.StackTrace(ex2, true);
                                continue;

                            }
                        }
                        trns.Commit();
                    }
                }


                tg.Assimilate();


            }

        }

        public static List<DriftRecrd> framesEndPointsDrifts(List<string> groupNames, double factorForWind, double factorForSeismic, List<string> LoadCaseKeys, int targetObjectType = 2 )
        {
            List<Beam> bms = new List<Beam>();
            List<Column> cols = new List<Column>();
            int ret = 0;

            SAP2000v1.cSapModel mysapModel = Get_sap_mode();

            List<string> frameIds = getIdsOfElementsInGroup(groupNames, targetObjectType);

            List<Frame> targetFrames = getSapFrames(ref bms, ref cols);
            targetFrames = targetFrames.Where(x => frameIds.Contains(x.ID)).ToList();


            //
            int NumberResults = 0;
            string[] objPointsNamesAssotiated = null;
            string[] elementPointsNamesAssotiated = null;
            string[] returnedLoadCombAndCases = null;
            string[] StepType = null;
            double[] StepNumber = null;

            double[] U1 = null;
            double[] U2 = null;
            double[] U3 = null;
            double[] R1 = null;
            double[] R2 = null;
            double[] R3 = null;



            int NumberResults_2 = 0;
            string[] objPointsNamesAssotiated_2 = null;
            string[] elementPointsNamesAssotiated_2 = null;
            string[] returnedLoadCombAndCases_2 = null;
            string[] StepType_2 = null;
            double[] StepNumber_2 = null;

            double[] U1_2 = null;
            double[] U2_2 = null;
            double[] U3_2 = null;
            double[] R1_2 = null;
            double[] R2_2 = null;
            double[] R3_2 = null;

            double Xdiff = 0;
            double Ydiff = 0;

            List<DriftRecrd> records = new List<DriftRecrd>();

            string[] allLoadCombs = null;
            int numberOfCombs = 0;
            bool selected = false;
            List<string> combsSelectedForOutput = new List<string>();


            #region handle load combs for output

            mysapModel.RespCombo.GetNameList(ref numberOfCombs, ref allLoadCombs);


            for (int i = 0; i < numberOfCombs; i++)
            {
                mysapModel.Results.Setup.GetComboSelectedForOutput(allLoadCombs[i], ref selected);

                if (selected)
                {
                    combsSelectedForOutput.Add(allLoadCombs[i]);
                }
            }


            mysapModel.Results.Setup.DeselectAllCasesAndCombosForOutput();


            allLoadCombs =  allLoadCombs.Where(x=>LoadCaseKeys.Any(a=> x.Contains(a))).ToArray();


            for (int i = 0; i < allLoadCombs.Length; i++)
            {
                mysapModel.Results.Setup.SetComboSelectedForOutput(allLoadCombs[i]);
            }

            #endregion

            foreach (var f in targetFrames)
            {
                ret = mysapModel.Results.JointDispl(f.StartPoint.Id, eItemTypeElm.ObjectElm, ref NumberResults, ref objPointsNamesAssotiated, ref elementPointsNamesAssotiated, ref returnedLoadCombAndCases, ref StepType, ref StepNumber, ref U1, ref U2, ref U3, ref R1, ref R2, ref R3);
                ret = mysapModel.Results.JointDispl(f.EndPoint.Id, eItemTypeElm.ObjectElm, ref NumberResults_2, ref objPointsNamesAssotiated_2, ref elementPointsNamesAssotiated_2, ref returnedLoadCombAndCases_2, ref StepType_2, ref StepNumber_2, ref U1_2, ref U2_2, ref U3_2, ref R1_2, ref R2_2, ref R3_2);

                for (int i = 0; i < NumberResults; i++)
                {

                         Xdiff = Math.Abs(U1_2[i] - U1[i]);
                         Ydiff = Math.Abs(U2_2[i] - U2[i]);

                        if ((Xdiff/ f.Length) > 1/ factorForWind)
                        {
                            DriftRecrd rec = new DriftRecrd();
                            rec.f = f;
                            rec.direction = "X";

                            Point pointWithHigherZ = f.StartPoint.Z > f.EndPoint.Z ? f.StartPoint : f.EndPoint;
                            rec.pt = pointWithHigherZ;
                            rec.loadCases = returnedLoadCombAndCases[i];
                            rec.factor = 1 / (Xdiff / f.Length);
                            records.Add(rec);
                        }

                        if ((Ydiff / f.Length) > 1 / factorForWind)
                        {
                            DriftRecrd rec = new DriftRecrd();
                            rec.f = f;
                            rec.direction = "Y";

                            Point pointWithHigherZ = f.StartPoint.Z > f.EndPoint.Z ? f.StartPoint : f.EndPoint;
                            rec.pt = pointWithHigherZ;
                            rec.loadCases = returnedLoadCombAndCases[i];
                            rec.factor = 1 / (Ydiff / f.Length);
                            records.Add(rec);


                        }

                    

                        Xdiff = Math.Abs(U1_2[i] - U1[i]);
                        Ydiff = Math.Abs(U2_2[i] - U2[i]);


                        if ((Xdiff / f.Length) > 1 / factorForSeismic)
                        {
                            DriftRecrd rec = new DriftRecrd();
                            rec.f = f;
                            rec.direction = "X";

                            Point pointWithHigherZ = f.StartPoint.Z > f.EndPoint.Z ? f.StartPoint : f.EndPoint;
                            rec.pt = pointWithHigherZ;
                            rec.loadCases = returnedLoadCombAndCases[i];
                            rec.factor = 1 / (Xdiff / f.Length);
                            records.Add(rec);

                        }

                        if ((Ydiff / f.Length) > 1 / factorForSeismic)
                        {
                            DriftRecrd rec = new DriftRecrd();
                            rec.f = f;
                            rec.direction = "Y";

                            Point pointWithHigherZ = f.StartPoint.Z > f.EndPoint.Z ? f.StartPoint : f.EndPoint;
                            rec.pt = pointWithHigherZ;
                            rec.loadCases = returnedLoadCombAndCases[i];
                            rec.factor = 1 / (Ydiff / f.Length);
                            records.Add(rec);


                        }
                    
                }

            }



            // re assign the old one again 
            mysapModel.Results.Setup.DeselectAllCasesAndCombosForOutput();

            for (int i = 0; i < allLoadCombs.Length; i++)
            {
                mysapModel.Results.Setup.SetComboSelectedForOutput(allLoadCombs[i]);
            }


            return records;
        }

        public static void deleteFrameObjects(List<string> ids)
        {
            SAP2000v1.cSapModel mysap =  SapUtils.Get_sap_mode();


            foreach (var item in ids)
            {
                mysap.FrameObj.Delete(item);
            }


        }

        public static void deleteAreaObjects(List<string> ids)
        {
            SAP2000v1.cSapModel mysap = SapUtils.Get_sap_mode();


            foreach (var item in ids)
            {
                mysap.AreaObj.Delete(item);
            }


        }

        public class DriftRecrd
        {
            public Frame f { get; set; }
            public Point pt { get; set; }
            public string loadCases { get; set; }

            public string direction { get; set; }

            public double factor { get; set; }
        }

        #endregion

        #region get Sap Areas

        public static List<string> getSelectedAreaIDS()
        {


            cSapModel mySapModel = Get_sap_mode();
            List<string> result = new List<string>();


            #region defintions 
            double x1, y1, z1;
            x1 = y1 = z1 = 0;
            int NumberItems = 0;
            int[] ObjectType = null;
            string[] ObjectName = null;
            string st_id = "";
            string end_id = "";
            int ret = 0;
            #endregion

            ret = mySapModel.SelectObj.GetSelected(ref NumberItems, ref ObjectType, ref ObjectName);

            for (int i = 0; i < NumberItems; i++)
            {
                if (ObjectType[i] == 5)
                {
                    result.Add(ObjectName[i]);
                }
            }




            return result;
        }


        public static List<Area> getSapAreas()
        {
            #region definitions
            double x, y, z;
            x = y = z = 0;
            int ret = 0;
            int NumberNames = 0;
            string[] AreaIDs = new string[0];
            int numberOfPoints = 0;
            string[] PointIds = null;
            int numberOfGroups = 0;
            string[] GroupsNames = null;
            double Ang = 0;
            bool Advanced = false;
            string PropName = null;
            #endregion


            cSapModel mySapModel = SapUtils.Get_sap_mode();



            #region get areaIDS
            mySapModel.AreaObj.GetNameList(ref NumberNames, ref AreaIDs);
            List<string> selectedAreaIDS = getSelectedAreaIDS();
            if (selectedAreaIDS.Count > 0)
                AreaIDs = selectedAreaIDS.ToArray();
            #endregion


            List<Area> sapAreas = new List<Area>();

            foreach (var areaId in AreaIDs)
            {
                Area A = new Area();
                A.ID = areaId;
                A.vertices = new List<Point>();
                #region set Points
                mySapModel.AreaObj.GetPoints(areaId, ref numberOfPoints, ref PointIds);
                foreach (var pointId in PointIds)
                {
                    Point pt = new Point();
                    ret = mySapModel.PointObj.GetCoordCartesian(pointId, ref x, ref y, ref z);
                    pt.Id = pointId;
                    pt.X = x;
                    pt.Y = y;
                    pt.Z = z;
                    A.vertices.Add(pt);

                }

                #endregion


                mySapModel.AreaObj.GetGroupAssign(areaId, ref numberOfGroups, ref GroupsNames);
                A.Groups = GroupsNames.ToList();


                mySapModel.AreaObj.GetLocalAxes(areaId, ref Ang, ref Advanced);
                A.Beta = Ang;


                mySapModel.AreaObj.GetProperty(areaId, ref PropName);
                A.Sec = PropName;

                if (A.Sec == "None")
                {
                    A.type = "None";

                }

                else if (A.Sec.ToUpper() == "WALL")
                {

                    A.type = "Wall";

                }

                else
                {
                    A.type = "Floor";

                }




                //mySapModel.AreaObj.GetProperty(areaId, ref PropName);
                //A.Sec = PropName;

                //if (A.Sec == "None")
                //{
                //    A.type = "None";

                //}

                //else if (Area.IsAreaVertical(A, 100))
                //{

                //    A.type = "Wall";

                //}

                //else
                //{
                //    A.type = "Floor";

                //}




                sapAreas.Add(A);
            }

            return sapAreas;
        }

        #endregion

        #region SapToEtabsClasses

        public class DataBase
        {

            #region SapToEtabs  

            // final method
            public static void ExportSapToEtabs(string pointSGroupForStories)
            {

                cSapModel mysap = Get_sap_mode();
                ETABSv1.cSapModel etab = EtabsUtils.Get_sap_mode();

                CreateStories(mysap, etab, pointSGroupForStories);

                #region definisions

                bool FillImportLog = true;
                int NumberFatalErrors = 0;
                int NumberErrorMsg = 0;
                int NumWarrinings = 0;
                string importLog = "";
                int NumberInfoMsg = 0;
                int ret = 0;
                ETABSv1.cSapModel myEtabs = EtabsUtils.Get_sap_mode();
                Dictionary<string, string> dic = null;

                #endregion

                #region materils and Groups definitions
                ret = mysap.DatabaseTables.GetTableForDisplayArray(MaterialProp, ref MaterialPropFileds, "", ref tableVersion, ref filedsReturned, ref numberOFrRectors, ref MaterialPropValues);
                ret = myEtabs.DatabaseTables.SetTableForEditingArray(EtabsUtils.DataBase.MaterialProp, ref tableVersion, ref EtabsUtils.DataBase.MaterialPropFileds, numberOFrRectors, ref MaterialPropValues);


                ret = mysap.DatabaseTables.GetTableForDisplayArray(MaterialMechProp, ref MaterialMechPropFileds, "", ref tableVersion, ref filedsReturned, ref numberOFrRectors, ref MaterialMechPropValues);
                ret = myEtabs.DatabaseTables.SetTableForEditingArray(EtabsUtils.DataBase.MaterialMechProp, ref tableVersion, ref EtabsUtils.DataBase.MaterialMechPropFileds, numberOFrRectors, ref MaterialMechPropValues);

                ret = mysap.DatabaseTables.GetTableForDisplayArray(MaterialPropConc, ref MaterialPropConcFileds, "", ref tableVersion, ref filedsReturned, ref numberOFrRectors, ref MaterialPropConcValues);
                ret = myEtabs.DatabaseTables.SetTableForEditingArray(EtabsUtils.DataBase.MaterialPropConc, ref tableVersion, ref EtabsUtils.DataBase.MaterialPropConcFileds, numberOFrRectors, ref MaterialPropConcValues);


                ret = mysap.DatabaseTables.GetTableForDisplayArray(MaterialPropSteel, ref MaterialPropSteelFileds, "", ref tableVersion, ref filedsReturned, ref numberOFrRectors, ref MaterialPropSteelValues);
                ret = myEtabs.DatabaseTables.SetTableForEditingArray(EtabsUtils.DataBase.MaterialPropSteel, ref tableVersion, ref EtabsUtils.DataBase.MaterialPropSteelFileds, numberOFrRectors, ref MaterialPropSteelValues);



                ret = mysap.DatabaseTables.GetTableForDisplayArray(GroupDefinitions, ref GroupDefinitionsFileds, "", ref tableVersion, ref filedsReturned, ref numberOFrRectors, ref GroupDefinitionsValues);
                ret = myEtabs.DatabaseTables.SetTableForEditingArray(EtabsUtils.DataBase.GroupDefinitions, ref tableVersion, ref EtabsUtils.DataBase.GroupDefinitionsFileds, numberOFrRectors, ref GroupDefinitionsValues);



                ret = myEtabs.DatabaseTables.ApplyEditedTables(FillImportLog, ref NumberFatalErrors, ref NumberErrorMsg, ref NumWarrinings, ref NumberInfoMsg, ref importLog);
                // Create a new text file and open it for writing
                //using (StreamWriter writer = new StreamWriter(filePath))
                //{

                //    writer.Write(importLog);
                //}


                #endregion

                List<Frame> sapFrames = getSapFramesE();
                List<Point> SapPoints = getSapPoints(mysap);
                SapToEtabsDefineLoads();
                SapToEtabsDefineSections();
                createDiaphrams(mysap, etab);
                SatpToEtabsDefineGroups(mysap, etab);


                createPoints(SapPoints, etab);
                SapToEtabsPointAssignmentsAndDiaphram(mysap, myEtabs, SapPoints);

                List<Frame> reversedFrames = SapToEtabsDrawFrames(sapFrames, mysap, etab);
                SapToEtabsFramesAssignments(sapFrames, mysap, etab, reversedFrames);


                SapToEtabsExportAreaWithAssignments2();


                // shadi modifications
                SapToEtabsSetEdgeConstainsForAreas(myEtabs);
                SatpToEtabsSetEndLengthOffsetForConcFrames(myEtabs, sapFrames);
                SapToEtabsSetEndLengthOffsetForSteelFrames(myEtabs, sapFrames);
                // SetFixedSupportsAtConcAreas(etab);
                //




                myEtabs.View.RefreshView();

                ret = mysap.DatabaseTables.GetTableForDisplayArray(EtabsUtils.DataBase.FrameLoadsDist, ref EtabsUtils.DataBase.FrameLoadsDistFileds, "", ref tableVersion, ref filedsReturned, ref numberOFrRectors, ref EtabsUtils.DataBase.FrameLoadsDistValues);
                ret = myEtabs.DatabaseTables.SetTableForEditingArray(EtabsUtils.DataBase.FrameLoadsDist, ref tableVersion, ref EtabsUtils.DataBase.FrameLoadsDistFileds, numberOFrRectors, ref EtabsUtils.DataBase.FrameLoadsDistValues);

                ret = myEtabs.DatabaseTables.ApplyEditedTables(FillImportLog, ref NumberFatalErrors, ref NumberErrorMsg, ref NumWarrinings, ref NumberInfoMsg, ref importLog);



            }

            public static List<Point> SapToEtabsExportAreaWithAssignments()
            {
                #region def

                string[] Groups = new string[0];
                double[] Offset1 = new double[0];
                double[] Offset2 = new double[0];

                int NumberItems = 0;
                string[] FrameName = new string[0];
                string[] LoadPat = new string[0];


                int[] MyType = new int[0];
                int[] Dir = new int[0];
                double[] RD1 = new double[0];
                double[] RD2 = new double[0];
                double[] Dist1 = new double[0];
                double[] Dist2 = new double[0];
                double[] Val1 = new double[0];
                double[] Val2 = new double[0];



                string[] PointName = new string[0];
                int[] LCStep = new int[0];
                string[] CSys3 = new string[0];
                double[] F1 = new double[0];
                double[] F2 = new double[0];
                double[] F3 = new double[0];
                double[] M1 = new double[0];
                double[] M2 = new double[0];
                double[] M3 = new double[0];
                double[] Value = new double[0];

                bool[] restValues = new bool[0];


                #endregion



                cSapModel mysap = Get_sap_mode();
                ETABSv1.cSapModel myetab = EtabsUtils.Get_sap_mode();
                List<Area> sapAreas = getSapAreas();
                List<Point> pts = new List<Point>();
                foreach (var ar in sapAreas)
                {
                    pts.AddRange(ar.vertices);
                }


                pts = pts.GroupBy(x => x.Id).Select(x => x.FirstOrDefault()).ToList();

                createPoints(pts, myetab);
                defineAreaSections(mysap, myetab, sapAreas);
                drawAreas(myetab, sapAreas);
                AreasAssignments(mysap, myetab, sapAreas);


                return pts;



            }

            public static void SapToEtabsExportAreaWithAssignments2()
            {
                #region def

                string[] Groups = new string[0];
                double[] Offset1 = new double[0];
                double[] Offset2 = new double[0];

                int NumberItems = 0;
                string[] FrameName = new string[0];
                string[] LoadPat = new string[0];


                int[] MyType = new int[0];
                int[] Dir = new int[0];
                double[] RD1 = new double[0];
                double[] RD2 = new double[0];
                double[] Dist1 = new double[0];
                double[] Dist2 = new double[0];
                double[] Val1 = new double[0];
                double[] Val2 = new double[0];



                string[] PointName = new string[0];
                int[] LCStep = new int[0];
                string[] CSys3 = new string[0];
                double[] F1 = new double[0];
                double[] F2 = new double[0];
                double[] F3 = new double[0];
                double[] M1 = new double[0];
                double[] M2 = new double[0];
                double[] M3 = new double[0];
                double[] Value = new double[0];

                bool[] restValues = new bool[0];


                #endregion



                cSapModel mysap = Get_sap_mode();
                ETABSv1.cSapModel myetab = EtabsUtils.Get_sap_mode();
                List<Area> sapAreas = getSapAreas();



                defineAreaSections(mysap, myetab, sapAreas);
                drawAreas(myetab, sapAreas);
                AreasAssignments(mysap, myetab, sapAreas);





            }


            #region loadDefinitions
            public static void SapToEtabsDefineLoads()
            {
                LoadPattern.SetLoadPatternsInEtabs(LoadPattern.getLoadPatternsFromSap());
                LoadCase.SetLinearLoadCasesInEtabs(LoadCase.GetLinearLoadcases());
                LoadCase.SetNonLinearLoadCasesInEtabs(LoadCase.GetNonLinearLoadcases());
                LOadComb.SetLoadCombsInEtabs(LOadComb.getLoadCombsFromSap());
            }

            #endregion

            #region Sections

            public static void SapToEtabsDefineSections()
            {
                cSapModel mySapModel = Get_sap_mode();
                ETABSv1.cSapModel myEtabs = EtabsUtils.Get_sap_mode();


                List<string> SapSectionS = SapUtils.Sections.GetAllNamesAssignedToFramesInSap(mySapModel);
                Dictionary<string, Section_type> mydic = SapUtils.Sections.getSectionTypes(SapSectionS);

                List<string> iSection = mydic.Where(x => x.Value == Section_type.I_SECTION_B || x.Value == Section_type.I_SECTION_H).ToDictionary(x => x.Key, x => x.Value).Keys.ToList();
                List<string> boxSection = mydic.Where(x => x.Value == Section_type.Hollow_box_B || x.Value == Section_type.Hollow_box_H).ToDictionary(x => x.Key, x => x.Value).Keys.ToList();
                List<string> pipeSection = mydic.Where(x => x.Value == Section_type.Hollow_pipe_B || x.Value == Section_type.Hollow_pipe_H).ToDictionary(x => x.Key, x => x.Value).Keys.ToList();
                List<string> channelSection = mydic.Where(x => x.Value == Section_type.Channel_B || x.Value == Section_type.Channel_H).ToDictionary(x => x.Key, x => x.Value).Keys.ToList();
                List<string> iTaperedSection = mydic.Where(x => x.Value == Section_type.I_Tapered).ToDictionary(x => x.Key, x => x.Value).Keys.ToList();
                List<string> boxTaperedSection = mydic.Where(x => x.Value == Section_type.Box_Tapered).ToDictionary(x => x.Key, x => x.Value).Keys.ToList();
                List<string> pipeTaperedSection = mydic.Where(x => x.Value == Section_type.Pipe_Tapered).ToDictionary(x => x.Key, x => x.Value).Keys.ToList();
                List<string> solidPipeSection = mydic.Where(x => x.Value == Section_type.Solid_pipe).ToDictionary(x => x.Key, x => x.Value).Keys.ToList();
                List<string> solidBoxSection = mydic.Where(x => x.Value == Section_type.Solid_box).ToDictionary(x => x.Key, x => x.Value).Keys.ToList();
                List<string> teeSection = mydic.Where(x => x.Value == Section_type.T_Section_B || x.Value == Section_type.T_Section_B).ToDictionary(x => x.Key, x => x.Value).Keys.ToList();
                List<string> singleAngleSection = mydic.Where(x => x.Value == Section_type.S_Angle_B || x.Value == Section_type.S_Angle_H).ToDictionary(x => x.Key, x => x.Value).Keys.ToList();
                List<string> dblAngleSection = mydic.Where(x => x.Value == Section_type.DblAngle || x.Value == Section_type.D_Angle_B).ToDictionary(x => x.Key, x => x.Value).Keys.ToList();

                List<string> concrBeamRectSection = mydic.Where(x => x.Value == Section_type.Concrete_rect).ToDictionary(x => x.Key, x => x.Value).Keys.ToList();
                List<string> concrBeamPipeSection = mydic.Where(x => x.Value == Section_type.Concrete_circ).ToDictionary(x => x.Key, x => x.Value).Keys.ToList();

                List<string> sectionDesignerSection = mydic.Where(x => x.Value == Section_type.Section_Designer).ToDictionary(x => x.Key, x => x.Value).Keys.ToList();



                Get_i_section_dimentions(mySapModel, iSection, myEtabs);
                Get_Box_section_dimentions(mySapModel, boxSection, myEtabs);
                Get_pipe_section_dimentions(mySapModel, pipeSection, myEtabs);
                Get_channel_section_dimentions(mySapModel, channelSection, myEtabs);
                Get_Tapered_i_section_dimentions(mySapModel, iTaperedSection, myEtabs);
                Get_Tapered_BOX_section_dimentions(mySapModel, boxTaperedSection, myEtabs);
                Get_Tapered_PIPE_section_dimentions(mySapModel, pipeTaperedSection, myEtabs);
                Get_solid_pipe_section_dimentions(mySapModel, solidPipeSection, myEtabs);
                Get_solid_rect_section_dimentions(mySapModel, solidBoxSection, myEtabs);
                Get_solid_rect_section_dimentions(mySapModel, concrBeamRectSection, myEtabs);
                Get_Sig_angle_section_dimentions(mySapModel, singleAngleSection, myEtabs);
                Get_dbl_angle_section_dimentions(mySapModel, dblAngleSection, myEtabs);
                Get_solid_pipe_section_dimentions(mySapModel, concrBeamPipeSection, myEtabs);

                SetModifiers(SapSectionS, mySapModel, myEtabs);

            }
            public static void SetModifiers(List<string> profiles, cSapModel mysapmodel, ETABSv1.cSapModel myEtabs)
            {
                double[] Value = null;
                foreach (var item in profiles)
                {
                    mysapmodel.PropFrame.GetModifiers(item, ref Value);

                    myEtabs.PropFrame.SetModifiers(item, ref Value);
                }


            }
            public static void Get_i_section_dimentions(cSapModel mySapModel, List<string> profiles, ETABSv1.cSapModel myEtabs)
            {

                #region defintions 
                string FileName = null;
                string MatProp = null;
                double t3 = 0;
                double t2 = 0;
                double tf = 0;
                double tw = 0;
                double t2b = 0;
                double tfb = 0;
                int Color = 0;
                string Notes = null;
                string GUID = null;
                profiles = profiles.Distinct().ToList();

                #endregion


                foreach (string profile in profiles)
                {
                    int ret = mySapModel.PropFrame.GetISection(profile, ref FileName, ref MatProp, ref t3, ref t2, ref tf, ref tw, ref t2b, ref tfb, ref Color, ref Notes, ref GUID);
                    if (FileName != "" && FileName != null)
                    {

                        FileName = DataBase.CapitalizeVariableName(FileName);
                        if (FileName.ToUpper().Contains("ARCELORMITTAL_BRITISH"))
                        {
                            FileName = "Arcelormittal_British";
                        }
                        ret = myEtabs.PropFrame.ImportProp(profile, MatProp, FileName + ".xml", profile, Color, Notes);
                        continue;
                    }

                    tw = System.Math.Round(tw, 2);
                    tf = System.Math.Round(tf, 2);
                    t3 = System.Math.Round(t3, 2);
                    t2 = System.Math.Round(t2, 2);


                    ret = myEtabs.PropFrame.SetISection(profile, MatProp, t3, t2, tf, tw, t2b, tfb, Color, Notes);



                }

            }

            public static void Get_Box_section_dimentions(cSapModel mySapModel, List<string> profiles, ETABSv1.cSapModel myEtabs)
            {

                #region defintions 
                string FileName = null;
                string MatProp = null;
                double t3 = 0;
                double t2 = 0;
                double tf = 0;
                double tw = 0;
                int Color = 0;
                string Notes = null;
                string GUID = null;
                profiles = profiles.Distinct().ToList();
                #endregion

                List<object> section_dimentions = new List<object>();

                foreach (string profile in profiles)
                {
                    int ret = mySapModel.PropFrame.GetTube(profile, ref FileName, ref MatProp, ref t3, ref t2, ref tf, ref tw, ref Color, ref Notes, ref GUID);
                    if (FileName != "" && FileName != null)
                    {

                        FileName = DataBase.CapitalizeVariableName(FileName);
                        ret = myEtabs.PropFrame.ImportProp(profile, MatProp, FileName + ".xml", profile, Color, Notes);
                        continue;
                    }

                    tw = System.Math.Round(tw, 2);
                    tf = System.Math.Round(tf, 2);
                    t3 = System.Math.Round(t3, 2);
                    t2 = System.Math.Round(t2, 2);




                    ret = myEtabs.PropFrame.SetTube(profile, MatProp, t3, t2, tf, tw, Color, Notes);


                }


            }

            public static void Get_pipe_section_dimentions(cSapModel mySapModel, List<string> profiles, ETABSv1.cSapModel myEtabs)
            {

                #region defintions 
                string FileName = null;
                string MatProp = null;
                double t3 = 0;
                double tw = 0;

                int Color = 0;
                string Notes = null;
                string GUID = null;
                profiles = profiles.Distinct().ToList();
                #endregion


                foreach (string profile in profiles)
                {
                    int ret = mySapModel.PropFrame.GetPipe(profile, ref FileName, ref MatProp, ref t3, ref tw, ref Color, ref Notes, ref GUID);

                    if (FileName != "" && FileName != null)
                    {

                        FileName = DataBase.CapitalizeVariableName(FileName);
                        ret = myEtabs.PropFrame.ImportProp(profile, MatProp, FileName + ".xml", profile, Color, Notes);
                        continue;
                    }
                    t3 = System.Math.Round(t3, 2);
                    tw = System.Math.Round(tw, 2);

                    ret = myEtabs.PropFrame.SetPipe(profile, MatProp, t3, tw, Color, Notes);



                    //handle if file name is not blank (improted section) you need to access xml files

                }

            }

            public static void Get_channel_section_dimentions(cSapModel mySapModel, List<string> profiles, ETABSv1.cSapModel myEtabs)
            {

                #region defintions 
                string FileName = null;
                string MatProp = null;
                double t3 = 0;
                double t2 = 0;
                double tf = 0;
                double tw = 0;
                int Color = 0;
                string Notes = null;
                string GUID = null;
                profiles = profiles.Distinct().ToList();
                #endregion


                foreach (string profile in profiles)
                {
                    int ret = mySapModel.PropFrame.GetChannel(profile, ref FileName, ref MatProp, ref t3, ref t2, ref tf, ref tw, ref Color, ref Notes, ref GUID);
                    if (FileName != "" && FileName != null)
                    {

                        FileName = DataBase.CapitalizeVariableName(FileName);
                        ret = myEtabs.PropFrame.ImportProp(profile, MatProp, FileName + ".xml", profile, Color, Notes);
                        continue;
                    }

                    tw = System.Math.Round(tw, 2);
                    tf = System.Math.Round(tf, 2);
                    t3 = System.Math.Round(t3, 2);
                    t2 = System.Math.Round(t2, 2);

                    ret = myEtabs.PropFrame.SetChannel(profile, MatProp, t3, t2, tf, tw, Color, Notes);




                }

            }

            public static void Get_dbl_angle_section_dimentions(cSapModel mySapModel, List<string> profiles, ETABSv1.cSapModel myEtabs)
            {

                #region defintions 
                string FileName = null;
                string MatProp = null;
                double t3 = 0;
                double t2 = 0;
                double tf = 0;
                double tw = 0;
                double dis = 0;
                int Color = 0;
                string Notes = null;
                string GUID = null;
                profiles = profiles.Distinct().ToList();
                #endregion


                foreach (string profile in profiles)
                {

                    int ret = mySapModel.PropFrame.GetDblAngle(profile, ref FileName, ref MatProp, ref t3, ref t2, ref tf, ref tw, ref dis, ref Color, ref Notes, ref GUID);
                    if (FileName != "" && FileName != null)
                    {

                        FileName = DataBase.CapitalizeVariableName(FileName);
                        ret = myEtabs.PropFrame.ImportProp(profile, MatProp, FileName + ".xml", profile, Color, Notes);
                        continue;
                    }

                    tw = System.Math.Round(tw, 2);
                    tf = System.Math.Round(tf, 2);
                    t3 = System.Math.Round(t3, 2);
                    t2 = System.Math.Round(t2, 2);

                    ret = myEtabs.PropFrame.SetDblAngle(profile, MatProp, t3, t2, tf, tw, dis, Color, Notes);
                }


            }

            public static void Get_Sig_angle_section_dimentions(cSapModel mySapModel, List<string> profiles, ETABSv1.cSapModel myEtabs)
            {

                #region defintions 
                string FileName = null;
                string MatProp = null;
                double t3 = 0;
                double t2 = 0;
                double tf = 0;
                double tw = 0;
                int Color = 0;
                string Notes = null;
                string GUID = null;
                profiles = profiles.Distinct().ToList();
                #endregion


                foreach (string profile in profiles)
                {
                    int ret = mySapModel.PropFrame.GetAngle(profile, ref FileName, ref MatProp, ref t3, ref t2, ref tf, ref tw, ref Color, ref Notes, ref GUID);

                    if (FileName != "" && FileName != null)
                    {

                        FileName = DataBase.CapitalizeVariableName(FileName);
                        ret = myEtabs.PropFrame.ImportProp(profile, MatProp, FileName + ".xml", profile, Color, Notes);
                        continue;
                    }

                    tw = System.Math.Round(tw, 2);
                    tf = System.Math.Round(tf, 2);
                    t3 = System.Math.Round(t3, 2);
                    t2 = System.Math.Round(t2, 2);


                    ret = myEtabs.PropFrame.SetAngle(profile, MatProp, t3, t2, tf, tw, Color, Notes);


                }

            }

            public static void Get_Tee_section_dimentions(cSapModel mySapModel, List<string> profiles, ETABSv1.cSapModel myEtabs)
            {

                #region defintions 
                string FileName = null;
                string MatProp = null;
                double t3 = 0;
                double t2 = 0;
                double tf = 0;
                double tw = 0;
                int Color = 0;
                string Notes = null;
                string GUID = null;
                profiles = profiles.Distinct().ToList();
                #endregion


                foreach (string profile in profiles)
                {
                    int ret = mySapModel.PropFrame.GetTee(profile, ref FileName, ref MatProp, ref t3, ref t2, ref tf, ref tw, ref Color, ref Notes, ref GUID);
                    if (FileName != "" && FileName != null)
                    {

                        FileName = DataBase.CapitalizeVariableName(FileName);
                        ret = myEtabs.PropFrame.ImportProp(profile, MatProp, FileName + ".xml", profile, Color, Notes);
                        continue;
                    }


                    tw = System.Math.Round(tw, 2);
                    tf = System.Math.Round(tf, 2);
                    t3 = System.Math.Round(t3, 2);
                    t2 = System.Math.Round(t2, 2);

                    ret = myEtabs.PropFrame.SetTee(profile, MatProp, t3, t2, tf, tw, Color, Notes);


                }

            }

            public static void Get_solid_pipe_section_dimentions(cSapModel mySapModel, List<string> profiles, ETABSv1.cSapModel myEtabs)
            {

                #region defintions 
                string FileName = null;
                string MatProp = null;
                double t3 = 0;
                int Color = 0;
                string Notes = null;
                string GUID = null;
                profiles = profiles.Distinct().ToList();
                #endregion


                foreach (string profile in profiles)
                {
                    int ret = mySapModel.PropFrame.GetCircle(profile, ref FileName, ref MatProp, ref t3, ref Color, ref Notes, ref GUID);
                    if (FileName != "" && FileName != null)
                    {

                        FileName = DataBase.CapitalizeVariableName(FileName);
                        ret = myEtabs.PropFrame.ImportProp(profile, MatProp, FileName + ".xml", profile, Color, Notes);
                        continue;
                    }
                    ret = myEtabs.PropFrame.SetCircle(profile, MatProp, t3, Color, Notes);
                }


            }




            public static void Get_solid_rect_section_dimentions(cSapModel mySapModel, List<string> profiles, ETABSv1.cSapModel myEtabs)
            {

                #region defintions 
                string FileName = null;
                string MatProp = null;
                double t3 = 0;
                double t2 = 0;
                int Color = 0;
                string Notes = null;
                string GUID = null;
                profiles = profiles.Distinct().ToList();
                #endregion




                foreach (string profile in profiles)
                {
                    int ret = mySapModel.PropFrame.GetRectangle(profile, ref FileName, ref MatProp, ref t3, ref t2, ref Color, ref Notes, ref GUID);
                    ret = myEtabs.PropFrame.SetRectangle(profile, MatProp, t3, t2, Color, Notes);
                }

            }

            public static void Get_Tapered_i_section_dimentions(cSapModel mySapModel, List<string> profiles, ETABSv1.cSapModel myEtabs)
            {

                #region defintions 
                string FileName = null;
                string MatProp = null;
                double t3 = 0;
                double t2 = 0;
                double tf = 0;
                double tw = 0;
                double t2b = 0;
                double tfb = 0;
                double t32 = 0;
                double t22 = 0;
                double tf2 = 0;
                double tw2 = 0;
                double t2b2 = 0;
                double tfb2 = 0;
                int ret = 0;
                int Color = 0;
                string Notes = null;
                string GUID = null;
                int NumberItems = 0;
                string[] StartSec = null;
                string[] EndSec = null;
                double[] MyLength = null;
                int[] MyType = null;
                int[] EI33 = null;
                int[] EI22 = null;

                profiles = profiles.Distinct().ToList();
                #endregion




                foreach (string profile in profiles)
                {
                    ret = mySapModel.PropFrame.GetNonPrismatic(profile, ref NumberItems, ref StartSec, ref EndSec, ref MyLength, ref MyType, ref EI33, ref EI22, ref Color, ref Notes, ref GUID);

                    ret = myEtabs.PropFrame.SetNonPrismatic(profile, NumberItems, ref StartSec, ref EndSec, ref MyLength, ref MyType, ref EI33, ref EI22, Color, Notes);


                }

            }

            public static void Get_Tapered_BOX_section_dimentions(cSapModel mySapModel, List<string> profiles, ETABSv1.cSapModel myEtabs)
            {

                #region defintions 
                string FileName = null;
                string MatProp = null;
                double t3 = 0;
                double t2 = 0;
                double tf = 0;
                double tw = 0;
                double t32 = 0;
                double t22 = 0;
                double tf2 = 0;
                double tw2 = 0;
                int ret = 0;
                int Color = 0;
                string Notes = null;
                string GUID = null;
                int NumberItems = 0;
                string[] StartSec = null;
                string[] EndSec = null;
                double[] MyLength = null;
                int[] MyType = null;
                int[] EI33 = null;
                int[] EI22 = null;

                profiles = profiles.Distinct().ToList();
                #endregion

                List<object> section_dimentions = new List<object>();


                foreach (string profile in profiles)
                {
                    ret = mySapModel.PropFrame.GetNonPrismatic(profile, ref NumberItems, ref StartSec, ref EndSec, ref MyLength, ref MyType, ref EI33, ref EI22, ref Color, ref Notes, ref GUID);

                    ret = myEtabs.PropFrame.SetNonPrismatic(profile, NumberItems, ref StartSec, ref EndSec, ref MyLength, ref MyType, ref EI33, ref EI22, Color, Notes);


                }

            }

            public static void Get_Tapered_PIPE_section_dimentions(cSapModel mySapModel, List<string> profiles, ETABSv1.cSapModel myEtabs)
            {

                #region defintions 
                string FileName = null;
                string MatProp = null;
                double t3 = 0;
                double tw = 0;
                double t32 = 0;
                double tw2 = 0;
                int ret = 0;
                int Color = 0;
                string Notes = null;
                string GUID = null;
                int NumberItems = 0;
                string[] StartSec = null;
                string[] EndSec = null;
                double[] MyLength = null;
                int[] MyType = null;
                int[] EI33 = null;
                int[] EI22 = null;

                profiles = profiles.Distinct().ToList();
                #endregion

                List<object> section_dimentions = new List<object>();




                foreach (string profile in profiles)
                {
                    ret = mySapModel.PropFrame.GetNonPrismatic(profile, ref NumberItems, ref StartSec, ref EndSec, ref MyLength, ref MyType, ref EI33, ref EI22, ref Color, ref Notes, ref GUID);

                    ret = myEtabs.PropFrame.SetNonPrismatic(profile, NumberItems, ref StartSec, ref EndSec, ref MyLength, ref MyType, ref EI33, ref EI22, Color, Notes);


                }

            }



            public static void Get_SecDesigner_section_dimentions(cSapModel mySapModel, List<string> profiles, ETABSv1.cSapModel myEtabs)
            {

                #region defintions 
                string FileName = null;
                string MatProp = null;
                double t3 = 0;
                int Color = 0;
                string Notes = null;
                string GUID = null;
                profiles = profiles.Distinct().ToList();
                int ret = 1;
                int NumberItems = 0;
                string[] ShapeName = null;
                int[] MyType = null;
                int DesignType = 0;


                string PropName = null;
                double XCenter = 11;
                double YCenter = 11;
                double h = 11;
                double bf = 11;
                double tf = 11;
                double tw = 11;
                double bfb = 11;
                double tfb = 11;
                double Rotation = 11;
                double dist = 0;

                #endregion


                foreach (string profile in profiles)
                {


                    ret = mySapModel.PropFrame.GetSDSection(profile, ref MatProp, ref NumberItems, ref ShapeName, ref MyType, ref DesignType,
                        ref Color, ref Notes, ref GUID);

                    for (int i = 0; i < NumberItems; i++)
                    {
                        // i sefction
                        if (MyType[i] == 1)
                        {
                            ret = mySapModel.PropFrame.SDShape.GetISection(profile, ShapeName[i], ref MatProp, ref PropName, ref Color, ref XCenter, ref YCenter, ref h, ref bf, ref tf, ref tw, ref bfb, ref tfb, ref Rotation);

                        }
                        // channel
                        else if (MyType[i] == 2)
                        {
                            ret = mySapModel.PropFrame.SDShape.GetChannel(profile, ShapeName[i], ref MatProp, ref PropName, ref Color, ref XCenter, ref YCenter, ref h, ref bf, ref tf, ref tw, ref Rotation);

                        }
                        // Tee
                        else if (MyType[i] == 3)
                        {
                            ret = mySapModel.PropFrame.SDShape.GetTee(profile, ShapeName[i], ref MatProp, ref PropName,
                                ref Color, ref XCenter, ref YCenter, ref h, ref bf, ref tf, ref tw, ref Rotation);

                        }

                        // Angle
                        else if (MyType[i] == 4)
                        {
                            ret = mySapModel.PropFrame.SDShape.GetAngle(profile, ShapeName[i], ref MatProp, ref PropName, ref Color, ref XCenter, ref YCenter, ref h, ref bf, ref tf, ref tw, ref Rotation);

                        }
                        // doubleAngle
                        else if (MyType[i] == 5)
                        {
                            ret = mySapModel.PropFrame.SDShape.GetDblAngle(profile, ShapeName[i], ref MatProp, ref PropName, ref Color, ref XCenter, ref YCenter, ref h, ref bf, ref tf, ref tw, ref dist, ref Rotation);

                        }
                        // Box
                        else if (MyType[i] == 6)
                        {
                            ret = mySapModel.PropFrame.SDShape.GetTube(profile, ShapeName[i], ref MatProp, ref PropName, ref Color, ref XCenter, ref YCenter, ref h, ref bf, ref tf, ref tw, ref Rotation);

                        }
                        // Pipe
                        else if (MyType[i] == 7)
                        {
                            ret = mySapModel.PropFrame.SDShape.GetPipe(profile, ShapeName[i], ref MatProp, ref PropName, ref Color, ref XCenter, ref YCenter, ref h, ref bf);

                        }
                        // plate
                        else if (MyType[i] == 8)
                        {
                            ret = mySapModel.PropFrame.SDShape.GetPlate(profile, ShapeName[i], ref MatProp, ref Color, ref XCenter, ref YCenter, ref h, ref bf, ref Rotation);

                        }

                        //SolidSeg
                        else if (MyType[i] == 103)
                        {
                            ret = mySapModel.PropFrame.SDShape.GetSolidSegment(profile, ShapeName[i], ref MatProp, ref Color, ref XCenter, ref YCenter, ref h, ref bf, ref Rotation);

                        }

                        // Solid Sector
                        else if (MyType[i] == 104)
                        {
                            ret = mySapModel.PropFrame.SDShape.GetSolidSector(profile, ShapeName[i], ref MatProp, ref Color, ref XCenter, ref YCenter, ref h, ref bf, ref Rotation);

                        }

                    }

                }


            }

            #endregion

            #region GroupDefinitons
            public static void SatpToEtabsDefineGroups(cSapModel mysap, ETABSv1.cSapModel myetab)
            {

                int NumberNames = 0;
                string[] myName = null;
                mysap.GroupDef.GetNameList(ref NumberNames, ref myName);


                foreach (var item in myName)
                {
                    myetab.GroupDef.SetGroup(item);

                }


            }


            #endregion

            #region drawFrames
            public static List<Frame> SapToEtabsDrawFrames(List<Frame> frames, cSapModel mysap, ETABSv1.cSapModel myEtabs)
            {
                string id1 = null;
                string id2 = null;

                string FrameId = null;
                List<Frame> ReversedFrames = new List<Frame>();
                int ret = 0;
                foreach (var item in frames)
                {
                    //ret = myEtabs.PointObj.AddCartesian(item.StartPoint.X, item.StartPoint.Y, item.StartPoint.Z, ref id1, item.StartPoint.Id);
                    //ret = myEtabs.PointObj.AddCartesian(item.EndPoint.X, item.EndPoint.Y, item.EndPoint.Z, ref id2, item.EndPoint.Id);
                    //ret = myEtabs.FrameObj.AddByPoint(id1, id2, ref FrameId, item.Sec, item.ID);

                    ret = myEtabs.FrameObj.AddByPoint(item.StartPoint.Id, item.EndPoint.Id, ref FrameId, item.Sec, item.ID);


                    myEtabs.FrameObj.GetPoints(FrameId, ref id1, ref id2);

                    if (id1 != item.StartPoint.Id)
                    {
                        ReversedFrames.Add(item);
                    }


                }

                return ReversedFrames;
            }

            #endregion

            #region Frames Assignments groups, insertion points, output stations and load assignments

            public static void SapToEtabsFramesAssignments(List<Frame> sapframes, cSapModel mysap, ETABSv1.cSapModel myetab, List<Frame> reversedFrames)
            {
                #region def
                int numberGroups = 0;
                string[] Groups = new string[0];
                int CardinalPoint = 0;
                bool Mirror2 = false;
                bool StiffTransform = false;
                double[] Offset1 = new double[0];
                double[] Offset2 = new double[0];
                string CSys = null;

                int myType = 0;
                double MaxSize = 0;
                int MinSect = 0;
                bool NoOut = false;
                bool NoutAndDesign = false;
                int NumberItems = 0;
                string[] FrameName = new string[0];
                string[] LoadPat = new string[0];
                string[] CSys2 = null;

                int[] MyType = new int[0];
                int[] Dir = new int[0];
                double[] RD1 = new double[0];
                double[] RD2 = new double[0];
                double[] Dist1 = new double[0];
                double[] Dist2 = new double[0];
                double[] Val1 = new double[0];
                double[] Val2 = new double[0];
                string[] patternName = null;


                bool[] II = null;
                bool[] JJ = null;
                double[] StartValue = null;
                double[] EndValue = null;

                string[] PointName = new string[0];
                int[] LCStep = new int[0];
                string[] CSys3 = new string[0];
                double[] F1 = new double[0];
                double[] F2 = new double[0];
                double[] F3 = new double[0];
                double[] M1 = new double[0];
                double[] M2 = new double[0];
                double[] M3 = new double[0];
                double[] Value = new double[0];

                bool[] restValues = new bool[0];

                eFramePropType e = eFramePropType.Circle;
                string[] StartSec = null;
                string[] EndSec = null;
                double[] MYLength = null;
                int[] Ei33 = null;
                int[] Ei22 = null;
                int color = 0;
                string Notes = null;
                string guid = null;
                int[] Mytype2 = null;
                #endregion

                List<string> ptIds = sapframes.Select(x => x.StartPoint.Id).ToList();
                ptIds.AddRange(sapframes.Select(x => x.EndPoint.Id).ToList());

                ptIds = ptIds.Distinct().ToList();

                foreach (var item in sapframes)
                {
                    // group assignments
                    mysap.FrameObj.GetGroupAssign(item.ID, ref numberGroups, ref Groups);
                    foreach (var g in Groups)
                    {
                        myetab.FrameObj.SetGroupAssign(item.ID, g, false);
                    }

                    // insertionpoints
                    mysap.FrameObj.GetInsertionPoint(item.ID, ref CardinalPoint, ref Mirror2, ref StiffTransform, ref Offset1, ref Offset2, ref CSys);
                    myetab.FrameObj.SetInsertionPoint(item.ID, CardinalPoint, Mirror2, StiffTransform, ref Offset1, ref Offset2, CSys);

                    //outputstations
                    mysap.FrameObj.GetOutputStations(item.ID, ref myType, ref MaxSize, ref MinSect, ref NoOut, ref NoutAndDesign);
                    myetab.FrameObj.SetOutputStations(item.ID, myType, MaxSize, MinSect, NoOut, NoutAndDesign);


                    // dist loads
                    mysap.FrameObj.GetLoadDistributed(item.ID, ref NumberItems, ref FrameName, ref LoadPat, ref MyType,
                        ref CSys2, ref Dir, ref RD1, ref RD2, ref Dist1, ref Dist2, ref Val1, ref Val2);

                    for (int i = 0; i < NumberItems; i++)
                    {
                        if (reversedFrames.Contains(item))
                        {

                            if (Dir[i] == 2 || Dir[i] == 1)
                            {
                                Val2[i] = Val2[i] * -1;
                                Val1[i] = Val1[i] * -1;

                            }
                            myetab.FrameObj.SetLoadDistributed(item.ID, LoadPat[i], MyType[i], Dir[i], 1 - RD1[i], 1 - RD2[i], Val2[i], Val1[i], CSys2[i], true, false);

                        }
                        else
                        {
                            myetab.FrameObj.SetLoadDistributed(item.ID, LoadPat[i], MyType[i], Dir[i], RD1[i], RD2[i], Val1[i], Val2[i], CSys2[i], true, false);

                        }

                    }


                    // points loads
                    mysap.FrameObj.GetLoadPoint(item.ID, ref NumberItems, ref FrameName, ref LoadPat, ref MyType, ref CSys2, ref Dir, ref RD1, ref Dist1, ref Val1);
                    for (int i = 0; i < NumberItems; i++)
                    {
                        if (reversedFrames.Contains(item))
                        {
                            if (Dir[i] == 2 || Dir[i] == 1)
                            {
                                Val1[i] = Val1[i] * -1;

                            }
                            myetab.FrameObj.SetLoadPoint(item.ID, LoadPat[i], MyType[i], Dir[i], 1 - RD1[i], Val1[i], CSys2[i], true, false);

                        }
                        else
                        {
                            myetab.FrameObj.SetLoadPoint(item.ID, LoadPat[i], MyType[i], Dir[i], RD1[i], Val1[i], CSys2[i], true, false);

                        }

                    }


                    // temp loads
                    mysap.FrameObj.GetLoadTemperature(item.ID, ref NumberItems, ref FrameName, ref LoadPat, ref MyType, ref Val1, ref patternName);

                    for (int i = 0; i < NumberItems; i++)
                    {
                        myetab.FrameObj.SetLoadTemperature(item.ID, LoadPat[i], MyType[i], Val1[i], patternName[i], false);

                    }

                    if (reversedFrames.Contains(item))
                    {
                        myetab.FrameObj.SetLocalAxes(item.ID, item.Beta + 180);

                    }

                    else
                    {
                        /// local axes
                        myetab.FrameObj.SetLocalAxes(item.ID, item.Beta);

                    }


                    // releases
                    mysap.FrameObj.GetReleases(item.ID, ref II, ref JJ, ref StartValue, ref EndValue);
                    if (reversedFrames.Contains(item))
                    {
                        myetab.FrameObj.SetReleases(item.ID, ref JJ, ref II, ref EndValue, ref StartValue);

                    }
                    else
                    {
                        myetab.FrameObj.SetReleases(item.ID, ref II, ref JJ, ref StartValue, ref EndValue);

                    }


                    if (reversedFrames.Contains(item))
                    {
                        string sec = item.Sec;
                        mysap.PropFrame.GetTypeOAPI(sec, ref e);
                        if (e == eFramePropType.Variable)
                        {

                            mysap.PropFrame.GetNonPrismatic(sec, ref NumberItems, ref StartSec, ref EndSec, ref MYLength, ref Mytype2, ref Ei33, ref Ei22, ref color, ref Notes, ref guid);
                            myetab.PropFrame.SetNonPrismatic(sec + "Rev", NumberItems, ref EndSec, ref StartSec, ref MYLength, ref Mytype2, ref Ei33, ref Ei22, color, Notes);

                            myetab.FrameObj.SetSection(item.ID, sec + "Rev");
                        }
                    }

                }




            }


            public static void SapToEtabsFramesAssignments2(List<Frame> sapframes, cSapModel mysap, ETABSv1.cSapModel myetab, List<Frame> reversedFrames)
            {
                #region def
                int numberGroups = 0;
                string[] Groups = new string[0];
                int CardinalPoint = 0;
                bool Mirror2 = false;
                bool StiffTransform = false;
                double[] Offset1 = new double[0];
                double[] Offset2 = new double[0];
                string CSys = null;

                int myType = 0;
                double MaxSize = 0;
                int MinSect = 0;
                bool NoOut = false;
                bool NoutAndDesign = false;
                int NumberItems = 0;
                string[] FrameName = new string[0];
                string[] LoadPat = new string[0];
                string[] CSys2 = null;

                int[] MyType = new int[0];
                int[] Dir = new int[0];
                double[] RD1 = new double[0];
                double[] RD2 = new double[0];
                double[] Dist1 = new double[0];
                double[] Dist2 = new double[0];
                double[] Val1 = new double[0];
                double[] Val2 = new double[0];
                string[] patternName = null;


                bool[] II = null;
                bool[] JJ = null;
                double[] StartValue = null;
                double[] EndValue = null;

                string[] PointName = new string[0];
                int[] LCStep = new int[0];
                string[] CSys3 = new string[0];
                double[] F1 = new double[0];
                double[] F2 = new double[0];
                double[] F3 = new double[0];
                double[] M1 = new double[0];
                double[] M2 = new double[0];
                double[] M3 = new double[0];
                double[] Value = new double[0];

                bool[] restValues = new bool[0];

                eFramePropType e = eFramePropType.Circle;
                string[] StartSec = null;
                string[] EndSec = null;
                double[] MYLength = null;
                int[] Ei33 = null;
                int[] Ei22 = null;
                int color = 0;
                string Notes = null;
                string guid = null;
                int[] Mytype2 = null;





                string[] elements = new string[0];
                double[] relativeDistancesI = new double[0];
                double[] relativeDistancesJ = new double[0];
                int NElm = 0;
                #endregion

                List<string> ptIds = sapframes.Select(x => x.StartPoint.Id).ToList();
                ptIds.AddRange(sapframes.Select(x => x.EndPoint.Id).ToList());

                ptIds = ptIds.Distinct().ToList();

                foreach (var item in sapframes)
                {
                    // group assignments
                    mysap.FrameObj.GetGroupAssign(item.ID, ref numberGroups, ref Groups);
                    foreach (var g in Groups)
                    {
                        myetab.FrameObj.SetGroupAssign(item.ID, g, false);
                    }

                    // insertionpoints
                    mysap.FrameObj.GetInsertionPoint(item.ID, ref CardinalPoint, ref Mirror2, ref StiffTransform, ref Offset1, ref Offset2, ref CSys);
                    myetab.FrameObj.SetInsertionPoint(item.ID, CardinalPoint, Mirror2, StiffTransform, ref Offset1, ref Offset2, CSys);

                    //outputstations
                    mysap.FrameObj.GetOutputStations(item.ID, ref myType, ref MaxSize, ref MinSect, ref NoOut, ref NoutAndDesign);
                    myetab.FrameObj.SetOutputStations(item.ID, myType, MaxSize, MinSect, NoOut, NoutAndDesign);


                    // dist loads
                    #region aa

                    int numberOfElements = mysap.FrameObj.GetElm(item.ID, ref NElm, ref elements, ref relativeDistancesI, ref relativeDistancesJ);

                    for (int i = 0; i < NElm; i++)
                    {

                        string elementName = elements[i];
                        double relativeDistanceI = relativeDistancesI[i];
                        double relativeDistanceJ = relativeDistancesJ[i];

                        // Retrieve the distributed loads for the specific load case
                        string[] lineNames = null;
                        string[] loadPatterns = null;
                        int[] types = null;
                        string[] coordinateSystems = null;
                        int[] directions = null;
                        double[] relativeDistancesStart = null;
                        double[] relativeDistancesEnd = null;
                        double[] distancesStart = null;
                        double[] distancesEnd = null;
                        double[] valuesStart = null;
                        double[] valuesEnd = null;
                        int numberofLoads = 0;



                        if (item.ID == "D2_178")
                        {
                            int ab = 0;
                        }


                        mysap.LineElm.GetLoadDistributed(elementName, ref numberofLoads, ref lineNames, ref loadPatterns, ref types, ref coordinateSystems,
                                             ref directions, ref relativeDistancesStart, ref relativeDistancesEnd, ref distancesStart, ref distancesEnd, ref valuesStart, ref valuesEnd);


                        for (int j = 0; j < numberofLoads; j++)
                        {
                            if (reversedFrames.Select(x => x.ID).ToList().Contains(item.ID))
                            {
                                string lineName = lineNames[j];
                                string loadPattern = loadPatterns[j];
                                int type = types[j];
                                string coordinateSystem = coordinateSystems[j];
                                int direction = directions[j];
                                double relativeDistanceStart = relativeDistancesStart[j];
                                double relativeDistanceEnd = relativeDistancesEnd[j];

                                // Calculate absolute distances based on relative distances
                                double absoluteDistanceStart = relativeDistanceI + relativeDistanceStart / NElm;
                                double absoluteDistanceEnd = absoluteDistanceStart + (relativeDistanceEnd - relativeDistanceStart) / NElm;


                                //   if (Math.Abs (absoluteDistanceEnd- absoluteDistanceStart)>0.05)
                                //  {
                                int ret = myetab.FrameObj.SetLoadDistributed(item.ID, loadPattern, type, direction, 1 - absoluteDistanceStart, 1 - absoluteDistanceEnd, valuesEnd[j], valuesStart[j],
                                  coordinateSystem, true, false);
                                //   }
                                //else
                                //{
                                //    int ret = myetab.FrameObj.SetLoadDistributed(item.ID, loadPattern, type, direction, 1 - absoluteDistanceStart, 1 - absoluteDistanceEnd, 0.00, 0.00,
                                //     coordinateSystem, true, false);
                                //}

                            }

                            else
                            {
                                string lineName = lineNames[j];
                                string loadPattern = loadPatterns[j];
                                int type = types[j];
                                string coordinateSystem = coordinateSystems[j];
                                int direction = directions[j];
                                double relativeDistanceStart = relativeDistancesStart[j];
                                double relativeDistanceEnd = relativeDistancesEnd[j];

                                // Calculate absolute distances based on relative distances

                                double absoluteDistanceStart = relativeDistanceI + relativeDistanceStart / NElm;
                                double absoluteDistanceEnd = absoluteDistanceStart + (relativeDistanceEnd - relativeDistanceStart) / NElm;


                                //   if (Math.Abs(absoluteDistanceEnd - absoluteDistanceStart) > 0.05)
                                //   {                         // Assign the distributed load to the object model using SetLoadDistributed method
                                int ret = myetab.FrameObj.SetLoadDistributed(item.ID, loadPattern, type, direction, absoluteDistanceStart, absoluteDistanceEnd, valuesStart[j], valuesEnd[j],
                                coordinateSystem, true, false);
                                //  }

                                //else
                                //{
                                //    int ret = myetab.FrameObj.SetLoadDistributed(item.ID, loadPattern, type, direction, absoluteDistanceStart, absoluteDistanceEnd,0.00, 0.00,
                                //   coordinateSystem, true, false);
                                //}
                            }
                        }




                    }




                    // points loads
                    mysap.FrameObj.GetLoadPoint(item.ID, ref NumberItems, ref FrameName, ref LoadPat, ref MyType, ref CSys2, ref Dir, ref RD1, ref Dist1, ref Val1);
                    for (int i = 0; i < NumberItems; i++)
                    {
                        if (reversedFrames.Contains(item))
                        {
                            myetab.FrameObj.SetLoadPoint(item.ID, LoadPat[i], MyType[i], Dir[i], 1 - RD1[i], Val1[i], CSys2[i], true, false);

                        }
                        else
                        {
                            myetab.FrameObj.SetLoadPoint(item.ID, LoadPat[i], MyType[i], Dir[i], RD1[i], Val1[i], CSys2[i], true, false);

                        }

                    }


                    // temp loads
                    mysap.FrameObj.GetLoadTemperature(item.ID, ref NumberItems, ref FrameName, ref LoadPat, ref MyType, ref Val1, ref patternName);

                    for (int i = 0; i < NumberItems; i++)
                    {
                        myetab.FrameObj.SetLoadTemperature(item.ID, LoadPat[i], MyType[i], Val1[i], patternName[i], false);

                    }
                    #endregion





                    /// local axes
                    myetab.FrameObj.SetLocalAxes(item.ID, item.Beta);


                    // releases
                    mysap.FrameObj.GetReleases(item.ID, ref II, ref JJ, ref StartValue, ref EndValue);
                    if (reversedFrames.Contains(item))
                    {
                        myetab.FrameObj.SetReleases(item.ID, ref JJ, ref II, ref EndValue, ref StartValue);

                    }
                    else
                    {
                        myetab.FrameObj.SetReleases(item.ID, ref II, ref JJ, ref StartValue, ref EndValue);

                    }


                    if (reversedFrames.Contains(item))
                    {
                        string sec = item.Sec;
                        mysap.PropFrame.GetTypeOAPI(sec, ref e);
                        if (e == eFramePropType.Variable)
                        {

                            mysap.PropFrame.GetNonPrismatic(sec, ref NumberItems, ref StartSec, ref EndSec, ref MYLength, ref Mytype2, ref Ei33, ref Ei22, ref color, ref Notes, ref guid);
                            myetab.PropFrame.SetNonPrismatic(sec + "Rev", NumberItems, ref EndSec, ref StartSec, ref MYLength, ref Mytype2, ref Ei33, ref Ei22, color, Notes);

                            myetab.FrameObj.SetSection(item.ID, sec + "Rev");
                        }
                    }

                }

                foreach (var pt in ptIds)
                {
                    #region Pointload
                    mysap.PointObj.GetLoadForce(pt, ref NumberItems, ref PointName, ref LoadPat, ref LCStep, ref CSys3, ref F1, ref F2, ref F3, ref M1, ref M2, ref M3);

                    for (int i = 0; i < NumberItems; i++)
                    {
                        double[] value = new double[] { F1[0], F2[0], F3[0], M1[0], M2[0], M3[0] };

                        myetab.PointObj.SetLoadForce(pt, LoadPat[i], ref value, false, CSys3[0]);
                    }
                    #endregion
                    #region restains

                    mysap.PointObj.GetRestraint(pt, ref restValues);
                    myetab.PointObj.SetRestraint(pt, ref restValues);
                    #endregion
                }


            }

            #endregion


            #region getFrames

            public static List<Frame> getSapFramesE()
            {

                cSapModel mySapModel = Get_sap_mode();
                List<string> frame_ids = getSapFrameIDS();

                if (getSelectedFramesIDS().Count > 0)
                {
                    frame_ids = getSelectedFramesIDS();
                }


                List<Frame> sapFrames = new List<Frame>();

                #region defintions 
                string st_point = null;
                string end_point = null;
                int ret;
                double x1, y1, z1, x2, y2, z2;
                x1 = y1 = z1 = x2 = y2 = z2 = 0;
                string sAuto = null;

                eFramePropType shape = eFramePropType.Angle;


                double rotation = 0;
                bool isAdvancedAxes = false;
                string materialProperty = null;
                string sectionName = null;
                string NameInFile = null;
                string FileName = null;
                int numberOfGroup = 0;
                string[] assignedGroups = null;
                #endregion


                foreach (string id in frame_ids)
                {
                    Frame frame = new Frame();


                    #region getPoints and length
                    ret = mySapModel.FrameObj.GetPoints(id, ref st_point, ref end_point);
                    ret = mySapModel.PointObj.GetCoordCartesian(st_point, ref x1, ref y1, ref z1);
                    ret = mySapModel.PointObj.GetCoordCartesian(end_point, ref x2, ref y2, ref z2);

                    frame.ID = id;
                    frame.StartPoint = new Point(st_point, x1, y1, z1);
                    frame.EndPoint = new Point(end_point, x2, y2, z2);
                    frame.MidPoint = Point.getMidPoint(frame.StartPoint, frame.EndPoint);
                    frame.Length = System.Numerics.Vector3.Distance(Point.ToVector(frame.StartPoint), Point.ToVector(frame.EndPoint));
                    #endregion


                    #region get material and group and rotation angle and sec

                    mySapModel.FrameObj.GetGroupAssign(id, ref numberOfGroup, ref assignedGroups);
                    frame.Group = assignedGroups[0];
                    frame.Groups = assignedGroups.ToList();


                    mySapModel.FrameObj.GetSection(id, ref sectionName, ref sAuto);
                    frame.Sec = sectionName;

                    mySapModel.PropFrame.GetMaterial(sectionName, ref materialProperty);
                    frame.Material = materialProperty;

                    mySapModel.FrameObj.GetLocalAxes(id, ref rotation, ref isAdvancedAxes);

                    #endregion

                    #region  get Adanced Rotation 

                    int[] mypDir = new int[1];
                    string[] myplpt = new string[1];
                    double[] myplvector = new double[2];


                    string[] myPoints = new string[2];
                    double[] MyPlVect = new double[3];
                    int[] MyPlDir = new int[2];
                    string st_id = "";
                    string end_id = "";
                    Boolean aa = true;
                    bool Advanced = true;
                    int plane = 0;
                    int plVect = 0;
                    string cys = null;

                    double angleInDegrees = 0;

                    angleInDegrees = rotation;

                    if (isAdvancedAxes)
                    {


                        ret = mySapModel.FrameObj.GetLocalAxesAdvanced(id, ref Advanced, ref plane, ref plVect, ref cys, ref MyPlDir, ref myPoints, ref MyPlVect);

                        st_id = myPoints[0];
                        end_id = myPoints[1];

                        mySapModel.PointObj.GetCoordCartesian(st_id, ref x1, ref y1, ref z1);
                        mySapModel.PointObj.GetCoordCartesian(end_id, ref x2, ref y2, ref z2);

                        Vector3 v1 = new Vector3((float)(x1 - x2), (float)(y1 - y2), 0);


                        angleInDegrees = Math.Atan2(y1 - y2, x1 - x2) * 180 / Math.PI + rotation;


                    }

                    frame.Beta = angleInDegrees;


                    #endregion

                    Point.getMemberType(frame);

                    sapFrames.Add(frame);
                }


                return sapFrames;
            }

            #endregion

            #region Areas assignments

            // getsapArea()

            public static void defineAreaSections(cSapModel mysap, ETABSv1.cSapModel myetab, List<Area> sapAreas)
            {
                int proptype = 0;
                int ShellLayertp = 0;
                string materialProp = null;
                double matAng = 0;
                double thick = 0;
                double Bending = 0;
                int color = 0;
                string notes = null;
                string guid = null;


                int ret = 0;
                int MyType = 0;
                string MatProp = string.Empty;
                double MatAng = 0.0;
                double Thickness = 0.0;
                bool Incompatible = false;
                int Color = 0;
                string Notes = string.Empty;
                string GUID = string.Empty;
                int deckType = 1;



                ETABSv1.eWallPropType WallPropType;
                ETABSv1.eShellType ShellType;


                double[] Value = new double[10];


                // None,Wall,Floor
                List<string> wallsSections = sapAreas.Where(x => x.type == "Wall").Select(x => x.Sec).Distinct().ToList();
                List<string> FloorSections = sapAreas.Where(x => x.type == "Floor").Select(x => x.Sec).Distinct().ToList();
                List<string> Nonsections = sapAreas.Where(x => x.type == "None").Select(x => x.Sec).Distinct().ToList();

                foreach (var item in wallsSections)
                {
                    /// 0 all  1 shell 2 plane 3 solid

                    mysap.PropArea.GetTypeOAPI(item, ref proptype);
                    if (proptype == 1)
                    {

                        ret = mysap.PropArea.GetShell(item, ref ShellLayertp, ref materialProp, ref matAng, ref thick, ref Bending, ref color, ref notes, ref guid);

                        ShellType = (ETABSv1.eShellType)ShellLayertp;

                        ret = myetab.PropArea.SetWall(item, ETABSv1.eWallPropType.Specified, ShellType, materialProp, thick, color, notes);


                        ret = mysap.PropArea.GetModifiers(item, ref Value);
                        ret = myetab.PropArea.SetModifiers(item, ref Value);
                    }

                    else if (proptype == 2)
                    {
                        ret = mysap.PropArea.GetPlane(item, ref MyType, ref MatProp, ref MatAng, ref Thickness, ref Incompatible, ref Color, ref Notes, ref GUID);

                        ShellType = (ETABSv1.eShellType.ShellThin);
                        ret = myetab.PropArea.SetWall(item, ETABSv1.eWallPropType.Specified, ShellType, MatProp, Thickness, color, notes);


                        ret = mysap.PropArea.GetModifiers(item, ref Value);
                        ret = myetab.PropArea.SetModifiers(item, ref Value);
                    }

                }

                foreach (var item in FloorSections)
                {
                    /// 0 all  1 shell 2 plane 3 solid

                    mysap.PropArea.GetTypeOAPI(item, ref proptype);
                    if (proptype == 1)
                    {

                        ret = mysap.PropArea.GetShell(item, ref ShellLayertp, ref materialProp, ref matAng, ref thick, ref Bending, ref color, ref notes, ref guid);
                        ShellType = (ETABSv1.eShellType)ShellLayertp;

                        ret = myetab.PropArea.SetDeck(item, ETABSv1.eDeckType.Filled, ShellType, materialProp, thick);
                        ret = mysap.PropArea.GetModifiers(item, ref Value);
                        ret = myetab.PropArea.SetModifiers(item, ref Value);
                    }

                    else if (proptype == 2)
                    {
                        ret = mysap.PropArea.GetPlane(item, ref MyType, ref MatProp, ref MatAng, ref Thickness, ref Incompatible, ref Color, ref Notes, ref GUID);

                        ShellType = (ETABSv1.eShellType.ShellThin);
                        ret = myetab.PropArea.SetDeck(item, ETABSv1.eDeckType.Filled, ShellType, MatProp, Thickness);


                        ret = mysap.PropArea.GetModifiers(item, ref Value);
                        ret = myetab.PropArea.SetModifiers(item, ref Value);
                    }

                }

                foreach (var item in Nonsections)
                {
                    /// 0 all  1 shell 2 plane 3 solid
                    int Numberitems = 0;
                    string[] MyName = null;
                    myetab.PropMaterial.GetNameList(ref Numberitems, ref MyName, ETABSv1.eMatType.Concrete);
                    ShellType = (ETABSv1.eShellType.ShellThin);
                    ret = myetab.PropArea.SetDeck("Deck" + item, ETABSv1.eDeckType.Filled, ShellType, MyName[0], 100);

                    Value[0] = 1;
                    Value[1] = 1;
                    Value[2] = 1;
                    Value[3] = 1;
                    Value[4] = 1;
                    Value[5] = 1;
                    Value[6] = 1;
                    Value[7] = 1;
                    Value[8] = 0.001;
                    Value[9] = 0.001;


                    ret = myetab.PropArea.SetModifiers(item, ref Value);


                }


            }



            public static void AreasAssignments(cSapModel mysap, ETABSv1.cSapModel myetab, List<Area> sapAreas)
            {
                int ret;
                int NumberItems = 0;
                string[] AreaName = null;
                string[] LoadPat = null;
                string[] CSys = null;
                int[] Dir = null;
                double[] Value = null;
                int[] DistType = null;
                int[] MyType = null;
                string[] PatternName = default;
                foreach (var item in sapAreas)
                {
                    if (item.type == "None")
                    {
                        myetab.AreaObj.SetProperty(item.ID, "Deck" + item.Sec);
                    }
                    else
                    {
                        myetab.AreaObj.SetProperty(item.ID, item.Sec);
                    }


                    ret = mysap.AreaObj.GetLoadUniformToFrame(item.ID, ref NumberItems, ref AreaName, ref LoadPat, ref CSys, ref Dir, ref Value, ref DistType);

                    for (int i = 0; i < NumberItems; i++)
                    {
                        myetab.AreaObj.SetLoadUniform(item.ID, LoadPat[i], Value[i], Dir[i], false, CSys[i]);
                    }

                    ret = mysap.AreaObj.GetLoadUniform(item.ID, ref NumberItems, ref AreaName, ref LoadPat, ref CSys, ref Dir, ref Value);

                    for (int i = 0; i < NumberItems; i++)
                    {
                        myetab.AreaObj.SetLoadUniform(item.ID, LoadPat[i], Value[i], Dir[i], false, CSys[i]);
                    }



                    ret = mysap.AreaObj.GetLoadTemperature(item.ID, ref NumberItems, ref AreaName, ref LoadPat, ref MyType, ref Value, ref PatternName);
                    for (int i = 0; i < NumberItems; i++)
                    {
                        myetab.AreaObj.SetLoadTemperature(item.ID, LoadPat[i], MyType[i], Value[i], PatternName[i], false);
                    }

                    foreach (var g in item.Groups)
                    {
                        myetab.AreaObj.SetGroupAssign(item.ID, g);
                    }


                    myetab.AreaObj.SetLocalAxes(item.ID, item.Beta);
                }



            }
            #endregion

            #region  createPoints

            public static void createPoints(List<Point> pts, ETABSv1.cSapModel myEtab)
            {
                string id = null;
                foreach (var item in pts)
                {
                    id = item.Id;

                    myEtab.PointObj.AddCartesian(item.X, item.Y, item.Z, ref id, item.Id);
                }
            }

            #endregion

            #region  drawAreas

            public static void drawAreas(ETABSv1.cSapModel myetab, List<Area> sapAreas)
            {
                int ret = 1;
                foreach (var item in sapAreas)
                {
                    string[] pointsIds = item.vertices.Select(x => x.Id).ToArray();
                    string id = item.ID;
                    ret = myetab.AreaObj.AddByPoint(item.vertices.Count, ref pointsIds, ref id, "", item.ID);
                }
            }


            #endregion

            #region  loadsDefinistions

            public class LoadPattern
            {

                string name = null;
                eLoadPatternType loadType = eLoadPatternType.Creep;
                double selfWeightMulti = 0;


                public static List<LoadPattern> getLoadPatternsFromSap()
                {
                    List<LoadPattern> patterns = new List<LoadPattern>();


                    int numberOfItems = 0;
                    string[] LoadPatterns = new string[0];

                    cSapModel mysapmodel = Get_sap_mode();
                    int ret = mysapmodel.LoadPatterns.GetNameList(ref numberOfItems, ref LoadPatterns);

                    foreach (var item in LoadPatterns)
                    {
                        LoadPattern p = new LoadPattern();

                        eLoadPatternType MYType = default;
                        mysapmodel.LoadPatterns.GetLoadType(item, ref MYType);
                        double selfWe = 0;
                        mysapmodel.LoadPatterns.GetSelfWTMultiplier(item, ref selfWe);

                        p.name = item;
                        p.loadType = MYType;
                        p.selfWeightMulti = selfWe;
                        patterns.Add(p);

                    }

                    return patterns;
                }


                public static void SetLoadPatternsInEtabs(List<LoadPattern> sapLoadPatterns)
                {

                    foreach (var item in sapLoadPatterns)
                    {
                        ETABSv1.cSapModel Etabs = EtabsUtils.Get_sap_mode();
                        Etabs.LoadPatterns.Add(item.name, loadpattenConvertor(item.loadType), item.selfWeightMulti, false);


                    }


                }



            }

            public class LoadCase
            {
                public string name = null;
                public string[] loadName = new string[0];
                public string[] loadType = new string[0];
                public double[] SF = new double[0];


                public static List<LoadCase> GetLinearLoadcases()
                {

                    List<LoadCase> Lcases = new List<LoadCase>();

                    cSapModel mysap = Get_sap_mode();

                    int numberOfItems = 0;
                    string[] NamesOfCaes = new string[0];
                    mysap.LoadCases.GetNameList(ref numberOfItems, ref NamesOfCaes, eLoadCaseType.LinearStatic);

                    foreach (var item in NamesOfCaes)
                    {
                        LoadCase l = new LoadCase();
                        l.name = item;
                        mysap.LoadCases.StaticLinear.GetLoads(l.name, ref numberOfItems, ref l.loadType, ref l.loadName, ref l.SF);

                        Lcases.Add(l);


                    }



                    return Lcases;
                }

                public static List<LoadCase> GetNonLinearLoadcases()
                {

                    List<LoadCase> Lcases = new List<LoadCase>();

                    cSapModel mysap = Get_sap_mode();

                    int numberOfItems = 0;
                    string[] NamesOfCaes = new string[0];
                    mysap.LoadCases.GetNameList(ref numberOfItems, ref NamesOfCaes, eLoadCaseType.NonlinearStatic);

                    foreach (var item in NamesOfCaes)
                    {
                        LoadCase l = new LoadCase();
                        l.name = item;
                        mysap.LoadCases.StaticNonlinear.GetLoads(l.name, ref numberOfItems, ref l.loadType, ref l.loadName, ref l.SF);

                        Lcases.Add(l);


                    }



                    return Lcases;
                }




                public static void SetLinearLoadCasesInEtabs(List<LoadCase> Lcases)
                {
                    ETABSv1.cSapModel etabs = EtabsUtils.Get_sap_mode();


                    foreach (var LC in Lcases)
                    {
                        etabs.LoadCases.StaticLinear.SetCase(LC.name);
                        etabs.LoadCases.StaticLinear.SetLoads(LC.name, LC.loadName.Length, ref LC.loadType, ref LC.loadName, ref LC.SF);

                    }




                }

                public static void SetNonLinearLoadCasesInEtabs(List<LoadCase> Lcases)
                {
                    ETABSv1.cSapModel etabs = EtabsUtils.Get_sap_mode();


                    foreach (var LC in Lcases)
                    {
                        etabs.LoadCases.StaticNonlinear.SetCase(LC.name);
                        etabs.LoadCases.StaticNonlinear.SetLoads(LC.name, LC.loadName.Length, ref LC.loadType, ref LC.loadName, ref LC.SF);

                    }




                }




            }

            public class LOadComb
            {
                public string Name;
                public eCNameType[] CType = null;
                public string[] CName = null;
                public double[] Sf = null;

                public int combType;

                public static List<LOadComb> getLoadCombsFromSap()
                {
                    List<LOadComb> loadCombs = new List<LOadComb>();
                    cSapModel mysap = Get_sap_mode();

                    int numberNames = 0;
                    string[] LoadCombsNames = new string[0];
                    mysap.RespCombo.GetNameList(ref numberNames, ref LoadCombsNames);

                    foreach (var item in LoadCombsNames)
                    {
                        LOadComb L = new LOadComb();
                        L.Name = item;

                        mysap.RespCombo.GetCaseList(L.Name, ref numberNames, ref L.CType, ref L.CName, ref L.Sf);

                        mysap.RespCombo.GetTypeOAPI(L.Name, ref L.combType);


                        loadCombs.Add(L);
                    }

                    return loadCombs;
                }

                public static void SetLoadCombsInEtabs(List<LOadComb> loadCombsSap)
                {
                    ETABSv1.cSapModel etabs = EtabsUtils.Get_sap_mode();


                    foreach (var LC in loadCombsSap)
                    {
                        etabs.RespCombo.Add(LC.Name, LC.combType);

                        for (int i = 0; i < LC.CName.Length; i++)
                        {
                            ETABSv1.eCNameType e;
                            if (LC.CType[i] == eCNameType.LoadCase)
                            {
                                e = ETABSv1.eCNameType.LoadCase;
                            }
                            else
                            {
                                e = ETABSv1.eCNameType.LoadCombo;
                            }

                            etabs.RespCombo.SetCaseList(LC.Name, ref e, LC.CName[i], LC.Sf[i]);
                        }




                    }

                }



            }

            public static ETABSv1.eLoadPatternType loadpattenConvertor(SAP2000v1.eLoadPatternType loadPatType)
            {

                switch (loadPatType)
                {
                    case eLoadPatternType.Dead:
                        return e.Dead;

                    case eLoadPatternType.SuperDead:
                        return e.SuperDead;

                    case eLoadPatternType.Live:
                        return e.Live;

                    case eLoadPatternType.ReduceLive:
                        return e.ReduceLive;
                    case eLoadPatternType.Quake:
                        return e.Quake;
                    case eLoadPatternType.Wind:
                        return e.Wind;
                    case eLoadPatternType.Snow:
                        return e.Snow;
                    case eLoadPatternType.Other:
                        return e.Other;
                    case eLoadPatternType.Move:
                        return e.Move;
                    case eLoadPatternType.Temperature:
                        return e.Temperature;
                    case eLoadPatternType.Rooflive:
                        return e.Rooflive;
                    case eLoadPatternType.Notional:
                        return e.Notional;
                    case eLoadPatternType.PatternLive:
                        return e.PatternLive;
                    case eLoadPatternType.Wave:
                        return e.Wave;
                    case eLoadPatternType.Braking:
                        return e.Braking;
                    case eLoadPatternType.Centrifugal:
                        return e.Centrifugal;
                    case eLoadPatternType.Friction:
                        return e.Friction;
                    case eLoadPatternType.Ice:
                        return e.Ice;
                    case eLoadPatternType.WindOnLiveLoad:
                        return e.WindOnLiveLoad;
                    case eLoadPatternType.HorizontalEarthPressure:
                        return e.HorizontalEarthPressure;
                    case eLoadPatternType.VerticalEarthPressure:
                        return e.VerticalEarthPressure;
                    case eLoadPatternType.EarthSurcharge:
                        return e.EarthSurcharge;
                    case eLoadPatternType.DownDrag:
                        return e.DownDrag;
                    case eLoadPatternType.VehicleCollision:
                        return e.VehicleCollision;
                    case eLoadPatternType.VesselCollision:
                        return e.VesselCollision;
                    case eLoadPatternType.TemperatureGradient:
                        return e.TemperatureGradient;
                    case eLoadPatternType.Settlement:
                        return e.Settlement;
                    case eLoadPatternType.Shrinkage:
                        return e.Shrinkage;
                    case eLoadPatternType.Creep:
                        return e.Creep;
                    case eLoadPatternType.WaterloadPressure:
                        return e.WaterloadPressure;
                    case eLoadPatternType.LiveLoadSurcharge:
                        return e.LiveLoadSurcharge;
                    case eLoadPatternType.LockedInForces:
                        return e.LockedInForces;
                    case eLoadPatternType.PedestrianLL:
                        return e.PedestrianLL;
                    case eLoadPatternType.Prestress:
                        return e.Prestress;
                    case eLoadPatternType.Hyperstatic:
                        return e.Hyperstatic;
                    case eLoadPatternType.Bouyancy:
                        return e.Bouyancy;
                    case eLoadPatternType.StreamFlow:
                        return e.StreamFlow;
                    case eLoadPatternType.Impact:
                        return e.Impact;
                    case eLoadPatternType.Construction:
                        return e.Construction;
                    case eLoadPatternType.DeadWearing:
                        return e.DeadWearing;
                    case eLoadPatternType.DeadWater:
                        return e.DeadWater;
                    case eLoadPatternType.DeadManufacture:
                        return e.DeadManufacture;
                    case eLoadPatternType.EarthHydrostatic:
                        return e.EarthHydrostatic;
                    case eLoadPatternType.PassiveEarthPressure:
                        return e.PassiveEarthPressure;
                    case eLoadPatternType.ActiveEarthPressure:
                        return e.ActiveEarthPressure;
                    case eLoadPatternType.PedestrianLLReduced:
                        return e.PedestrianLLReduced;
                    case eLoadPatternType.SnowHighAltitude:
                        return e.SnowHighAltitude;
                    case eLoadPatternType.EuroLm1Char:
                        return e.EuroLm1Char;
                    case eLoadPatternType.EuroLm1Freq:
                        return e.EuroLm1Freq;
                    case eLoadPatternType.EuroLm2:
                        return e.EuroLm2;
                    case eLoadPatternType.EuroLm3:
                        return e.EuroLm3;
                    case eLoadPatternType.EuroLm4:
                        return e.EuroLm4;
                    case eLoadPatternType.SeaState:
                        return e.SeaState;
                    case eLoadPatternType.Permit:
                        return e.Permit;
                    case eLoadPatternType.MoveFatigue:
                        return e.MoveFatigue;
                    case eLoadPatternType.MoveFatiguePermit:
                        return e.MoveFatiguePermit;
                    case eLoadPatternType.MoveDeflection:
                        return e.MoveDeflection;
                    default:
                        return e.Other;
                }



            }


            #endregion

            #region  create Stories

            public static void CreateStories(cSapModel mySap, ETABSv1.cSapModel myEtabs, string group)
            {

                double baseElevation = 0.0;


                double X = 0;
                double Y = 0;
                double Z = 0;

                int NumberItems = 0;
                int[] objType = null;
                string[] ObjectName = null;
                mySap.GroupDef.GetAssignments(group, ref NumberItems, ref objType, ref ObjectName);
                List<string> pointsOfLevels = new List<string>();
                List<Point> pts = new List<Point>();

                for (int i = 0; i < NumberItems; i++)
                {
                    if (objType[i] == 1)
                    {
                        pointsOfLevels.Add(ObjectName[i]);
                    }
                }

                foreach (var pt in pointsOfLevels)
                {
                    Point pp = new Point();
                    mySap.PointObj.GetCoordCartesian(pt, ref X, ref Y, ref Z);
                    pp.Id = pt;
                    pp.X = X;
                    pp.Y = Y;
                    pp.Z = Z;
                    pts.Add(pp);
                }

                pts = pts.GroupBy(x => x.Z).Select(x => x.FirstOrDefault()).ToList();
                pts.OrderBy(x => x.Z);
                Point minZPt = pts.FirstOrDefault();
                if (minZPt != null)
                {
                    baseElevation = minZPt.Z;
                }
                pts.ForEach(x => x.Z = Math.Abs(x.Z - baseElevation));

                for (int i = 1; i < pts.Count; i++)
                {
                    pts[i].Z -= pts[i - 1].Z;
                }


                int numberOfStories = pts.Count;
                string[] stroiesNames = new string[numberOfStories];
                double[] stroiesHeights = new double[numberOfStories];
                bool[] IsMasterStory = new bool[numberOfStories];
                string[] SimilarStory = new string[numberOfStories];
                bool[] IsSpliceAbove = new bool[numberOfStories];
                double[] spliesHeights = new double[numberOfStories];
                int[] colors = new int[numberOfStories];

                for (int i = 0; i < numberOfStories; i++)
                {
                    stroiesNames[i] = "Story_ " + pts[i].Z.ToString();
                    stroiesHeights[i] = pts[i].Z;
                    IsMasterStory[i] = false;
                    SimilarStory[i] = "None";
                    IsSpliceAbove[i] = false;
                    spliesHeights[i] = 0;
                    colors[i] = 0;

                }

                if (numberOfStories > 1)
                {
                    IsMasterStory[numberOfStories - 1] = true;

                }
                else if (numberOfStories == 1)
                {
                    IsMasterStory[0] = true;

                }

                int ret = myEtabs.Story.SetStories_2(baseElevation, numberOfStories, ref stroiesNames, ref stroiesHeights, ref IsMasterStory,
                    ref SimilarStory, ref IsSpliceAbove, ref spliesHeights, ref colors);


            }

            #endregion



            #region assignments Points 
            public static void SapToEtabsPointAssignmentsAndDiaphram(cSapModel mysap, ETABSv1.cSapModel myEtabs, List<Point> myPoints)
            {
                #region def

                string[] Groups = new string[0];
                double[] Offset1 = new double[0];
                double[] Offset2 = new double[0];

                int NumberItems = 0;
                string[] FrameName = new string[0];
                string[] LoadPat = new string[0];


                int[] MyType = new int[0];
                int[] Dir = new int[0];
                double[] RD1 = new double[0];
                double[] RD2 = new double[0];
                double[] Dist1 = new double[0];
                double[] Dist2 = new double[0];
                double[] Val1 = new double[0];
                double[] Val2 = new double[0];



                string[] PointName = new string[0];
                int[] LCStep = new int[0];
                string[] CSys3 = new string[0];
                double[] F1 = new double[0];
                double[] F2 = new double[0];
                double[] F3 = new double[0];
                double[] M1 = new double[0];
                double[] M2 = new double[0];
                double[] M3 = new double[0];
                double[] Value = new double[0];

                bool[] restValues = new bool[0];

                List<Point> PointsDiaphram = new List<Point>();
                string[] constrainNames = null;

                eConstraintType e = eConstraintType.Line;
                #endregion


                foreach (var pt in myPoints)
                {
                    #region Pointload
                    mysap.PointObj.GetLoadForce(pt.Id, ref NumberItems, ref PointName, ref LoadPat, ref LCStep, ref CSys3, ref F1, ref F2, ref F3, ref M1, ref M2, ref M3);

                    for (int i = 0; i < NumberItems; i++)
                    {
                        double[] value = new double[] { F1[0], F2[0], F3[0], M1[0], M2[0], M3[0] };

                        myEtabs.PointObj.SetLoadForce(pt.Id, LoadPat[i], ref value, false, CSys3[0]);
                    }
                    #endregion
                    #region restains

                    mysap.PointObj.GetRestraint(pt.Id, ref restValues);
                    myEtabs.PointObj.SetRestraint(pt.Id, ref restValues);




                    #endregion

                    #region Diaphram
                    //mysap.PointObj.GetConstraint(pt.Id, ref NumberItems, ref PointName, ref constrainNames);

                    //for (int i = 0; i < NumberItems; i++)
                    //{
                    //    mysap.ConstraintDef.GetConstraintType(constrainNames[i], ref e);

                    //    if (e == eConstraintType.Diaphragm)
                    //    {
                    //        myEtabs.PointObj.SetDiaphragm(pt.Id, ETABSv1.eDiaphragmOption.DefinedDiaphragm, constrainNames[i]);
                    //        PointsDiaphram.Add(pt);

                    //    }

                    //}
                    #endregion
                }


                #region CreationOfFloors

                //PointsDiaphram = PointsDiaphram.GroupBy(x => x.Z).Select(x => x.FirstOrDefault()).ToList();
                //Point TopPoint = myPoints.FirstOrDefault(x => x.Z == myPoints.Max(y => y.Z));
                //if (!PointsDiaphram.Any(x=>x.Z == TopPoint.Z))
                //{
                //    PointsDiaphram.Add(TopPoint);

                //}
                //List<object> Stories = new List<object>();

                //int numberOfStories = PointsDiaphram.Count;
                //string[] stroiesNames = new string[numberOfStories];
                //double[] stroiesHeights = new double[numberOfStories];
                //bool[] IsMasterStory = new bool[numberOfStories];
                //string[] SimilarStory = new string[numberOfStories];
                //bool[] IsSpliceAbove = new bool[numberOfStories];
                //double[] spliesHeights = new double[numberOfStories];
                //int[] colors = new int[numberOfStories];


                //for (int i = 0; i < numberOfStories; i++)
                //{
                //    stroiesNames[i] = "Story_ " + PointsDiaphram[i].Z.ToString();
                //    stroiesHeights[i] = PointsDiaphram[i].Z;
                //    IsMasterStory[i] = true;
                //    SimilarStory[i] = stroiesNames[0];
                //    IsSpliceAbove[i] = false;
                //    spliesHeights[i] = 0;
                //    colors[i] = 0;
                //}
                //IsMasterStory[numberOfStories-1] = true;


                //int ret =  myEtabs.Story.SetStories_2(100, numberOfStories, ref stroiesNames, ref stroiesHeights, ref IsMasterStory,
                //                       ref SimilarStory, ref IsSpliceAbove, ref spliesHeights, ref colors);




                #endregion

            }



            #endregion

            #region  methods from sh

            public static void SapToEtabsSetEdgeConstainsForAreas(ETABSv1.cSapModel myEtabs)
            {
                int ret = myEtabs.AreaObj.SetEdgeConstraint("ALL", true, ETABSv1.eItemType.Group);

            }

            public static void SatpToEtabsSetEndLengthOffsetForConcFrames(ETABSv1.cSapModel myEtabs, List<Frame> etabsFrames)
            {
                int ret = 1;
                string sec = null;

                string stId = null;
                string endId = null;

                string[] points = null;
                int numberITems = 0;
                bool[] RestValues = null;

                bool[] FixedRstValues = new bool[] { true, true, true, true, true, true };
                myEtabs.PointObj.GetNameList(ref numberITems, ref points);






                foreach (var f in etabsFrames)
                {
                    sec = f.Sec;
                    string MatProp = null;
                    myEtabs.PropFrame.GetMaterial(sec, ref MatProp);


                    ETABSv1.eMatType MatType = ETABSv1.eMatType.Aluminum;
                    int symType = 1;
                    myEtabs.PropMaterial.GetTypeOAPI(MatProp, ref MatType, ref symType);

                    if (MatType == ETABSv1.eMatType.Concrete)
                    {
                        ret = myEtabs.FrameObj.SetEndLengthOffset(f.ID, false, 1, 1, 0.5);
                        ret = myEtabs.FrameObj.SetEndLengthOffset(f.ID, true, 1, 1, 0.5);


                        myEtabs.FrameObj.GetPoints(f.ID, ref stId, ref endId);

                        myEtabs.PointObj.GetRestraint(stId, ref RestValues);
                        if (RestValues[0] == true)
                        {
                            myEtabs.PointObj.SetRestraint(stId, ref FixedRstValues);
                        }


                        myEtabs.PointObj.GetRestraint(endId, ref RestValues);
                        if (RestValues[0] == true)
                        {
                            myEtabs.PointObj.SetRestraint(endId, ref FixedRstValues);
                        }




                    }



                }



            }
            public static void SapToEtabsSetEndLengthOffsetForSteelFrames(ETABSv1.cSapModel myEtabs, List<Frame> etabsFrames)
            {
                int ret = 1;
                string sec = null;

                foreach (var f in etabsFrames)
                {
                    sec = f.Sec;
                    string MatProp = null;
                    myEtabs.PropFrame.GetMaterial(sec, ref MatProp);


                    ETABSv1.eMatType MatType = ETABSv1.eMatType.Aluminum;
                    int symType = 1;
                    myEtabs.PropMaterial.GetTypeOAPI(MatProp, ref MatType, ref symType);

                    if (MatType == ETABSv1.eMatType.Steel)
                    {
                        ret = myEtabs.FrameObj.SetEndLengthOffset(f.ID, false, 0, 0, 0);

                    }



                }

            }

            public static void SetFixedSupportsAtConcAreas(ETABSv1.cSapModel myETabs)
            {
                int num = 0;
                string[] Points = null;
                bool[] resValues = null;
                bool[] fixedRestvalues = new bool[] { true, true, true, true, true, true };
                string sec = null;
                int propType = 0;

                List<Area> areas = getSapAreas();

                foreach (var item in areas)
                {

                    // myETabs.AreaObj.GetProperty(item.ID, ref sec);
                    // myETabs.PropArea.GetTypeOAPI(sec,ref propType);


                    //ETABSv1. eWallPropType WallPropType = default;
                    // ETABSv1.eShellType ShellType = ETABSv1.eShellType.Layered;
                    // string matProp = null;
                    // double thick = 0;
                    // string noted = null;
                    // string guid = null;
                    // int color = 0;
                    // myETabs.PropArea.GetWall(sec, ref WallPropType, ref ShellType, ref matProp, ref thick, ref color, ref noted, ref guid);

                    if (item.Sec.ToUpper() == "WALL")
                    {
                        myETabs.AreaObj.GetPoints(item.ID, ref num, ref Points);
                        foreach (var pt in Points)
                        {

                            myETabs.PointObj.GetRestraint(pt, ref resValues);
                            if (resValues[0] == true)
                            {
                                myETabs.PointObj.SetRestraint(pt, ref fixedRestvalues);
                            }


                        }
                    }



                }


            }

            public static void createDiaphrams(cSapModel mySap, ETABSv1.cSapModel myEtabs)
            {

                int numberItems = 0;
                string[] ConstrainsNames = null;
                List<string> Diaphram = new List<string>();
                mySap.ConstraintDef.GetNameList(ref numberItems, ref ConstrainsNames);

                eConstraintAxis e1 = eConstraintAxis.AutoAxis;
                ETABSv1.eConstraintAxis e1Etabs = ETABSv1.eConstraintAxis.Y;


                string CYS = null;


                eConstraintType e = eConstraintType.Local;
                for (int i = 0; i < numberItems; i++)
                {
                    mySap.ConstraintDef.GetConstraintType(ConstrainsNames[i], ref e);

                    if (e == eConstraintType.Diaphragm)
                    {
                        Diaphram.Add(ConstrainsNames[i]);
                    }



                }

                foreach (var item in Diaphram)
                {
                    mySap.ConstraintDef.GetDiaphragm(item, ref e1, ref CYS);


                    switch (e1)
                    {
                        case eConstraintAxis.X:
                            e1Etabs = ETABSv1.eConstraintAxis.X;
                            break;
                        case eConstraintAxis.Y:
                            e1Etabs = ETABSv1.eConstraintAxis.Y;

                            break;
                        case eConstraintAxis.Z:
                            e1Etabs = ETABSv1.eConstraintAxis.Y;

                            break;
                        case eConstraintAxis.AutoAxis:
                            e1Etabs = ETABSv1.eConstraintAxis.AutoAxis;
                            break;
                        default:
                            e1Etabs = ETABSv1.eConstraintAxis.AutoAxis;
                            break;
                    }




                    myEtabs.Diaphragm.SetDiaphragm(item, true);

                }




            }
            #endregion

            #endregion

            #region for tables
            public static string[] filedsReturned = null;
            public static int numberOFrRectors = 1;
            public static int tableVersion = 1;

            public static string Connectevity = "Joint Coordinates";
            public static string[] ConnectevityFileds = new string[] { "Joint", "XorR", "Y", "Z", "SpecialJt" };
            public static string[] ConnectevityValues = null;


            public static string JointRestAss = "Joint Restraint Assignments";
            public static string[] JointRestAssFileds = new string[] { "Joint", "U1", "U2", "U3", "R1", "R2", "R3" };
            public static string[] JointRestAssValues = null;



            public static string JointLoads = "Joint Loads Assignments - Force";
            public static string[] JointLoadsFileds = new string[] { "Joint", "LoadPat", "F1", "F2", "F3", "M1", "M2", "M3" };
            public static string[] JointLoadsValues = null;


            public static string JointLoadGroundDisp = "Joint Loads - Ground Displacement";
            public static string[] JointLoadGroundDispFileds = new string[] { "Joint", "LoadPat", "U1", "U2", "U3", "R1", "R2", "R3" };
            public static string[] JointLoadGroundDispValues = null;
















            public static string ConnectevityBM = "Connectivity - Frame";
            public static string[] ConnectevBMityFileds = new string[] { "Frame", "JointI", "JointJ" };
            public static string[] ConnectevityBMValues = null;



            public static string FrameSectionAssnignmetns = "Frame Section Assignments";
            public static string[] FrameSectionAssnignmetnsFileds = new string[] { "Frame", "AutoSelect", "AnalSect" };
            public static string[] FrameSectionAssnignmetnsValues = null;

            public static string FrameLoadsDist = "Frame Loads - Distributed";
            public static string[] FrameLoadsDistFileds = new string[] { "Frame", "LoadPat", "Type", "Dir", "DistType", "RelDistA", "RelDistB", "AbsDistA", "AbsDistB", "FOverLA", "FOverLB", "MOverLA", "MOverLB" };
            public static string[] FrameLoadsDistValues = null;

            public static string FrameLoadTepm = "Frame Loads - Temperature";
            public static string[] FrameLoadTepmFileds = new string[] { "Frame", "LoadPat", "Type", "Temp", "TempGrad2", "TempGrad3" };
            public static string[] FrameLoadTepmValues = null;


            public static string FrameLoadPoint = "Frame Loads - Point";
            public static string[] FrameLoadPointFileds = new string[] { "Frame", "LoadPat", "Type", "Dir", "DistType", "RelDist", "AbsDist", "Force", "Moment" };
            public static string[] FrameLoadPointValues = null;


            public static string FrameLocalAxes = "Frame Local Axes Assignments 1 - Typical";
            public static string[] FrameLocalAxesFileds = new string[] { "Frame", "Angle" };
            public static string[] FrameLocalAxesValues = null;



            public static string FrameAutoMeshOPtion = "Frame Auto Mesh Assignments";
            public static string[] FrameAutoMeshOPtionFileds = new string[] { "Frame", "AutoMesh", "AtJoints", "AtFrames", "NumSegments", "MaxLength" };
            public static string[] FrameAutoMeshOPtionValues = null;



            public static string FrameOutPutStation = "Frame Output Station Assignments";
            public static string[] FrameOutPutStationFileds = new string[] { "UniqueName", "MinStations", "MaxSpacing" };
            public static string[] FrameOutPutStationValues = null;



            public static string FrameRealeses = "Frame Release Assignments 1 - General";
            public static string[] FrameRealesesFileds = new string[] { "Frame", "PI", "V2I", "V3I", "TI", "M2I", "M3I", "PJ", "V2J", "V3J", "TJ", "M2J", "M3J" };
            public static string[] FrameRealesesValues = null;



            public static string insertionPoints = "Frame Insertion Point Assignments";
            public static string[] insertionPointsFileds = new string[] { "Frame", "CardinalPt", "CoordSys", "JtOffsetXI", "JtOffsetYI", "JtOffsetZI", "JtOffsetXJ", "JtOffsetYJ", "JtOffsetZJ", "Mirror2", "Mirror3", "Transform" };
            public static string[] insertionPointsValues = null;




            public static string GroupDefinitions = "Groups 1 - Definitions";
            public static string[] GroupDefinitionsFileds = new string[] { "GroupName", "SelDesSteel", "Color", "Concrete", "Aluminum" };
            public static string[] GroupDefinitionsValues = null;


            public static string GroupAssignments = "Groups 2 - Assignments";
            public static string[] GroupAssignmentsFileds = new string[] { "GroupName", "ObjectType", "ObjectLabel" };
            public static string[] GroupAssignmentsValues = null;




            public static string MaterialProp = "Material Properties 01 - General";
            public static string[] MaterialPropFileds = new string[] { "Material", "Type", "Grade", "SymType", "Color", "Notes" };
            public static string[] MaterialPropValues = null;




            public static string MaterialMechProp = "Material Properties 02 - Basic Mechanical Properties";
            public static string[] MaterialMechPropFileds = new string[] { "Material", "UnitWeight", "UnitMass", "E1", "G12", "U12", "A1" };
            public static string[] MaterialMechPropValues = null;




            public static string MaterialPropSteel = "Material Properties 03a - Steel Data";
            public static string[] MaterialPropSteelFileds = new string[] { "Material", "Fy", "Fu", "EffFy", "EffFu", "SSCurveOpt", "SSHysType", "SHard", "SMax", "SRup", "FinalSlope" };
            public static string[] MaterialPropSteelValues = null;


            public static string MaterialPropConc = "Material Properties 03b - Concrete Data";
            public static string[] MaterialPropConcFileds = new string[] { "Material", "Fc", "LtWtConc", "LtWtFact", "SSCurveOpt", "SSHysType", "SFc", "SCap", "FinalSlope", "FAngle", "DAngle", "TimeType", "TimeE", "EFact", "TimeCreep", "CreepFact", "TimeShrink", "ShrinkFact", "CreepType", "CreepTerms" };
            public static string[] MaterialPropConcValues = null;








            public static string GeneralFramesSections = "Frame Section Properties 01 - General";
            public static string[] GeneralFramesSectionsFileds = new string[] { "SectionName", "Material", "Shape", "t3", "t2", "tf", "tw", "t2b", "tfb", "FilletRadius", "Color", "FromFile", "AMod", "A2Mod", "A3Mod", "JMod", "I2Mod", "I3Mod", "MMod", "WMod", "SectInFile", "FileName", "GUID", "Notes" };
            public static string[] GeneralFramesSectionsValues = null;



            public static string NonPresmaticSections = "Frame Section Properties 05 - Nonprismatic";
            public static string[] NonPresmaticSectionsFileds = new string[] { "SectionName", "NumSegment", "StartSect", "EndSect", "LengthType", "AbsLength", "VarLength", "EI33Var", "EI22Var", "Color" };
            public static string[] NonPresmaticSectionsValues = null;





            public static void SapToEtabsNew()
            {

                #region definisions

                bool FillImportLog = true;
                int NumberFatalErrors = 0;
                int NumberErrorMsg = 0;
                int NumWarrinings = 0;
                string importLog = "";
                int NumberInfoMsg = 0;
                int ret = 0;
                cSapModel mysap = Get_sap_mode();
                ETABSv1.cSapModel myEtabs = EtabsUtils.Get_sap_mode();
                Dictionary<string, string> dic = null;
                string filePath = @"C:\Users\Ahmed.Ramzy\Desktop\New Text Document.txt";

                #endregion

                #region materils and Groups definitions
                ret = mysap.DatabaseTables.GetTableForDisplayArray(MaterialProp, ref MaterialPropFileds, "", ref tableVersion, ref filedsReturned, ref numberOFrRectors, ref MaterialPropValues);
                ret = myEtabs.DatabaseTables.SetTableForEditingArray(EtabsUtils.DataBase.MaterialProp, ref tableVersion, ref EtabsUtils.DataBase.MaterialPropFileds, numberOFrRectors, ref MaterialPropValues);


                ret = mysap.DatabaseTables.GetTableForDisplayArray(MaterialMechProp, ref MaterialMechPropFileds, "", ref tableVersion, ref filedsReturned, ref numberOFrRectors, ref MaterialMechPropValues);
                ret = myEtabs.DatabaseTables.SetTableForEditingArray(EtabsUtils.DataBase.MaterialMechProp, ref tableVersion, ref EtabsUtils.DataBase.MaterialMechPropFileds, numberOFrRectors, ref MaterialMechPropValues);

                ret = mysap.DatabaseTables.GetTableForDisplayArray(MaterialPropConc, ref MaterialPropConcFileds, "", ref tableVersion, ref filedsReturned, ref numberOFrRectors, ref MaterialPropConcValues);
                ret = myEtabs.DatabaseTables.SetTableForEditingArray(EtabsUtils.DataBase.MaterialPropConc, ref tableVersion, ref EtabsUtils.DataBase.MaterialPropConcFileds, numberOFrRectors, ref MaterialPropConcValues);


                ret = mysap.DatabaseTables.GetTableForDisplayArray(MaterialPropSteel, ref MaterialPropSteelFileds, "", ref tableVersion, ref filedsReturned, ref numberOFrRectors, ref MaterialPropSteelValues);
                ret = myEtabs.DatabaseTables.SetTableForEditingArray(EtabsUtils.DataBase.MaterialPropSteel, ref tableVersion, ref EtabsUtils.DataBase.MaterialPropSteelFileds, numberOFrRectors, ref MaterialPropSteelValues);



                ret = mysap.DatabaseTables.GetTableForDisplayArray(GroupDefinitions, ref GroupDefinitionsFileds, "", ref tableVersion, ref filedsReturned, ref numberOFrRectors, ref GroupDefinitionsValues);
                ret = myEtabs.DatabaseTables.SetTableForEditingArray(EtabsUtils.DataBase.GroupDefinitions, ref tableVersion, ref EtabsUtils.DataBase.GroupDefinitionsFileds, numberOFrRectors, ref GroupDefinitionsValues);




                #endregion

                #region Frame Sections

                List<string> secSapDistnct = Sections.GetAllNamesAssignedToFramesInSap(mysap);

                ret = mysap.DatabaseTables.GetTableForDisplayArray(GeneralFramesSections, ref GeneralFramesSectionsFileds, "", ref tableVersion, ref filedsReturned, ref numberOFrRectors, ref GeneralFramesSectionsValues);
                List<object> data3 = ConvertTableData(filedsReturned, GeneralFramesSectionsValues);
                data3 = data3.Where(x => secSapDistnct.Contains(((dynamic)x).SectionName)).ToList();

                List<string> DataOfISection = new List<string>();
                List<string> DataOfBoxSection = new List<string>();
                List<string> DataOfPipeSection = new List<string>();


                int numberOfIsections = 0;
                int numberOfBoxsections = 0;
                int numberOfPipesections = 0;


                foreach (dynamic item in data3)
                {
                    if (item.Shape == "I/Wide Flange")
                    {
                        numberOfIsections++;
                        if (filedsReturned.Contains("FileName"))
                        {

                            DataOfISection.Add(item.SectionName);
                            DataOfISection.Add(item.Material);
                            DataOfISection.Add(item.FromFile);

                            if (item.FileName != null)
                            {
                                DataOfISection.Add(CapitalizeVariableName(item.FileName));
                            }
                            else
                            {
                                DataOfISection.Add(item.FileName);

                            }
                            DataOfISection.Add(item.SectInFile);
                            DataOfISection.Add(item.t3);
                            DataOfISection.Add(item.t2);
                            DataOfISection.Add(item.tf);
                            DataOfISection.Add(item.tw);
                            DataOfISection.Add(item.t2b);
                            DataOfISection.Add(item.tfb);
                            DataOfISection.Add(item.FilletRadius);
                            DataOfISection.Add(item.AMod);
                            DataOfISection.Add(item.A2Mod);
                            DataOfISection.Add(item.A3Mod);
                            DataOfISection.Add(item.JMod);
                            DataOfISection.Add(item.I2Mod);
                            DataOfISection.Add(item.I3Mod);
                            DataOfISection.Add(item.MMod);
                            DataOfISection.Add(item.WMod);
                            DataOfISection.Add(item.Color);
                            DataOfISection.Add(item.Notes);

                        }


                        else
                        {
                            DataOfISection.Add(item.SectionName);
                            DataOfISection.Add(item.Material);
                            DataOfISection.Add(item.FromFile);
                            DataOfISection.Add(item.t3);
                            DataOfISection.Add(item.t2);
                            DataOfISection.Add(item.tf);
                            DataOfISection.Add(item.tw);
                            DataOfISection.Add(item.t2b);
                            DataOfISection.Add(item.tfb);
                            DataOfISection.Add(item.FilletRadius);
                            DataOfISection.Add(item.AMod);
                            DataOfISection.Add(item.A2Mod);
                            DataOfISection.Add(item.A3Mod);
                            DataOfISection.Add(item.JMod);
                            DataOfISection.Add(item.I2Mod);
                            DataOfISection.Add(item.I3Mod);
                            DataOfISection.Add(item.MMod);
                            DataOfISection.Add(item.WMod);
                            DataOfISection.Add(item.Color);
                            DataOfISection.Add(item.Notes);

                        }





                    }
                    if (item.Shape == "Box/Tube")
                    {
                        numberOfBoxsections++;
                        if (filedsReturned.Contains("FileName"))
                        {

                            DataOfBoxSection.Add(item.SectionName);
                            DataOfBoxSection.Add(item.Material);
                            DataOfBoxSection.Add(item.FromFile);

                            if (item.FileName != null)
                            {
                                DataOfBoxSection.Add(CapitalizeVariableName(item.FileName));
                            }
                            else
                            {
                                DataOfBoxSection.Add(item.FileName);

                            }
                            DataOfBoxSection.Add(item.SectInFile);
                            DataOfBoxSection.Add(item.t3);
                            DataOfBoxSection.Add(item.t2);
                            DataOfBoxSection.Add(item.tf);
                            DataOfBoxSection.Add(item.tw);
                            DataOfBoxSection.Add(item.FilletRadius);
                            DataOfBoxSection.Add(item.AMod);
                            DataOfBoxSection.Add(item.A2Mod);
                            DataOfBoxSection.Add(item.A3Mod);
                            DataOfBoxSection.Add(item.JMod);
                            DataOfBoxSection.Add(item.I2Mod);
                            DataOfBoxSection.Add(item.I3Mod);
                            DataOfBoxSection.Add(item.MMod);
                            DataOfBoxSection.Add(item.WMod);
                            DataOfBoxSection.Add(item.Color);
                            DataOfBoxSection.Add(item.Notes);

                        }


                        else
                        {
                            DataOfBoxSection.Add(item.SectionName);
                            DataOfBoxSection.Add(item.Material);
                            DataOfBoxSection.Add(item.FromFile);
                            DataOfBoxSection.Add(item.t3);
                            DataOfBoxSection.Add(item.t2);
                            DataOfBoxSection.Add(item.tf);
                            DataOfBoxSection.Add(item.tw);
                            DataOfBoxSection.Add(item.t2b);
                            DataOfBoxSection.Add(item.tfb);
                            DataOfBoxSection.Add(item.FilletRadius);
                            DataOfBoxSection.Add(item.AMod);
                            DataOfBoxSection.Add(item.A2Mod);
                            DataOfBoxSection.Add(item.A3Mod);
                            DataOfBoxSection.Add(item.JMod);
                            DataOfBoxSection.Add(item.I2Mod);
                            DataOfBoxSection.Add(item.I3Mod);
                            DataOfBoxSection.Add(item.MMod);
                            DataOfBoxSection.Add(item.WMod);
                            DataOfBoxSection.Add(item.Color);
                            DataOfBoxSection.Add(item.Notes);

                        }


                    }
                    if (item.Shape == "Pipe")
                    {
                        numberOfPipesections++;
                        if (filedsReturned.Contains("FileName"))
                        {

                            DataOfPipeSection.Add(item.SectionName);
                            DataOfPipeSection.Add(item.Material);
                            DataOfPipeSection.Add(item.FromFile);

                            if (item.FileName != null)
                            {
                                DataOfPipeSection.Add(CapitalizeVariableName(item.FileName));
                            }
                            else
                            {
                                DataOfPipeSection.Add(item.FileName);

                            }
                            DataOfPipeSection.Add(item.SectInFile);
                            DataOfPipeSection.Add(item.t3);
                            DataOfPipeSection.Add(item.tw);
                            DataOfPipeSection.Add(item.AMod);
                            DataOfPipeSection.Add(item.A2Mod);
                            DataOfPipeSection.Add(item.A3Mod);
                            DataOfPipeSection.Add(item.JMod);
                            DataOfPipeSection.Add(item.I2Mod);
                            DataOfPipeSection.Add(item.I3Mod);
                            DataOfPipeSection.Add(item.MMod);
                            DataOfPipeSection.Add(item.WMod);
                            DataOfPipeSection.Add(item.Color);
                            DataOfPipeSection.Add(item.Notes);

                        }


                        else
                        {
                            DataOfPipeSection.Add(item.SectionName);
                            DataOfPipeSection.Add(item.Material);
                            DataOfPipeSection.Add(item.FromFile);
                            DataOfPipeSection.Add(item.t3);
                            DataOfPipeSection.Add(item.tw);
                            DataOfPipeSection.Add(item.AMod);
                            DataOfPipeSection.Add(item.A2Mod);
                            DataOfPipeSection.Add(item.A3Mod);
                            DataOfPipeSection.Add(item.JMod);
                            DataOfPipeSection.Add(item.I2Mod);
                            DataOfPipeSection.Add(item.I3Mod);
                            DataOfPipeSection.Add(item.MMod);
                            DataOfPipeSection.Add(item.WMod);
                            DataOfPipeSection.Add(item.Color);
                            DataOfPipeSection.Add(item.Notes);

                        }


                    }

                }

                string[] DataOfIsectionArray = DataOfISection.ToArray();
                string[] DataOfBoxsectionArray = DataOfBoxSection.ToArray();
                string[] DataOfPipesectionArray = DataOfPipeSection.ToArray();




                ret = myEtabs.DatabaseTables.SetTableForEditingArray(EtabsUtils.DataBase.FramesISections, ref tableVersion, ref EtabsUtils.DataBase.FramesISectionsFileds, numberOfIsections, ref DataOfIsectionArray);
                ret = myEtabs.DatabaseTables.SetTableForEditingArray(EtabsUtils.DataBase.FramesBoxSections, ref tableVersion, ref EtabsUtils.DataBase.FramesBoxSectionsFileds, numberOfIsections, ref DataOfBoxsectionArray);
                ret = myEtabs.DatabaseTables.SetTableForEditingArray(EtabsUtils.DataBase.FramesPipeSections, ref tableVersion, ref EtabsUtils.DataBase.FramesPipeSectionsFileds, numberOfIsections, ref DataOfPipesectionArray);




                #endregion

                #region getNonPresmatic

                ret = mysap.DatabaseTables.GetTableForDisplayArray(NonPresmaticSections, ref NonPresmaticSectionsFileds, "", ref tableVersion, ref filedsReturned, ref numberOFrRectors, ref NonPresmaticSectionsValues);
                List<object> data4 = ConvertTableData(filedsReturned, NonPresmaticSectionsValues);
                data4 = data4.Where(x => secSapDistnct.Contains(((dynamic)x).SectionName)).ToList();

                List<string> dataNonPresmatic = new List<string>();
                int numberOfNunPresmatic = 0;
                foreach (dynamic item in data4)
                {
                    //if ((int.Parse((string)item.NumSegment)) > 1)
                    //{
                    //    continue;
                    //}
                    numberOfNunPresmatic++;
                    dataNonPresmatic.Add(item.SectionName);
                    dataNonPresmatic.Add(item.StartSect);
                    dataNonPresmatic.Add(item.EndSect);


                    if (item.LengthType == "Variable")
                    {
                        dataNonPresmatic.Add("Proportional");
                    }
                    else
                    {
                        dataNonPresmatic.Add("Absolute");
                    }


                    if (item.LengthType == "Variable")
                    {
                        dataNonPresmatic.Add(item.VarLength);
                    }
                    else
                    {
                        dataNonPresmatic.Add(item.AbsLength);
                    }

                    dataNonPresmatic.Add(item.EI33Var);
                    dataNonPresmatic.Add(item.EI22Var);

                    // dataNonPresmatic.Add(item.Color);


                }

                string[] dataNonPresmaticArray = dataNonPresmatic.ToArray();
                ret = myEtabs.DatabaseTables.SetTableForEditingArray(EtabsUtils.DataBase.NonPresmaticSections, ref tableVersion, ref EtabsUtils.DataBase.NonPresmaticSectionsFileds, numberOfNunPresmatic, ref dataNonPresmaticArray);

                #endregion

                #region points

                ret = mysap.DatabaseTables.GetTableForDisplayArray(Connectevity, ref ConnectevityFileds, "", ref tableVersion, ref filedsReturned, ref numberOFrRectors, ref ConnectevityValues);
                ret = myEtabs.DatabaseTables.SetTableForEditingArray(EtabsUtils.DataBase.Connectevity, ref tableVersion, ref EtabsUtils.DataBase.ConnectevityFileds, numberOFrRectors, ref ConnectevityValues);


                ret = mysap.DatabaseTables.GetTableForDisplayArray(JointRestAss, ref JointRestAssFileds, "", ref tableVersion, ref filedsReturned, ref numberOFrRectors, ref JointRestAssValues);
                ret = myEtabs.DatabaseTables.SetTableForEditingArray(EtabsUtils.DataBase.JointRestAss, ref tableVersion, ref EtabsUtils.DataBase.JointRestAssFileds, numberOFrRectors, ref JointRestAssValues);


                #endregion

                #region beams

                ret = mysap.DatabaseTables.GetTableForDisplayArray(ConnectevityBM, ref ConnectevBMityFileds, "", ref tableVersion, ref filedsReturned, ref numberOFrRectors, ref ConnectevityBMValues);
                #region convertor
                List<Frame> myframes = getSapFramesE();
                data3 = ConvertTableData(filedsReturned, ConnectevityBMValues);

                List<string> dataOfBeams = new List<string>();
                string[] dataOfBeamsArra = null;
                foreach (Frame item in myframes)
                {
                    dataOfBeams.Add(item.ID);
                    dataOfBeams.Add(item.StartPoint.Id);
                    dataOfBeams.Add(item.EndPoint.Id);

                }
                dataOfBeamsArra = dataOfBeams.ToArray();
                #endregion
                ret = myEtabs.DatabaseTables.SetTableForEditingArray(EtabsUtils.DataBase.ConnectevityBM, ref tableVersion, ref EtabsUtils.DataBase.ConnectevBMityFileds, numberOFrRectors, ref dataOfBeamsArra);

                ret = mysap.DatabaseTables.GetTableForDisplayArray(FrameSectionAssnignmetns, ref FrameSectionAssnignmetnsFileds, "", ref tableVersion, ref filedsReturned, ref numberOFrRectors, ref FrameSectionAssnignmetnsValues);
                ret = myEtabs.DatabaseTables.SetTableForEditingArray(EtabsUtils.DataBase.FrameSectionAssnignmetns, ref tableVersion, ref EtabsUtils.DataBase.FrameSectionAssnignmetnsFileds, numberOFrRectors, ref FrameSectionAssnignmetnsValues);




                //ret = mysap.DatabaseTables.GetTableForDisplayArray(FrameLoadsDist, ref FrameLoadsDistFileds, "", ref tableVersion, ref filedsReturned, ref numberOFrRectors, ref FrameLoadsDistValues);
                //#region convertor
                //dic = new Dictionary<string, string>();
                //dic.Add("1", "Local-1");
                //dic.Add("2", "Local-2");
                //dic.Add("3", "Local-3");
                //dic.Add("X", "Global-X");
                //dic.Add("Y", "Global-Y");
                //dic.Add("Z", "Gravity");
                //dic.Add("X Projected", "Global-X Proj");
                //dic.Add("Y Projected", "Global-Y Proj");
                //dic.Add("Z Projected", "Gravity Proj");
                //dic.Add("Gravity Projected", "Gravity Proj");
                //convertorFromSapToEtabs(ref FrameLoadsDistValues, dic, 3, 13);
                //#endregion
                //ret = myEtabs.DatabaseTables.SetTableForEditingArray(EtabsUtils.DataBase.FrameLoadsDist, ref tableVersion, ref EtabsUtils.DataBase.FrameLoadsDistFileds, numberOFrRectors, ref FrameLoadsDistValues);



                //ret = mysap.DatabaseTables.GetTableForDisplayArray(FrameLoadTepm, ref FrameLoadTepmFileds, "", ref tableVersion, ref filedsReturned, ref numberOFrRectors, ref FrameLoadTepmValues);
                //ret = myEtabs.DatabaseTables.SetTableForEditingArray(EtabsUtils.DataBase.FrameLoadTepm, ref tableVersion, ref EtabsUtils.DataBase.FrameLoadTepmFileds, numberOFrRectors, ref FrameLoadTepmValues);




                //ret = mysap.DatabaseTables.GetTableForDisplayArray(FrameLoadPoint, ref FrameLoadPointFileds, "", ref tableVersion, ref filedsReturned, ref numberOFrRectors, ref FrameLoadPointValues);
                //ret = myEtabs.DatabaseTables.SetTableForEditingArray(EtabsUtils.DataBase.FrameLoadPoint, ref tableVersion, ref EtabsUtils.DataBase.FrameLoadPointFileds, numberOFrRectors, ref FrameLoadPointValues);



                //ret = mysap.DatabaseTables.GetTableForDisplayArray(FrameOutPutStation, ref FrameOutPutStationFileds, "", ref tableVersion, ref filedsReturned, ref numberOFrRectors, ref FrameOutPutStationValues);
                //ret = myEtabs.DatabaseTables.SetTableForEditingArray(EtabsUtils.DataBase.FrameOutPutStation, ref tableVersion, ref EtabsUtils.DataBase.FrameOutPutStationFileds, numberOFrRectors, ref FrameOutPutStationValues);


                //ret = mysap.DatabaseTables.GetTableForDisplayArray(FrameRealeses, ref FrameRealesesFileds, "", ref tableVersion, ref filedsReturned, ref numberOFrRectors, ref FrameRealesesValues);
                //ret = myEtabs.DatabaseTables.SetTableForEditingArray(EtabsUtils.DataBase.FrameRealeses, ref tableVersion, ref EtabsUtils.DataBase.FrameRealesesFileds, numberOFrRectors, ref FrameRealesesValues);


                //ret = mysap.DatabaseTables.GetTableForDisplayArray(insertionPoints, ref insertionPointsFileds, "", ref tableVersion, ref filedsReturned, ref numberOFrRectors, ref insertionPointsValues);
                //#region convertor
                //dic = new Dictionary<string, string>();
                //dic.Add("Yes", "NO");
                //dic.Add("No", "Yes");

                //convertorFromSapToEtabs(ref insertionPointsValues, dic, 11, 12);
                //#endregion
                //ret = myEtabs.DatabaseTables.SetTableForEditingArray(EtabsUtils.DataBase.insertionPoints, ref tableVersion, ref EtabsUtils.DataBase.insertionPointsFileds, numberOFrRectors, ref insertionPointsValues);


                #endregion


                ret = myEtabs.DatabaseTables.ApplyEditedTables(FillImportLog, ref NumberFatalErrors, ref NumberErrorMsg, ref NumWarrinings, ref NumberInfoMsg, ref importLog);
                // Create a new text file and open it for writing
                using (StreamWriter writer = new StreamWriter(filePath))
                {

                    writer.Write(importLog);
                }



                ret = mysap.DatabaseTables.GetTableForDisplayArray(GroupAssignments, ref GroupAssignmentsFileds, "", ref tableVersion, ref filedsReturned, ref numberOFrRectors, ref GroupAssignmentsValues);
                GroupAssignmentsValues = GroupAssignmentsValues.ToList().ConvertAll(item => item == "Joint" ? "Point" : item).ToArray();
                GroupAssignmentsValues = GroupAssignmentsValues.ToList().ConvertAll(item => item == "Frame" ? "Line" : item).ToArray();
                ret = myEtabs.DatabaseTables.SetTableForEditingArray(EtabsUtils.DataBase.GroupAssignments, ref tableVersion, ref EtabsUtils.DataBase.GroupAssignmentsFileds, numberOFrRectors, ref GroupAssignmentsValues);











                ret = myEtabs.DatabaseTables.ApplyEditedTables(FillImportLog, ref NumberFatalErrors, ref NumberErrorMsg, ref NumWarrinings, ref NumberInfoMsg, ref importLog);


                // Create a new text file and open it for writing
                using (StreamWriter writer = new StreamWriter(filePath))
                {

                    writer.Write(importLog);
                }

            }




            public static void aa()
            {

                #region definisions

                bool FillImportLog = true;
                int NumberFatalErrors = 0;
                int NumberErrorMsg = 0;
                int NumWarrinings = 0;
                string importLog = "";
                int NumberInfoMsg = 0;
                int ret = 0;
                cSapModel mysap = Get_sap_mode();
                ETABSv1.cSapModel myEtabs = EtabsUtils.Get_sap_mode();
                Dictionary<string, string> dic = null;
                string filePath = @"C:\Users\Ahmed.Ramzy\Desktop\New Text Document.txt";

                #endregion

                #region materils and Groups definitions
                ret = mysap.DatabaseTables.GetTableForDisplayArray(MaterialProp, ref MaterialPropFileds, "", ref tableVersion, ref filedsReturned, ref numberOFrRectors, ref MaterialPropValues);
                ret = myEtabs.DatabaseTables.SetTableForEditingArray(EtabsUtils.DataBase.MaterialProp, ref tableVersion, ref EtabsUtils.DataBase.MaterialPropFileds, numberOFrRectors, ref MaterialPropValues);


                ret = mysap.DatabaseTables.GetTableForDisplayArray(MaterialMechProp, ref MaterialMechPropFileds, "", ref tableVersion, ref filedsReturned, ref numberOFrRectors, ref MaterialMechPropValues);
                ret = myEtabs.DatabaseTables.SetTableForEditingArray(EtabsUtils.DataBase.MaterialMechProp, ref tableVersion, ref EtabsUtils.DataBase.MaterialMechPropFileds, numberOFrRectors, ref MaterialMechPropValues);

                ret = mysap.DatabaseTables.GetTableForDisplayArray(MaterialPropConc, ref MaterialPropConcFileds, "", ref tableVersion, ref filedsReturned, ref numberOFrRectors, ref MaterialPropConcValues);
                ret = myEtabs.DatabaseTables.SetTableForEditingArray(EtabsUtils.DataBase.MaterialPropConc, ref tableVersion, ref EtabsUtils.DataBase.MaterialPropConcFileds, numberOFrRectors, ref MaterialPropConcValues);


                ret = mysap.DatabaseTables.GetTableForDisplayArray(MaterialPropSteel, ref MaterialPropSteelFileds, "", ref tableVersion, ref filedsReturned, ref numberOFrRectors, ref MaterialPropSteelValues);
                ret = myEtabs.DatabaseTables.SetTableForEditingArray(EtabsUtils.DataBase.MaterialPropSteel, ref tableVersion, ref EtabsUtils.DataBase.MaterialPropSteelFileds, numberOFrRectors, ref MaterialPropSteelValues);



                ret = mysap.DatabaseTables.GetTableForDisplayArray(GroupDefinitions, ref GroupDefinitionsFileds, "", ref tableVersion, ref filedsReturned, ref numberOFrRectors, ref GroupDefinitionsValues);
                ret = myEtabs.DatabaseTables.SetTableForEditingArray(EtabsUtils.DataBase.GroupDefinitions, ref tableVersion, ref EtabsUtils.DataBase.GroupDefinitionsFileds, numberOFrRectors, ref GroupDefinitionsValues);




                ret = myEtabs.DatabaseTables.ApplyEditedTables(FillImportLog, ref NumberFatalErrors, ref NumberErrorMsg, ref NumWarrinings, ref NumberInfoMsg, ref importLog);
                // Create a new text file and open it for writing
                using (StreamWriter writer = new StreamWriter(filePath))
                {

                    writer.Write(importLog);
                }


                #endregion

                #region Frame Sections
                ret = mysap.DatabaseTables.GetTableForDisplayArray(GeneralFramesSections, ref GeneralFramesSectionsFileds, "", ref tableVersion, ref filedsReturned, ref numberOFrRectors, ref GeneralFramesSectionsValues);
                List<object> data3 = ConvertTableData(filedsReturned, GeneralFramesSectionsValues);

                foreach (dynamic item in data3)
                {
                    if (item.Shape == "I/Wide Flange")
                    {

                        if (filedsReturned.Contains("FileName"))
                        {

                            if (item.FileName != null)
                            {
                                item.FileName = CapitalizeVariableName(item.FileName);

                                ret = myEtabs.PropFrame.ImportProp(item.SectionName, item.Material, item.FileName, item.SectionName, item.Color, item.Notes);
                            }

                            else
                            {
                                ret = myEtabs.PropFrame.SetISection(item.SectionName, item.Material, double.Parse(item.t3), double.Parse(item.t2), double.Parse(item.tf), double.Parse(item.tw), double.Parse(item.t2b), double.Parse(item.tfb));

                            }





                        }


                        else
                        {
                            ret = myEtabs.PropFrame.SetISection(item.SectionName, item.Material, double.Parse(item.t3), double.Parse(item.t2), double.Parse(item.tf), double.Parse(item.tw), double.Parse(item.t2b), double.Parse(item.tfb));

                        }





                    }
                    if (item.Shape == "Box/Tube")
                    {
                        if (filedsReturned.Contains("FileName"))
                        {
                            if (item.FileName != null)
                            {
                                item.FileName = CapitalizeVariableName(item.FileName);
                                ret = myEtabs.PropFrame.ImportProp(item.SectionName, item.Material, item.FileName, item.SectionName);
                            }

                            else
                            {

                                ret = myEtabs.PropFrame.SetTube(item.SectionName, item.Material, double.Parse(item.t3), double.Parse(item.t2), double.Parse(item.tf), double.Parse(item.tw));

                            }



                        }


                        else
                        {
                            ret = myEtabs.PropFrame.SetTube(item.SectionName, item.Material, double.Parse(item.t3), double.Parse(item.t2), double.Parse(item.tf), double.Parse(item.tw));

                        }


                    }
                    if (item.Shape == "Pipe")
                    {
                        if (filedsReturned.Contains("FileName"))
                        {
                            if (item.FileName != null)
                            {
                                item.FileName = CapitalizeVariableName(item.FileName);
                                ret = myEtabs.PropFrame.ImportProp(item.SectionName, item.Material, item.FileName, item.SectionName);
                            }

                            else
                            {

                                ret = myEtabs.PropFrame.SetPipe(item.SectionName, item.Material, double.Parse(item.t3), double.Parse(item.tw));

                            }



                        }


                        else
                        {
                            ret = myEtabs.PropFrame.SetPipe(item.SectionName, item.Material, double.Parse(item.t3), double.Parse(item.tw));

                        }

                    }

                }

                #region getNonPresmatic

                ret = mysap.DatabaseTables.GetTableForDisplayArray(NonPresmaticSections, ref NonPresmaticSectionsFileds, "", ref tableVersion, ref filedsReturned, ref numberOFrRectors, ref NonPresmaticSectionsValues);
                List<object> data4 = ConvertTableData(filedsReturned, NonPresmaticSectionsValues);
                List<string> dataNonPresmatic = new List<string>();
                int numberOfNunPresmatic = 0;
                foreach (dynamic item in data4)
                {
                    int numberOfITems = 0;
                    string GUID = null;
                    string notes = null;
                    int color = 0;
                    string[] StartSec = null;
                    string[] Endsec = null;
                    double[] MyLength = null;
                    int[] MyType = null;
                    int[] EI33 = null;
                    int[] EI22 = null;
                    ret = mysap.PropFrame.GetNonPrismatic(item.SectionName, ref numberOfITems, ref StartSec, ref Endsec, ref MyLength, ref MyType, ref EI33, ref EI22, ref color, ref notes, ref GUID);
                    ret = myEtabs.PropFrame.SetNonPrismatic(item.SectionName, numberOfITems, ref StartSec, ref Endsec, ref MyLength, ref MyType, ref EI33, ref EI22, color, notes);



                }


                #endregion


                #endregion

                #region points
                ret = mysap.DatabaseTables.GetTableForDisplayArray(Connectevity, ref ConnectevityFileds, "", ref tableVersion, ref filedsReturned, ref numberOFrRectors, ref ConnectevityValues);
                ret = myEtabs.DatabaseTables.SetTableForEditingArray(EtabsUtils.DataBase.Connectevity, ref tableVersion, ref EtabsUtils.DataBase.ConnectevityFileds, numberOFrRectors, ref ConnectevityValues);


                ret = mysap.DatabaseTables.GetTableForDisplayArray(JointRestAss, ref JointRestAssFileds, "", ref tableVersion, ref filedsReturned, ref numberOFrRectors, ref JointRestAssValues);
                ret = myEtabs.DatabaseTables.SetTableForEditingArray(EtabsUtils.DataBase.JointRestAss, ref tableVersion, ref EtabsUtils.DataBase.JointRestAssFileds, numberOFrRectors, ref JointRestAssValues);


                ret = myEtabs.DatabaseTables.ApplyEditedTables(FillImportLog, ref NumberFatalErrors, ref NumberErrorMsg, ref NumWarrinings, ref NumberInfoMsg, ref importLog);
                // Create a new text file and open it for writing
                using (StreamWriter writer = new StreamWriter(filePath))
                {

                    writer.Write(importLog);
                }

                #endregion

                #region beams
                ret = mysap.DatabaseTables.GetTableForDisplayArray(ConnectevityBM, ref ConnectevBMityFileds, "", ref tableVersion, ref filedsReturned, ref numberOFrRectors, ref ConnectevityBMValues);
                ret = myEtabs.DatabaseTables.SetTableForEditingArray(EtabsUtils.DataBase.ConnectevityBM, ref tableVersion, ref EtabsUtils.DataBase.ConnectevBMityFileds, numberOFrRectors, ref ConnectevityBMValues);

                ret = myEtabs.DatabaseTables.ApplyEditedTables(FillImportLog, ref NumberFatalErrors, ref NumberErrorMsg, ref NumWarrinings, ref NumberInfoMsg, ref importLog);
                // Create a new text file and open it for writing
                using (StreamWriter writer = new StreamWriter(filePath))
                {

                    writer.Write(importLog);
                }

                ret = mysap.DatabaseTables.GetTableForDisplayArray(FrameSectionAssnignmetns, ref FrameSectionAssnignmetnsFileds, "", ref tableVersion, ref filedsReturned, ref numberOFrRectors, ref FrameSectionAssnignmetnsValues);
                ret = myEtabs.DatabaseTables.SetTableForEditingArray(EtabsUtils.DataBase.FrameSectionAssnignmetns, ref tableVersion, ref EtabsUtils.DataBase.FrameSectionAssnignmetnsFileds, numberOFrRectors, ref FrameSectionAssnignmetnsValues);

                ret = myEtabs.DatabaseTables.ApplyEditedTables(FillImportLog, ref NumberFatalErrors, ref NumberErrorMsg, ref NumWarrinings, ref NumberInfoMsg, ref importLog);
                // Create a new text file and open it for writing
                using (StreamWriter writer = new StreamWriter(filePath))
                {

                    writer.Write(importLog);
                }


                ret = mysap.DatabaseTables.GetTableForDisplayArray(FrameLoadsDist, ref FrameLoadsDistFileds, "", ref tableVersion, ref filedsReturned, ref numberOFrRectors, ref FrameLoadsDistValues);
                #region convertor
                dic = new Dictionary<string, string>();
                dic.Add("1", "Local-1");
                dic.Add("2", "Local-2");
                dic.Add("3", "Local-3");
                dic.Add("X", "Global-X");
                dic.Add("Y", "Global-Y");
                dic.Add("Z", "Gravity");
                dic.Add("X Projected", "Global-X Proj");
                dic.Add("Y Projected", "Global-Y Proj");
                dic.Add("Z Projected", "Gravity Proj");
                dic.Add("Gravity Projected", "Gravity Proj");
                convertorFromSapToEtabs(ref FrameLoadsDistValues, dic, 3, 13);
                #endregion
                ret = myEtabs.DatabaseTables.SetTableForEditingArray(EtabsUtils.DataBase.FrameLoadsDist, ref tableVersion, ref EtabsUtils.DataBase.FrameLoadsDistFileds, numberOFrRectors, ref FrameLoadsDistValues);



                ret = mysap.DatabaseTables.GetTableForDisplayArray(FrameLoadTepm, ref FrameLoadTepmFileds, "", ref tableVersion, ref filedsReturned, ref numberOFrRectors, ref FrameLoadTepmValues);
                ret = myEtabs.DatabaseTables.SetTableForEditingArray(EtabsUtils.DataBase.FrameLoadTepm, ref tableVersion, ref EtabsUtils.DataBase.FrameLoadTepmFileds, numberOFrRectors, ref FrameLoadTepmValues);




                ret = mysap.DatabaseTables.GetTableForDisplayArray(FrameLoadPoint, ref FrameLoadPointFileds, "", ref tableVersion, ref filedsReturned, ref numberOFrRectors, ref FrameLoadPointValues);
                ret = myEtabs.DatabaseTables.SetTableForEditingArray(EtabsUtils.DataBase.FrameLoadPoint, ref tableVersion, ref EtabsUtils.DataBase.FrameLoadPointFileds, numberOFrRectors, ref FrameLoadPointValues);



                ret = mysap.DatabaseTables.GetTableForDisplayArray(FrameOutPutStation, ref FrameOutPutStationFileds, "", ref tableVersion, ref filedsReturned, ref numberOFrRectors, ref FrameOutPutStationValues);
                ret = myEtabs.DatabaseTables.SetTableForEditingArray(EtabsUtils.DataBase.FrameOutPutStation, ref tableVersion, ref EtabsUtils.DataBase.FrameOutPutStationFileds, numberOFrRectors, ref FrameOutPutStationValues);


                ret = mysap.DatabaseTables.GetTableForDisplayArray(FrameRealeses, ref FrameRealesesFileds, "", ref tableVersion, ref filedsReturned, ref numberOFrRectors, ref FrameRealesesValues);
                ret = myEtabs.DatabaseTables.SetTableForEditingArray(EtabsUtils.DataBase.FrameRealeses, ref tableVersion, ref EtabsUtils.DataBase.FrameRealesesFileds, numberOFrRectors, ref FrameRealesesValues);


                ret = mysap.DatabaseTables.GetTableForDisplayArray(insertionPoints, ref insertionPointsFileds, "", ref tableVersion, ref filedsReturned, ref numberOFrRectors, ref insertionPointsValues);
                #region convertor
                dic = new Dictionary<string, string>();
                dic.Add("Yes", "NO");
                dic.Add("No", "Yes");

                convertorFromSapToEtabs(ref insertionPointsValues, dic, 11, 12);
                #endregion
                ret = myEtabs.DatabaseTables.SetTableForEditingArray(EtabsUtils.DataBase.insertionPoints, ref tableVersion, ref EtabsUtils.DataBase.insertionPointsFileds, numberOFrRectors, ref insertionPointsValues);


                #endregion




                ret = mysap.DatabaseTables.GetTableForDisplayArray(GroupAssignments, ref GroupAssignmentsFileds, "", ref tableVersion, ref filedsReturned, ref numberOFrRectors, ref GroupAssignmentsValues);
                ret = myEtabs.DatabaseTables.SetTableForEditingArray(EtabsUtils.DataBase.GroupAssignments, ref tableVersion, ref EtabsUtils.DataBase.GroupAssignmentsFileds, numberOFrRectors, ref GroupAssignmentsValues);











                ret = myEtabs.DatabaseTables.ApplyEditedTables(FillImportLog, ref NumberFatalErrors, ref NumberErrorMsg, ref NumWarrinings, ref NumberInfoMsg, ref importLog);


                // Create a new text file and open it for writing
                using (StreamWriter writer = new StreamWriter(filePath))
                {

                    writer.Write(importLog);
                }

            }

            public static void aa7()
            {


                #region definisions

                bool FillImportLog = true;
                int NumberFatalErrors = 0;
                int NumberErrorMsg = 0;
                int NumWarrinings = 0;
                string importLog = "";
                int NumberInfoMsg = 0;
                int ret = 0;
                cSapModel mysap = Get_sap_mode();
                ETABSv1.cSapModel myEtabs = EtabsUtils.Get_sap_mode();
                Dictionary<string, string> dic = null;
                string filePath = @"C:\Users\Ahmed.Ramzy\Desktop\New Text Document.txt";

                #endregion

                #region materils and Groups definitions
                ret = mysap.DatabaseTables.GetTableForDisplayArray(MaterialProp, ref MaterialPropFileds, "", ref tableVersion, ref filedsReturned, ref numberOFrRectors, ref MaterialPropValues);
                ret = myEtabs.DatabaseTables.SetTableForEditingArray(EtabsUtils.DataBase.MaterialProp, ref tableVersion, ref EtabsUtils.DataBase.MaterialPropFileds, numberOFrRectors, ref MaterialPropValues);


                ret = mysap.DatabaseTables.GetTableForDisplayArray(MaterialMechProp, ref MaterialMechPropFileds, "", ref tableVersion, ref filedsReturned, ref numberOFrRectors, ref MaterialMechPropValues);
                ret = myEtabs.DatabaseTables.SetTableForEditingArray(EtabsUtils.DataBase.MaterialMechProp, ref tableVersion, ref EtabsUtils.DataBase.MaterialMechPropFileds, numberOFrRectors, ref MaterialMechPropValues);

                ret = mysap.DatabaseTables.GetTableForDisplayArray(MaterialPropConc, ref MaterialPropConcFileds, "", ref tableVersion, ref filedsReturned, ref numberOFrRectors, ref MaterialPropConcValues);
                ret = myEtabs.DatabaseTables.SetTableForEditingArray(EtabsUtils.DataBase.MaterialPropConc, ref tableVersion, ref EtabsUtils.DataBase.MaterialPropConcFileds, numberOFrRectors, ref MaterialPropConcValues);


                ret = mysap.DatabaseTables.GetTableForDisplayArray(MaterialPropSteel, ref MaterialPropSteelFileds, "", ref tableVersion, ref filedsReturned, ref numberOFrRectors, ref MaterialPropSteelValues);
                ret = myEtabs.DatabaseTables.SetTableForEditingArray(EtabsUtils.DataBase.MaterialPropSteel, ref tableVersion, ref EtabsUtils.DataBase.MaterialPropSteelFileds, numberOFrRectors, ref MaterialPropSteelValues);



                ret = mysap.DatabaseTables.GetTableForDisplayArray(GroupDefinitions, ref GroupDefinitionsFileds, "", ref tableVersion, ref filedsReturned, ref numberOFrRectors, ref GroupDefinitionsValues);
                ret = myEtabs.DatabaseTables.SetTableForEditingArray(EtabsUtils.DataBase.GroupDefinitions, ref tableVersion, ref EtabsUtils.DataBase.GroupDefinitionsFileds, numberOFrRectors, ref GroupDefinitionsValues);




                #endregion

                #region Frame Sections

                List<string> secSapDistnct = Sections.GetAllNamesAssignedToFramesInSap(mysap);

                ret = mysap.DatabaseTables.GetTableForDisplayArray(GeneralFramesSections, ref GeneralFramesSectionsFileds, "", ref tableVersion, ref filedsReturned, ref numberOFrRectors, ref GeneralFramesSectionsValues);
                List<object> data3 = ConvertTableData(filedsReturned, GeneralFramesSectionsValues);
                data3 = data3.Where(x => secSapDistnct.Contains(((dynamic)x).SectionName)).ToList();

                List<string> DataOfISection = new List<string>();
                List<string> DataOfBoxSection = new List<string>();
                List<string> DataOfPipeSection = new List<string>();


                int numberOfIsections = 0;
                int numberOfBoxsections = 0;
                int numberOfPipesections = 0;


                foreach (dynamic item in data3)
                {
                    if (item.Shape == "I/Wide Flange")
                    {
                        numberOfIsections++;
                        if (filedsReturned.Contains("FileName"))
                        {

                            DataOfISection.Add(item.SectionName);
                            DataOfISection.Add(item.Material);
                            DataOfISection.Add(item.FromFile);

                            if (item.FileName != null)
                            {
                                DataOfISection.Add(CapitalizeVariableName(item.FileName));
                            }
                            else
                            {
                                DataOfISection.Add(item.FileName);

                            }
                            DataOfISection.Add(item.SectInFile);
                            DataOfISection.Add(item.t3);
                            DataOfISection.Add(item.t2);
                            DataOfISection.Add(item.tf);
                            DataOfISection.Add(item.tw);
                            DataOfISection.Add(item.t2b);
                            DataOfISection.Add(item.tfb);
                            DataOfISection.Add(item.FilletRadius);
                            DataOfISection.Add(item.AMod);
                            DataOfISection.Add(item.A2Mod);
                            DataOfISection.Add(item.A3Mod);
                            DataOfISection.Add(item.JMod);
                            DataOfISection.Add(item.I2Mod);
                            DataOfISection.Add(item.I3Mod);
                            DataOfISection.Add(item.MMod);
                            DataOfISection.Add(item.WMod);
                            DataOfISection.Add(item.Color);
                            DataOfISection.Add(item.Notes);

                        }


                        else
                        {
                            DataOfISection.Add(item.SectionName);
                            DataOfISection.Add(item.Material);
                            DataOfISection.Add(item.FromFile);
                            DataOfISection.Add(item.t3);
                            DataOfISection.Add(item.t2);
                            DataOfISection.Add(item.tf);
                            DataOfISection.Add(item.tw);
                            DataOfISection.Add(item.t2b);
                            DataOfISection.Add(item.tfb);
                            DataOfISection.Add(item.FilletRadius);
                            DataOfISection.Add(item.AMod);
                            DataOfISection.Add(item.A2Mod);
                            DataOfISection.Add(item.A3Mod);
                            DataOfISection.Add(item.JMod);
                            DataOfISection.Add(item.I2Mod);
                            DataOfISection.Add(item.I3Mod);
                            DataOfISection.Add(item.MMod);
                            DataOfISection.Add(item.WMod);
                            DataOfISection.Add(item.Color);
                            DataOfISection.Add(item.Notes);

                        }





                    }
                    if (item.Shape == "Box/Tube")
                    {
                        numberOfBoxsections++;
                        if (filedsReturned.Contains("FileName"))
                        {

                            DataOfBoxSection.Add(item.SectionName);
                            DataOfBoxSection.Add(item.Material);
                            DataOfBoxSection.Add(item.FromFile);

                            if (item.FileName != null)
                            {
                                DataOfBoxSection.Add(CapitalizeVariableName(item.FileName));
                            }
                            else
                            {
                                DataOfBoxSection.Add(item.FileName);

                            }
                            DataOfBoxSection.Add(item.SectInFile);
                            DataOfBoxSection.Add(item.t3);
                            DataOfBoxSection.Add(item.t2);
                            DataOfBoxSection.Add(item.tf);
                            DataOfBoxSection.Add(item.tw);
                            DataOfBoxSection.Add(item.FilletRadius);
                            DataOfBoxSection.Add(item.AMod);
                            DataOfBoxSection.Add(item.A2Mod);
                            DataOfBoxSection.Add(item.A3Mod);
                            DataOfBoxSection.Add(item.JMod);
                            DataOfBoxSection.Add(item.I2Mod);
                            DataOfBoxSection.Add(item.I3Mod);
                            DataOfBoxSection.Add(item.MMod);
                            DataOfBoxSection.Add(item.WMod);
                            DataOfBoxSection.Add(item.Color);
                            DataOfBoxSection.Add(item.Notes);

                        }


                        else
                        {
                            DataOfBoxSection.Add(item.SectionName);
                            DataOfBoxSection.Add(item.Material);
                            DataOfBoxSection.Add(item.FromFile);
                            DataOfBoxSection.Add(item.t3);
                            DataOfBoxSection.Add(item.t2);
                            DataOfBoxSection.Add(item.tf);
                            DataOfBoxSection.Add(item.tw);
                            DataOfBoxSection.Add(item.t2b);
                            DataOfBoxSection.Add(item.tfb);
                            DataOfBoxSection.Add(item.FilletRadius);
                            DataOfBoxSection.Add(item.AMod);
                            DataOfBoxSection.Add(item.A2Mod);
                            DataOfBoxSection.Add(item.A3Mod);
                            DataOfBoxSection.Add(item.JMod);
                            DataOfBoxSection.Add(item.I2Mod);
                            DataOfBoxSection.Add(item.I3Mod);
                            DataOfBoxSection.Add(item.MMod);
                            DataOfBoxSection.Add(item.WMod);
                            DataOfBoxSection.Add(item.Color);
                            DataOfBoxSection.Add(item.Notes);

                        }


                    }
                    if (item.Shape == "Pipe")
                    {
                        numberOfPipesections++;
                        if (filedsReturned.Contains("FileName"))
                        {

                            DataOfPipeSection.Add(item.SectionName);
                            DataOfPipeSection.Add(item.Material);
                            DataOfPipeSection.Add(item.FromFile);

                            if (item.FileName != null)
                            {
                                DataOfPipeSection.Add(CapitalizeVariableName(item.FileName));
                            }
                            else
                            {
                                DataOfPipeSection.Add(item.FileName);

                            }
                            DataOfPipeSection.Add(item.SectInFile);
                            DataOfPipeSection.Add(item.t3);
                            DataOfPipeSection.Add(item.tw);
                            DataOfPipeSection.Add(item.AMod);
                            DataOfPipeSection.Add(item.A2Mod);
                            DataOfPipeSection.Add(item.A3Mod);
                            DataOfPipeSection.Add(item.JMod);
                            DataOfPipeSection.Add(item.I2Mod);
                            DataOfPipeSection.Add(item.I3Mod);
                            DataOfPipeSection.Add(item.MMod);
                            DataOfPipeSection.Add(item.WMod);
                            DataOfPipeSection.Add(item.Color);
                            DataOfPipeSection.Add(item.Notes);

                        }


                        else
                        {
                            DataOfPipeSection.Add(item.SectionName);
                            DataOfPipeSection.Add(item.Material);
                            DataOfPipeSection.Add(item.FromFile);
                            DataOfPipeSection.Add(item.t3);
                            DataOfPipeSection.Add(item.tw);
                            DataOfPipeSection.Add(item.AMod);
                            DataOfPipeSection.Add(item.A2Mod);
                            DataOfPipeSection.Add(item.A3Mod);
                            DataOfPipeSection.Add(item.JMod);
                            DataOfPipeSection.Add(item.I2Mod);
                            DataOfPipeSection.Add(item.I3Mod);
                            DataOfPipeSection.Add(item.MMod);
                            DataOfPipeSection.Add(item.WMod);
                            DataOfPipeSection.Add(item.Color);
                            DataOfPipeSection.Add(item.Notes);

                        }


                    }

                }

                string[] DataOfIsectionArray = DataOfISection.ToArray();
                string[] DataOfBoxsectionArray = DataOfBoxSection.ToArray();
                string[] DataOfPipesectionArray = DataOfPipeSection.ToArray();




                ret = myEtabs.DatabaseTables.SetTableForEditingArray(EtabsUtils.DataBase.FramesISections, ref tableVersion, ref EtabsUtils.DataBase.FramesISectionsFileds, numberOfIsections, ref DataOfIsectionArray);
                ret = myEtabs.DatabaseTables.SetTableForEditingArray(EtabsUtils.DataBase.FramesBoxSections, ref tableVersion, ref EtabsUtils.DataBase.FramesBoxSectionsFileds, numberOfIsections, ref DataOfBoxsectionArray);
                ret = myEtabs.DatabaseTables.SetTableForEditingArray(EtabsUtils.DataBase.FramesPipeSections, ref tableVersion, ref EtabsUtils.DataBase.FramesPipeSectionsFileds, numberOfIsections, ref DataOfPipesectionArray);




                #endregion

                #region getNonPresmatic

                ret = mysap.DatabaseTables.GetTableForDisplayArray(NonPresmaticSections, ref NonPresmaticSectionsFileds, "", ref tableVersion, ref filedsReturned, ref numberOFrRectors, ref NonPresmaticSectionsValues);
                List<object> data4 = ConvertTableData(filedsReturned, NonPresmaticSectionsValues);
                data4 = data4.Where(x => secSapDistnct.Contains(((dynamic)x).SectionName)).ToList();

                List<string> dataNonPresmatic = new List<string>();
                int numberOfNunPresmatic = 0;
                foreach (dynamic item in data4)
                {
                    //if ((int.Parse((string)item.NumSegment)) > 1)
                    //{
                    //    continue;
                    //}
                    numberOfNunPresmatic++;
                    dataNonPresmatic.Add(item.SectionName);
                    dataNonPresmatic.Add(item.StartSect);
                    dataNonPresmatic.Add(item.EndSect);


                    if (item.LengthType == "Variable")
                    {
                        dataNonPresmatic.Add("Proportional");
                    }
                    else
                    {
                        dataNonPresmatic.Add("Absolute");
                    }


                    if (item.LengthType == "Variable")
                    {
                        dataNonPresmatic.Add(item.VarLength);
                    }
                    else
                    {
                        dataNonPresmatic.Add(item.AbsLength);
                    }

                    dataNonPresmatic.Add(item.EI33Var);
                    dataNonPresmatic.Add(item.EI22Var);

                    // dataNonPresmatic.Add(item.Color);


                }

                string[] dataNonPresmaticArray = dataNonPresmatic.ToArray();
                ret = myEtabs.DatabaseTables.SetTableForEditingArray(EtabsUtils.DataBase.NonPresmaticSections, ref tableVersion, ref EtabsUtils.DataBase.NonPresmaticSectionsFileds, numberOfNunPresmatic, ref dataNonPresmaticArray);

                #endregion

                #region points

                ret = mysap.DatabaseTables.GetTableForDisplayArray(Connectevity, ref ConnectevityFileds, "", ref tableVersion, ref filedsReturned, ref numberOFrRectors, ref ConnectevityValues);
                ret = myEtabs.DatabaseTables.SetTableForEditingArray(EtabsUtils.DataBase.Connectevity, ref tableVersion, ref EtabsUtils.DataBase.ConnectevityFileds, numberOFrRectors, ref ConnectevityValues);


                ret = mysap.DatabaseTables.GetTableForDisplayArray(JointRestAss, ref JointRestAssFileds, "", ref tableVersion, ref filedsReturned, ref numberOFrRectors, ref JointRestAssValues);
                ret = myEtabs.DatabaseTables.SetTableForEditingArray(EtabsUtils.DataBase.JointRestAss, ref tableVersion, ref EtabsUtils.DataBase.JointRestAssFileds, numberOFrRectors, ref JointRestAssValues);


                #endregion

                #region beams

                ret = mysap.DatabaseTables.GetTableForDisplayArray(ConnectevityBM, ref ConnectevBMityFileds, "", ref tableVersion, ref filedsReturned, ref numberOFrRectors, ref ConnectevityBMValues);
                ret = myEtabs.DatabaseTables.SetTableForEditingArray(EtabsUtils.DataBase.ConnectevityBM, ref tableVersion, ref EtabsUtils.DataBase.ConnectevBMityFileds, numberOFrRectors, ref ConnectevityBMValues);

                ret = mysap.DatabaseTables.GetTableForDisplayArray(FrameSectionAssnignmetns, ref FrameSectionAssnignmetnsFileds, "", ref tableVersion, ref filedsReturned, ref numberOFrRectors, ref FrameSectionAssnignmetnsValues);
                ret = myEtabs.DatabaseTables.SetTableForEditingArray(EtabsUtils.DataBase.FrameSectionAssnignmetns, ref tableVersion, ref EtabsUtils.DataBase.FrameSectionAssnignmetnsFileds, numberOFrRectors, ref FrameSectionAssnignmetnsValues);




                ret = mysap.DatabaseTables.GetTableForDisplayArray(FrameLoadsDist, ref FrameLoadsDistFileds, "", ref tableVersion, ref filedsReturned, ref numberOFrRectors, ref FrameLoadsDistValues);
                #region convertor
                dic = new Dictionary<string, string>();
                dic.Add("1", "Local-1");
                dic.Add("2", "Local-2");
                dic.Add("3", "Local-3");
                dic.Add("X", "Global-X");
                dic.Add("Y", "Global-Y");
                dic.Add("Z", "Gravity");
                dic.Add("X Projected", "Global-X Proj");
                dic.Add("Y Projected", "Global-Y Proj");
                dic.Add("Z Projected", "Gravity Proj");
                dic.Add("Gravity Projected", "Gravity Proj");
                convertorFromSapToEtabs(ref FrameLoadsDistValues, dic, 3, 13);
                #endregion
                ret = myEtabs.DatabaseTables.SetTableForEditingArray(EtabsUtils.DataBase.FrameLoadsDist, ref tableVersion, ref EtabsUtils.DataBase.FrameLoadsDistFileds, numberOFrRectors, ref FrameLoadsDistValues);



                ret = mysap.DatabaseTables.GetTableForDisplayArray(FrameLoadTepm, ref FrameLoadTepmFileds, "", ref tableVersion, ref filedsReturned, ref numberOFrRectors, ref FrameLoadTepmValues);
                ret = myEtabs.DatabaseTables.SetTableForEditingArray(EtabsUtils.DataBase.FrameLoadTepm, ref tableVersion, ref EtabsUtils.DataBase.FrameLoadTepmFileds, numberOFrRectors, ref FrameLoadTepmValues);




                ret = mysap.DatabaseTables.GetTableForDisplayArray(FrameLoadPoint, ref FrameLoadPointFileds, "", ref tableVersion, ref filedsReturned, ref numberOFrRectors, ref FrameLoadPointValues);
                ret = myEtabs.DatabaseTables.SetTableForEditingArray(EtabsUtils.DataBase.FrameLoadPoint, ref tableVersion, ref EtabsUtils.DataBase.FrameLoadPointFileds, numberOFrRectors, ref FrameLoadPointValues);



                ret = mysap.DatabaseTables.GetTableForDisplayArray(FrameOutPutStation, ref FrameOutPutStationFileds, "", ref tableVersion, ref filedsReturned, ref numberOFrRectors, ref FrameOutPutStationValues);
                ret = myEtabs.DatabaseTables.SetTableForEditingArray(EtabsUtils.DataBase.FrameOutPutStation, ref tableVersion, ref EtabsUtils.DataBase.FrameOutPutStationFileds, numberOFrRectors, ref FrameOutPutStationValues);


                ret = mysap.DatabaseTables.GetTableForDisplayArray(FrameRealeses, ref FrameRealesesFileds, "", ref tableVersion, ref filedsReturned, ref numberOFrRectors, ref FrameRealesesValues);
                ret = myEtabs.DatabaseTables.SetTableForEditingArray(EtabsUtils.DataBase.FrameRealeses, ref tableVersion, ref EtabsUtils.DataBase.FrameRealesesFileds, numberOFrRectors, ref FrameRealesesValues);


                ret = mysap.DatabaseTables.GetTableForDisplayArray(insertionPoints, ref insertionPointsFileds, "", ref tableVersion, ref filedsReturned, ref numberOFrRectors, ref insertionPointsValues);
                #region convertor
                dic = new Dictionary<string, string>();
                dic.Add("Yes", "NO");
                dic.Add("No", "Yes");

                convertorFromSapToEtabs(ref insertionPointsValues, dic, 11, 12);
                #endregion
                ret = myEtabs.DatabaseTables.SetTableForEditingArray(EtabsUtils.DataBase.insertionPoints, ref tableVersion, ref EtabsUtils.DataBase.insertionPointsFileds, numberOFrRectors, ref insertionPointsValues);


                #endregion


                ret = myEtabs.DatabaseTables.ApplyEditedTables(FillImportLog, ref NumberFatalErrors, ref NumberErrorMsg, ref NumWarrinings, ref NumberInfoMsg, ref importLog);
                // Create a new text file and open it for writing
                using (StreamWriter writer = new StreamWriter(filePath))
                {

                    writer.Write(importLog);
                }



                ret = mysap.DatabaseTables.GetTableForDisplayArray(GroupAssignments, ref GroupAssignmentsFileds, "", ref tableVersion, ref filedsReturned, ref numberOFrRectors, ref GroupAssignmentsValues);
                GroupAssignmentsValues = GroupAssignmentsValues.ToList().ConvertAll(item => item == "Joint" ? "Point" : item).ToArray();
                GroupAssignmentsValues = GroupAssignmentsValues.ToList().ConvertAll(item => item == "Frame" ? "Line" : item).ToArray();
                ret = myEtabs.DatabaseTables.SetTableForEditingArray(EtabsUtils.DataBase.GroupAssignments, ref tableVersion, ref EtabsUtils.DataBase.GroupAssignmentsFileds, numberOFrRectors, ref GroupAssignmentsValues);











                ret = myEtabs.DatabaseTables.ApplyEditedTables(FillImportLog, ref NumberFatalErrors, ref NumberErrorMsg, ref NumWarrinings, ref NumberInfoMsg, ref importLog);


                // Create a new text file and open it for writing
                using (StreamWriter writer = new StreamWriter(filePath))
                {

                    writer.Write(importLog);
                }

            }

            public static void aa6()
            {


                #region definisions

                bool FillImportLog = true;
                int NumberFatalErrors = 0;
                int NumberErrorMsg = 0;
                int NumWarrinings = 0;
                string importLog = "";
                int NumberInfoMsg = 0;
                int ret = 0;
                cSapModel mysap = Get_sap_mode();
                ETABSv1.cSapModel myEtabs = EtabsUtils.Get_sap_mode();
                Dictionary<string, string> dic = null;
                string filePath = @"C:\Users\Ahmed.Ramzy\Desktop\New Text Document.txt";

                #endregion

                #region materils and Groups definitions
                ret = mysap.DatabaseTables.GetTableForDisplayArray(MaterialProp, ref MaterialPropFileds, "", ref tableVersion, ref filedsReturned, ref numberOFrRectors, ref MaterialPropValues);
                ret = myEtabs.DatabaseTables.SetTableForEditingArray(EtabsUtils.DataBase.MaterialProp, ref tableVersion, ref EtabsUtils.DataBase.MaterialPropFileds, numberOFrRectors, ref MaterialPropValues);


                ret = mysap.DatabaseTables.GetTableForDisplayArray(MaterialMechProp, ref MaterialMechPropFileds, "", ref tableVersion, ref filedsReturned, ref numberOFrRectors, ref MaterialMechPropValues);
                ret = myEtabs.DatabaseTables.SetTableForEditingArray(EtabsUtils.DataBase.MaterialMechProp, ref tableVersion, ref EtabsUtils.DataBase.MaterialMechPropFileds, numberOFrRectors, ref MaterialMechPropValues);

                ret = mysap.DatabaseTables.GetTableForDisplayArray(MaterialPropConc, ref MaterialPropConcFileds, "", ref tableVersion, ref filedsReturned, ref numberOFrRectors, ref MaterialPropConcValues);
                ret = myEtabs.DatabaseTables.SetTableForEditingArray(EtabsUtils.DataBase.MaterialPropConc, ref tableVersion, ref EtabsUtils.DataBase.MaterialPropConcFileds, numberOFrRectors, ref MaterialPropConcValues);


                ret = mysap.DatabaseTables.GetTableForDisplayArray(MaterialPropSteel, ref MaterialPropSteelFileds, "", ref tableVersion, ref filedsReturned, ref numberOFrRectors, ref MaterialPropSteelValues);
                ret = myEtabs.DatabaseTables.SetTableForEditingArray(EtabsUtils.DataBase.MaterialPropSteel, ref tableVersion, ref EtabsUtils.DataBase.MaterialPropSteelFileds, numberOFrRectors, ref MaterialPropSteelValues);



                ret = mysap.DatabaseTables.GetTableForDisplayArray(GroupDefinitions, ref GroupDefinitionsFileds, "", ref tableVersion, ref filedsReturned, ref numberOFrRectors, ref GroupDefinitionsValues);
                ret = myEtabs.DatabaseTables.SetTableForEditingArray(EtabsUtils.DataBase.GroupDefinitions, ref tableVersion, ref EtabsUtils.DataBase.GroupDefinitionsFileds, numberOFrRectors, ref GroupDefinitionsValues);




                ret = myEtabs.DatabaseTables.ApplyEditedTables(FillImportLog, ref NumberFatalErrors, ref NumberErrorMsg, ref NumWarrinings, ref NumberInfoMsg, ref importLog);
                // Create a new text file and open it for writing
                using (StreamWriter writer = new StreamWriter(filePath))
                {

                    writer.Write(importLog);
                }


                #endregion

                #region Frame Sections


                List<string> secSapDistnct = Sections.GetAllNamesAssignedToFramesInSap(mysap);

                ret = mysap.DatabaseTables.GetTableForDisplayArray(GeneralFramesSections, ref GeneralFramesSectionsFileds, "", ref tableVersion, ref filedsReturned, ref numberOFrRectors, ref GeneralFramesSectionsValues);
                List<object> data3 = ConvertTableData(filedsReturned, GeneralFramesSectionsValues);
                data3 = data3.Where(x => secSapDistnct.Contains(((dynamic)x).SectionName)).ToList();

                List<string> DataOfISection = new List<string>();
                List<string> DataOfBoxSection = new List<string>();
                List<string> DataOfPipeSection = new List<string>();


                int numberOfIsections = 0;
                int numberOfBoxsections = 0;
                int numberOfPipesections = 0;


                foreach (dynamic item in data3)
                {
                    if (item.Shape == "I/Wide Flange")
                    {
                        numberOfIsections++;
                        if (filedsReturned.Contains("FileName"))
                        {

                            DataOfISection.Add(item.SectionName);
                            DataOfISection.Add(item.Material);
                            DataOfISection.Add(item.FromFile);

                            if (item.FileName != null)
                            {
                                DataOfISection.Add(CapitalizeVariableName(item.FileName));
                            }
                            else
                            {
                                DataOfISection.Add(item.FileName);

                            }
                            DataOfISection.Add(item.SectInFile);
                            DataOfISection.Add(item.t3);
                            DataOfISection.Add(item.t2);
                            DataOfISection.Add(item.tf);
                            DataOfISection.Add(item.tw);
                            DataOfISection.Add(item.t2b);
                            DataOfISection.Add(item.tfb);
                            DataOfISection.Add(item.FilletRadius);
                            DataOfISection.Add(item.AMod);
                            DataOfISection.Add(item.A2Mod);
                            DataOfISection.Add(item.A3Mod);
                            DataOfISection.Add(item.JMod);
                            DataOfISection.Add(item.I2Mod);
                            DataOfISection.Add(item.I3Mod);
                            DataOfISection.Add(item.MMod);
                            DataOfISection.Add(item.WMod);
                            DataOfISection.Add(item.Color);
                            DataOfISection.Add(item.Notes);

                        }


                        else
                        {
                            DataOfISection.Add(item.SectionName);
                            DataOfISection.Add(item.Material);
                            DataOfISection.Add(item.FromFile);
                            DataOfISection.Add(item.t3);
                            DataOfISection.Add(item.t2);
                            DataOfISection.Add(item.tf);
                            DataOfISection.Add(item.tw);
                            DataOfISection.Add(item.t2b);
                            DataOfISection.Add(item.tfb);
                            DataOfISection.Add(item.FilletRadius);
                            DataOfISection.Add(item.AMod);
                            DataOfISection.Add(item.A2Mod);
                            DataOfISection.Add(item.A3Mod);
                            DataOfISection.Add(item.JMod);
                            DataOfISection.Add(item.I2Mod);
                            DataOfISection.Add(item.I3Mod);
                            DataOfISection.Add(item.MMod);
                            DataOfISection.Add(item.WMod);
                            DataOfISection.Add(item.Color);
                            DataOfISection.Add(item.Notes);

                        }





                    }
                    if (item.Shape == "Box/Tube")
                    {
                        numberOfBoxsections++;
                        if (filedsReturned.Contains("FileName"))
                        {

                            DataOfBoxSection.Add(item.SectionName);
                            DataOfBoxSection.Add(item.Material);
                            DataOfBoxSection.Add(item.FromFile);

                            if (item.FileName != null)
                            {
                                DataOfBoxSection.Add(CapitalizeVariableName(item.FileName));
                            }
                            else
                            {
                                DataOfBoxSection.Add(item.FileName);

                            }
                            DataOfBoxSection.Add(item.SectInFile);
                            DataOfBoxSection.Add(item.t3);
                            DataOfBoxSection.Add(item.t2);
                            DataOfBoxSection.Add(item.tf);
                            DataOfBoxSection.Add(item.tw);
                            DataOfBoxSection.Add(item.FilletRadius);
                            DataOfBoxSection.Add(item.AMod);
                            DataOfBoxSection.Add(item.A2Mod);
                            DataOfBoxSection.Add(item.A3Mod);
                            DataOfBoxSection.Add(item.JMod);
                            DataOfBoxSection.Add(item.I2Mod);
                            DataOfBoxSection.Add(item.I3Mod);
                            DataOfBoxSection.Add(item.MMod);
                            DataOfBoxSection.Add(item.WMod);
                            DataOfBoxSection.Add(item.Color);
                            DataOfBoxSection.Add(item.Notes);

                        }


                        else
                        {
                            DataOfBoxSection.Add(item.SectionName);
                            DataOfBoxSection.Add(item.Material);
                            DataOfBoxSection.Add(item.FromFile);
                            DataOfBoxSection.Add(item.t3);
                            DataOfBoxSection.Add(item.t2);
                            DataOfBoxSection.Add(item.tf);
                            DataOfBoxSection.Add(item.tw);
                            DataOfBoxSection.Add(item.t2b);
                            DataOfBoxSection.Add(item.tfb);
                            DataOfBoxSection.Add(item.FilletRadius);
                            DataOfBoxSection.Add(item.AMod);
                            DataOfBoxSection.Add(item.A2Mod);
                            DataOfBoxSection.Add(item.A3Mod);
                            DataOfBoxSection.Add(item.JMod);
                            DataOfBoxSection.Add(item.I2Mod);
                            DataOfBoxSection.Add(item.I3Mod);
                            DataOfBoxSection.Add(item.MMod);
                            DataOfBoxSection.Add(item.WMod);
                            DataOfBoxSection.Add(item.Color);
                            DataOfBoxSection.Add(item.Notes);

                        }


                    }
                    if (item.Shape == "Pipe")
                    {
                        numberOfPipesections++;
                        if (filedsReturned.Contains("FileName"))
                        {

                            DataOfPipeSection.Add(item.SectionName);
                            DataOfPipeSection.Add(item.Material);
                            DataOfPipeSection.Add(item.FromFile);

                            if (item.FileName != null)
                            {
                                DataOfPipeSection.Add(CapitalizeVariableName(item.FileName));
                            }
                            else
                            {
                                DataOfPipeSection.Add(item.FileName);

                            }
                            DataOfPipeSection.Add(item.SectInFile);
                            DataOfPipeSection.Add(item.t3);
                            DataOfPipeSection.Add(item.tw);
                            DataOfPipeSection.Add(item.AMod);
                            DataOfPipeSection.Add(item.A2Mod);
                            DataOfPipeSection.Add(item.A3Mod);
                            DataOfPipeSection.Add(item.JMod);
                            DataOfPipeSection.Add(item.I2Mod);
                            DataOfPipeSection.Add(item.I3Mod);
                            DataOfPipeSection.Add(item.MMod);
                            DataOfPipeSection.Add(item.WMod);
                            DataOfPipeSection.Add(item.Color);
                            DataOfPipeSection.Add(item.Notes);

                        }


                        else
                        {
                            DataOfPipeSection.Add(item.SectionName);
                            DataOfPipeSection.Add(item.Material);
                            DataOfPipeSection.Add(item.FromFile);
                            DataOfPipeSection.Add(item.t3);
                            DataOfPipeSection.Add(item.tw);
                            DataOfPipeSection.Add(item.AMod);
                            DataOfPipeSection.Add(item.A2Mod);
                            DataOfPipeSection.Add(item.A3Mod);
                            DataOfPipeSection.Add(item.JMod);
                            DataOfPipeSection.Add(item.I2Mod);
                            DataOfPipeSection.Add(item.I3Mod);
                            DataOfPipeSection.Add(item.MMod);
                            DataOfPipeSection.Add(item.WMod);
                            DataOfPipeSection.Add(item.Color);
                            DataOfPipeSection.Add(item.Notes);

                        }


                    }

                }

                string[] DataOfIsectionArray = DataOfISection.ToArray();
                string[] DataOfBoxsectionArray = DataOfBoxSection.ToArray();
                string[] DataOfPipesectionArray = DataOfPipeSection.ToArray();




                ret = myEtabs.DatabaseTables.SetTableForEditingArray(EtabsUtils.DataBase.FramesISections, ref tableVersion, ref EtabsUtils.DataBase.FramesISectionsFileds, numberOfIsections, ref DataOfIsectionArray);
                ret = myEtabs.DatabaseTables.SetTableForEditingArray(EtabsUtils.DataBase.FramesBoxSections, ref tableVersion, ref EtabsUtils.DataBase.FramesBoxSectionsFileds, numberOfIsections, ref DataOfBoxsectionArray);
                ret = myEtabs.DatabaseTables.SetTableForEditingArray(EtabsUtils.DataBase.FramesPipeSections, ref tableVersion, ref EtabsUtils.DataBase.FramesPipeSectionsFileds, numberOfIsections, ref DataOfPipesectionArray);



                ret = myEtabs.DatabaseTables.ApplyEditedTables(FillImportLog, ref NumberFatalErrors, ref NumberErrorMsg, ref NumWarrinings, ref NumberInfoMsg, ref importLog);


                // Create a new text file and open it for writing
                using (StreamWriter writer = new StreamWriter(filePath))
                {

                    writer.Write(importLog);
                }

                #endregion

                #region getNonPresmatic

                ret = mysap.DatabaseTables.GetTableForDisplayArray(NonPresmaticSections, ref NonPresmaticSectionsFileds, "", ref tableVersion, ref filedsReturned, ref numberOFrRectors, ref NonPresmaticSectionsValues);
                List<object> data4 = ConvertTableData(filedsReturned, NonPresmaticSectionsValues);
                data4 = data4.Where(x => secSapDistnct.Contains(((dynamic)x).SectionName)).ToList();

                List<string> dataNonPresmatic = new List<string>();
                int numberOfNunPresmatic = 0;
                foreach (dynamic item in data4)
                {
                    //if ((int.Parse((string)item.NumSegment)) > 1)
                    //{
                    //    continue;
                    //}
                    numberOfNunPresmatic++;
                    dataNonPresmatic.Add(item.SectionName);
                    dataNonPresmatic.Add(item.StartSect);
                    dataNonPresmatic.Add(item.EndSect);


                    if (item.LengthType == "Variable")
                    {
                        dataNonPresmatic.Add("Proportional");
                    }
                    else
                    {
                        dataNonPresmatic.Add("Absolute");
                    }


                    if (item.LengthType == "Variable")
                    {
                        dataNonPresmatic.Add(item.VarLength);
                    }
                    else
                    {
                        dataNonPresmatic.Add(item.AbsLength);
                    }

                    dataNonPresmatic.Add(item.EI33Var);
                    dataNonPresmatic.Add(item.EI22Var);

                    // dataNonPresmatic.Add(item.Color);


                }

                string[] dataNonPresmaticArray = dataNonPresmatic.ToArray();
                ret = myEtabs.DatabaseTables.SetTableForEditingArray(EtabsUtils.DataBase.NonPresmaticSections, ref tableVersion, ref EtabsUtils.DataBase.NonPresmaticSectionsFileds, numberOfNunPresmatic, ref dataNonPresmaticArray);



                ret = myEtabs.DatabaseTables.ApplyEditedTables(FillImportLog, ref NumberFatalErrors, ref NumberErrorMsg, ref NumWarrinings, ref NumberInfoMsg, ref importLog);
                // Create a new text file and open it for writing
                using (StreamWriter writer = new StreamWriter(filePath))
                {

                    writer.Write(importLog);
                }

                #endregion



                #region points

                ret = mysap.DatabaseTables.GetTableForDisplayArray(Connectevity, ref ConnectevityFileds, "", ref tableVersion, ref filedsReturned, ref numberOFrRectors, ref ConnectevityValues);
                ret = myEtabs.DatabaseTables.SetTableForEditingArray(EtabsUtils.DataBase.Connectevity, ref tableVersion, ref EtabsUtils.DataBase.ConnectevityFileds, numberOFrRectors, ref ConnectevityValues);


                ret = mysap.DatabaseTables.GetTableForDisplayArray(JointRestAss, ref JointRestAssFileds, "", ref tableVersion, ref filedsReturned, ref numberOFrRectors, ref JointRestAssValues);
                ret = myEtabs.DatabaseTables.SetTableForEditingArray(EtabsUtils.DataBase.JointRestAss, ref tableVersion, ref EtabsUtils.DataBase.JointRestAssFileds, numberOFrRectors, ref JointRestAssValues);


                #endregion

                ret = myEtabs.DatabaseTables.ApplyEditedTables(FillImportLog, ref NumberFatalErrors, ref NumberErrorMsg, ref NumWarrinings, ref NumberInfoMsg, ref importLog);


                #region beams
                ret = mysap.DatabaseTables.GetTableForDisplayArray(ConnectevityBM, ref ConnectevBMityFileds, "", ref tableVersion, ref filedsReturned, ref numberOFrRectors, ref ConnectevityBMValues);
                #region convertor
                List<Frame> myframes = getSapFramesE();
                data3 = ConvertTableData(filedsReturned, ConnectevityBMValues);

                List<string> dataOfBeams = new List<string>();
                string[] dataOfBeamsArra = null;
                foreach (Frame item in myframes)
                {
                    dataOfBeams.Add(item.ID);
                    dataOfBeams.Add(item.StartPoint.Id);
                    dataOfBeams.Add(item.EndPoint.Id);

                }
                dataOfBeamsArra = dataOfBeams.ToArray();
                #endregion
                ret = myEtabs.DatabaseTables.SetTableForEditingArray(EtabsUtils.DataBase.ConnectevityBM, ref tableVersion, ref EtabsUtils.DataBase.ConnectevBMityFileds, numberOFrRectors, ref dataOfBeamsArra);

                ret = myEtabs.DatabaseTables.ApplyEditedTables(FillImportLog, ref NumberFatalErrors, ref NumberErrorMsg, ref NumWarrinings, ref NumberInfoMsg, ref importLog);
                // Create a new text file and open it for writing
                using (StreamWriter writer = new StreamWriter(filePath))
                {

                    writer.Write(importLog);
                }


                ret = mysap.DatabaseTables.GetTableForDisplayArray(FrameSectionAssnignmetns, ref FrameSectionAssnignmetnsFileds, "", ref tableVersion, ref filedsReturned, ref numberOFrRectors, ref FrameSectionAssnignmetnsValues);
                ret = myEtabs.DatabaseTables.SetTableForEditingArray(EtabsUtils.DataBase.FrameSectionAssnignmetns, ref tableVersion, ref EtabsUtils.DataBase.FrameSectionAssnignmetnsFileds, numberOFrRectors, ref FrameSectionAssnignmetnsValues);

                ret = myEtabs.DatabaseTables.ApplyEditedTables(FillImportLog, ref NumberFatalErrors, ref NumberErrorMsg, ref NumWarrinings, ref NumberInfoMsg, ref importLog);
                // Create a new text file and open it for writing
                using (StreamWriter writer = new StreamWriter(filePath))
                {

                    writer.Write(importLog);
                }



                ret = mysap.DatabaseTables.GetTableForDisplayArray(FrameLoadsDist, ref FrameLoadsDistFileds, "", ref tableVersion, ref filedsReturned, ref numberOFrRectors, ref FrameLoadsDistValues);
                #region convertor
                dic = new Dictionary<string, string>();
                dic.Add("1", "Local-1");
                dic.Add("2", "Local-2");
                dic.Add("3", "Local-3");
                dic.Add("X", "Global-X");
                dic.Add("Y", "Global-Y");
                dic.Add("Z", "Gravity");
                dic.Add("X Projected", "Global-X Proj");
                dic.Add("Y Projected", "Global-Y Proj");
                dic.Add("Z Projected", "Gravity Proj");
                dic.Add("Gravity Projected", "Gravity Proj");
                convertorFromSapToEtabs(ref FrameLoadsDistValues, dic, 3, 13);
                #endregion
                ret = myEtabs.DatabaseTables.SetTableForEditingArray(EtabsUtils.DataBase.FrameLoadsDist, ref tableVersion, ref EtabsUtils.DataBase.FrameLoadsDistFileds, numberOFrRectors, ref FrameLoadsDistValues);



                ret = mysap.DatabaseTables.GetTableForDisplayArray(FrameLoadTepm, ref FrameLoadTepmFileds, "", ref tableVersion, ref filedsReturned, ref numberOFrRectors, ref FrameLoadTepmValues);
                ret = myEtabs.DatabaseTables.SetTableForEditingArray(EtabsUtils.DataBase.FrameLoadTepm, ref tableVersion, ref EtabsUtils.DataBase.FrameLoadTepmFileds, numberOFrRectors, ref FrameLoadTepmValues);




                ret = mysap.DatabaseTables.GetTableForDisplayArray(FrameLoadPoint, ref FrameLoadPointFileds, "", ref tableVersion, ref filedsReturned, ref numberOFrRectors, ref FrameLoadPointValues);
                ret = myEtabs.DatabaseTables.SetTableForEditingArray(EtabsUtils.DataBase.FrameLoadPoint, ref tableVersion, ref EtabsUtils.DataBase.FrameLoadPointFileds, numberOFrRectors, ref FrameLoadPointValues);



                ret = mysap.DatabaseTables.GetTableForDisplayArray(FrameOutPutStation, ref FrameOutPutStationFileds, "", ref tableVersion, ref filedsReturned, ref numberOFrRectors, ref FrameOutPutStationValues);
                ret = myEtabs.DatabaseTables.SetTableForEditingArray(EtabsUtils.DataBase.FrameOutPutStation, ref tableVersion, ref EtabsUtils.DataBase.FrameOutPutStationFileds, numberOFrRectors, ref FrameOutPutStationValues);


                ret = mysap.DatabaseTables.GetTableForDisplayArray(FrameRealeses, ref FrameRealesesFileds, "", ref tableVersion, ref filedsReturned, ref numberOFrRectors, ref FrameRealesesValues);
                ret = myEtabs.DatabaseTables.SetTableForEditingArray(EtabsUtils.DataBase.FrameRealeses, ref tableVersion, ref EtabsUtils.DataBase.FrameRealesesFileds, numberOFrRectors, ref FrameRealesesValues);


                ret = mysap.DatabaseTables.GetTableForDisplayArray(insertionPoints, ref insertionPointsFileds, "", ref tableVersion, ref filedsReturned, ref numberOFrRectors, ref insertionPointsValues);
                #region convertor
                dic = new Dictionary<string, string>();
                dic.Add("Yes", "NO");
                dic.Add("No", "Yes");

                convertorFromSapToEtabs(ref insertionPointsValues, dic, 11, 12);
                #endregion
                ret = myEtabs.DatabaseTables.SetTableForEditingArray(EtabsUtils.DataBase.insertionPoints, ref tableVersion, ref EtabsUtils.DataBase.insertionPointsFileds, numberOFrRectors, ref insertionPointsValues);


                #endregion




                ret = mysap.DatabaseTables.GetTableForDisplayArray(GroupAssignments, ref GroupAssignmentsFileds, "", ref tableVersion, ref filedsReturned, ref numberOFrRectors, ref GroupAssignmentsValues);
                GroupAssignmentsValues = GroupAssignmentsValues.ToList().ConvertAll(item => item == "Joint" ? "Point" : item).ToArray();
                GroupAssignmentsValues = GroupAssignmentsValues.ToList().ConvertAll(item => item == "Frame" ? "Line" : item).ToArray();
                ret = myEtabs.DatabaseTables.SetTableForEditingArray(EtabsUtils.DataBase.GroupAssignments, ref tableVersion, ref EtabsUtils.DataBase.GroupAssignmentsFileds, numberOFrRectors, ref GroupAssignmentsValues);











                ret = myEtabs.DatabaseTables.ApplyEditedTables(FillImportLog, ref NumberFatalErrors, ref NumberErrorMsg, ref NumWarrinings, ref NumberInfoMsg, ref importLog);


                // Create a new text file and open it for writing
                using (StreamWriter writer = new StreamWriter(filePath))
                {

                    writer.Write(importLog);
                }

            }


            #region SectionFilter

            public static string CapitalizeVariableName(string path)
            {
                string filename = Path.GetFileNameWithoutExtension(path);
                string[] parts = filename.Split('\\');
                string variableName = parts[parts.Length - 1];

                // Capitalize the first letter
                string capitalizedVariable = char.ToUpper(variableName[0]) + variableName.Substring(1);

                return capitalizedVariable;
            }

            public static List<object> ConvertTableData(string[] fieldKeysIncluded, string[] tableData)
            {
                List<object> convertedData = new List<object>();

                // Assuming each record has the same number of fields as fieldKeysIncluded
                int recordSize = fieldKeysIncluded.Length;

                for (int i = 0; i < tableData.Length / recordSize; i++)
                {
                    string[] recordData = new string[recordSize];
                    Array.Copy(tableData, i * recordSize, recordData, 0, recordSize);

                    var recordObject = new ExpandoObject() as IDictionary<string, object>;

                    for (int j = 0; j < recordSize; j++)
                    {
                        string fieldName = fieldKeysIncluded[j];
                        recordObject[fieldName] = recordData[j];
                    }

                    convertedData.Add(recordObject);
                }

                return convertedData;
            }









            public class TableRecord
            {
                public string[] Data { get; set; }
            }

            // Define the filter condition function
            public static bool FilterByShape(TableRecord record, string shape)
            {
                // Assuming "Shape" field is at index 21 in the Data array
                return record.Data[2] == shape;
            }

            // Get filtered table data
            public static List<TableRecord> GetFilteredTableData(string[] fieldKeysIncluded, string[] tableData, string shape)
            {
                List<TableRecord> filteredData = new List<TableRecord>();

                // Assuming each record has 22 fields, adjust the number if necessary
                int recordSize = 22;

                for (int i = 0; i < tableData.Length / recordSize; i++)
                {
                    string[] recordData = new string[recordSize];
                    Array.Copy(tableData, i * recordSize, recordData, 0, recordSize);

                    TableRecord record = new TableRecord
                    {
                        Data = recordData
                    };

                    if (FilterByShape(record, shape))
                    {
                        filteredData.Add(record);
                    }
                }

                return filteredData;
            }
            #endregion


            public static void convertorFromSapToEtabs(ref string[] data, Dictionary<string, string> map, int columnToReplace, int noOfcolumns)
            {
                int step = noOfcolumns; // Calculate the step size for column indexes
                int maxIndex = data.Length; // Maximum index in the array

                for (int i = columnToReplace; i <= maxIndex; i += step)
                {
                    if (map.ContainsKey(data[i]))
                    {
                        data[i] = map[data[i]];
                    }
                }
            }

            #endregion

        }

        #endregion


        #region HydraSimulator

        public static List<Frame> getFramesForHydra()
        {

            List<Column> col = new List<Column>();
            List<Beam> bm = new List<Beam>();
            List<Frame> frames = getSapFramesForHydra(ref bm, ref col);
            List<Frame> result = new List<Frame>();
            List<string> sections = frames.Select(x => x.Sec).Distinct().ToList();

            Dictionary<string, string> libDic = null;
            Dictionary<string, Section_type> dic = Sections.getSectionTypesForHydraTest(sections, ref libDic);


            List<string> notFound = sections.Except(dic.Keys.ToList()).ToList();

            foreach (Frame f in frames)
            {
                Section_type t = dic[f.Sec];

                try
                {
                    switch (t)
                    {
                        case Section_type.Pipe_H:
                            f.sectionTypeName = "Pipe";
                            f.IsBUS = false;
                            f.Lib = "N.A";
                            f.Shape = "Pipe";
                            break;
                        case Section_type.Pipe_B:
                            f.sectionTypeName = "Pipe";
                            f.IsBUS = true;
                            f.Lib = "N.A";
                            f.Shape = "Pipe";
                            break;
                        case Section_type.Pipe_T:
                            f.sectionTypeName = "VARPipe";
                            f.IsBUS = true;
                            f.Lib = "N.A";
                            f.Shape = "VARPipe";
                            break;
                        case Section_type.Tube_H:
                            f.sectionTypeName = "Box";
                            f.IsBUS = false;
                            f.Lib = "N.A";
                            f.Shape = "Box";
                            break;
                        case Section_type.Tube_B:
                            f.sectionTypeName = "Box";
                            f.IsBUS = true;
                            f.Lib = "N.A";
                            f.Shape = "Box";
                            break;
                        case Section_type.Tube_T:
                            f.sectionTypeName = "VARBox";
                            f.IsBUS = true;
                            f.Lib = "N.A";
                            f.Shape = "VARBox";
                            break;
                        case Section_type.HEA:
                            f.sectionTypeName = "I";
                            f.IsBUS = false;
                            f.Lib = "N.A";
                            f.Shape = "I";

                            break;
                        case Section_type.HEAA:
                            f.sectionTypeName = "I";
                            f.IsBUS = false;
                            f.Lib = "N.A";
                            f.Shape = "I";
                            break;
                        case Section_type.HEB:
                            f.sectionTypeName = "I";
                            f.IsBUS = false;
                            f.Lib = "N.A";
                            f.Shape = "I";
                            break;
                        case Section_type.HEM:
                            f.sectionTypeName = "I";
                            f.IsBUS = false;
                            f.Lib = "N.A";
                            f.Shape = "I";
                            break;
                        case Section_type.IPE:
                            f.sectionTypeName = "I";
                            f.IsBUS = false;
                            f.Lib = "N.A";
                            f.Shape = "I";
                            break;
                        case Section_type.IPEA:
                            f.sectionTypeName = "I";
                            f.IsBUS = false;
                            f.Lib = "N.A";
                            f.Shape = "I";
                            break;
                        case Section_type.IPEO:
                            f.sectionTypeName = "I";
                            f.IsBUS = false;
                            f.Lib = "N.A";
                            f.Shape = "I";
                            break;
                        case Section_type.I_B:
                            f.sectionTypeName = "I";
                            f.IsBUS = true;
                            f.Lib = "N.A";
                            f.Shape = "I";
                            break;
                        case Section_type.I_T:
                            f.sectionTypeName = "VARI";
                            f.IsBUS = true;
                            f.Lib = "N.A";
                            f.Shape = "I";
                            break;
                        case Section_type.C_H:
                            f.sectionTypeName = "Channel";
                            f.IsBUS = false;
                            f.Lib = "N.A";
                            f.Shape = "Channel";
                            break;
                        case Section_type.DblAngle:
                            f.sectionTypeName = "DblAngle";
                            f.IsBUS = false;
                            f.Lib = "N.A";
                            f.Shape = "DblAngle";

                            break;
                        case Section_type.EQAngle:
                            f.sectionTypeName = "Angle";
                            f.IsBUS = false;
                            f.Lib = "N.A";
                            f.Shape = "Angle";

                            break;
                        case Section_type.UQAngle:
                            f.sectionTypeName = "Angle";
                            f.IsBUS = false;
                            f.Lib = "N.A";
                            f.Shape = "Angle";

                            break;
                        case Section_type.I_SECTION_B:
                            f.sectionTypeName = "I";
                            f.IsBUS = true;
                            f.Lib = "N.A";
                            f.Shape = "I";

                            break;
                        case Section_type.I_SECTION_H:
                            f.sectionTypeName = "I";
                            f.IsBUS = false;
                            f.Lib = libDic[f.Sec];
                            f.Shape = "I";

                            break;


                        case Section_type.Channel_B:
                            f.sectionTypeName = "Channel";
                            f.IsBUS = true;
                            f.Lib = "N.A";
                            f.Shape = "Channel";

                            break;
                        case Section_type.Channel_H:
                            f.sectionTypeName = "Channel";
                            f.IsBUS = false;
                            f.Lib = "N.A";
                            f.Shape = "Channel";

                            break;

                        default:
                            f.sectionTypeName = "Dummy";
                            f.IsBUS = false;
                            f.Lib = "N.A";
                            f.Shape = "Dummy";

                            break;
                    }

                }
                catch (Exception ex)
                {
                    continue;
                }

                f.Length = Point.getDistance(f.StartPoint, f.EndPoint);


            }

            object[,] arr = new object[frames.Count, 20];


            for (int i = 0; i < frames.Count; i++)
            {
                arr[i, 0] = frames[i].ID;
                arr[i, 1] = frames[i].StartPoint.Id;
                arr[i, 2] = frames[i].EndPoint.Id;
                arr[i, 3] = frames[i].StartPoint.X;
                arr[i, 4] = frames[i].StartPoint.Y;
                arr[i, 5] = frames[i].StartPoint.Z;
                arr[i, 6] = frames[i].EndPoint.X;
                arr[i, 7] = frames[i].EndPoint.Y;
                arr[i, 8] = frames[i].EndPoint.Z;
                arr[i, 9] = frames[i].Sec;
                arr[i, 10] = frames[i].Lib;
                arr[i, 11] = frames[i].sectionTypeName;
                arr[i, 12] = frames[i].IsBUS;
                arr[i, 13] = true;
                arr[i, 14] = true;
                arr[i, 15] = true;
                arr[i, 16] = frames[i].Group;
                arr[i, 17] = frames[i].Material;
                arr[i, 18] = frames[i].Length;
                arr[i, 19] = frames[i].Beta;

            }


            Excel.Application app = null;
            Excel._Workbook wb = null;
            Excel._Worksheet ws = null;

            ExcelUtils.CreateExcelForSAP(ref app, ref wb, ref ws, "Sap");

            Range stCell = (Range)ws.Cells[7, 1];
            Range endCell = (Range)ws.Cells[frames.Count + 7 - 1, 20];
            Range totalRange = ws.Range[stCell, endCell];

            totalRange.Value2 = arr;
            ws.Range["F5"].Value = frames.Count;
            wb.Save();

            wb.Close();

            return result;
        }

        public static List<Frame> getSapFramesForHydra(ref List<Beam> beams, ref List<Column> columns)
        {

            cSapModel mySapModel = Get_sap_mode();
            List<string> frame_ids = getSapFrameIDS();

            if (getSelectedFramesIDS().Count > 0)
            {
                frame_ids = getSelectedFramesIDS();
            }


            List<Frame> sapFrames = new List<Frame>();

            #region defintions 
            string st_point = null;
            string end_point = null;
            int ret;
            double x1, y1, z1, x2, y2, z2;
            x1 = y1 = z1 = x2 = y2 = z2 = 0;
            String sectionName = null;
            string auto = null;
            string shapeName = null;
            eFramePropType shape = eFramePropType.Angle;
            int numberOfGroup = 0;
            string[] assignedGroups = null;
            string materialProperty = null;

            double rotation = 0;
            bool isAdvancedAxes = false;

            string NameInFile = null;
            string FileName = null;

            #endregion


            foreach (string id in frame_ids)
            {
                Frame frame = new Frame();


                #region getPoints and length
                ret = mySapModel.FrameObj.GetPoints(id, ref st_point, ref end_point);
                ret = mySapModel.PointObj.GetCoordCartesian(st_point, ref x1, ref y1, ref z1);
                ret = mySapModel.PointObj.GetCoordCartesian(end_point, ref x2, ref y2, ref z2);

                frame.ID = id;
                frame.StartPoint = new Point(st_point, x1, y1, z1);
                frame.EndPoint = new Point(end_point, x2, y2, z2);
                frame.MidPoint = Point.getMidPoint(frame.StartPoint, frame.EndPoint);
                frame.Length = System.Numerics.Vector3.Distance(Point.ToVector(frame.StartPoint), Point.ToVector(frame.EndPoint));
                #endregion

                #region get section
                mySapModel.FrameObj.GetSection(id, ref sectionName, ref auto);
                frame.Sec = sectionName;

                mySapModel.PropFrame.GetTypeOAPI(sectionName, ref shape);

                if (shape == eFramePropType.Variable)
                {
                    #region defintions
                    int NumberItems = 0;
                    string[] StartSec = null;
                    string[] EndSec = null;
                    double[] MyLength = null;
                    int[] MyType = null;
                    int[] EI33 = null;
                    int[] EI22 = null;
                    int Color = 0;
                    string Notes = null;
                    string GUID = null;
                    eFramePropType shape2 = eFramePropType.Angle;
                    #endregion

                    ret = mySapModel.PropFrame.GetNonPrismatic(sectionName, ref NumberItems, ref StartSec, ref EndSec, ref MyLength, ref MyType, ref EI33, ref EI22, ref Color, ref Notes, ref GUID);
                    ret = mySapModel.PropFrame.GetTypeOAPI(StartSec[0], ref shape2);
                    sectionName = StartSec[0];

                    //tapered i or pipe or box
                    switch (shape2)
                    {
                        case eFramePropType.I:
                            shapeName = "VARI";

                            break;

                        case eFramePropType.Box:
                            shapeName = "VARBox";
                            break;

                        case eFramePropType.Pipe:
                            shapeName = "VARPipe";
                            break;

                    }



                }

                else
                {
                    shapeName = shape.ToString();
                }

                frame.Shape = shapeName;
                #endregion

                #region get material and group and rotation angle
                mySapModel.FrameObj.GetGroupAssign(id, ref numberOfGroup, ref assignedGroups);
                frame.Group = assignedGroups[0];
                frame.Groups = assignedGroups.ToList();
                mySapModel.PropFrame.GetMaterial(sectionName, ref materialProperty);
                frame.Material = materialProperty;

                mySapModel.FrameObj.GetLocalAxes(id, ref rotation, ref isAdvancedAxes);

                frame.Beta = rotation;
                #endregion

                #region getBusOption
                ret = mySapModel.PropFrame.GetNameInPropFile(sectionName, ref NameInFile, ref FileName, ref materialProperty, ref shape);

                if (NameInFile == "")
                {
                    frame.IsBUS = true;
                }
                else
                {
                    frame.IsBUS = false;
                    frame.Lib = FileName;
                }

                #endregion



                #region  get Adanced Rotation 

                int[] mypDir = new int[1];
                string[] myplpt = new string[1];
                double[] myplvector = new double[2];


                string[] myPoints = new string[2];
                double[] MyPlVect = new double[3];
                int[] MyPlDir = new int[2];
                string st_id = "";
                string end_id = "";
                Boolean aa = true;
                bool Advanced = true;
                int plane = 0;
                int plVect = 0;
                string cys = null;

                double angleInDegrees = 0;

                angleInDegrees = rotation;



                #endregion

                Point.getMemberType(frame);



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

                sapFrames.Add(frame);
            }


            return sapFrames;
        }




        #endregion


    }

}