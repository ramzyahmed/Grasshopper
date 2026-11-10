using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using ETABSv1;
using Microsoft.Office.Interop.Excel;
using MyGrasshopperAssembly2;
using Rhino;
using Rhino.DocObjects;
using Rhino.Geometry;
//using GH_IO.Serialization;
//using Rhino.Input.Custom;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Numerics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Xml.Linq;
using Application = Microsoft.Office.Interop.Excel.Application;
using Excel = Microsoft.Office.Interop.Excel;
using Line = Autodesk.Revit.DB.Line;
using Parameter = Autodesk.Revit.DB.Parameter;
//using SAP2000v1;

namespace BIM_Integration
{
    public class EtabsUtils

    {
        public static string EuroPath = @"C:\Program Files\Computers and Structures\ETABS 20\Property Libraries\Euro.xml";
        public static string AmericanPath = @"C:\Program Files\Computers and Structures\ETABS 20\Property Libraries\AISC15.xml";
        public static string BiritishPath = @"C:\Program Files\Computers and Structures\ETABS 20\Property Libraries\BSShapes2006.xml";

        public static string path2 = @"";
        private static string[] filedsKeysIncluded;

        public static cSapModel Get_sap_mode()
        {
            ETABSv1.cOAPI mySapObject2 = (ETABSv1.cOAPI)Marshal.GetActiveObject("CSI.ETABS.API.ETABSObject");
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



            return mySapObject2.SapModel;


            //ETABSv1.cHelper cHelper = new ETABSv1.Helper();
            //cOAPI mySapObject2 = cHelper.GetObject("CSI.ETABS.API.ETABSObject");
            //ETABSv1.cSapModel mysapmodel = mySapObject2.SapModel;



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
                cSapModel mySapModel = Get_sap_mode();

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

                            if (NameInFile != "")
                            {
                                mysectiontype = Section_type.I_SECTION_H;
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
                cSapModel mySapModel = Get_sap_mode();



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
        public static List<string> getSelectedPointIDS()
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
                if (ObjectType[i] == 1)
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


            List<Frame> etabsFrames = new List<Frame>();


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
            #region getStories
            int NumberNames = 0;
            string[] MyName = null;
            string[] PropName = null;
            string[] StoryName = null;
            string[] PointName1 = null;
            string[] PointName2 = null;
            double[] Point1X = null;
            double[] Point1Y = null;
            double[] Point1Z = null;
            double[] Point2X = null;
            double[] Point2Y = null;
            double[] Point2Z = null;
            double[] Angle = null;
            double[] Offset1X = null;
            double[] Offset2X = null;
            double[] Offset1Y = null;
            double[] Offset2Y = null;
            double[] Offset1Z = null;
            double[] Offset2Z = null;
            int[] CardinalPoint = null;
            string csys = "Global";


            int result = mySapModel.FrameObj.GetAllFrames(
                            ref NumberNames,
                            ref MyName,
                            ref PropName,
                            ref StoryName,
                            ref PointName1,
                            ref PointName2,
                            ref Point1X,
                            ref Point1Y,
                            ref Point1Z,
                            ref Point2X,
                            ref Point2Y,
                            ref Point2Z,
                            ref Angle,
                            ref Offset1X,
                            ref Offset2X,
                            ref Offset1Y,
                            ref Offset2Y,
                            ref Offset1Z,
                            ref Offset2Z,
                            ref CardinalPoint
);

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

                #region  get Rotation 

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

                #region Releases

                bool[] ii = new bool[6];
                bool[] jj = new bool[6];
                double[] StartValue = new double[6];
                double[] EndValue = new double[6];
                mySapModel.FrameObj.GetReleases(id, ref ii, ref jj, ref StartValue, ref EndValue);



                if (ii.Any(x => x))
                {
                    frame.StartReleased = true;
                }

                else
                {
                    frame.StartReleased = false;

                }


                if (jj.Any(x => x))
                {
                    frame.EndReleased = true;
                }

                else
                {
                    frame.EndReleased = false;

                }

                #endregion Releases

                #region lengtOffsets
                double stLength = 0;
                double endLength = 0;
                double RZ = 0;
                bool Auto = false;

                mySapModel.FrameObj.GetEndLengthOffset(id, ref Auto, ref stLength, ref endLength, ref RZ);

                frame.StartLengthOffset = stLength;
                frame.EndLengthOffset = endLength;

                #endregion

                #region getStory
                int index = Array.IndexOf(MyName, id);
                string story = StoryName[index];
                string topStory = null;
                string botStory = null;
                string label = null;


                frame.Story = story;

                mySapModel.PointObj.GetLabelFromName(frame.StartPoint.Id, ref label, ref botStory);
                mySapModel.PointObj.GetLabelFromName(frame.EndPoint.Id, ref label, ref topStory);

                frame.EtabsTopStory = topStory;
                frame.EtabsBotStory = botStory;
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

                #region getMemberType
                eFrameDesignOrientation myDesignOriegnationType = eFrameDesignOrientation.Beam;

                mySapModel.FrameObj.GetDesignOrientation(id, ref myDesignOriegnationType);

                if (myDesignOriegnationType == eFrameDesignOrientation.Column)
                {
                    frame.MemberType = MemberType.Column;
                }

                else
                {
                    frame.MemberType = MemberType.Beam;

                }

                #endregion


                etabsFrames.Add(frame);
            }


            #region validateGroups

            try
            {
                List<Frame> sapFramesAssignedToGroup = etabsFrames.Where(x => x.Groups.Any(y => y[0] == '*')).ToList();

                foreach (var item in sapFramesAssignedToGroup)
                {
                    item.Group = item.Groups.First(x => x[0] == '*');
                    item.MainGroup = item.Groups.First(x => x[0] == '*');
                }

                var sapFramesGroups = sapFramesAssignedToGroup.GroupBy(x => x.MainGroup);

                foreach (var framesGroup in sapFramesGroups)
                {
                    List<string> sections = framesGroup.Select(x => x.Sec).Distinct().ToList();

                    foreach (var sapFrame in framesGroup)
                    {
                        sapFrame.SubGroup = sapFrame.Group + " _" + sapFrame.Sec;
                        sapFrame.Group = sapFrame.Group + " _" + sapFrame.Sec;

                        //sapFrame.SubGroup = sapFrame.Group + " _" + sections.IndexOf(sapFrame.Sec);
                        //sapFrame.SubGroupUsage = HardCodedData.MemberSubGroupDic[framesGroup.Key];


                    }
                }
            }

            catch (Exception)
            {

                throw;
            }




            #endregion


            #region getColumnReinforcementData
            ColumnRebarDesignData.RetrieveDataFromETABS(ref etabsFrames);
            #endregion

            return etabsFrames;
        }


        #region ColumnMArks
        public class GroupInfo
        {
            public List<string> Sections { get; set; }
            public List<double> StartZ { get; set; }
            public List<double> EndZ { get; set; }
            public int CountOfMembers { get; set; }

            public GroupInfo(List<string> sections, List<double> startZ, List<double> endZ, int countOfMembers)
            {
                Sections = sections;
                StartZ = startZ;
                EndZ = endZ;
                CountOfMembers = countOfMembers;
            }

            // Override Equals and GetHashCode for proper dictionary key comparison
            public override bool Equals(object obj)
            {
                if (obj == null || GetType() != obj.GetType())
                    return false;

                GroupInfo other = (GroupInfo)obj;

                return Sections.SequenceEqual(other.Sections)
                    && StartZ.SequenceEqual(other.StartZ)
                    && EndZ.SequenceEqual(other.EndZ)
                    && CountOfMembers == other.CountOfMembers;
            }

            public override int GetHashCode()
            {
                int hash = 17;
                hash = hash * 31 + Sections.Aggregate(0, (acc, x) => acc + x.GetHashCode());
                hash = hash * 31 + StartZ.Aggregate(0, (acc, x) => acc + x.GetHashCode());
                hash = hash * 31 + EndZ.Aggregate(0, (acc, x) => acc + x.GetHashCode());
                hash = hash * 31 + CountOfMembers.GetHashCode();
                return hash;
            }
        }

        public class GroupInfo2
        {
            public List<string> Sections { get; set; }
            public List<string> Stories { get; set; }
            public List<double> Length { get; set; }
            public int CountOfMembers { get; set; }

            public GroupInfo2(List<string> sections, List<string> stories, List<double> lengths, int countOfMembers)
            {
                Sections = sections;
                Stories = stories;
                Length = lengths;
                CountOfMembers = countOfMembers;
            }

            // Override Equals and GetHashCode for proper dictionary key comparison
            public override bool Equals(object obj)
            {
                if (obj == null || GetType() != obj.GetType())
                    return false;

                GroupInfo2 other = (GroupInfo2)obj;

                // Check if sections and stories are equal
                if (!Sections.SequenceEqual(other.Sections) || !Stories.SequenceEqual(other.Stories))
                    return false;

                // Check if lengths are approximately equal within tolerance
                if (Length.Count != other.Length.Count)
                    return false;

                for (int i = 0; i < Length.Count; i++)
                {
                    if (Math.Abs(Length[i] - other.Length[i]) > 1000)
                        return false;
                }

                // Check if count of members is equal
                return CountOfMembers == other.CountOfMembers;
            }

            public override int GetHashCode()
            {
                unchecked
                {
                    int hash = 17;

                    // Hash sections
                    hash = hash * 31 + Sections.Aggregate(0, (acc, x) => acc + x.GetHashCode());

                    // Hash stories
                    hash = hash * 31 + Stories.Aggregate(0, (acc, x) => acc + x.GetHashCode());

                    // Hash lengths within tolerance
                    foreach (var len in Length)
                    {
                        long lengthHash = BitConverter.DoubleToInt64Bits(Math.Round(len / 1000) * 1000); // Round to tolerance
                        hash = hash * 31 + (int)(lengthHash ^ (lengthHash >> 32));
                    }

                    // Hash count of members
                    hash = hash * 31 + CountOfMembers;

                    return hash;
                }
            }
            public static int HasSameValues(List<GroupInfo2> currentList, List<string> sections, List<string> stories, List<double> lengths, int countOfMembers, double tolerance)
            {
                foreach (var item in currentList)
                {
                    try
                    {
                        if ((item.Sections.Count != sections.Count) || (item.Stories.Count != stories.Count) || (item.CountOfMembers != countOfMembers))
                        {

                            return -1;

                        }

                        foreach (var sec in item.Sections)
                        {
                            int index = item.Sections.IndexOf(sec);
                            if (sec != sections[index])
                            {
                                return -1;
                            }

                        }

                        foreach (var len in item.Length)
                        {
                            int index = item.Length.IndexOf(len);
                            if (Math.Abs(len - lengths[index]) > tolerance)
                            {
                                return -1;
                            }

                        }


                        foreach (var story in item.Stories)
                        {
                            int index = item.Stories.IndexOf(story);
                            if (story != stories[index])
                            {
                                return -1;
                            }

                        }

                        return currentList.IndexOf(item);

                    }

                    catch
                    {
                        return -1;
                    }

                }
                return -1;

            }
        }


        public class ParentGroupInfo
        {
            public List<string> Sections { get; set; }
            public List<string> Stories { get; set; }
            public int CountOfMembers { get; set; }

            public ParentGroupInfo(List<string> sections, List<string> stories, int countOfMembers)
            {
                Sections = sections;
                Stories = stories;
                CountOfMembers = countOfMembers;
            }

            // Override Equals and GetHashCode for proper dictionary key comparison
            public override bool Equals(object obj)
            {
                if (obj == null || GetType() != obj.GetType())
                    return false;

                ParentGroupInfo other = (ParentGroupInfo)obj;

                // Check if sections and stories are equal
                if (!Sections.SequenceEqual(other.Sections) || !Stories.SequenceEqual(other.Stories))
                    return false;

                // Check if count of members is equal
                return CountOfMembers == other.CountOfMembers;
            }

            public override int GetHashCode()
            {
                unchecked
                {
                    int hash = 17;

                    // Hash sections
                    hash = hash * 31 + Sections.Aggregate(0, (acc, x) => acc + x.GetHashCode());

                    // Hash stories
                    hash = hash * 31 + Stories.Aggregate(0, (acc, x) => acc + x.GetHashCode());

                    // Hash count of members
                    hash = hash * 31 + CountOfMembers;

                    return hash;
                }
            }
        }


        public class LengthGroupInfo
        {
            public List<double> Length { get; set; }

            public LengthGroupInfo(List<double> lengths)
            {
                Length = lengths;
            }

            // Override Equals and GetHashCode for proper dictionary key comparison
            public override bool Equals(object obj)
            {
                if (obj == null || GetType() != obj.GetType())
                    return false;

                LengthGroupInfo other = (LengthGroupInfo)obj;

                // Check if lengths are approximately equal within tolerance
                if (Length.Count != other.Length.Count)
                    return false;

                for (int i = 0; i < Length.Count; i++)
                {
                    if (Math.Abs(Length[i] - other.Length[i]) > 1000)
                        return false;
                }

                // Check if count of members is equal
                return true;
            }

            public override int GetHashCode()
            {
                // Define a hash code seed
                int hash = 17;

                // Use a fixed tolerance value for hash computation
                double tolerance = 1000.0;

                foreach (double length in Length)
                {
                    // Normalize the length to the tolerance
                    double normalizedLength = Math.Round(length / tolerance);
                    hash = hash * 31 + normalizedLength.GetHashCode();
                }

                return hash;
            }
        }
        public class LengthGroupInfoWithoutLastLengthConsideration
        {
            public List<double> Length { get; set; }

            public LengthGroupInfoWithoutLastLengthConsideration(List<double> lengths)
            {
                Length = lengths;
            }

            // Override Equals and GetHashCode for proper dictionary key comparison
            public override bool Equals(object obj)
            {
                if (obj == null || GetType() != obj.GetType())
                    return false;

                LengthGroupInfoWithoutLastLengthConsideration other = (LengthGroupInfoWithoutLastLengthConsideration)obj;

                // Check if lengths are approximately equal within tolerance
                if (Length.Count != other.Length.Count)
                    return false;

                for (int i = 0; i < Length.Count; i++)
                {
                    if (i == Length.Count - 1)
                        continue;
                    if (Math.Abs(Length[i] - other.Length[i]) > 1000)
                        return false;
                }

                // Check if count of members is equal
                return true;
            }

            public override int GetHashCode()
            {
                // Define a hash code seed
                int hash = 17;

                // Use a fixed tolerance value for hash computation
                double tolerance = 1000.0;

                // Loop through the Length list, excluding the last element
                for (int i = 0; i < Length.Count - 1; i++)
                {
                    double normalizedLength = Math.Round(Length[i] / tolerance);
                    hash = hash * 31 + normalizedLength.GetHashCode();
                }

                return hash;
            }
        }


        #region NewForDammamSKiPark

        public class SectionsOnly
        {
            public List<string> Sections { get; set; }
            public int CountOfMembers { get; set; }

            public SectionsOnly(List<string> sections, int countOfMembers)
            {
                Sections = sections;
                CountOfMembers = countOfMembers;
            }

            // Override Equals and GetHashCode for proper dictionary key comparison
            public override bool Equals(object obj)
            {
                if (obj == null || GetType() != obj.GetType())
                    return false;

                SectionsOnly other = (SectionsOnly)obj;


                // Check if count of members is equal
                return CountOfMembers == other.CountOfMembers;
            }

            public override int GetHashCode()
            {
                unchecked
                {
                    int hash = 17;

                    // Hash sections
                    hash = hash * 31 + Sections.Aggregate(0, (acc, x) => acc + x.GetHashCode());

                    // Hash count of members
                    hash = hash * 31 + CountOfMembers;

                    return hash;
                }
            }
        }


        #endregion

        #region GrouppingAlgorithms



        //Main with consideration to all segmentsLengths
        public static void AssignColumnsMarks5(ref List<Frame> etabsFrames)
        {
            double tolerance = 1000;
            // Filter out only the columns
            List<Frame> columns = etabsFrames.Where(x => x.MemberType == MemberType.Column).ToList();

            // Group columns based on their start coordinates (X, Y)
            var columnsGroupedBasedOnStartXandY = columns.GroupBy(x => new { x.StartPoint.X, x.StartPoint.Y });

            // Dictionary to store the grouped groups
            var groupedGroups = new Dictionary<ParentGroupInfo, List<List<Frame>>>();


            // Process each group based on start coordinates
            foreach (var group in columnsGroupedBasedOnStartXandY)
            {
                // Order frames within the group by their start Z coordinate
                var framesInGroup = group.OrderBy(x => x.StartPoint.Z).ToList();

                // Extract sections, startZ, endZ, and countOfMembers for the group

                List<string> sections = framesInGroup.GroupBy(y => y.ExistSec).Select(x => x.Key).ToList();
                List<string> distinctStories = framesInGroup.GroupBy(y => y.ExistSec)
                                                            .SelectMany(g => g.Select(e => e.Story).Distinct())
                                                            .Distinct()
                                                            .ToList();

                int countOfMembers = framesInGroup.GroupBy(y => y.ExistSec).ToList().Count;

                // Create a GroupInfo object
                var groupInfo = new ParentGroupInfo(sections, distinctStories, countOfMembers);

                // Add the group to the dictionary
                if (!groupedGroups.ContainsKey(groupInfo))
                {
                    groupedGroups[groupInfo] = new List<List<Frame>>();
                }

                groupedGroups[groupInfo].Add(framesInGroup);
            }






            int i = 1;
            // Loop through each entry in the groupedGroups dictionary
            foreach (var entry in groupedGroups)
            {
                // Get the GroupInfo key and the list of groups
                ParentGroupInfo groupInfo = entry.Key;
                List<List<Frame>> InstancesOfEachGroup = entry.Value;


                // Dictionary to store the grouped groups
                var groupedSubGroups = new Dictionary<LengthGroupInfo, List<List<Frame>>>();

                // Loop through each group in the list of groups
                foreach (var instance in InstancesOfEachGroup)
                {
                    List<double> lengths = instance.GroupBy(y => y.ExistSec).Select(x => x.Sum(e => e.Length)).ToList();
                    LengthGroupInfo subGroup = new LengthGroupInfo(lengths);

                    // Add the group to the dictionary
                    if (!groupedSubGroups.ContainsKey(subGroup))
                    {
                        groupedSubGroups[subGroup] = new List<List<Frame>>();
                    }

                    groupedSubGroups[subGroup].Add(instance);
                }




                // Loop through each subGroup in the list of groups

                if (groupedSubGroups.Keys.Count == 1)
                {

                    // Loop through each group in the list of groups
                    foreach (var instance in groupedSubGroups)
                    {
                        List<List<Frame>> InstancesOfEachSubGroup = entry.Value;

                        foreach (var item in InstancesOfEachSubGroup.SelectMany(x => x))
                        {
                            item.Mark = "C" + i.ToString();

                        }
                    }

                    // Additional processing for each entry in the dictionary can be done here

                }

                else
                {
                    char a = 'A';

                    foreach (var item in groupedSubGroups)
                    {
                        List<List<Frame>> InstancesOfEachSubGroup = item.Value;

                        foreach (var instance in InstancesOfEachSubGroup)
                        {

                            foreach (var item1 in instance)
                            {
                                item1.Mark = "C" + i.ToString() + a;

                            }
                        }

                        a++;

                    }
                }


                i++;

            }


        }


        //Main but withoutConsdireingLastSegmentsLength
        public static void AssignColumnsMarks6(ref List<Frame> etabsFrames)
        {
            double tolerance = 1000;
            // Filter out only the columns
            List<Frame> columns = etabsFrames.Where(x => x.MemberType == MemberType.Column).ToList();

            // Group columns based on their start coordinates (X, Y)
            var columnsGroupedBasedOnStartXandY = columns.GroupBy(x => new { x.StartPoint.X, x.StartPoint.Y });

            // Dictionary to store the grouped groups
            var groupedGroups = new Dictionary<SectionsOnly, List<List<Frame>>>();


            // Process each group based on start coordinates
            foreach (var group in columnsGroupedBasedOnStartXandY)
            {
                // Order frames within the group by their start Z coordinate
                var framesInGroup = group.OrderBy(x => x.StartPoint.Z).ToList();

                // Extract sections, startZ, endZ, and countOfMembers for the group

                List<string> sections = framesInGroup.GroupBy(y => y.ExistSec).Select(x => x.Key).ToList();
                List<string> distinctStories = framesInGroup.GroupBy(y => y.ExistSec)
                                                            .SelectMany(g => g.Select(e => e.Story).Distinct())
                                                            .Distinct()
                                                            .ToList();

                int countOfMembers = framesInGroup.GroupBy(y => y.ExistSec).ToList().Count;

                // Create a GroupInfo object
                var groupInfo = new SectionsOnly(sections, countOfMembers);

                // Add the group to the dictionary
                if (!groupedGroups.ContainsKey(groupInfo))
                {
                    groupedGroups[groupInfo] = new List<List<Frame>>();
                }

                groupedGroups[groupInfo].Add(framesInGroup);
            }






            int i = 1;
            // Loop through each entry in the groupedGroups dictionary
            foreach (var entry in groupedGroups)
            {
                // Get the GroupInfo key and the list of groups
                SectionsOnly groupInfo = entry.Key;
                List<List<Frame>> InstancesOfEachGroup = entry.Value;


                // Dictionary to store the grouped groups
                var groupedSubGroups = new Dictionary<LengthGroupInfoWithoutLastLengthConsideration, List<List<Frame>>>();

                // Loop through each group in the list of groups
                foreach (var instance in InstancesOfEachGroup)
                {
                    List<double> lengths = instance.GroupBy(y => y.ExistSec).Select(x => x.Sum(e => e.Length)).ToList();
                    LengthGroupInfoWithoutLastLengthConsideration subGroup = new LengthGroupInfoWithoutLastLengthConsideration(lengths);

                    // Add the group to the dictionary
                    if (!groupedSubGroups.ContainsKey(subGroup))
                    {
                        groupedSubGroups[subGroup] = new List<List<Frame>>();
                    }

                    groupedSubGroups[subGroup].Add(instance);
                }




                // Loop through each subGroup in the list of groups

                if (groupedSubGroups.Keys.Count == 1)
                {

                    // Loop through each group in the list of groups
                    foreach (var instance in groupedSubGroups)
                    {
                        List<List<Frame>> InstancesOfEachSubGroup = entry.Value;

                        foreach (var item in InstancesOfEachSubGroup.SelectMany(x => x))
                        {
                            item.Mark = "C" + i.ToString();

                        }
                    }

                    // Additional processing for each entry in the dictionary can be done here

                }

                else
                {
                    char a = 'A';

                    foreach (var item in groupedSubGroups)
                    {
                        List<List<Frame>> InstancesOfEachSubGroup = item.Value;

                        foreach (var instance in InstancesOfEachSubGroup)
                        {

                            foreach (var item1 in instance)
                            {
                                item1.Mark = "C" + i.ToString() + a;

                            }
                        }

                        a++;

                    }
                }


                i++;

            }


        }


        //Main but withoutConsdireinganyLengthsAtAll
        public static void AssignColumnsMarks7(ref List<Frame> etabsFrames)
        {
            double tolerance = 1000;
            // Filter out only the columns
            List<Frame> columns = etabsFrames.Where(x => x.MemberType == MemberType.Column).ToList();

            // Group columns based on their start coordinates (X, Y)
            var columnsGroupedBasedOnStartXandY = columns.GroupBy(x => new { x.StartPoint.X, x.StartPoint.Y });

            // Dictionary to store the grouped groups
            var groupedGroups = new Dictionary<SectionsOnly, List<List<Frame>>>();


            // Process each group based on start coordinates
            foreach (var group in columnsGroupedBasedOnStartXandY)
            {
                // Order frames within the group by their start Z coordinate
                var framesInGroup = group.OrderBy(x => x.StartPoint.Z).ToList();

                // Extract sections, startZ, endZ, and countOfMembers for the group

                List<string> sections = framesInGroup.GroupBy(y => y.ExistSec).Select(x => x.Key).ToList();
                List<string> distinctStories = framesInGroup.GroupBy(y => y.ExistSec)
                                                            .SelectMany(g => g.Select(e => e.Story).Distinct())
                                                            .Distinct()
                                                            .ToList();

                int countOfMembers = framesInGroup.GroupBy(y => y.ExistSec).ToList().Count;

                // Create a GroupInfo object
                var groupInfo = new SectionsOnly(sections, countOfMembers);

                // Add the group to the dictionary
                if (!groupedGroups.ContainsKey(groupInfo))
                {
                    groupedGroups[groupInfo] = new List<List<Frame>>();
                }

                groupedGroups[groupInfo].Add(framesInGroup);
            }


            int i = 1;
            // Loop through each entry in the groupedGroups dictionary
            foreach (var entry in groupedGroups)
            {
                // Get the GroupInfo key and the list of groups
                SectionsOnly groupInfo = entry.Key;
                List<List<Frame>> InstancesOfEachGroup = entry.Value;


                // Loop through each group in the list of groups
                foreach (var instance in InstancesOfEachGroup)
                {
                    List<List<Frame>> InstancesOfEachSubGroup = entry.Value;

                    foreach (var item in InstancesOfEachSubGroup.SelectMany(x => x))
                    {
                        item.Mark = "C" + i.ToString();

                    }
                }

                i++;

            }


        }



        #region OLD
        public static void AssignColumnsMarks()
        {
            List<Frame> etabsFrames = getSapFrames();
            // Filter out only the columns
            List<Frame> columns = etabsFrames.Where(x => x.MemberType == MemberType.Column).ToList();

            // Group columns based on their start coordinates (X, Y)
            var columnsGroupedBasedOnStartXandY = columns.GroupBy(x => new { x.StartPoint.X, x.StartPoint.Y });

            // Dictionary to store unique group objects and their corresponding group marks
            Dictionary<object, int> groupDictionary = new Dictionary<object, int>();

            int groupIndex = 1;

            // Process each group based on start coordinates
            foreach (var group in columnsGroupedBasedOnStartXandY)
            {
                // Order frames within the group by their start Z coordinate
                var framesInGroup = group.OrderBy(x => x.StartPoint.Z).ToList();

                // Extract sections, startZ, endZ, and countOfMembers for the group
                List<string> sections = framesInGroup.Select(x => x.Sec).ToList();
                List<double> startZ = framesInGroup.Select(x => x.StartPoint.Z).ToList();
                List<double> endZ = framesInGroup.Select(x => x.EndPoint.Z).ToList();
                int countOfMembers = framesInGroup.Count;

                // Create a unique object representing the group
                var groupObject = new { sections, startZ, endZ, countOfMembers };

                // Check if the group object already exists in the dictionary
                if (!groupDictionary.ContainsKey(groupObject))
                {
                    // Assign a new group index and add it to the dictionary
                    groupDictionary.Add(groupObject, groupIndex);
                    groupIndex++;
                }
            }

            // Assign group marks to each column based on the group dictionary
            foreach (var group in columnsGroupedBasedOnStartXandY)
            {
                var framesInGroup = group.ToList();

                // Order frames within the group by their start Z coordinate
                var framesOrderedByStartZ = framesInGroup.OrderBy(x => x.StartPoint.Z).ToList();

                // Extract sections, startZ, endZ, and countOfMembers for the group
                List<string> sections = framesOrderedByStartZ.Select(x => x.Sec).ToList();
                List<double> startZ = framesOrderedByStartZ.Select(x => x.StartPoint.Z).ToList();
                List<double> endZ = framesOrderedByStartZ.Select(x => x.EndPoint.Z).ToList();
                int countOfMembers = framesOrderedByStartZ.Count;

                // Create a unique object representing the group
                var groupObject = new { sections, startZ, endZ, countOfMembers };

                // Retrieve the group mark from the dictionary
                if (groupDictionary.TryGetValue(groupObject, out int groupMark))
                {
                    // Assign the group mark to each frame in the current group
                    foreach (var frame in framesInGroup)
                    {
                        frame.Group = groupMark.ToString();
                    }
                }
            }
        }
        public static void AssignColumnsMarks2(ref List<Frame> etabsFrames)
        {
            // Filter out only the columns
            List<Frame> columns = etabsFrames.Where(x => x.MemberType == MemberType.Column).ToList();

            // Group columns based on their start coordinates (X, Y)
            var columnsGroupedBasedOnStartXandY = columns.GroupBy(x => new { x.StartPoint.X, x.StartPoint.Y });

            // Dictionary to store the grouped groups
            var groupedGroups = new Dictionary<GroupInfo, List<List<Frame>>>();


            // Process each group based on start coordinates
            foreach (var group in columnsGroupedBasedOnStartXandY)
            {
                // Order frames within the group by their start Z coordinate
                var framesInGroup = group.OrderBy(x => x.StartPoint.Z).ToList();

                // Extract sections, startZ, endZ, and countOfMembers for the group
                List<string> sections = framesInGroup.Select(x => x.Sec).ToList();
                List<double> startZ = framesInGroup.Select(x => x.StartPoint.Z).ToList();
                List<double> endZ = framesInGroup.Select(x => x.EndPoint.Z).ToList();
                int countOfMembers = framesInGroup.Count;

                // Create a GroupInfo object
                var groupInfo = new GroupInfo(sections, startZ, endZ, countOfMembers);

                // Add the group to the dictionary
                if (!groupedGroups.ContainsKey(groupInfo))
                {
                    groupedGroups[groupInfo] = new List<List<Frame>>();
                }

                groupedGroups[groupInfo].Add(framesInGroup);
            }

            int i = 1;

            // Loop through each entry in the groupedGroups dictionary
            foreach (var entry in groupedGroups)
            {
                // Get the GroupInfo key and the list of groups
                GroupInfo groupInfo = entry.Key;
                List<List<Frame>> InstancesOfEachGroup = entry.Value;

                // Access the attributes from GroupInfo
                List<string> sections = groupInfo.Sections;
                List<double> startZ = groupInfo.StartZ;
                List<double> endZ = groupInfo.EndZ;
                int countOfMembers = groupInfo.CountOfMembers;

                // Loop through each group in the list of groups
                foreach (var instance in InstancesOfEachGroup)
                {
                    foreach (var item in instance)
                    {
                        item.Mark = "C" + i.ToString();
                    }
                }

                // Additional processing for each entry in the dictionary can be done here
                i++;
            }

        }
        public static void AssignColumnsMarks3(ref List<Frame> etabsFrames)
        {
            double tolerance = 1000;
            // Filter out only the columns
            List<Frame> columns = etabsFrames.Where(x => x.MemberType == MemberType.Column).ToList();

            // Group columns based on their start coordinates (X, Y)
            var columnsGroupedBasedOnStartXandY = columns.GroupBy(x => new { x.StartPoint.X, x.StartPoint.Y });

            // Dictionary to store the grouped groups
            var groupedGroups = new Dictionary<GroupInfo2, List<List<Frame>>>();


            // Process each group based on start coordinates
            foreach (var group in columnsGroupedBasedOnStartXandY)
            {
                // Order frames within the group by their start Z coordinate
                var framesInGroup = group.OrderBy(x => x.StartPoint.Z).ToList();

                // Extract sections, startZ, endZ, and countOfMembers for the group
                List<string> sections = framesInGroup.Select(x => x.Sec).ToList();
                List<string> stories = framesInGroup.Select(x => x.Story).ToList();
                List<double> lengths = framesInGroup.Select(x => Math.Round(x.Length)).ToList();
                int countOfMembers = framesInGroup.Count;

                // Create a GroupInfo object
                var groupInfo = new GroupInfo2(sections, stories, lengths, countOfMembers);

                // Add the group to the dictionary
                if (!groupedGroups.ContainsKey(groupInfo))
                {
                    groupedGroups[groupInfo] = new List<List<Frame>>();
                }

                groupedGroups[groupInfo].Add(framesInGroup);
            }

            int i = 1;

            // Loop through each entry in the groupedGroups dictionary
            foreach (var entry in groupedGroups)
            {
                // Get the GroupInfo key and the list of groups
                GroupInfo2 groupInfo = entry.Key;
                List<List<Frame>> InstancesOfEachGroup = entry.Value;

                // Loop through each group in the list of groups
                foreach (var instance in InstancesOfEachGroup)
                {
                    foreach (var item in instance)
                    {
                        item.Mark = "C" + i.ToString();
                    }
                }

                // Additional processing for each entry in the dictionary can be done here
                i++;
            }

        }
        public static void AssignColumnsMarks4(ref List<Frame> etabsFrames)
        {
            double tolerance = 1000;
            // Filter out only the columns
            List<Frame> columns = etabsFrames.Where(x => x.MemberType == MemberType.Column).ToList();

            // Group columns based on their start coordinates (X, Y)
            var columnsGroupedBasedOnStartXandY = columns.GroupBy(x => new { x.StartPoint.X, x.StartPoint.Y });

            // Dictionary to store the grouped groups
            var groupedGroups = new Dictionary<GroupInfo2, List<List<Frame>>>();


            // Process each group based on start coordinates
            foreach (var group in columnsGroupedBasedOnStartXandY)
            {
                // Order frames within the group by their start Z coordinate
                var framesInGroup = group.OrderBy(x => x.StartPoint.Z).ToList();

                // Extract sections, startZ, endZ, and countOfMembers for the group

                List<string> sections = framesInGroup.GroupBy(y => y.ExistSec).Select(x => x.Key).ToList();
                List<double> lengths = framesInGroup.GroupBy(y => y.ExistSec).Select(x => x.Sum(e => e.Length)).ToList();
                List<string> distinctStories = framesInGroup.GroupBy(y => y.ExistSec)
                                                            .SelectMany(g => g.Select(e => e.Story).Distinct())
                                                            .Distinct()
                                                            .ToList();

                int countOfMembers = framesInGroup.GroupBy(y => y.ExistSec).ToList().Count;

                // Create a GroupInfo object
                var groupInfo = new GroupInfo2(sections, distinctStories, lengths, countOfMembers);

                // Add the group to the dictionary
                if (!groupedGroups.ContainsKey(groupInfo))
                {
                    groupedGroups[groupInfo] = new List<List<Frame>>();
                }

                groupedGroups[groupInfo].Add(framesInGroup);
            }




            int i = 1;

            // Loop through each entry in the groupedGroups dictionary
            foreach (var entry in groupedGroups)
            {
                // Get the GroupInfo key and the list of groups
                GroupInfo2 groupInfo = entry.Key;
                List<List<Frame>> InstancesOfEachGroup = entry.Value;

                // Loop through each group in the list of groups
                foreach (var instance in InstancesOfEachGroup)
                {
                    foreach (var item in instance)
                    {
                        item.Mark = "C" + i.ToString();
                    }
                }

                // Additional processing for each entry in the dictionary can be done here
                i++;
            }

        }
        #endregion

        #endregion

        public static void AssignColumnsAndBeamsDimentions(ref List<Frame> etabsFrames)
        {
            #region defintions 
            string FileName = null;
            string MatProp = null;
            double t3 = 0;
            double t2 = 0;
            int Color = 0;
            string Notes = null;
            string GUID = null;
            #endregion


            cSapModel myEtabs = Get_sap_mode();
            List<string> sections = etabsFrames.Select(x => x.Sec).ToList();
            sections = sections.Distinct().ToList();

            eFramePropType shape = eFramePropType.Angle;

            foreach (var sec in sections)
            {
                myEtabs.PropFrame.GetTypeOAPI(sec, ref shape);

                if (shape == eFramePropType.Rectangular)
                {
                    int ret = myEtabs.PropFrame.GetRectangle(sec, ref FileName, ref MatProp, ref t3, ref t2, ref Color, ref Notes, ref GUID);


                    foreach (var item in etabsFrames.Where(x => x.Sec == sec))
                    {
                        item.Depth = t3;
                        item.Width = t2;
                    }
                }

                else if (shape == eFramePropType.Circle)
                {
                    int ret = myEtabs.PropFrame.GetCircle(sec, ref FileName, ref MatProp, ref t3, ref Color, ref Notes, ref GUID);


                    foreach (var item in etabsFrames.Where(x => x.Sec == sec))
                    {
                        item.Diameter = t3;
                    }
                }


            }



        }
        #endregion
        #region CreateExtrudedShapes
        public static void addColumnRigidZone(List<Frame> frames)
        {

            #region definitions
            string flieName = null;
            string matProp = null;
            double depth = 0;
            double width = 0;
            int color = 1;
            string notes = null;
            string guid = null;
            Rhino.RhinoDoc doc = Rhino.RhinoDoc.ActiveDoc;


            #endregion
            List<Rectangle3d> resultRec = new List<Rectangle3d>();
            List<Circle> resultCircle = new List<Circle>();
            List<double> rotations = new List<double>();


            ETABSv1.cSapModel etabsModel = EtabsUtils.Get_sap_mode();

            List<Frame> targetRectFrames = frames.Where(x => x.MemberType == MemberType.Column && x.ColumnsRebarDesign.Pattern == "R").ToList();
            List<Frame> targetCircFrames = frames.Where(x => x.MemberType == MemberType.Column && x.ColumnsRebarDesign.Pattern == "C").ToList();

            foreach (var frame in targetRectFrames)
            {
                Point pt;
                pt = frame.StartPoint;

                Point3d origin = new Point3d(pt.X, pt.Y, pt.Z);
                Rhino.Geometry.Plane pl = new Rhino.Geometry.Plane(origin, Vector3d.ZAxis);

                Point3d pt1 = new Point3d(pt.X - frame.Width / 2, pt.Y - frame.Depth / 2, pt.Z);
                Point3d pt2 = new Point3d(pt.X + frame.Width / 2, pt.Y - frame.Depth / 2, pt.Z);
                Point3d pt3 = new Point3d(pt.X + frame.Width / 2, pt.Y + frame.Depth / 2, pt.Z);
                Point3d pt4 = new Point3d(pt.X - frame.Width / 2, pt.Y + frame.Depth / 2, pt.Z);

                Rectangle3d rec = new Rectangle3d(pl, pt1, pt3);
                // Create a transformation for rotation
                double angleInRadians = RhinoMath.ToRadians(frame.Beta);
                Rhino.Geometry.Transform rotationTransform = Rhino.Geometry.Transform.Rotation(angleInRadians, rec.Center);

                // Apply the rotation transformation to the rectangle
                Rectangle3d rotatedRectangle = rec;
                rotatedRectangle.Transform(rotationTransform);


                // Create a planar surface from the rectangle
                NurbsSurface surface2 = NurbsSurface.CreateFromCorners(rotatedRectangle.Corner(0), rotatedRectangle.Corner(1), rotatedRectangle.Corner(2), rotatedRectangle.Corner(3));

                // Define the extrusion vector (height)
                Vector3d extrusionVector = new Vector3d(0, 0, frame.EndPoint.Z - frame.StartPoint.Z); // Extrude 20 units in Z direction


                Rhino.Geometry.Curve curve = rotatedRectangle.ToNurbsCurve();
                Rhino.Geometry.Surface extrusions = Rhino.Geometry.Surface.CreateExtrusion(curve, extrusionVector);




                #region Attributes
                ObjectAttributes attr = new ObjectAttributes();

                attr.SetUserString("Section", frame.Sec);
                attr.SetUserString("Story", frame.Story);
                attr.SetUserString("Shape", frame.Shape);
                attr.SetUserString("StartPointId", frame.StartPoint.Id);
                attr.SetUserString("EndPointId", frame.EndPoint.Id);
                attr.SetUserString("Angle", frame.Beta.ToString());
                attr.SetUserString("Length", frame.Length.ToString());
                attr.SetUserString("ID", frame.ID);
                attr.SetUserString("MainGroup", frame.MainGroup);
                attr.SetUserString("SubGroup", frame.SubGroup);
                attr.SetUserString("ColumnMark", frame.Mark);
                attr.SetUserString("Deignation", ColumnRebarDesignData.getDesignationMark(frame));
                attr.SetUserString("ColumnType", frame.ColumnsRebarDesign.Pattern);

                #endregion
                //doc.Objects.AddRectangle(rotatedRectangle);
                doc.Objects.AddSurface(extrusions, attr);
            }


            foreach (var frame in targetCircFrames)
            {
                Point pt;
                pt = frame.StartPoint;


                Point3d origin = new Point3d(pt.X, pt.Y, pt.Z);
                Rhino.Geometry.Plane pl = new Rhino.Geometry.Plane(origin, Vector3d.ZAxis);

                double diameter = 0;


                Circle circle = new Circle(origin, frame.Diameter / 2);
                // Define the extrusion vector (height)
                Vector3d extrusionVector = new Vector3d(0, 0, frame.EndPoint.Z - frame.StartPoint.Z); // Extrude 20 units in Z direction
                Rhino.Geometry.Curve curve = circle.ToNurbsCurve();
                Rhino.Geometry.Surface extrusions = Rhino.Geometry.Surface.CreateExtrusion(curve, extrusionVector);


                #region Attributes
                ObjectAttributes attr = new ObjectAttributes();

                attr.SetUserString("Section", frame.Sec);
                attr.SetUserString("Story", frame.Story);
                attr.SetUserString("Shape", frame.Shape);
                attr.SetUserString("StartPointId", frame.StartPoint.Id);
                attr.SetUserString("EndPointId", frame.EndPoint.Id);
                attr.SetUserString("Angle", frame.Beta.ToString());
                attr.SetUserString("Length", frame.Length.ToString());
                attr.SetUserString("ID", frame.ID);
                attr.SetUserString("MainGroup", frame.MainGroup);
                attr.SetUserString("SubGroup", frame.SubGroup);
                attr.SetUserString("ColumnMark", frame.Mark);
                attr.SetUserString("Deignation", ColumnRebarDesignData.getDesignationMark(frame));
                attr.SetUserString("ColumnType", frame.ColumnsRebarDesign.Pattern);

                #endregion
                doc.Objects.AddSurface(extrusions, attr);

            }

            // RectColumns = resultRec;
            //  CircColumns = resultCircle;
            // rotationsForRecColumns = rotations;

        }


        //public static void getRigidZoneOfBeams3(string Story, ref List<Polyline> RectColumns)
        //{

        //    #region definitions
        //    string flieName = null;
        //    string matProp = null;
        //    double depth = 0;
        //    double width = 0;
        //    int color = 1;
        //    string notes = null;
        //    string guid = null;



        //    #endregion
        //    List<Polyline> resultRec = new List<Polyline>();
        //    List<Circle> resultCircle = new List<Circle>();



        //    List<Frame> f = EtabsUtils.getSapFrames2();

        //    ETABSv1.cSapModel etabsModel = EtabsUtils.Get_sap_mode();
        //    Dictionary<string, double> StoryElevDic = EtabsUtils.getStoriesElevation(etabsModel);
        //    double Elev = StoryElevDic[Story];


        //    //List<Frame> targetRectFrames = f.Where(x => x.MemberType == MemberType.Beam || x.MemberType == MemberType.VLBracing).Where(x => x.Shape == "Rectangular" && x.Sec != "None").Where(x => x.Story == Story).ToList();
        //    List<Frame> targetRectFrames = f.Where(x => x.MemberType == MemberType.Beam || x.MemberType == MemberType.VLBracing).Where(x => x.Shape == "Rectangular" && x.Sec != "None").Where(x => Math.Abs(x.StartPoint.Z - Elev) <= 150 || Math.Abs(x.EndPoint.Z - Elev) <= 150).ToList();



        //    foreach (var frame in targetRectFrames)
        //    {
        //        Point pt;
        //        if (Math.Abs(frame.StartPoint.Z - Elev) <= 150)
        //        {
        //            pt = frame.StartPoint;
        //        }

        //        else
        //        {
        //            pt = frame.StartPoint;

        //        }



        //        Point3d origin = new Point3d(pt.X, pt.Y, Elev);
        //        Rhino.Geometry.Plane pl = new Rhino.Geometry.Plane(origin, Vector3d.ZAxis);


        //        etabsModel.PropFrame.GetRectangle(frame.Sec, ref flieName, ref matProp, ref depth, ref width, ref color, ref notes, ref guid);
        //        depth = frame.Length;

        //        Line l = new Line(frame.StartPoint.X, frame.StartPoint.Y, Elev, frame.EndPoint.X, frame.EndPoint.Y, Elev);
        //        Polyline Pl = CreatePolylineFromLineAndWidth2(l, width);



        //        resultRec.Add(Pl);
        //    }


        //    RectColumns = resultRec;



        //}


        #endregion

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

        public static void setDataInExcel(List<Frame> sapFrames)
        {

            Microsoft.Office.Interop.Excel.Application Excel = null;
            Microsoft.Office.Interop.Excel._Workbook wb = null;
            Microsoft.Office.Interop.Excel._Worksheet ws = null;
            ExcelUtils.CreateExcelForSAP(ref Excel, ref wb, ref ws, "Etabs");

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
                RevitUtils.setParametersEtabs(instance, newBeam);

                trns.Commit();
            }


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
                RevitUtils.setParametersEtabs(instance, newColumn);

                trns.Commit();
            }

            return id;
        }

        public static List<List<Frame>> assignEtabsFramesGroups()
        {
            cSapModel mySapModel = Get_sap_mode();
            List<Beam> etabsBeams = new List<Beam>();
            List<Column> etabsColumns = new List<Column>();


            List<Frame> etabsFrames = getSapFrames(ref etabsBeams, ref etabsColumns);

            foreach (var f in etabsFrames)
            {

                List<string> MainGroups = f.Groups.Where(x => x.Length >= 4 && x.Substring(1, 3) == "SKI" && x.Count(y => y == '-') == 2).ToList();

                string Maingroup = f.Groups.FirstOrDefault(x => x.Length >= 4 && x.Substring(1, 3) == "SKI" && x.Count(y => y == '-') == 2);

                if (Maingroup != null)
                    f.Group = Maingroup;

                if (MainGroups != null)
                {
                    f.Groups2 = MainGroups;
                }

            }

            var GrouppedetabsFrames = etabsFrames.GroupBy(x => x.Group);



            List<List<Frame>> etabsFramesGroups = new List<List<Frame>>();
            foreach (var Framesgroup in GrouppedetabsFrames)
            {
                List<Frame> frames = new List<Frame>();
                frames = Framesgroup.ToList();
                etabsFramesGroups.Add(frames);

            }


            return etabsFramesGroups;

        }

        public static object extractEtabsBeamsForPres2()
        {
            cSapModel mySapModel = Get_sap_mode();
            List<Beam> etabsBeams = new List<Beam>();
            List<Column> etabsColumns = new List<Column>();


            List<Frame> etabsFrames = getSapFrames(ref etabsBeams, ref etabsColumns);

            foreach (var f in etabsFrames)
            {

                string Maingroup = f.Groups.FirstOrDefault(x => x.Length >= 3 && x.Substring(0, 3) == "SKI" && x.Contains("-"));


                if (Maingroup != null)
                {
                    f.Group = Maingroup;

                    if (Maingroup.Contains("HB"))
                    {
                        f.Lib = "HB";
                    }
                    else if (Maingroup.Contains("SB"))
                    {
                        f.Lib = "SB";

                    }
                    else if (Maingroup.Contains("VB"))
                    {
                        f.Lib = "VB";

                    }
                    else if (Maingroup.Contains("BC"))
                    {
                        f.Lib = "BC";

                    }

                    else if (Maingroup.Contains("SC"))
                    {
                        f.Lib = "SC";

                    }

                    else if (Maingroup.Contains("TC"))
                    {
                        f.Lib = "TC ";

                    }
                    else if (Maingroup.Contains("V"))
                    {
                        f.Lib = "V";

                    }
                    else if (Maingroup.Contains("D"))
                    {
                        f.Lib = "D";

                    }

                    else
                    {
                        f.Lib = "not";
                    }


                }

            }


            var GrouppedetabsFrames = etabsFrames.GroupBy(x => x.Group);


            List<object> objects = new List<object>();
            List<List<Frame>> etabsFramesGroups = new List<List<Frame>>();
            foreach (var Framesgroup in GrouppedetabsFrames)
            {


                List<Frame> frames = new List<Frame>();
                frames = Framesgroup.ToList();
                etabsFramesGroups.Add(frames);

                string sec = string.Join(",", frames.Select(x => x.Sec).Distinct());
                string Group = Framesgroup.Key;
                string MainGroup = frames[0].Lib;
                int count = frames.Count;

                var obj = new { MainGroup = MainGroup, Group = Group, Section = sec, Count = count };
                objects.Add(obj);
            }

            // Create Excel application
            Application excelApp = new Application();
            excelApp.Visible = true;

            // Create a new workbook
            Workbook workbook = excelApp.Workbooks.Add();

            // Add data to the first worksheet
            AddDataToWorksheet(workbook, objects, "BeamEndForces");

            return objects;

        }

        public static void AddDataToWorksheet(Workbook workbook, List<object> records, string sheetName)
        {
            // Create a new worksheet
            Worksheet worksheet = (Worksheet)workbook.Sheets.Add();
            worksheet.Name = sheetName;

            // Add header row with property names
            int rowIndex = 1;
            int colIndex = 1;
            foreach (PropertyInfo property in records[0].GetType().GetProperties())
            {
                worksheet.Range[rowIndex, colIndex].Value = property.Name;
                colIndex++;
            }

            // Add data rows
            rowIndex++;
            foreach (var record in records)
            {
                colIndex = 1;
                foreach (PropertyInfo property in record.GetType().GetProperties())
                {
                    worksheet.Range[rowIndex, colIndex].Value = property.GetValue(record);
                    colIndex++;
                }
                rowIndex++;
            }


            // Add a final row for column sums
            int lastRow = rowIndex;

            // Calculate column sums
            for (int i = 1; i <= colIndex - 1; i++) // Loop through each column
            {
                double sum = 0;
                for (int j = 2; j <= lastRow; j++) // Start from the second row to exclude header
                {
                    double cellValue;
                    if (double.TryParse(worksheet.Range[j, i].Value?.ToString(), out cellValue)) // Try to parse cell value to double
                    {
                        sum += cellValue;
                    }
                }

                if (sum != 0)
                {
                    worksheet.Range[lastRow + 1, i].Value = sum; // Insert sum in the last row

                }
            }





            // Autofit columns
            Excel.Range usedRange = worksheet.UsedRange;
            usedRange.Columns.AutoFit();
        }

        #region  SapGroupUtils
        public static void createGroupinSap(List<string> frameIds, cSapModel mySapModel, string groupName)

        {
            mySapModel.GroupDef.Delete(groupName);
            mySapModel.GroupDef.SetGroup(groupName);


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
            string[] PointIDs = null;
            int noOfPoints = 0;
            int ret;

            #endregion

            ret = mySapModel.PointObj.GetNameList(ref noOfPoints, ref PointIDs);

            if (noOfPoints == 0)
            {
                return new List<string>();
            }

            if (getSelectedPointIDS().Count > 0)
            {
                PointIDs = getSelectedPointIDS().ToArray();
            }


                return PointIDs.ToList();
            
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


        // used for rhino
        public static List<string> addNewPointsNew(ref List<Point> points)
        {
            cSapModel mySapModel = Get_sap_mode();
            List<string> Ids = new List<string>();
            string id = null;

            foreach (var item in points)
            {
                mySapModel.PointObj.AddCartesian(item.X, item.Y, item.Z, ref id, "");
                Ids.Add(id);
                item.Id = id;
            }

            return Ids;
        }

        #endregion

        #region RelatedToSections

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

            if (material == null || material == "")
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

                        ret = mysapmodel.PropFrame.ImportProp(revitFrame.Sec, material, EuroPath, revitFrame.Sec);
                        return revitFrame.Sec;


                    case SectType.BM_HEAA:
                        numeric = new String(revitFrame.RvtName.Where(Char.IsDigit).ToArray());
                        revitFrame.Sec = "HE" + numeric + "AA";
                        ret = mysapmodel.PropFrame.ImportProp(revitFrame.Sec, material, EuroPath, revitFrame.Sec);
                        return revitFrame.Sec;



                    case SectType.BM_HEB:
                        numeric = new String(revitFrame.RvtName.Where(Char.IsDigit).ToArray());
                        revitFrame.Sec = "HE" + numeric + "B";
                        ret = mysapmodel.PropFrame.ImportProp(revitFrame.Sec, material, EuroPath, revitFrame.Sec);
                        return revitFrame.Sec;


                    case SectType.BM_HEM:
                        numeric = new String(revitFrame.RvtName.Where(Char.IsDigit).ToArray());
                        revitFrame.Sec = "HE" + numeric + "M";
                        ret = mysapmodel.PropFrame.ImportProp(revitFrame.Sec, material, EuroPath, revitFrame.Sec);
                        return revitFrame.Sec;


                    case SectType.BM_IPE:
                        numeric = new String(revitFrame.RvtName.Where(Char.IsDigit).ToArray());
                        revitFrame.Sec = "IPE" + numeric;

                        ret = mysapmodel.PropFrame.ImportProp(revitFrame.Sec, material, EuroPath, revitFrame.Sec);
                        return revitFrame.Sec;


                    case SectType.BM_IPEA:
                        numeric = new String(revitFrame.RvtName.Where(Char.IsDigit).ToArray());
                        revitFrame.Sec = "IPEA" + numeric;
                        ret = mysapmodel.PropFrame.ImportProp(revitFrame.Sec, material, EuroPath, revitFrame.Sec);
                        return revitFrame.Sec;


                    case SectType.BM_IPEO:
                        numeric = new String(revitFrame.RvtName.Where(Char.IsDigit).ToArray());
                        revitFrame.Sec = "IPEO" + numeric;
                        ret = mysapmodel.PropFrame.ImportProp(revitFrame.Sec, material, EuroPath, revitFrame.Sec);
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
                        ret = mysapmodel.PropFrame.ImportProp(revitFrame.Sec, material, EuroPath, revitFrame.Sec);
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
                    //ret = mysapmodel.PropFrame.ImportProp(revitFrame.Sec, material, "AISC14.xml", revitFrame.Sec);
                    //return revitFrame.Sec;

                    #endregion

                    #region UB
                    case SectType.BM_UB:
                        cc = revitFrame.RvtName.SkipWhile(c => !char.IsDigit(c)).ToList()[0];
                        lastSpacecharacterIndex = revitFrame.RvtName.IndexOf(cc);
                        dimentionswithX = revitFrame.RvtName.Substring(lastSpacecharacterIndex);
                        revitFrame.Sec = "UKB" + dimentionswithX;
                        ret = mysapmodel.PropFrame.ImportProp(revitFrame.Sec, material, BiritishPath, revitFrame.Sec);
                        return revitFrame.Sec;
                    #endregion

                    #region UC
                    case SectType.BM_UC:
                        cc = revitFrame.RvtName.SkipWhile(c => !char.IsDigit(c)).ToList()[0];
                        lastSpacecharacterIndex = revitFrame.RvtName.IndexOf(cc);
                        dimentionswithX = revitFrame.RvtName.Substring(lastSpacecharacterIndex);
                        revitFrame.Sec = "UKC" + dimentionswithX;
                        ret = mysapmodel.PropFrame.ImportProp(revitFrame.Sec, material, BiritishPath, revitFrame.Sec);
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
                            ret = mysapmodel.PropFrame.ImportProp(revitFrame.Sec, material, EuroPath, revitFrame.Sec);

                            return revitFrame.Sec;
                            //revitFrame.Sec = "L" + firsapimention + "X" + secondDimention + "X" + angleThick;

                        }



                        else if (noOfX == 1)
                        {

                            dimentionswithX = new string(revitFrame.RvtName.SkipWhile(c => !char.IsDigit(c)).ToArray());
                            firsapimention = new string(dimentionswithX.TakeWhile(c => char.IsDigit(c)).ToArray());
                            revitFrame.Sec = "L" + firsapimention + "X" + angleThick;
                            ret = mysapmodel.PropFrame.ImportProp(revitFrame.Sec, material, EuroPath, revitFrame.Sec);

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

                        ret = mysapmodel.PropFrame.ImportProp(revitFrame.Sec, material, EuroPath, revitFrame.Sec);

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

                        ret = mysapmodel.PropFrame.ImportProp(revitFrame.Sec, material, EuroPath, revitFrame.Sec);

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

            if (material == null || material == "")
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
                        ret = mysapmodel.PropFrame.ImportProp(revitFrame.Sec, material, EuroPath, revitFrame.Sec);
                        return revitFrame.Sec;


                    case SectType.COL_HEAA:
                        numeric = new String(revitFrame.RvtName.Where(Char.IsDigit).ToArray());
                        revitFrame.Sec = "HE" + numeric + "AA";
                        ret = mysapmodel.PropFrame.ImportProp(revitFrame.Sec, material, EuroPath, revitFrame.Sec);
                        return revitFrame.Sec;



                    case SectType.COL_HEB:
                        numeric = new String(revitFrame.RvtName.Where(Char.IsDigit).ToArray());
                        revitFrame.Sec = "HE" + numeric + "B";
                        ret = mysapmodel.PropFrame.ImportProp(revitFrame.Sec, material, EuroPath, revitFrame.Sec);
                        return revitFrame.Sec;


                    case SectType.COL_HEM:
                        numeric = new String(revitFrame.RvtName.Where(Char.IsDigit).ToArray());
                        revitFrame.Sec = "HE" + numeric + "M";
                        ret = mysapmodel.PropFrame.ImportProp(revitFrame.Sec, material, EuroPath, revitFrame.Sec);
                        return revitFrame.Sec;


                    case SectType.COL_IPE:
                        numeric = new String(revitFrame.RvtName.Where(Char.IsDigit).ToArray());
                        revitFrame.Sec = "IPE" + numeric;
                        ret = mysapmodel.PropFrame.ImportProp(revitFrame.Sec, material, EuroPath, revitFrame.Sec);
                        return revitFrame.Sec;


                    case SectType.COL_IPEA:
                        numeric = new String(revitFrame.RvtName.Where(Char.IsDigit).ToArray());
                        revitFrame.Sec = "IPEA" + numeric;
                        ret = mysapmodel.PropFrame.ImportProp(revitFrame.Sec, material, EuroPath, revitFrame.Sec);
                        return revitFrame.Sec;


                    case SectType.COL_IPEO:
                        numeric = new String(revitFrame.RvtName.Where(Char.IsDigit).ToArray());
                        revitFrame.Sec = "IPEO" + numeric;
                        ret = mysapmodel.PropFrame.ImportProp(revitFrame.Sec, material, EuroPath, revitFrame.Sec);
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
                        ret = mysapmodel.PropFrame.ImportProp(revitFrame.Sec, material, EuroPath, revitFrame.Sec);
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
                    //ret = mysapmodel.PropFrame.ImportProp(revitFrame.Sec, material, "AISC14.xml", revitFrame.Sec);
                    //return revitFrame.Sec;

                    #endregion

                    #region UB
                    case SectType.COL_UB:
                        cc = revitFrame.RvtName.SkipWhile(c => !char.IsDigit(c)).ToList()[0];
                        lastSpacecharacterIndex = revitFrame.RvtName.IndexOf(cc);
                        dimentionswithX = revitFrame.RvtName.Substring(lastSpacecharacterIndex);
                        revitFrame.Sec = "UKB" + dimentionswithX;
                        ret = mysapmodel.PropFrame.ImportProp(revitFrame.Sec, material, BiritishPath, revitFrame.Sec);
                        return revitFrame.Sec;
                    #endregion

                    #region UC
                    case SectType.COL_UC:
                        cc = revitFrame.RvtName.SkipWhile(c => !char.IsDigit(c)).ToList()[0];
                        lastSpacecharacterIndex = revitFrame.RvtName.IndexOf(cc);
                        dimentionswithX = revitFrame.RvtName.Substring(lastSpacecharacterIndex);
                        revitFrame.Sec = "UKC" + dimentionswithX;
                        ret = mysapmodel.PropFrame.ImportProp(revitFrame.Sec, material, BiritishPath, revitFrame.Sec);
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
                            ret = mysapmodel.PropFrame.ImportProp(revitFrame.Sec, material, EuroPath, revitFrame.Sec);

                            return revitFrame.Sec;
                            //revitFrame.Sec = "L" + firsapimention + "X" + secondDimention + "X" + angleThick;

                        }



                        else if (noOfX == 1)
                        {

                            dimentionswithX = new string(revitFrame.RvtName.SkipWhile(c => !char.IsDigit(c)).ToArray());
                            firsapimention = new string(dimentionswithX.TakeWhile(c => char.IsDigit(c)).ToArray());
                            revitFrame.Sec = "L" + firsapimention + "X" + angleThick;
                            ret = mysapmodel.PropFrame.ImportProp(revitFrame.Sec, material, EuroPath, revitFrame.Sec);

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

                        ret = mysapmodel.PropFrame.ImportProp(revitFrame.Sec, material, EuroPath, revitFrame.Sec);

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

                        ret = mysapmodel.PropFrame.ImportProp(revitFrame.Sec, material, EuroPath, revitFrame.Sec);

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


        public static Dictionary<string, string> createSapSectionForBeams2(List<Frame> revitFrmaes, cSapModel mysapmodel)
        {

            #region definitions 
            string sec;
            string sapSection = null;
            #endregion

            List<object> sapSecNameModifiedSapSecName = ChangeEtabsSectionNames.GetSapSectionsDictionary(mysapmodel);
            List<string> updatedProfilesNameForBus = sapSecNameModifiedSapSecName.Cast<dynamic>().Select(x => x.newSec).Cast<string>().ToList();




            List<Frame> filteredRevitFrames = revitFrmaes.GroupBy(x => x.RvtName).Select(g => g.First()).ToList();

            Dictionary<string, string> dic = new Dictionary<string, string>();

            foreach (Frame revitFrame in filteredRevitFrames)
            {
                try
                {
                    sec = updatedProfilesNameForBus.AsParallel().FirstOrDefault(x => x == revitFrame.RvtName);
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
            string sec2;
            string sapSection = null;
            #endregion


            List<object> sapSecNameModifiedSapSecName = ChangeEtabsSectionNames.GetSapSectionsDictionary(mysapmodel);
            List<string> updatedProfilesNameForBus = sapSecNameModifiedSapSecName.Cast<dynamic>().Select(x => x.newSec).Cast<string>().ToList();
            List<string> ExistingEtabsNames = EtabsUtils.Sections.GetAllNamesAssignedToFramesInSap(mysapmodel);

            List<Frame> filteredRevitFrames = revitFrmaes.GroupBy(x => x.RvtName).Select(g => g.First()).ToList();

            Dictionary<string, string> dic = new Dictionary<string, string>();


            foreach (Frame revitFrame in filteredRevitFrames)
            {
                try
                {
                    sec = updatedProfilesNameForBus.AsParallel().FirstOrDefault(x => x == revitFrame.RvtName);
                    sec2 = ExistingEtabsNames.AsParallel().FirstOrDefault(x => x == revitFrame.RvtName);

                    if (sec != null)
                    {
                        revitFrame.Sec = sec;
                        sapSection = sec;
                        dic.Add(revitFrame.RvtName, sapSecNameModifiedSapSecName.Cast<dynamic>().First(x => x.newSec == sec).SapSec);
                        continue;
                    }

                    if (sec2 != null)
                    {
                        revitFrame.Sec = sec2;
                        sapSection = sec2;
                        dic.Add(revitFrame.RvtName, sapSecNameModifiedSapSecName.Cast<dynamic>().First(x => x.newSec == sec2).SapSec);
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
        #endregion

        #region rhino
        public static void createGroupinSap(List<string> frameIds, string groupName)

        {
            cSapModel mySapModel = Get_sap_mode();
            mySapModel.GroupDef.SetGroup(groupName);


            foreach (var id in frameIds.AsParallel())
            {
                mySapModel.FrameObj.SetGroupAssign(id, groupName);

            }


        }

        public static void createGroupOfPointsinSap(List<string> frameIds, string groupName)

        {
            cSapModel mySapModel = Get_sap_mode();
            mySapModel.GroupDef.SetGroup(groupName);


            foreach (var id in frameIds.AsParallel())
            {
                mySapModel.PointObj.SetGroupAssign(id, groupName);

            }


        }

        public static void createGroupOfAreasinSap(List<string> frameIds, string groupName)

        {
            cSapModel mySapModel = Get_sap_mode();
            mySapModel.GroupDef.SetGroup(groupName);


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

                if (A.Sec == "None")
                {
                    mySapModel.AreaObj.SetOpening(A.ID, true);
                }


                #endregion
            }


            #region Setting Etabs Modification in Floor object Connectivity Table

            // This Name is Obtained from the Excel Sheet Produced from Sap2000 when exporting the data to excel
            int tableVersion_Table_01 = 1;
            string[] filedKeyList_Table_01 = null;
            int numberOfRectords_Table_01 = 0;
            string[] data_Table_01 = null;

            //This Sets all Tables to be Shown when you fetch them from the Database (Atleast in SAP2000)

            mySapModel.DatabaseTables.GetTableForEditingArray(
                TableKey: "Floor Object Connectivity",
                GroupName: "All",
                TableVersion: ref tableVersion_Table_01,
                FieldsKeysIncluded: ref filedKeyList_Table_01,
                NumberRecords: ref numberOfRectords_Table_01,
                TableData: ref data_Table_01);

            List<CSIFLoorConnectivityDB> cSIFloorConnectivityDBs_Table_01 = new List<CSIFLoorConnectivityDB>();

            for (int i = 0; i < data_Table_01.Count(); i += 6)
            {
                CSIFLoorConnectivityDB cSIfloorConnectivity = new CSIFLoorConnectivityDB();
                cSIfloorConnectivity.UniqueName = data_Table_01[i];
                cSIfloorConnectivity.UniquePt1 = data_Table_01[i + 1];
                cSIfloorConnectivity.UniquePt2 = data_Table_01[i + 2];
                cSIfloorConnectivity.UniquePt3 = data_Table_01[i + 3];
                cSIfloorConnectivity.UniquePt4 = data_Table_01[i + 4];
                cSIfloorConnectivity.GUID = data_Table_01[i + 5];

                cSIfloorConnectivity.AllPointIds.Add(cSIfloorConnectivity.UniquePt1);
                cSIfloorConnectivity.AllPointIds.Add(cSIfloorConnectivity.UniquePt2);
                cSIfloorConnectivity.AllPointIds.Add(cSIfloorConnectivity.UniquePt3);
                cSIfloorConnectivity.AllPointIds.Add(cSIfloorConnectivity.UniquePt4);
                cSIfloorConnectivity.AllPointIds.RemoveAll(x => String.IsNullOrWhiteSpace(x));

                cSIFloorConnectivityDBs_Table_01.Add(cSIfloorConnectivity);
            }



            //// Add point ids to parent rows 
            //for (int i = 0; i < cSIFloorConnectivityDBs_Table_01.Count(); i++)
            //{
            //    List<CSIFLoorConnectivityDB> itemsWithEmptyGUID = cSIFloorConnectivityDBs_Table_01.SkipWhile((item, index) => index < i || !string.IsNullOrEmpty(item.GUID))
            //                                                                   .TakeWhile(item => string.IsNullOrEmpty(item.GUID))
            //                                                                   .ToList();
            //    foreach (var item in itemsWithEmptyGUID)
            //    {
            //        cSIFloorConnectivityDBs_Table_01[i].AllPointIds.Add(item.UniquePt1);
            //        cSIFloorConnectivityDBs_Table_01[i].AllPointIds.Add(item.UniquePt2);
            //        cSIFloorConnectivityDBs_Table_01[i].AllPointIds.Add(item.UniquePt3);
            //        cSIFloorConnectivityDBs_Table_01[i].AllPointIds.Add(item.UniquePt4);
            //    }
            //}


            //var groupps = cSIFloorConnectivityDBs_Table_01.GroupBy(x => x.GUID);

            var groupps = cSIFloorConnectivityDBs_Table_01.GroupBy(x => x.UniqueName);
            foreach (var group in groupps)
            {
                var elements = group.ToList();
                if (elements.Count > 1)
                {
                    elements.FirstOrDefault().AllPointIds.AddRange(elements.Skip(1).SelectMany(x => x.AllPointIds));

                }

            }

            #region Trial 01
            // take all parent rows only 
            List<CSIFLoorConnectivityDB> cSIFLoorConnectivityDBs_Filtered = cSIFloorConnectivityDBs_Table_01.Where(item => !String.IsNullOrWhiteSpace(item.GUID)).ToList();

            List<CSIFLoorConnectivityDB> cSIFLoorConnectivityDBs_Written = cSIFLoorConnectivityDBs_Filtered.ToList();

            List<CSIFLoorConnectivityDB> cSIFLoorConnectivityDBs_ToBeDeleted = new List<CSIFLoorConnectivityDB>();


            foreach (var connectivityDB in cSIFLoorConnectivityDBs_Filtered)
            {
                Area area = Areas.Where(item => item.ID == connectivityDB.UniqueName).FirstOrDefault();
                if (connectivityDB.UniqueName == "1337")
                {
                    int aa = 0;
                }

                CSIFLoorConnectivityDB ParentRow = new CSIFLoorConnectivityDB();

                if (area != null)
                {
                    List<string> PointsIDs = area.vertices.Select(x => x.Id).ToList();
                    if (PointsIDs.Count > 4)
                    {
                        cSIFLoorConnectivityDBs_ToBeDeleted.Add(connectivityDB);
                        ParentRow.UniquePt1 = PointsIDs[0];
                        ParentRow.UniquePt2 = PointsIDs[1];
                        ParentRow.UniquePt3 = PointsIDs[2];
                        ParentRow.UniquePt4 = PointsIDs[3];
                        ParentRow.UniqueName = connectivityDB.UniqueName;
                        ParentRow.GUID = connectivityDB.GUID;

                        ParentRow.AllPointIds.Add(ParentRow.UniquePt1);
                        ParentRow.AllPointIds.Add(ParentRow.UniquePt2);
                        ParentRow.AllPointIds.Add(ParentRow.UniquePt3);
                        ParentRow.AllPointIds.Add(ParentRow.UniquePt4);


                        for (int i = 4; i < PointsIDs.Count; i += 4)
                        {
                            CSIFLoorConnectivityDB FollowerRow = new CSIFLoorConnectivityDB();

                            if (PointsIDs.Count - i == 3)
                            {
                                FollowerRow.UniquePt1 = PointsIDs[i];
                                FollowerRow.UniquePt2 = PointsIDs[i + 1];
                                FollowerRow.UniquePt3 = PointsIDs[i + 2];

                                FollowerRow.UniqueName = connectivityDB.UniqueName;
                                FollowerRow.GUID = "";
                                cSIFLoorConnectivityDBs_Written.Add(FollowerRow);

                                ParentRow.AllPointIds.Add(FollowerRow.UniquePt1);
                                ParentRow.AllPointIds.Add(FollowerRow.UniquePt2);
                                ParentRow.AllPointIds.Add(FollowerRow.UniquePt3);
                                break;
                            }
                            else if (PointsIDs.Count - i == 2)
                            {
                                FollowerRow.UniquePt1 = PointsIDs[i];
                                FollowerRow.UniquePt2 = PointsIDs[i + 1];

                                FollowerRow.UniqueName = connectivityDB.UniqueName;
                                FollowerRow.GUID = "";
                                cSIFLoorConnectivityDBs_Written.Add(FollowerRow);


                                ParentRow.AllPointIds.Add(FollowerRow.UniquePt1);
                                ParentRow.AllPointIds.Add(FollowerRow.UniquePt2);
                                break;
                            }
                            else if (PointsIDs.Count - i == 1)
                            {
                                FollowerRow.UniquePt1 = PointsIDs[i];

                                FollowerRow.UniqueName = connectivityDB.UniqueName;
                                FollowerRow.GUID = "";
                                cSIFLoorConnectivityDBs_Written.Add(FollowerRow);

                                ParentRow.AllPointIds.Add(FollowerRow.UniquePt1);
                                break;
                            }
                            else
                            {
                                CSIFLoorConnectivityDB TempPoint = new CSIFLoorConnectivityDB();

                                TempPoint.UniquePt1 = PointsIDs[i];
                                TempPoint.UniquePt2 = PointsIDs[i + 1];
                                TempPoint.UniquePt3 = PointsIDs[i + 2];
                                TempPoint.UniquePt4 = PointsIDs[i + 3];

                                TempPoint.UniqueName = connectivityDB.UniqueName;
                                TempPoint.GUID = "";
                                cSIFLoorConnectivityDBs_Written.Add(TempPoint);

                                ParentRow.AllPointIds.Add(TempPoint.UniquePt1);
                                ParentRow.AllPointIds.Add(TempPoint.UniquePt2);
                                ParentRow.AllPointIds.Add(TempPoint.UniquePt3);
                                ParentRow.AllPointIds.Add(TempPoint.UniquePt4);
                            }
                        }


                        cSIFLoorConnectivityDBs_Written.Add(ParentRow);
                    }
                    else if (PointsIDs.Count == 4)
                    {
                        cSIFLoorConnectivityDBs_ToBeDeleted.Add(connectivityDB);
                        ParentRow.UniquePt1 = PointsIDs[0];
                        ParentRow.UniquePt2 = PointsIDs[1];
                        ParentRow.UniquePt3 = PointsIDs[2];
                        ParentRow.UniquePt4 = PointsIDs[3];
                        ParentRow.UniqueName = connectivityDB.UniqueName;
                        ParentRow.GUID = connectivityDB.GUID;

                        ParentRow.AllPointIds.Add(ParentRow.UniquePt1);
                        ParentRow.AllPointIds.Add(ParentRow.UniquePt2);
                        ParentRow.AllPointIds.Add(ParentRow.UniquePt3);
                        ParentRow.AllPointIds.Add(ParentRow.UniquePt4);

                        cSIFLoorConnectivityDBs_Written.Add(ParentRow);
                    }
                    else if (PointsIDs.Count == 3)
                    {
                        cSIFLoorConnectivityDBs_ToBeDeleted.Add(connectivityDB);
                        ParentRow.UniquePt1 = PointsIDs[0];
                        ParentRow.UniquePt2 = PointsIDs[1];
                        ParentRow.UniquePt3 = PointsIDs[2];
                        ParentRow.UniqueName = connectivityDB.UniqueName;
                        ParentRow.GUID = connectivityDB.GUID;

                        ParentRow.AllPointIds.Add(ParentRow.UniquePt1);
                        ParentRow.AllPointIds.Add(ParentRow.UniquePt2);
                        ParentRow.AllPointIds.Add(ParentRow.UniquePt3);

                        cSIFLoorConnectivityDBs_Written.Add(ParentRow);
                    }
                    else if (PointsIDs.Count < 3)
                    {
                        // Throw an exception for cases with less than 3 points
                        cSIFLoorConnectivityDBs_ToBeDeleted.Add(connectivityDB);
                    }
                }

                else
                {
                    // if Area != null then the Elemenet Doesn't need Upade of Assigns
                    //cSIFLoorConnectivityDBs_Filtered.Add(connectivityDB);
                }
            }

            foreach (var item in cSIFLoorConnectivityDBs_ToBeDeleted)
            {
                cSIFLoorConnectivityDBs_Written.Remove(item);
            }
            #endregion


            List<CSIFLoorConnectivityDB> cSIFLoorConnectivityDBs_Final = new List<CSIFLoorConnectivityDB>();

            foreach (var connectivityDB in cSIFLoorConnectivityDBs_Written)
            {
                CSIFLoorConnectivityDB newObject = new CSIFLoorConnectivityDB();


                if (connectivityDB.AllPointIds.Count() > 4)
                {
                    newObject.UniquePt1 = connectivityDB.AllPointIds[0];
                    newObject.UniquePt2 = connectivityDB.AllPointIds[1];
                    newObject.UniquePt3 = connectivityDB.AllPointIds[2];
                    newObject.UniquePt4 = connectivityDB.AllPointIds[3];
                    newObject.UniqueName = connectivityDB.UniqueName;
                    newObject.GUID = connectivityDB.GUID;
                    newObject.AllPointIds = connectivityDB.AllPointIds;
                    cSIFLoorConnectivityDBs_Final.Add(newObject);
                    // Create new cSILoadPattern objects for the remaining points

                    // Unhandled Excpetion in Case the the Points were 3 not 4 in that next loop
                    for (int i = 4; i < connectivityDB.AllPointIds.Count; i += 4)
                    {
                        CSIFLoorConnectivityDB ChiledObject = new CSIFLoorConnectivityDB();
                        if (connectivityDB.AllPointIds.Count - i == 3)
                        {

                            ChiledObject.UniquePt1 = connectivityDB.AllPointIds[i];
                            ChiledObject.UniquePt2 = connectivityDB.AllPointIds[i + 1];
                            ChiledObject.UniquePt3 = connectivityDB.AllPointIds[i + 2];
                            ChiledObject.UniqueName = connectivityDB.UniqueName;
                            ChiledObject.GUID = "";
                            ChiledObject.AllPointIds = connectivityDB.AllPointIds;
                            cSIFLoorConnectivityDBs_Final.Add(ChiledObject);
                            break;
                        }
                        else if (connectivityDB.AllPointIds.Count - i == 2)
                        {
                            ChiledObject.UniquePt1 = connectivityDB.AllPointIds[i];
                            ChiledObject.UniquePt2 = connectivityDB.AllPointIds[i + 1];
                            ChiledObject.UniqueName = connectivityDB.UniqueName;
                            ChiledObject.GUID = "";
                            ChiledObject.AllPointIds = connectivityDB.AllPointIds;
                            cSIFLoorConnectivityDBs_Final.Add(ChiledObject);
                            break;
                        }
                        else if (connectivityDB.AllPointIds.Count - i == 1)
                        {
                            ChiledObject.UniquePt1 = connectivityDB.AllPointIds[i];
                            ChiledObject.UniqueName = connectivityDB.UniqueName;
                            ChiledObject.GUID = "";
                            cSIFLoorConnectivityDBs_Final.Add(ChiledObject);
                            break;
                        }
                        else
                        {
                            CSIFLoorConnectivityDB TempPoint = new CSIFLoorConnectivityDB();

                            TempPoint.UniquePt1 = connectivityDB.AllPointIds[i];
                            TempPoint.UniquePt2 = connectivityDB.AllPointIds[i + 1];
                            TempPoint.UniquePt3 = connectivityDB.AllPointIds[i + 2];
                            TempPoint.UniquePt4 = connectivityDB.AllPointIds[i + 3];
                            TempPoint.UniqueName = connectivityDB.UniqueName;
                            TempPoint.GUID = "";
                            cSIFLoorConnectivityDBs_Final.Add(TempPoint);
                        }
                    }
                }
                else if (connectivityDB.AllPointIds.Count() == 4)
                {
                    newObject.UniquePt1 = connectivityDB.AllPointIds[0];
                    newObject.UniquePt2 = connectivityDB.AllPointIds[1];
                    newObject.UniquePt3 = connectivityDB.AllPointIds[2];
                    newObject.UniquePt4 = connectivityDB.AllPointIds[3];
                    newObject.UniqueName = connectivityDB.UniqueName;
                    newObject.GUID = connectivityDB.GUID;
                    cSIFLoorConnectivityDBs_Final.Add(newObject);
                }
                else if (connectivityDB.AllPointIds.Count() == 3)
                {
                    newObject.UniquePt1 = connectivityDB.AllPointIds[0];
                    newObject.UniquePt2 = connectivityDB.AllPointIds[1];
                    newObject.UniquePt3 = connectivityDB.AllPointIds[2];
                    newObject.UniqueName = connectivityDB.UniqueName;
                    newObject.GUID = connectivityDB.GUID;
                    cSIFLoorConnectivityDBs_Final.Add(newObject);
                }
            }

            string[] reconstructedStrings_Table_01 = new string[cSIFLoorConnectivityDBs_Final.Count * 6]; // 6 is the number of properties in the class
            int Counter = 0;
            foreach (var cSILoadPattern in cSIFLoorConnectivityDBs_Final)
            {
                reconstructedStrings_Table_01[Counter] = cSILoadPattern.UniqueName.ToString();
                reconstructedStrings_Table_01[Counter + 1] = !String.IsNullOrWhiteSpace(cSILoadPattern.UniquePt1) ? cSILoadPattern.UniquePt1.ToString() : string.Empty;
                reconstructedStrings_Table_01[Counter + 2] = !String.IsNullOrWhiteSpace(cSILoadPattern.UniquePt2) ? cSILoadPattern.UniquePt2.ToString() : string.Empty;
                reconstructedStrings_Table_01[Counter + 3] = !String.IsNullOrWhiteSpace(cSILoadPattern.UniquePt3) ? cSILoadPattern.UniquePt3.ToString() : string.Empty;
                reconstructedStrings_Table_01[Counter + 4] = !String.IsNullOrWhiteSpace(cSILoadPattern.UniquePt4) ? cSILoadPattern.UniquePt4.ToString() : string.Empty;
                reconstructedStrings_Table_01[Counter + 5] = !String.IsNullOrWhiteSpace(cSILoadPattern.GUID) ? cSILoadPattern.GUID.ToString() : string.Empty;
                Counter += 6;
            }

            int ret = mySapModel.DatabaseTables.SetTableForEditingArray(
                                                                 TableKey: "Floor Object Connectivity",
                                                                 TableVersion: ref tableVersion_Table_01,
                                                                 FieldsKeysIncluded: ref filedKeyList_Table_01,
                                                                 NumberRecords: numberOfRectords_Table_01,
                                                                 TableData: ref reconstructedStrings_Table_01)
                                                               ;



            #endregion

            #region Walls

            //This Sets all Tables to be Shown when you fetch them from the Database (Atleast in SAP2000)
            mySapModel.DatabaseTables.GetTableForEditingArray(
                TableKey: "Wall Object Connectivity",
                GroupName: "All",
                TableVersion: ref tableVersion_Table_01,
                FieldsKeysIncluded: ref filedKeyList_Table_01,
                NumberRecords: ref numberOfRectords_Table_01,
                TableData: ref data_Table_01);

            List<CSIFLoorConnectivityDB> cSIWallsConnectivityDBs_Table_01 = new List<CSIFLoorConnectivityDB>();

            for (int i = 0; i < data_Table_01.Count(); i += 6)
            {
                CSIFLoorConnectivityDB cSIWallConnectivity = new CSIFLoorConnectivityDB();
                cSIWallConnectivity.UniqueName = data_Table_01[i];
                cSIWallConnectivity.UniquePt1 = data_Table_01[i + 1];
                cSIWallConnectivity.UniquePt2 = data_Table_01[i + 2];
                cSIWallConnectivity.UniquePt3 = data_Table_01[i + 3];
                cSIWallConnectivity.UniquePt4 = data_Table_01[i + 4];
                cSIWallConnectivity.GUID = data_Table_01[i + 5];

                cSIWallConnectivity.AllPointIds.Add(cSIWallConnectivity.UniquePt1);
                cSIWallConnectivity.AllPointIds.Add(cSIWallConnectivity.UniquePt2);
                cSIWallConnectivity.AllPointIds.Add(cSIWallConnectivity.UniquePt3);
                cSIWallConnectivity.AllPointIds.Add(cSIWallConnectivity.UniquePt4);
                cSIWallConnectivity.AllPointIds.RemoveAll(x => String.IsNullOrWhiteSpace(x));

                cSIWallsConnectivityDBs_Table_01.Add(cSIWallConnectivity);
            }



            for (int i = 0; i < cSIWallsConnectivityDBs_Table_01.Count(); i++)
            {
                List<CSIFLoorConnectivityDB> itemsWithEmptyGUID = cSIWallsConnectivityDBs_Table_01.SkipWhile((item, index) => index < i || !string.IsNullOrEmpty(item.GUID))
                                                                               .TakeWhile(item => string.IsNullOrEmpty(item.GUID))
                                                                               .ToList();
                foreach (var item in itemsWithEmptyGUID)
                {
                    cSIWallsConnectivityDBs_Table_01[i].AllPointIds.Add(item.UniquePt1);
                    cSIWallsConnectivityDBs_Table_01[i].AllPointIds.Add(item.UniquePt2);
                    cSIWallsConnectivityDBs_Table_01[i].AllPointIds.Add(item.UniquePt3);
                    cSIWallsConnectivityDBs_Table_01[i].AllPointIds.Add(item.UniquePt4);
                }
            }


            #region Trial 01
            List<CSIFLoorConnectivityDB> cSIWallConnectivityDBs_Filtered = cSIWallsConnectivityDBs_Table_01.Where(item => !String.IsNullOrWhiteSpace(item.GUID)).ToList();

            List<CSIFLoorConnectivityDB> cSIWallConnectivityDBs_Written = cSIWallConnectivityDBs_Filtered.ToList();

            List<CSIFLoorConnectivityDB> cSIWallConnectivityDBs_ToBeDeleted = new List<CSIFLoorConnectivityDB>();


            foreach (var connectivityDB in cSIWallConnectivityDBs_Filtered)
            {
                Area area = Areas.Where(item => item.ID == connectivityDB.UniqueName).FirstOrDefault();

                CSIFLoorConnectivityDB ParentRow = new CSIFLoorConnectivityDB();

                if (area != null)
                {
                    List<string> PointsIDs = area.vertices.Select(x => x.Id).ToList();
                    if (PointsIDs.Count > 4)
                    {
                        cSIWallConnectivityDBs_ToBeDeleted.Add(connectivityDB);
                        ParentRow.UniquePt1 = PointsIDs[0];
                        ParentRow.UniquePt2 = PointsIDs[1];
                        ParentRow.UniquePt3 = PointsIDs[2];
                        ParentRow.UniquePt4 = PointsIDs[3];
                        ParentRow.UniqueName = connectivityDB.UniqueName;
                        ParentRow.GUID = connectivityDB.GUID;

                        ParentRow.AllPointIds.Add(ParentRow.UniquePt1);
                        ParentRow.AllPointIds.Add(ParentRow.UniquePt2);
                        ParentRow.AllPointIds.Add(ParentRow.UniquePt3);
                        ParentRow.AllPointIds.Add(ParentRow.UniquePt4);


                        for (int i = 4; i < PointsIDs.Count; i += 4)
                        {
                            CSIFLoorConnectivityDB FollowerRow = new CSIFLoorConnectivityDB();

                            if (PointsIDs.Count - i == 3)
                            {
                                FollowerRow.UniquePt1 = PointsIDs[i];
                                FollowerRow.UniquePt2 = PointsIDs[i + 1];
                                FollowerRow.UniquePt3 = PointsIDs[i + 2];

                                FollowerRow.UniqueName = connectivityDB.UniqueName;
                                FollowerRow.GUID = "";
                                cSIWallConnectivityDBs_Written.Add(FollowerRow);

                                ParentRow.AllPointIds.Add(FollowerRow.UniquePt1);
                                ParentRow.AllPointIds.Add(FollowerRow.UniquePt2);
                                ParentRow.AllPointIds.Add(FollowerRow.UniquePt3);
                                break;
                            }
                            else if (PointsIDs.Count - i == 2)
                            {
                                FollowerRow.UniquePt1 = PointsIDs[i];
                                FollowerRow.UniquePt2 = PointsIDs[i + 1];

                                FollowerRow.UniqueName = connectivityDB.UniqueName;
                                FollowerRow.GUID = "";
                                cSIWallConnectivityDBs_Written.Add(FollowerRow);


                                ParentRow.AllPointIds.Add(FollowerRow.UniquePt1);
                                ParentRow.AllPointIds.Add(FollowerRow.UniquePt2);
                                break;
                            }
                            else if (PointsIDs.Count - i == 1)
                            {
                                FollowerRow.UniquePt1 = PointsIDs[i];

                                FollowerRow.UniqueName = connectivityDB.UniqueName;
                                FollowerRow.GUID = "";
                                cSIWallConnectivityDBs_Written.Add(FollowerRow);

                                ParentRow.AllPointIds.Add(FollowerRow.UniquePt1);
                                break;
                            }
                            else
                            {
                                CSIFLoorConnectivityDB TempPoint = new CSIFLoorConnectivityDB();

                                TempPoint.UniquePt1 = PointsIDs[i];
                                TempPoint.UniquePt2 = PointsIDs[i + 1];
                                TempPoint.UniquePt3 = PointsIDs[i + 2];
                                TempPoint.UniquePt4 = PointsIDs[i + 3];

                                TempPoint.UniqueName = connectivityDB.UniqueName;
                                TempPoint.GUID = "";
                                cSIWallConnectivityDBs_Written.Add(TempPoint);

                                ParentRow.AllPointIds.Add(TempPoint.UniquePt1);
                                ParentRow.AllPointIds.Add(TempPoint.UniquePt2);
                                ParentRow.AllPointIds.Add(TempPoint.UniquePt3);
                                ParentRow.AllPointIds.Add(TempPoint.UniquePt4);
                            }
                        }


                        cSIWallConnectivityDBs_Written.Add(ParentRow);
                    }
                    else if (PointsIDs.Count == 4)
                    {
                        cSIWallConnectivityDBs_ToBeDeleted.Add(connectivityDB);
                        ParentRow.UniquePt1 = PointsIDs[0];
                        ParentRow.UniquePt2 = PointsIDs[1];
                        ParentRow.UniquePt3 = PointsIDs[2];
                        ParentRow.UniquePt4 = PointsIDs[3];
                        ParentRow.UniqueName = connectivityDB.UniqueName;
                        ParentRow.GUID = connectivityDB.GUID;

                        ParentRow.AllPointIds.Add(ParentRow.UniquePt1);
                        ParentRow.AllPointIds.Add(ParentRow.UniquePt2);
                        ParentRow.AllPointIds.Add(ParentRow.UniquePt3);
                        ParentRow.AllPointIds.Add(ParentRow.UniquePt4);

                        cSIWallConnectivityDBs_Written.Add(ParentRow);
                    }
                    else if (PointsIDs.Count == 3)
                    {
                        cSIWallConnectivityDBs_ToBeDeleted.Add(connectivityDB);
                        ParentRow.UniquePt1 = PointsIDs[0];
                        ParentRow.UniquePt2 = PointsIDs[1];
                        ParentRow.UniquePt3 = PointsIDs[2];
                        ParentRow.UniqueName = connectivityDB.UniqueName;
                        ParentRow.GUID = connectivityDB.GUID;

                        ParentRow.AllPointIds.Add(ParentRow.UniquePt1);
                        ParentRow.AllPointIds.Add(ParentRow.UniquePt2);
                        ParentRow.AllPointIds.Add(ParentRow.UniquePt3);

                        cSIWallConnectivityDBs_Written.Add(ParentRow);
                    }
                    else if (PointsIDs.Count < 3)
                    {
                        // Throw an exception for cases with less than 3 points
                        cSIWallConnectivityDBs_ToBeDeleted.Add(connectivityDB);
                    }
                }

                else
                {
                    // if Area != null then the Elemenet Doesn't need Upade of Assigns
                    //cSIWallConnectivityDBs_Filtered.Add(connectivityDB);
                }
            }

            foreach (var item in cSIWallConnectivityDBs_ToBeDeleted)
            {
                cSIWallConnectivityDBs_Written.Remove(item);
            }
            #endregion


            List<CSIFLoorConnectivityDB> cSIWallConnectivityDBs_Final = new List<CSIFLoorConnectivityDB>();

            foreach (var connectivityDB in cSIWallConnectivityDBs_Written)
            {
                CSIFLoorConnectivityDB newObject = new CSIFLoorConnectivityDB();


                if (connectivityDB.AllPointIds.Count() > 4)
                {
                    newObject.UniquePt1 = connectivityDB.AllPointIds[0];
                    newObject.UniquePt2 = connectivityDB.AllPointIds[1];
                    newObject.UniquePt3 = connectivityDB.AllPointIds[2];
                    newObject.UniquePt4 = connectivityDB.AllPointIds[3];
                    newObject.UniqueName = connectivityDB.UniqueName;
                    newObject.GUID = connectivityDB.GUID;
                    newObject.AllPointIds = connectivityDB.AllPointIds;
                    cSIWallConnectivityDBs_Final.Add(newObject);
                    // Create new cSILoadPattern objects for the remaining points

                    // Unhandled Excpetion in Case the the Points were 3 not 4 in that next loop
                    for (int i = 4; i < connectivityDB.AllPointIds.Count; i += 4)
                    {
                        CSIFLoorConnectivityDB ChiledObject = new CSIFLoorConnectivityDB();
                        if (connectivityDB.AllPointIds.Count - i == 3)
                        {

                            ChiledObject.UniquePt1 = connectivityDB.AllPointIds[i];
                            ChiledObject.UniquePt2 = connectivityDB.AllPointIds[i + 1];
                            ChiledObject.UniquePt3 = connectivityDB.AllPointIds[i + 2];
                            ChiledObject.UniqueName = connectivityDB.UniqueName;
                            ChiledObject.GUID = "";
                            ChiledObject.AllPointIds = connectivityDB.AllPointIds;
                            cSIWallConnectivityDBs_Final.Add(ChiledObject);
                            break;
                        }
                        else if (connectivityDB.AllPointIds.Count - i == 2)
                        {
                            ChiledObject.UniquePt1 = connectivityDB.AllPointIds[i];
                            ChiledObject.UniquePt2 = connectivityDB.AllPointIds[i + 1];
                            ChiledObject.UniqueName = connectivityDB.UniqueName;
                            ChiledObject.GUID = "";
                            ChiledObject.AllPointIds = connectivityDB.AllPointIds;
                            cSIWallConnectivityDBs_Final.Add(ChiledObject);
                            break;
                        }
                        else if (connectivityDB.AllPointIds.Count - i == 1)
                        {
                            ChiledObject.UniquePt1 = connectivityDB.AllPointIds[i];
                            ChiledObject.UniqueName = connectivityDB.UniqueName;
                            ChiledObject.GUID = "";
                            cSIWallConnectivityDBs_Final.Add(ChiledObject);
                            break;
                        }
                        else
                        {
                            CSIFLoorConnectivityDB TempPoint = new CSIFLoorConnectivityDB();

                            TempPoint.UniquePt1 = connectivityDB.AllPointIds[i];
                            TempPoint.UniquePt2 = connectivityDB.AllPointIds[i + 1];
                            TempPoint.UniquePt3 = connectivityDB.AllPointIds[i + 2];
                            TempPoint.UniquePt4 = connectivityDB.AllPointIds[i + 3];
                            TempPoint.UniqueName = connectivityDB.UniqueName;
                            TempPoint.GUID = "";
                            cSIWallConnectivityDBs_Final.Add(TempPoint);
                        }
                    }
                }
                else if (connectivityDB.AllPointIds.Count() == 4)
                {
                    newObject.UniquePt1 = connectivityDB.AllPointIds[0];
                    newObject.UniquePt2 = connectivityDB.AllPointIds[1];
                    newObject.UniquePt3 = connectivityDB.AllPointIds[2];
                    newObject.UniquePt4 = connectivityDB.AllPointIds[3];
                    newObject.UniqueName = connectivityDB.UniqueName;
                    newObject.GUID = connectivityDB.GUID;
                    cSIWallConnectivityDBs_Final.Add(newObject);
                }
                else if (connectivityDB.AllPointIds.Count() == 3)
                {
                    newObject.UniquePt1 = connectivityDB.AllPointIds[0];
                    newObject.UniquePt2 = connectivityDB.AllPointIds[1];
                    newObject.UniquePt3 = connectivityDB.AllPointIds[2];
                    newObject.UniqueName = connectivityDB.UniqueName;
                    newObject.GUID = connectivityDB.GUID;
                    cSIWallConnectivityDBs_Final.Add(newObject);
                }
            }

            string[] reconstructedStringsWalls_Table_01 = new string[cSIWallConnectivityDBs_Final.Count * 6]; // 6 is the number of properties in the class
            int Counter2 = 0;
            foreach (var cSILoadPattern in cSIWallConnectivityDBs_Final)
            {
                reconstructedStringsWalls_Table_01[Counter2] = cSILoadPattern.UniqueName.ToString();
                reconstructedStringsWalls_Table_01[Counter2 + 1] = !String.IsNullOrWhiteSpace(cSILoadPattern.UniquePt1) ? cSILoadPattern.UniquePt1.ToString() : string.Empty;
                reconstructedStringsWalls_Table_01[Counter2 + 2] = !String.IsNullOrWhiteSpace(cSILoadPattern.UniquePt2) ? cSILoadPattern.UniquePt2.ToString() : string.Empty;
                reconstructedStringsWalls_Table_01[Counter2 + 3] = !String.IsNullOrWhiteSpace(cSILoadPattern.UniquePt3) ? cSILoadPattern.UniquePt3.ToString() : string.Empty;
                reconstructedStringsWalls_Table_01[Counter2 + 4] = !String.IsNullOrWhiteSpace(cSILoadPattern.UniquePt4) ? cSILoadPattern.UniquePt4.ToString() : string.Empty;
                reconstructedStringsWalls_Table_01[Counter2 + 5] = !String.IsNullOrWhiteSpace(cSILoadPattern.GUID) ? cSILoadPattern.GUID.ToString() : string.Empty;
                Counter2 += 6;
            }



            mySapModel.DatabaseTables.SetTableForEditingArray(
                                                             TableKey: "Wall Object Connectivity",
                                                             TableVersion: ref tableVersion_Table_01,
                                                             FieldsKeysIncluded: ref filedKeyList_Table_01,
                                                             NumberRecords: numberOfRectords_Table_01,
                                                             TableData: ref reconstructedStringsWalls_Table_01
                                                             );



            #endregion

            #region OpeningsAreas

            //This Sets all Tables to be Shown when you fetch them from the Database (Atleast in SAP2000)
            mySapModel.DatabaseTables.GetTableForEditingArray(
                TableKey: "Null Area Object Connectivity",
                GroupName: "All",
                TableVersion: ref tableVersion_Table_01,
                FieldsKeysIncluded: ref filedKeyList_Table_01,
                NumberRecords: ref numberOfRectords_Table_01,
                TableData: ref data_Table_01);

            List<CSIFLoorConnectivityDB> cSIOpeningsConnectivityDBs_Table_01 = new List<CSIFLoorConnectivityDB>();

            for (int i = 0; i < data_Table_01.Count(); i += 6)
            {
                CSIFLoorConnectivityDB cSIOpeningConnectivity = new CSIFLoorConnectivityDB();
                cSIOpeningConnectivity.UniqueName = data_Table_01[i];
                cSIOpeningConnectivity.UniquePt1 = data_Table_01[i + 1];
                cSIOpeningConnectivity.UniquePt2 = data_Table_01[i + 2];
                cSIOpeningConnectivity.UniquePt3 = data_Table_01[i + 3];
                cSIOpeningConnectivity.UniquePt4 = data_Table_01[i + 4];
                cSIOpeningConnectivity.GUID = data_Table_01[i + 5];

                cSIOpeningConnectivity.AllPointIds.Add(cSIOpeningConnectivity.UniquePt1);
                cSIOpeningConnectivity.AllPointIds.Add(cSIOpeningConnectivity.UniquePt2);
                cSIOpeningConnectivity.AllPointIds.Add(cSIOpeningConnectivity.UniquePt3);
                cSIOpeningConnectivity.AllPointIds.Add(cSIOpeningConnectivity.UniquePt4);
                cSIOpeningConnectivity.AllPointIds.RemoveAll(x => String.IsNullOrWhiteSpace(x));

                cSIOpeningsConnectivityDBs_Table_01.Add(cSIOpeningConnectivity);
            }



            for (int i = 0; i < cSIOpeningsConnectivityDBs_Table_01.Count(); i++)
            {
                List<CSIFLoorConnectivityDB> itemsWithEmptyGUID = cSIOpeningsConnectivityDBs_Table_01.SkipWhile((item, index) => index < i || !string.IsNullOrEmpty(item.GUID))
                                                                               .TakeWhile(item => string.IsNullOrEmpty(item.GUID))
                                                                               .ToList();
                foreach (var item in itemsWithEmptyGUID)
                {
                    cSIOpeningsConnectivityDBs_Table_01[i].AllPointIds.Add(item.UniquePt1);
                    cSIOpeningsConnectivityDBs_Table_01[i].AllPointIds.Add(item.UniquePt2);
                    cSIOpeningsConnectivityDBs_Table_01[i].AllPointIds.Add(item.UniquePt3);
                    cSIOpeningsConnectivityDBs_Table_01[i].AllPointIds.Add(item.UniquePt4);
                }
            }


            #region Trial 01
            List<CSIFLoorConnectivityDB> cSIOpeningConnectivityDBs_Filtered = cSIOpeningsConnectivityDBs_Table_01.Where(item => !String.IsNullOrWhiteSpace(item.GUID)).ToList();

            List<CSIFLoorConnectivityDB> cSIOpeningConnectivityDBs_Written = cSIOpeningConnectivityDBs_Filtered.ToList();

            List<CSIFLoorConnectivityDB> cSIOpeningConnectivityDBs_ToBeDeleted = new List<CSIFLoorConnectivityDB>();


            foreach (var connectivityDB in cSIOpeningConnectivityDBs_Filtered)
            {
                Area area = Areas.Where(item => item.ID == connectivityDB.UniqueName).FirstOrDefault();

                CSIFLoorConnectivityDB ParentRow = new CSIFLoorConnectivityDB();

                if (area != null)
                {
                    List<string> PointsIDs = area.vertices.Select(x => x.Id).ToList();
                    if (PointsIDs.Count > 4)
                    {
                        cSIOpeningConnectivityDBs_ToBeDeleted.Add(connectivityDB);
                        ParentRow.UniquePt1 = PointsIDs[0];
                        ParentRow.UniquePt2 = PointsIDs[1];
                        ParentRow.UniquePt3 = PointsIDs[2];
                        ParentRow.UniquePt4 = PointsIDs[3];
                        ParentRow.UniqueName = connectivityDB.UniqueName;
                        ParentRow.GUID = connectivityDB.GUID;

                        ParentRow.AllPointIds.Add(ParentRow.UniquePt1);
                        ParentRow.AllPointIds.Add(ParentRow.UniquePt2);
                        ParentRow.AllPointIds.Add(ParentRow.UniquePt3);
                        ParentRow.AllPointIds.Add(ParentRow.UniquePt4);


                        for (int i = 4; i < PointsIDs.Count; i += 4)
                        {
                            CSIFLoorConnectivityDB FollowerRow = new CSIFLoorConnectivityDB();

                            if (PointsIDs.Count - i == 3)
                            {
                                FollowerRow.UniquePt1 = PointsIDs[i];
                                FollowerRow.UniquePt2 = PointsIDs[i + 1];
                                FollowerRow.UniquePt3 = PointsIDs[i + 2];

                                FollowerRow.UniqueName = connectivityDB.UniqueName;
                                FollowerRow.GUID = "";
                                cSIOpeningConnectivityDBs_Written.Add(FollowerRow);

                                ParentRow.AllPointIds.Add(FollowerRow.UniquePt1);
                                ParentRow.AllPointIds.Add(FollowerRow.UniquePt2);
                                ParentRow.AllPointIds.Add(FollowerRow.UniquePt3);
                                break;
                            }
                            else if (PointsIDs.Count - i == 2)
                            {
                                FollowerRow.UniquePt1 = PointsIDs[i];
                                FollowerRow.UniquePt2 = PointsIDs[i + 1];

                                FollowerRow.UniqueName = connectivityDB.UniqueName;
                                FollowerRow.GUID = "";
                                cSIOpeningConnectivityDBs_Written.Add(FollowerRow);


                                ParentRow.AllPointIds.Add(FollowerRow.UniquePt1);
                                ParentRow.AllPointIds.Add(FollowerRow.UniquePt2);
                                break;
                            }
                            else if (PointsIDs.Count - i == 1)
                            {
                                FollowerRow.UniquePt1 = PointsIDs[i];

                                FollowerRow.UniqueName = connectivityDB.UniqueName;
                                FollowerRow.GUID = "";
                                cSIOpeningConnectivityDBs_Written.Add(FollowerRow);

                                ParentRow.AllPointIds.Add(FollowerRow.UniquePt1);
                                break;
                            }
                            else
                            {
                                CSIFLoorConnectivityDB TempPoint = new CSIFLoorConnectivityDB();

                                TempPoint.UniquePt1 = PointsIDs[i];
                                TempPoint.UniquePt2 = PointsIDs[i + 1];
                                TempPoint.UniquePt3 = PointsIDs[i + 2];
                                TempPoint.UniquePt4 = PointsIDs[i + 3];

                                TempPoint.UniqueName = connectivityDB.UniqueName;
                                TempPoint.GUID = "";
                                cSIOpeningConnectivityDBs_Written.Add(TempPoint);

                                ParentRow.AllPointIds.Add(TempPoint.UniquePt1);
                                ParentRow.AllPointIds.Add(TempPoint.UniquePt2);
                                ParentRow.AllPointIds.Add(TempPoint.UniquePt3);
                                ParentRow.AllPointIds.Add(TempPoint.UniquePt4);
                            }
                        }


                        cSIOpeningConnectivityDBs_Written.Add(ParentRow);
                    }
                    else if (PointsIDs.Count == 4)
                    {
                        cSIOpeningConnectivityDBs_ToBeDeleted.Add(connectivityDB);
                        ParentRow.UniquePt1 = PointsIDs[0];
                        ParentRow.UniquePt2 = PointsIDs[1];
                        ParentRow.UniquePt3 = PointsIDs[2];
                        ParentRow.UniquePt4 = PointsIDs[3];
                        ParentRow.UniqueName = connectivityDB.UniqueName;
                        ParentRow.GUID = connectivityDB.GUID;

                        ParentRow.AllPointIds.Add(ParentRow.UniquePt1);
                        ParentRow.AllPointIds.Add(ParentRow.UniquePt2);
                        ParentRow.AllPointIds.Add(ParentRow.UniquePt3);
                        ParentRow.AllPointIds.Add(ParentRow.UniquePt4);

                        cSIOpeningConnectivityDBs_Written.Add(ParentRow);
                    }
                    else if (PointsIDs.Count == 3)
                    {
                        cSIOpeningConnectivityDBs_ToBeDeleted.Add(connectivityDB);
                        ParentRow.UniquePt1 = PointsIDs[0];
                        ParentRow.UniquePt2 = PointsIDs[1];
                        ParentRow.UniquePt3 = PointsIDs[2];
                        ParentRow.UniqueName = connectivityDB.UniqueName;
                        ParentRow.GUID = connectivityDB.GUID;

                        ParentRow.AllPointIds.Add(ParentRow.UniquePt1);
                        ParentRow.AllPointIds.Add(ParentRow.UniquePt2);
                        ParentRow.AllPointIds.Add(ParentRow.UniquePt3);

                        cSIOpeningConnectivityDBs_Written.Add(ParentRow);
                    }
                    else if (PointsIDs.Count < 3)
                    {
                        // Throw an exception for cases with less than 3 points
                        cSIOpeningConnectivityDBs_ToBeDeleted.Add(connectivityDB);
                    }
                }

                else
                {
                    // if Area != null then the Elemenet Doesn't need Upade of Assigns
                    //cSIOpeningConnectivityDBs_Filtered.Add(connectivityDB);
                }
            }

            foreach (var item in cSIOpeningConnectivityDBs_ToBeDeleted)
            {
                cSIOpeningConnectivityDBs_Written.Remove(item);
            }
            #endregion


            List<CSIFLoorConnectivityDB> cSIOpeningConnectivityDBs_Final = new List<CSIFLoorConnectivityDB>();

            foreach (var connectivityDB in cSIOpeningConnectivityDBs_Written)
            {
                CSIFLoorConnectivityDB newObject = new CSIFLoorConnectivityDB();


                if (connectivityDB.AllPointIds.Count() > 4)
                {
                    newObject.UniquePt1 = connectivityDB.AllPointIds[0];
                    newObject.UniquePt2 = connectivityDB.AllPointIds[1];
                    newObject.UniquePt3 = connectivityDB.AllPointIds[2];
                    newObject.UniquePt4 = connectivityDB.AllPointIds[3];
                    newObject.UniqueName = connectivityDB.UniqueName;
                    newObject.GUID = connectivityDB.GUID;
                    newObject.AllPointIds = connectivityDB.AllPointIds;
                    cSIOpeningConnectivityDBs_Final.Add(newObject);
                    // Create new cSILoadPattern objects for the remaining points

                    // Unhandled Excpetion in Case the the Points were 3 not 4 in that next loop
                    for (int i = 4; i < connectivityDB.AllPointIds.Count; i += 4)
                    {
                        CSIFLoorConnectivityDB ChiledObject = new CSIFLoorConnectivityDB();
                        if (connectivityDB.AllPointIds.Count - i == 3)
                        {

                            ChiledObject.UniquePt1 = connectivityDB.AllPointIds[i];
                            ChiledObject.UniquePt2 = connectivityDB.AllPointIds[i + 1];
                            ChiledObject.UniquePt3 = connectivityDB.AllPointIds[i + 2];
                            ChiledObject.UniqueName = connectivityDB.UniqueName;
                            ChiledObject.GUID = "";
                            ChiledObject.AllPointIds = connectivityDB.AllPointIds;
                            cSIOpeningConnectivityDBs_Final.Add(ChiledObject);
                            break;
                        }
                        else if (connectivityDB.AllPointIds.Count - i == 2)
                        {
                            ChiledObject.UniquePt1 = connectivityDB.AllPointIds[i];
                            ChiledObject.UniquePt2 = connectivityDB.AllPointIds[i + 1];
                            ChiledObject.UniqueName = connectivityDB.UniqueName;
                            ChiledObject.GUID = "";
                            ChiledObject.AllPointIds = connectivityDB.AllPointIds;
                            cSIOpeningConnectivityDBs_Final.Add(ChiledObject);
                            break;
                        }
                        else if (connectivityDB.AllPointIds.Count - i == 1)
                        {
                            ChiledObject.UniquePt1 = connectivityDB.AllPointIds[i];
                            ChiledObject.UniqueName = connectivityDB.UniqueName;
                            ChiledObject.GUID = "";
                            cSIOpeningConnectivityDBs_Final.Add(ChiledObject);
                            break;
                        }
                        else
                        {
                            CSIFLoorConnectivityDB TempPoint = new CSIFLoorConnectivityDB();

                            TempPoint.UniquePt1 = connectivityDB.AllPointIds[i];
                            TempPoint.UniquePt2 = connectivityDB.AllPointIds[i + 1];
                            TempPoint.UniquePt3 = connectivityDB.AllPointIds[i + 2];
                            TempPoint.UniquePt4 = connectivityDB.AllPointIds[i + 3];
                            TempPoint.UniqueName = connectivityDB.UniqueName;
                            TempPoint.GUID = "";
                            cSIOpeningConnectivityDBs_Final.Add(TempPoint);
                        }
                    }
                }
                else if (connectivityDB.AllPointIds.Count() == 4)
                {
                    newObject.UniquePt1 = connectivityDB.AllPointIds[0];
                    newObject.UniquePt2 = connectivityDB.AllPointIds[1];
                    newObject.UniquePt3 = connectivityDB.AllPointIds[2];
                    newObject.UniquePt4 = connectivityDB.AllPointIds[3];
                    newObject.UniqueName = connectivityDB.UniqueName;
                    newObject.GUID = connectivityDB.GUID;
                    cSIOpeningConnectivityDBs_Final.Add(newObject);
                }
                else if (connectivityDB.AllPointIds.Count() == 3)
                {
                    newObject.UniquePt1 = connectivityDB.AllPointIds[0];
                    newObject.UniquePt2 = connectivityDB.AllPointIds[1];
                    newObject.UniquePt3 = connectivityDB.AllPointIds[2];
                    newObject.UniqueName = connectivityDB.UniqueName;
                    newObject.GUID = connectivityDB.GUID;
                    cSIOpeningConnectivityDBs_Final.Add(newObject);
                }
            }

            string[] reconstructedStringsOpenings_Table_01 = new string[cSIOpeningConnectivityDBs_Final.Count * 6]; // 6 is the number of properties in the class
            Counter2 = 0;
            foreach (var cSILoadPattern in cSIOpeningConnectivityDBs_Final)
            {
                reconstructedStringsOpenings_Table_01[Counter2] = cSILoadPattern.UniqueName.ToString();
                reconstructedStringsOpenings_Table_01[Counter2 + 1] = !String.IsNullOrWhiteSpace(cSILoadPattern.UniquePt1) ? cSILoadPattern.UniquePt1.ToString() : string.Empty;
                reconstructedStringsOpenings_Table_01[Counter2 + 2] = !String.IsNullOrWhiteSpace(cSILoadPattern.UniquePt2) ? cSILoadPattern.UniquePt2.ToString() : string.Empty;
                reconstructedStringsOpenings_Table_01[Counter2 + 3] = !String.IsNullOrWhiteSpace(cSILoadPattern.UniquePt3) ? cSILoadPattern.UniquePt3.ToString() : string.Empty;
                reconstructedStringsOpenings_Table_01[Counter2 + 4] = !String.IsNullOrWhiteSpace(cSILoadPattern.UniquePt4) ? cSILoadPattern.UniquePt4.ToString() : string.Empty;
                reconstructedStringsOpenings_Table_01[Counter2 + 5] = !String.IsNullOrWhiteSpace(cSILoadPattern.GUID) ? cSILoadPattern.GUID.ToString() : string.Empty;
                Counter2 += 6;
            }



            mySapModel.DatabaseTables.SetTableForEditingArray(
                                                             TableKey: "Null Area Object Connectivity",
                                                             TableVersion: ref tableVersion_Table_01,
                                                             FieldsKeysIncluded: ref filedKeyList_Table_01,
                                                             NumberRecords: numberOfRectords_Table_01,
                                                             TableData: ref reconstructedStringsOpenings_Table_01
                                                             );



            #endregion




            #region ApplyChangesInEtabs
            #region Definistions

            int NumFatalErrors = 0;
            int NumErrorMsgs = 0;
            int NumWarnMsgs = 0;
            int NumInfoMsgs = 0;
            string ImportLog = "";
            #endregion

            ret = mySapModel.DatabaseTables.ApplyEditedTables(true,
            ref NumFatalErrors,
            ref NumErrorMsgs,
            ref NumInfoMsgs,
            ref NumWarnMsgs,
            ref ImportLog);

            // Get the path to the desktop
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            // Define the file path and name
            string filePath = Path.Combine(desktopPath, "output.txt");


            mySapModel.DatabaseTables.CancelTableEditing();
            // Write the string to the text file
            File.WriteAllText(filePath, ImportLog);
            #endregion

            #region Add to Excel sheet

            List<object> Floorrecords = new List<object>();
            List<object> Wallsrecords = new List<object>();


            Counter = 0;
            foreach (var cSILoadPattern in cSIFLoorConnectivityDBs_Final)
            {
                object record = new
                {
                    UniqueName = reconstructedStrings_Table_01[Counter],
                    UniquePt1 = reconstructedStrings_Table_01[Counter + 1],
                    UniquePt2 = reconstructedStrings_Table_01[Counter + 2],
                    UniquePt3 = reconstructedStrings_Table_01[Counter + 3],
                    UniquePt4 = reconstructedStrings_Table_01[Counter + 4],
                    GUID = reconstructedStrings_Table_01[Counter + 5],
                };

                Counter += 6;

                Floorrecords.Add(record);
            }






            // Create Excel application
            Application excelApp = new Application();
            excelApp.Visible = true;

            // Create a new workbook
            Workbook workbook = excelApp.Workbooks.Add();

            if (Floorrecords.Count > 0)
            {
                ExcelUtils.AddDataToWorksheet(workbook, Floorrecords, "Floors");

            }

            // Add data to the first worksheet












            Counter = 0;
            foreach (var cSILoadPattern in cSIWallConnectivityDBs_Final)
            {
                object record = new
                {
                    UniqueName = reconstructedStringsWalls_Table_01[Counter],
                    UniquePt1 = reconstructedStringsWalls_Table_01[Counter + 1],
                    UniquePt2 = reconstructedStringsWalls_Table_01[Counter + 2],
                    UniquePt3 = reconstructedStringsWalls_Table_01[Counter + 3],
                    UniquePt4 = reconstructedStringsWalls_Table_01[Counter + 4],
                    GUID = reconstructedStringsWalls_Table_01[Counter + 5],
                };

                Counter += 6;

                Wallsrecords.Add(record);
            }

            if (Wallsrecords.Count > 0)
            {
                ExcelUtils.AddDataToWorksheet(workbook, Wallsrecords, "Walls");

            }



            #endregion

        }


        public static void updataAssignmentsFromRhinoToSap2(List<Area> Areas)
        {
            cSapModel mySapModel = Get_sap_mode();
            List<string> Allgroups = getGroupsList();
            CSiAPIv1.cSapModel aa = (CSiAPIv1.cSapModel)mySapModel;


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

                if (A.Sec == "None")
                {
                    mySapModel.AreaObj.SetOpening(A.ID, true);
                }


                string[] pointIDS = A.vertices.Select(x => x.Id).ToArray();
                int ret =   aa.EditArea.ChangeConnectivity(A.ID, A.vertices.Count,ref pointIDS);


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
                                Autodesk.Revit.DB.Curve line = Line.CreateBound(StartPoint, EndPoint);
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

        public static List<DriftRecrd> framesEndPointsDrifts(List<string> groupNames, double factorForWind, double factorForSeismic, List<string> LoadCaseKeys, int targetObjectType = 2)
        {
            List<Beam> bms = new List<Beam>();
            List<Column> cols = new List<Column>();
            int ret = 0;

            cSapModel mysapModel = Get_sap_mode();

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


            allLoadCombs = allLoadCombs.Where(x => LoadCaseKeys.Any(a => x.Contains(a))).ToArray();


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

                    if ((Xdiff / f.Length) > 1 / factorForWind)
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
            ETABSv1.cSapModel mysap = EtabsUtils.Get_sap_mode();


            foreach (var item in ids)
            {
                mysap.FrameObj.Delete(item);
            }


        }

        public static void deleteAreaObjects(List<string> ids)
        {
            ETABSv1.cSapModel mysap = Get_sap_mode();


            foreach (var item in ids)
            {
                int ret = mysap.AreaObj.Delete(item);
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
        public static List<Area> getEtabsAreas()
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


            cSapModel myEtabsModel = EtabsUtils.Get_sap_mode();



            #region get areaIDS
            myEtabsModel.AreaObj.GetNameList(ref NumberNames, ref AreaIDs);
            List<string> selectedAreaIDS = getSelectedAreaIDS();
            if (selectedAreaIDS.Count > 0)
                AreaIDs = selectedAreaIDS.ToArray();
            #endregion


            List<Area> EtabsAreas = new List<Area>();

            foreach (var areaId in AreaIDs)
            {
                Area A = new Area();
                A.ID = areaId;
                A.vertices = new List<Point>();


                #region set Points
                myEtabsModel.AreaObj.GetPoints(areaId, ref numberOfPoints, ref PointIds);


                foreach (var pointId in PointIds)
                {
                    Point pt = new Point();
                    ret = myEtabsModel.PointObj.GetCoordCartesian(pointId, ref x, ref y, ref z);
                    pt.Id = pointId;
                    pt.X = x;
                    pt.Y = y;
                    pt.Z = z;
                    A.vertices.Add(pt);

                }

                #endregion


                myEtabsModel.AreaObj.GetGroupAssign(areaId, ref numberOfGroups, ref GroupsNames);
                A.Groups = GroupsNames.ToList();


                myEtabsModel.AreaObj.GetLocalAxes(areaId, ref Ang, ref Advanced);
                A.Beta = Ang;


                myEtabsModel.AreaObj.GetProperty(areaId, ref PropName);
                A.Sec = PropName;


                eAreaDesignOrientation a = eAreaDesignOrientation.Floor;
                myEtabsModel.AreaObj.GetDesignOrientation(areaId, ref a);


                if (a == eAreaDesignOrientation.Floor)
                {
                    A.type = "Floor";

                }

                else if (a == eAreaDesignOrientation.Wall)
                {

                    A.type = "Wall";

                }

                else if (a == eAreaDesignOrientation.Null)
                {
                    A.type = "Null";

                }

                else
                {
                    A.type = "Other";

                }



                //myEtabsModel.AreaObj.GetProperty(areaId, ref PropName);
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




                EtabsAreas.Add(A);
            }



            #region validateGroups

            try
            {
                List<Area> EtabsAreasAssignedToGroup = EtabsAreas.Where(f => f != null).Where(a => a.Groups.Any(b => b[0] == '*')).ToList();

                foreach (var item in EtabsAreasAssignedToGroup)
                {
                    item.Group = item.Groups.First(a => a[0] == '*');
                    item.MainGroup = item.Groups.First(a => a[0] == '*');
                }

                var EtabsAreasGroups = EtabsAreasAssignedToGroup.GroupBy(a => a.MainGroup);

                foreach (var EtabsAreaGroup in EtabsAreasGroups)
                {
                    List<string> sections = EtabsAreaGroup.Select(a => a.Sec).Distinct().ToList();

                    foreach (var aera in EtabsAreaGroup)
                    {
                        aera.SubGroup = aera.Group + " _" + aera.Sec;
                        aera.Group = aera.Group + " _" + aera.Sec;

                    }
                }
            }

            catch (Exception)
            {

                throw;
            }




            #endregion


            return EtabsAreas;
        }

        #endregion

        #region related to segements 

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



            cSapModel mySapModel = Get_sap_mode();

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


                Frames.Add("Beam Object Connectivity");
                Frames.Add("Column Object Connectivity");
                Frames.Add("Frame Loads Assignments - Distributed");
                Frames.Add("Frame Loads Assignments - Temperature");
                Frames.Add("Frame Loads Assignments - Gravity");
                Frames.Add("Frame Loads Assignments - Point");


                Frames.Add("Frame Assignments - End Length Offsets");
                Frames.Add("Frame Assignments - Frame Auto Mesh Options");
                Frames.Add("Frame Assignments - Insertion Point");
                Frames.Add("Frame Assignments - Local Axes");
                Frames.Add("Frame Assignments - Output Stations");
                Frames.Add("Frame Assignments - Section Properties");
                Frames.Add("Frame Assignments - Releases and Partial Fixity");



                Frames.Add("Steel Frame Design Overwrites - AISC 360-16");
                Frames.Add("Steel Frame Design Preferences - AISC 360-16");



                // advanced local axes to be handeled
                // to be transefered using Api
                Frames.Add("Frame Section Property Definitions - Concrete Rectangular");
                Frames.Add("Frame Section Property Definitions - Nonprismatic");
                Frames.Add("Frame Section Property Definitions - Steel Pipe");
                Frames.Add("Frame Section Property Definitions - Steel Tube");




                //////
                Frames.Add("Frame Local Axes Assignments 2 - Advanced");
                Frames.Add("Frame Load Transfer Options");
                Frames.Add("Frame Loads - Deformation");
                #endregion

                #region joint Tables

                Joints.Add("Point Object Connectivity");
                Joints.Add("Joint Assignments - Restraints");

                Joints.Add("Joint Loads Assignments - Force");
                Joints.Add("Joint Loads Assignments - Ground Displacement");

                Joints.Add("Joint Pattern Definitions");

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



        #endregion

        #region DataBaseClasss


        public class DataBase
        {


            public static string Connectevity = "Point Object Connectivity";
            public static string[] ConnectevityFileds = new string[] { "UniqueName", "X", "Y", "Z", "IsSpecial" };
            public static string[] ConnectevityValues = null;


            public static string JointRestAss = "Joint Assignments - Restraints";
            public static string[] JointRestAssFileds = new string[] { "UniqueName", "UX", "UY", "UZ", "RX", "RY", "RZ" };
            public static string[] JointRestAssValues = null;



            public static string JointLoads = "Joint Loads Assignments - Force";
            public static string[] JointLoadsFileds = new string[] { "UniqueName", "LoadPattern", "FX", "FY", "FZ", "MX", "MY", "MZ" };
            public static string[] JointLoadsValues = null;


            public static string JointLoadGroundDisp = "Joint Loads Assignments - Ground Displacement";
            public static string[] JointLoadGroundDispFileds = new string[] { "UniqueName", "LoadPattern", "UX", "UY", "UZ", "RX", "RY", "RZ" };
            public static string[] JointLoadGroundDispValues = null;













            public static string ConnectevityBM = "Beam Object Connectivity";
            public static string[] ConnectevBMityFileds = new string[] { "UniqueName", "UniquePtI", "UniquePtJ" };
            public static string[] ConnectevityBMValues = null;


            public static string FrameSectionAssnignmetns = "Frame Assignments - Section Properties";
            public static string[] FrameSectionAssnignmetnsFileds = new string[] { "UniqueName", "AutoSelect", "SectProp" };
            public static string[] FrameSectionAssnignmetnsValues = null;


            public static string ConnectevityCL = "Column Object Connectivity";
            public static string[] ConnectevityCLFileds = new string[] { "UniqueName", "UniquePtI", "UniquePtJ" };
            public static string[] ConnectevityCLValues = null;


            public static string FrameLoadsDist = "Frame Loads Assignments - Distributed";
            public static string[] FrameLoadsDistFileds = new string[] { "UniqueName", "LoadPattern", "LoadType", "Dir", "DistType", "RelDistA", "RelDistB", "AbsDistA", "AbsDistB", "ForceA", "ForceB", "MomentA", "MomentB" };
            public static string[] FrameLoadsDistValues = null;

            public static string FrameLoadTepm = "Frame Loads Assignments - Temperature";
            public static string[] FrameLoadTepmFileds = new string[] { "UniqueName", "LoadPattern", "TempType", "Temp", "TempGrad2", "TempGrad3" };
            public static string[] FrameLoadTepmValues = null;


            public static string FrameLoadPoint = "Frame Loads Assignments - Point";
            public static string[] FrameLoadPointFileds = new string[] { "UniqueName", "LoadPattern", "LoadType", "Dir", "DistType", "RelDist", "AbsDist", "Force", "Moment" };
            public static string[] FrameLoadPointValues = null;


            public static string FrameLocalAxes = "Frame Assignments - Local Axes";
            public static string[] FrameLocalAxesFileds = new string[] { "UniqueName", "Angle" };
            public static string[] FrameLocalAxesValues = null;



            public static string FrameAutoMeshOPtion = "Frame Assignments - Frame Auto Mesh Options";
            public static string[] FrameAutoMeshOPtionFileds = new string[] { "UniqueName", "AutoMesh", "AtIntJoints", "AtIntersect", "NumSegments", "MaxLength" };
            public static string[] FrameAutoMeshOPtionValues = null;


            /// <summary>
            /// to be handeled
            /// </summary>
            public static string FrameOutPutStation = "Frame Assignments - Output Stations";
            public static string[] FrameOutPutStationFileds = new string[] { "UniqueName", "MinStations", "MaxSpacing" };
            public static string[] FrameOutPutStationValues = null;




            public static string FrameRealeses = "Frame Assignments - Releases and Partial Fixity";
            public static string[] FrameRealesesFileds = new string[] { "UniqueName", "PI", "V2I", "V3I", "TI", "M2I", "M3I", "PJ", "V2J", "V3J", "TJ", "M2J", "M3J" };
            public static string[] FrameRealesesValues = null;



            public static string insertionPoints = "Frame Assignments - Insertion Point";
            public static string[] insertionPointsFileds = new string[] { "UniqueName", "CardinalPt", "OffsetCSys", "XI", "YI", "ZI", "XJ", "YJ", "ZJ", "Mirror2", "Mirror3", "NoTransform" };
            public static string[] insertionPointsValues = null;








            public static string GroupDefinitions = "Group Definitions";
            public static string[] GroupDefinitionsFileds = new string[] { "Name", "Steel", "Color", "Concrete", "Composite" };
            public static string[] GroupDefinitionsValues = null;



            public static string GroupAssignments = "Group Assignments";
            public static string[] GroupAssignmentsFileds = new string[] { "UniqueName", "ObjectType", "ObjectLabel" };
            public static string[] GroupAssignmentsValues = null;



            public static string MaterialProp = "Material Properties - General";
            public static string[] MaterialPropFileds = new string[] { "Material", "Type", "Grade", "SymType", "Color", "Notes" };
            public static string[] MaterialPropValues = null;


            public static string MaterialMechProp = "Material Properties - Basic Mechanical Properties";
            public static string[] MaterialMechPropFileds = new string[] { "Material", "UnitWeight", "UnitMass", "E1", "G12", "U12", "A1" };
            public static string[] MaterialMechPropValues = null;


            public static string MaterialPropSteel = "Material Properties - Steel Data";
            public static string[] MaterialPropSteelFileds = new string[] { "Material", "Fy", "Fu", "Fye", "Fue", "SSCurveOpt", "SSHysType", "SHard", "SMax", "SRup", "FinalSlope" };
            public static string[] MaterialPropSteelValues = null;


            public static string MaterialPropConc = "Material Properties - Concrete Data";
            public static string[] MaterialPropConcFileds = new string[] { "Material", "Fc", "LtWtConc", "LtWtFact", "SSCurveOpt", "SSHysType", "SFc", "SCap", "FinalSlope", "FAngle", "DAngle", "TimeType", "TimeE", "EFact", "TimeCreep", "CreepFact", "TimeShrink", "ShrinkFact", "CreepType", "CreepTerms" };
            public static string[] MaterialPropConcValues = null;



            /// sections

            public static string FramesISections = "Frame Section Property Definitions - Steel I/Wide Flange";
            public static string[] FramesISectionsFileds = new string[] { "Name", "Material", "FromFile", "FileName", "SectInFile", "t3", "t2", "tf", "tw", "t2b", "tfb", "FilletRadius", "AMod", "A2Mod", "A3Mod", "JMod", "I2Mod", "I3Mod", "MMod", "WMod", "Color", "Notes", "GUID" };
            public static string[] FramesISectionsValues = null;


            public static string FramesBoxSections = "Frame Section Property Definitions - Steel Tube";
            public static string[] FramesBoxSectionsFileds = new string[] { "Name", "Material", "FromFile", "FileName", "SectInFile", "t3", "t2", "tf", "tw", "CornerRad", "AMod", "A2Mod", "A3Mod", "JMod", "I2Mod", "I3Mod", "MMod", "WMod", "Color", "Notes" };
            public static string[] FramesBoxSectionsValues = null;



            public static string FramesPipeSections = "Frame Section Property Definitions - Steel Pipe";
            public static string[] FramesPipeSectionsFileds = new string[] { "Name", "Material", "FromFile", "FileName", "SectInFile", "t3", "tw", "AMod", "A2Mod", "A3Mod", "JMod", "I2Mod", "I3Mod", "MMod", "WMod", "Color", "Notes" };
            public static string[] FramesPipeSectionsValues = null;

            public static string FramesRectConcSections = "Frame Section Property Definitions - Steel Pipe";
            public static string[] FramesRectConcSectionsFileds = new string[] { "Name", "Material", "FromFile", "FileName", "SectInFile", "t3", "tw", "AMod", "A2Mod", "A3Mod", "JMod", "I2Mod", "I3Mod", "MMod", "WMod", "Color", "Notes" };
            public static string[] FramesRectConcSectionsValues = null;


            public static string NonPresmaticSections = "Frame Section Property Definitions - Nonprismatic";
            public static string[] NonPresmaticSectionsFileds = new string[] { "Name", "StartSect", "EndSect", "LengthType", "Length", "EI33Var", "EI22Var" };
            public static string[] NonPresmaticSectionsValues = null;






            public static string AreaDistLoad = "Area Load Assignments - Uniform";
            public static string[] AreaDistLoadFileds = new string[] { "UniqueName", "LoadPattern", "Dir", "Load" };
            public static string[] AreaDistLoadValues = null;



        }


        public class Joint
        {



        }

        public class Fm
        {




        }



        #endregion


    }

}