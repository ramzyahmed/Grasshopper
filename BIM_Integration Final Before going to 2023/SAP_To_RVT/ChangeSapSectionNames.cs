using BIM_Integration;
using SAP2000v1;
using System;
using System.Collections.Generic;
using System.Linq;

namespace BIM_Integration
{
    public class ChangeSapSectionNames
    {
        public static void GetSectionsDimentions2(cSapModel mySapModel, ref List<Frame> frames, List<string> OptionalFrameIDS = null)
        {
            List<object> secDic = new List<object>();

            //List<string> SapSectionS = SapUtils.Sections.getSections(mySapModel);
            List<string> SapSectionS = SapUtils.Sections.GetAllNamesAssignedToFramesInSap(mySapModel, OptionalFrameIDS);
            Dictionary<string, Section_type> mydic = SapUtils.Sections.getSectionTypes(SapSectionS);


            List<string> iSection = mydic.Where(x => x.Value == Section_type.I_SECTION_B).ToDictionary(x => x.Key, x => x.Value).Keys.ToList();
            List<Object> IDimentions = SapUtils.Sections.Get_i_section_dimentions(mySapModel, iSection) as List<Object>;
            secDic = secDic.Union(reNameIsection2(IDimentions)).ToList();



            List<string> boxSection = mydic.Where(x => x.Value == Section_type.Hollow_box_B || x.Value == Section_type.Hollow_box_H).ToDictionary(x => x.Key, x => x.Value).Keys.ToList();
            List<Object> boxDimentions = SapUtils.Sections.Get_Box_section_dimentions(mySapModel, boxSection) as List<Object>;
            secDic = secDic.Union(reNameBoxsections2(boxDimentions)).ToList();



            List<string> pipeSection = mydic.Where(x => x.Value == Section_type.Hollow_pipe_B || x.Value == Section_type.Hollow_pipe_H).ToDictionary(x => x.Key, x => x.Value).Keys.ToList();
            List<Object> pipeDimentions = SapUtils.Sections.Get_pipe_section_dimentions(mySapModel, pipeSection) as List<Object>;
            secDic = secDic.Union(reNamePipesections2(pipeDimentions)).ToList();


            List<string> channelSection = mydic.Where(x => x.Value == Section_type.Channel_B).ToDictionary(x => x.Key, x => x.Value).Keys.ToList();
            List<Object> channelDimentions = SapUtils.Sections.Get_channel_section_dimentions(mySapModel, channelSection) as List<Object>;
            secDic = secDic.Union(reNameChannelsections2(channelDimentions)).ToList();


            List<string> iTaperedSection = mydic.Where(x => x.Value == Section_type.I_Tapered).ToDictionary(x => x.Key, x => x.Value).Keys.ToList();
            List<Object> iTaperedDimentions = SapUtils.Sections.Get_Tapered_i_section_dimentions(mySapModel, iTaperedSection) as List<Object>;
            secDic = secDic.Union(reNameITaperedsections2(iTaperedDimentions)).ToList();



            List<string> boxTaperedSection = mydic.Where(x => x.Value == Section_type.Box_Tapered).ToDictionary(x => x.Key, x => x.Value).Keys.ToList();
            List<Object> boxTaperedDimentions = SapUtils.Sections.Get_Tapered_BOX_section_dimentions(mySapModel, boxTaperedSection) as List<Object>;
            secDic = secDic.Union(reNameboxTaperedsections3(boxTaperedDimentions)).ToList();



            List<string> pipeTaperedSection = mydic.Where(x => x.Value == Section_type.Pipe_Tapered).ToDictionary(x => x.Key, x => x.Value).Keys.ToList();
            List<Object> pipeTaperedDimentions = SapUtils.Sections.Get_Tapered_PIPE_section_dimentions(mySapModel, pipeTaperedSection) as List<Object>;
            secDic = secDic.Union(reNamePipeTaperedsections2(pipeTaperedDimentions)).ToList();



            List<string> solidPipeSection = mydic.Where(x => x.Value == Section_type.Solid_pipe).ToDictionary(x => x.Key, x => x.Value).Keys.ToList();
            List<Object> sloidPipeDimentions = SapUtils.Sections.Get_solid_pipe_section_dimentions(mySapModel, solidPipeSection) as List<Object>;
            secDic = secDic.Union(reNameSolidPipesections2(sloidPipeDimentions)).ToList();



            List<string> solidBoxSection = mydic.Where(x => x.Value == Section_type.Solid_box).ToDictionary(x => x.Key, x => x.Value).Keys.ToList();
            List<Object> sloidBoxDimentions = SapUtils.Sections.Get_solid_rect_section_dimentions(mySapModel, solidBoxSection) as List<Object>;
            secDic = secDic.Union(reNameSolidBoxsections2(sloidBoxDimentions)).ToList();



            List<string> teeSection = mydic.Where(x => x.Value == Section_type.T_Section_B).ToDictionary(x => x.Key, x => x.Value).Keys.ToList();
            List<Object> teeDimentions = SapUtils.Sections.Get_Tee_section_dimentions(mySapModel, teeSection) as List<Object>;
            secDic = secDic.Union(reNameTsections2(teeDimentions)).ToList();



            List<string> singleAngleSection = mydic.Where(x => x.Value == Section_type.S_Angle_B).ToDictionary(x => x.Key, x => x.Value).Keys.ToList();
            List<Object> singleAngleDimentions = SapUtils.Sections.Get_Sig_angle_section_dimentions(mySapModel, singleAngleSection) as List<Object>;
            secDic = secDic.Union(reNameSingleAngle2(singleAngleDimentions)).ToList();



            List<string> concrBeamRectSection = mydic.Where(x => x.Value == Section_type.Concrete_rect).ToDictionary(x => x.Key, x => x.Value).Keys.ToList();
            List<Object> ConcrBeamRectDimentions = SapUtils.Sections.Get_solid_rect_section_dimentions(mySapModel, concrBeamRectSection) as List<Object>;
            secDic = secDic.Union(reNameSolidBoxsections2(ConcrBeamRectDimentions)).ToList();



            List<string> concrBeamPipeSection = mydic.Where(x => x.Value == Section_type.Concrete_circ).ToDictionary(x => x.Key, x => x.Value).Keys.ToList();
            List<Object> ConcrBeamPipeDimentions = SapUtils.Sections.Get_solid_pipe_section_dimentions(mySapModel, concrBeamPipeSection) as List<Object>;
            secDic = secDic.Union(reNameSolidPipesections2(ConcrBeamPipeDimentions)).ToList();


            foreach (dynamic secObjec in secDic)
            {
                frames.Where(x => x.Sec == secObjec.sec).ToList().ForEach(y => y.Sec = secObjec.newSec);

            }


        }


        public static void RenameSections(cSapModel mySapModel)
        {
            List<object> secDic = new List<object>();

            //List<string> SapSectionS = SapUtils.Sections.getSections(mySapModel);
            List<string> SapSectionS = SapUtils.Sections.GetAllNamesAssignedToFramesInSap(mySapModel);
            Dictionary<string, Section_type> mydic = SapUtils.Sections.getSectionTypes(SapSectionS);


            List<string> iSection = mydic.Where(x => x.Value == Section_type.I_SECTION_B).ToDictionary(x => x.Key, x => x.Value).Keys.ToList();
            List<Object> IDimentions = SapUtils.Sections.Get_i_section_dimentions(mySapModel, iSection) as List<Object>;
            secDic = secDic.Union(reNameIsection2(IDimentions)).ToList();



            List<string> boxSection = mydic.Where(x => x.Value == Section_type.Hollow_box_B || x.Value == Section_type.Hollow_box_H).ToDictionary(x => x.Key, x => x.Value).Keys.ToList();
            List<Object> boxDimentions = SapUtils.Sections.Get_Box_section_dimentions(mySapModel, boxSection) as List<Object>;
            secDic = secDic.Union(reNameBoxsections2(boxDimentions)).ToList();



            List<string> pipeSection = mydic.Where(x => x.Value == Section_type.Hollow_pipe_B || x.Value == Section_type.Hollow_pipe_H).ToDictionary(x => x.Key, x => x.Value).Keys.ToList();
            List<Object> pipeDimentions = SapUtils.Sections.Get_pipe_section_dimentions(mySapModel, pipeSection) as List<Object>;
            secDic = secDic.Union(reNamePipesections2(pipeDimentions)).ToList();


            List<string> channelSection = mydic.Where(x => x.Value == Section_type.Channel_B).ToDictionary(x => x.Key, x => x.Value).Keys.ToList();
            List<Object> channelDimentions = SapUtils.Sections.Get_channel_section_dimentions(mySapModel, channelSection) as List<Object>;
            secDic = secDic.Union(reNameChannelsections2(channelDimentions)).ToList();


            List<string> iTaperedSection = mydic.Where(x => x.Value == Section_type.I_Tapered).ToDictionary(x => x.Key, x => x.Value).Keys.ToList();
            List<Object> iTaperedDimentions = SapUtils.Sections.Get_Tapered_i_section_dimentions(mySapModel, iTaperedSection) as List<Object>;
            secDic = secDic.Union(reNameITaperedsections2(iTaperedDimentions)).ToList();



            List<string> boxTaperedSection = mydic.Where(x => x.Value == Section_type.Box_Tapered).ToDictionary(x => x.Key, x => x.Value).Keys.ToList();
            List<Object> boxTaperedDimentions = SapUtils.Sections.Get_Tapered_BOX_section_dimentions(mySapModel, boxTaperedSection) as List<Object>;
            secDic = secDic.Union(reNameboxTaperedsections2(boxTaperedDimentions)).ToList();



            List<string> pipeTaperedSection = mydic.Where(x => x.Value == Section_type.Pipe_Tapered).ToDictionary(x => x.Key, x => x.Value).Keys.ToList();
            List<Object> pipeTaperedDimentions = SapUtils.Sections.Get_Tapered_PIPE_section_dimentions(mySapModel, pipeTaperedSection) as List<Object>;
            secDic = secDic.Union(reNamePipeTaperedsections2(pipeTaperedDimentions)).ToList();



            List<string> solidPipeSection = mydic.Where(x => x.Value == Section_type.Solid_pipe).ToDictionary(x => x.Key, x => x.Value).Keys.ToList();
            List<Object> sloidPipeDimentions = SapUtils.Sections.Get_solid_pipe_section_dimentions(mySapModel, solidPipeSection) as List<Object>;
            secDic = secDic.Union(reNameSolidPipesections2(sloidPipeDimentions)).ToList();



            List<string> solidBoxSection = mydic.Where(x => x.Value == Section_type.Solid_box).ToDictionary(x => x.Key, x => x.Value).Keys.ToList();
            List<Object> sloidBoxDimentions = SapUtils.Sections.Get_solid_rect_section_dimentions(mySapModel, solidBoxSection) as List<Object>;
            secDic = secDic.Union(reNameSolidBoxsections2(sloidBoxDimentions)).ToList();



            List<string> teeSection = mydic.Where(x => x.Value == Section_type.T_Section_B).ToDictionary(x => x.Key, x => x.Value).Keys.ToList();
            List<Object> teeDimentions = SapUtils.Sections.Get_Tee_section_dimentions(mySapModel, teeSection) as List<Object>;
            secDic = secDic.Union(reNameTsections2(teeDimentions)).ToList();



            List<string> singleAngleSection = mydic.Where(x => x.Value == Section_type.S_Angle_B).ToDictionary(x => x.Key, x => x.Value).Keys.ToList();
            List<Object> singleAngleDimentions = SapUtils.Sections.Get_Sig_angle_section_dimentions(mySapModel, singleAngleSection) as List<Object>;
            secDic = secDic.Union(reNameSingleAngle2(singleAngleDimentions)).ToList();



            List<string> concrBeamRectSection = mydic.Where(x => x.Value == Section_type.Concrete_rect).ToDictionary(x => x.Key, x => x.Value).Keys.ToList();
            List<Object> ConcrBeamRectDimentions = SapUtils.Sections.Get_solid_rect_section_dimentions(mySapModel, concrBeamRectSection) as List<Object>;
            secDic = secDic.Union(reNameSolidBoxsections2(ConcrBeamRectDimentions)).ToList();



            List<string> concrBeamPipeSection = mydic.Where(x => x.Value == Section_type.Concrete_circ).ToDictionary(x => x.Key, x => x.Value).Keys.ToList();
            List<Object> ConcrBeamPipeDimentions = SapUtils.Sections.Get_solid_pipe_section_dimentions(mySapModel, concrBeamPipeSection) as List<Object>;
            secDic = secDic.Union(reNameSolidPipesections2(ConcrBeamPipeDimentions)).ToList();


            foreach (dynamic secObjec in secDic)
            {

                mySapModel.PropFrame.ChangeName(secObjec.sec, secObjec.newSec);


            }


        }


        public static List<object> GetSapSectionsDictionary(cSapModel mySapModel, ref List<Frame> frames)
        {
            List<object> sapSecNewSecObject = new List<object>();

            List<string> sapSections = frames.Select(x => x.Sec).Distinct().ToList();

            List<object> secDic = new List<object>();

            //List<string> SapSectionS = SapUtils.Sections.getSections(mySapModel);
            List<string> SapSectionS = SapUtils.Sections.getSections(mySapModel);
            Dictionary<string, Section_type> mydic = SapUtils.Sections.getSectionTypes(SapSectionS);



            List<string> iSection = mydic.Where(x => x.Value == Section_type.I_SECTION_B).ToDictionary(x => x.Key, x => x.Value).Keys.ToList();
            List<Object> IDimentions = SapUtils.Sections.Get_i_section_dimentions(mySapModel, iSection) as List<Object>;
            secDic = secDic.Union(reNameIsection2(IDimentions)).ToList();



            List<string> boxSection = mydic.Where(x => x.Value == Section_type.Hollow_box_B || x.Value == Section_type.Hollow_box_H).ToDictionary(x => x.Key, x => x.Value).Keys.ToList();
            List<Object> boxDimentions = SapUtils.Sections.Get_Box_section_dimentions(mySapModel, boxSection) as List<Object>;
            secDic = secDic.Union(reNameBoxsections2(boxDimentions)).ToList();



            List<string> pipeSection = mydic.Where(x => x.Value == Section_type.Hollow_pipe_B || x.Value == Section_type.Hollow_pipe_H).ToDictionary(x => x.Key, x => x.Value).Keys.ToList();
            List<Object> pipeDimentions = SapUtils.Sections.Get_pipe_section_dimentions(mySapModel, pipeSection) as List<Object>;
            secDic = secDic.Union(reNamePipesections2(pipeDimentions)).ToList();


            List<string> channelSection = mydic.Where(x => x.Value == Section_type.Channel_B).ToDictionary(x => x.Key, x => x.Value).Keys.ToList();
            List<Object> channelDimentions = SapUtils.Sections.Get_channel_section_dimentions(mySapModel, channelSection) as List<Object>;
            secDic = secDic.Union(reNameChannelsections2(channelDimentions)).ToList();


            List<string> iTaperedSection = mydic.Where(x => x.Value == Section_type.I_Tapered).ToDictionary(x => x.Key, x => x.Value).Keys.ToList();
            List<Object> iTaperedDimentions = SapUtils.Sections.Get_Tapered_i_section_dimentions(mySapModel, iTaperedSection) as List<Object>;
            secDic = secDic.Union(reNameITaperedsections2(iTaperedDimentions)).ToList();



            List<string> boxTaperedSection = mydic.Where(x => x.Value == Section_type.Box_Tapered).ToDictionary(x => x.Key, x => x.Value).Keys.ToList();
            List<Object> boxTaperedDimentions = SapUtils.Sections.Get_Tapered_BOX_section_dimentions(mySapModel, boxTaperedSection) as List<Object>;
            secDic = secDic.Union(reNameboxTaperedsections2(boxTaperedDimentions)).ToList();



            List<string> pipeTaperedSection = mydic.Where(x => x.Value == Section_type.Pipe_Tapered).ToDictionary(x => x.Key, x => x.Value).Keys.ToList();
            List<Object> pipeTaperedDimentions = SapUtils.Sections.Get_Tapered_PIPE_section_dimentions(mySapModel, pipeTaperedSection) as List<Object>;
            secDic = secDic.Union(reNamePipeTaperedsections2(pipeTaperedDimentions)).ToList();



            List<string> solidPipeSection = mydic.Where(x => x.Value == Section_type.Solid_pipe).ToDictionary(x => x.Key, x => x.Value).Keys.ToList();
            List<Object> sloidPipeDimentions = SapUtils.Sections.Get_solid_pipe_section_dimentions(mySapModel, solidPipeSection) as List<Object>;
            secDic = secDic.Union(reNameSolidPipesections2(sloidPipeDimentions)).ToList();



            List<string> solidBoxSection = mydic.Where(x => x.Value == Section_type.Solid_box).ToDictionary(x => x.Key, x => x.Value).Keys.ToList();
            List<Object> sloidBoxDimentions = SapUtils.Sections.Get_solid_rect_section_dimentions(mySapModel, solidBoxSection) as List<Object>;
            secDic = secDic.Union(reNameSolidBoxsections2(sloidBoxDimentions)).ToList();



            List<string> teeSection = mydic.Where(x => x.Value == Section_type.T_Section_B).ToDictionary(x => x.Key, x => x.Value).Keys.ToList();
            List<Object> teeDimentions = SapUtils.Sections.Get_Tee_section_dimentions(mySapModel, teeSection) as List<Object>;
            secDic = secDic.Union(reNameTsections2(teeDimentions)).ToList();



            List<string> singleAngleSection = mydic.Where(x => x.Value == Section_type.S_Angle_B).ToDictionary(x => x.Key, x => x.Value).Keys.ToList();
            List<Object> singleAngleDimentions = SapUtils.Sections.Get_Sig_angle_section_dimentions(mySapModel, singleAngleSection) as List<Object>;
            secDic = secDic.Union(reNameSingleAngle2(singleAngleDimentions)).ToList();



            List<string> concrBeamRectSection = mydic.Where(x => x.Value == Section_type.Concrete_rect).ToDictionary(x => x.Key, x => x.Value).Keys.ToList();
            List<Object> ConcrBeamRectDimentions = SapUtils.Sections.Get_solid_rect_section_dimentions(mySapModel, concrBeamRectSection) as List<Object>;
            secDic = secDic.Union(reNameSolidBoxsections2(ConcrBeamRectDimentions)).ToList();



            List<string> concrBeamPipeSection = mydic.Where(x => x.Value == Section_type.Concrete_circ).ToDictionary(x => x.Key, x => x.Value).Keys.ToList();
            List<Object> ConcrBeamPipeDimentions = SapUtils.Sections.Get_solid_pipe_section_dimentions(mySapModel, concrBeamPipeSection) as List<Object>;
            secDic = secDic.Union(reNameSolidPipesections2(ConcrBeamPipeDimentions)).ToList();


            foreach (dynamic secObjec in secDic)
            {

                string SapSec = SapSectionS.FirstOrDefault(x => x == secObjec.sec);

                sapSecNewSecObject.Add(new { SapSec = SapSec, newSec = secObjec.newSec });

            }

            List<string> sapSecFromFunction = secDic.Cast<dynamic>().Select(x => x.sec).Cast<string>().ToList();
            List<string> notFoundSapSection = SapSectionS.Except(sapSecFromFunction).ToList();

            foreach (var SapSec in notFoundSapSection)
            {

                sapSecNewSecObject.Add(new { SapSec = SapSec, newSec = SapSec });
            }

            return sapSecNewSecObject;
        }

        public static List<object> reNameIsection2(List<object> iSectionsDimentions)
        {
            List<object> result = new List<object>();
            foreach (dynamic section in iSectionsDimentions.AsParallel())
            {
                //var obj2 = new { name = profile, depth = t3, width = t2, flg_thick = tf, web_thick = tw, rf = 0, mat = MatProp };
                int ret = 0;

                double webDepth = section.depth - 2 * section.flg_thick;

                string newName = "2FL. PL. " + section.width + "X" + section.flg_thick + " +WEB PL. " + webDepth + "X" + section.web_thick;

                object secDic = new { sec = section.name, newSec = newName };

                result.Add(secDic);

            }

            return result;

        }

        public static List<object> reNameBoxsections2(List<object> boxSectionsDimentions)
        {
            // var obj = new { name = profile, depth = t3, width = t2, flg_thick = tf, web_thick = tw, rf = 0, mat = MatProp };
            //TUBE 500X500X40
            List<object> result = new List<object>();

            foreach (dynamic section in boxSectionsDimentions.AsParallel())
            {
                //var obj2 = new { name = profile, depth = t3, width = t2, flg_thick = tf, web_thick = tw, rf = 0, mat = MatProp };
                int ret = 0;
                string newName = null;

                if (section.flg_thick == section.web_thick)
                {
                    newName = "TUBE " + section.depth + "X" + section.width + "X" + section.flg_thick;
                }

                else
                {
                    newName = "TUBE " + section.depth + "X" + section.width + "X" + section.web_thick + "X" + section.flg_thick;
                }


                object secDic = new { sec = section.name, newSec = newName };

                result.Add(secDic);
            }

            return result;
        }

        public static List<object> reNamePipesections2(List<object> pipeSectionsDimentions)
        {
            List<object> result = new List<object>();
            foreach (dynamic section in pipeSectionsDimentions.AsParallel())
            {
                int ret = 0;
                string newName = "PIPE " + section.diameter + "X" + section.thick;
                object secDic = new { sec = section.name, newSec = newName };

                result.Add(secDic);

            }

            return result;



        }

        public static List<object> reNameITaperedsections2(List<object> iTaperdSectionsDimentions)
        {
            List<object> result = new List<object>();
            //var obj = new { name = profile, st_depth = t3, st_width = t2, flg_thick = tf, web_thick = tw, end_depth = t32, end_width = t22, mat = MatProp };
            //2FL.PL. 300X20 +WEB PL. 710/460X10

            double largeDepth = 0;
            double largeWebDepth = 0;
            double smallDeph = 0;
            double smallWebDepth = 0;
            int ret = 0;
            foreach (dynamic section in iTaperdSectionsDimentions.AsParallel())
            {

                if (section.st_depth > section.end_depth)
                {
                    largeDepth = section.st_depth;
                    smallDeph = section.end_depth;
                }

                else
                {
                    largeDepth = section.st_depth;
                    smallDeph = section.end_depth;
                }

                largeWebDepth = largeDepth - 2 * section.flg_thick;
                smallWebDepth = smallDeph - 2 * section.flg_thick;


                string newName = "2FL. PL. " + section.st_width + "X" + section.flg_thick + " +WEB PL. " + largeWebDepth + "/" + smallWebDepth + "X" + section.web_thick;
                object secDic = new { sec = section.name, newSec = newName };

                result.Add(secDic);

            }

            return result;

        }

        public static List<object> reNameboxTaperedsections2(List<object> boxTaperedSectionsDimentions)
        {
            List<object> result = new List<object>();

            //var obj = new { name = profile, st_depth = t3, st_width = t2, flg_thick = tf, web_thick = tw, end_depth = t32, end_width = t22, mat = MatProp };
            //  800/500X400/300X20

            double largeDepth = 0;
            double smallDepth = 0;
            double constDepth = 0;

            double largeWidth = 0;
            double smallWidth = 0;
            double constWidth = 0;

            bool isConstDepth = false;
            bool isConstWidth = false;
            int ret = 0;

            string newName = null;


            foreach (dynamic section in boxTaperedSectionsDimentions)
            {
                if (section.st_depth == section.end_depth)
                {
                    constDepth = section.st_depth;
                    isConstDepth = true;
                }

                else
                {
                    if (section.st_depth > section.end_depth)
                    {
                        largeDepth = section.st_depth;
                        smallDepth = section.end_depth;
                    }

                    else
                    {
                        largeDepth = section.st_depth;
                        smallDepth = section.end_depth;
                    }
                }


                if (section.st_width == section.end_width)
                {
                    constWidth = section.st_width;
                    isConstWidth = true;
                }

                else
                {
                    if (section.st_width > section.end_width)
                    {
                        largeWidth = section.st_width;
                        smallWidth = section.end_width;
                    }

                    else
                    {
                        largeWidth = section.st_width;
                        smallWidth = section.end_width;
                    }
                }


                if (largeDepth == largeWidth && smallDepth == smallWidth)
                {

                    newName = "TUBE " + largeDepth + "/" + smallDepth + "X" + section.flg_thick;
                }



                else if (!isConstDepth && !isConstWidth)
                {
                    newName = "TUBE " + largeDepth + "/" + smallDepth + "X" + largeWidth + "/" + smallWidth + "X" + section.flg_thick;

                }

                else if (!isConstDepth && isConstWidth)
                {
                    newName = "TUBE " + largeDepth + "/" + smallDepth + "X" + constWidth + "X" + section.flg_thick;
                }


                else if (isConstDepth && !isConstWidth)
                {
                    newName = "TUBE " + constDepth + "X" + largeWidth + "/" + smallWidth + "X" + section.flg_thick;
                }

                else if (isConstDepth && isConstWidth)
                {
                    newName = "TUBE " + constDepth + "X" + constWidth + "X" + section.flg_thick;
                }


                object secDic = new { sec = section.name, newSec = newName };

                result.Add(secDic);

            }

            return result;

        }

        public static List<object> reNameboxTaperedsections3(List<object> boxTaperedSectionsDimentions)
        {
            List<object> result = new List<object>();

            //var obj = new { name = profile, st_depth = t3, st_width = t2, flg_thick = tf, web_thick = tw, end_depth = t32, end_width = t22, mat = MatProp };
            //  800/500X400/300X20

            double largeDepth = 0;
            double smallDepth = 0;
            double constDepth = 0;

            double largeWidth = 0;
            double smallWidth = 0;
            double constWidth = 0;

            bool isConstDepth = false;
            bool isConstWidth = false;
            int ret = 0;

            string newName = null;


            foreach (dynamic section in boxTaperedSectionsDimentions)
            {
                if (section.st_depth == section.end_depth)
                {
                    constDepth = section.st_depth;
                    isConstDepth = true;
                }

                else
                {
                    if (section.st_depth > section.end_depth)
                    {
                        largeDepth = section.st_depth;
                        smallDepth = section.end_depth;
                    }

                    else
                    {
                        largeDepth = section.st_depth;
                        smallDepth = section.end_depth;
                    }
                }


                if (section.st_width == section.end_width)
                {
                    constWidth = section.st_width;
                    isConstWidth = true;
                }

                else
                {
                    if (section.st_width > section.end_width)
                    {
                        largeWidth = section.st_width;
                        smallWidth = section.end_width;
                    }

                    else
                    {
                        largeWidth = section.st_width;
                        smallWidth = section.end_width;
                    }
                }


                if (largeDepth == largeWidth && smallDepth == smallWidth)
                {

                    newName = "TUBE " + largeDepth + "/" + smallDepth + "X" + section.flg_thick;
                }



                else if (!isConstDepth && !isConstWidth && (section.flg_thick != section.web_thick))
                {
                    newName = "TUBE " + largeDepth + "/" + smallDepth + "X" + largeWidth + "/" + smallWidth + "X" + section.flg_thick + "X" + section.web_thick;

                }

                else if (!isConstDepth && !isConstWidth && (section.flg_thick == section.web_thick))
                {
                    newName = "TUBE " + largeDepth + "/" + smallDepth + "X" + largeWidth + "/" + smallWidth + "X" + section.flg_thick;

                }

                else if (!isConstDepth && isConstWidth)
                {
                    newName = "TUBE " + largeDepth + "/" + smallDepth + "X" + constWidth + "X" + section.flg_thick;
                }


                else if (isConstDepth && !isConstWidth)
                {
                    newName = "TUBE " + constDepth + "X" + largeWidth + "/" + smallWidth + "X" + section.flg_thick;
                }

                else if (isConstDepth && isConstWidth)
                {
                    newName = "TUBE " + constDepth + "X" + constWidth + "X" + section.flg_thick;
                }


                object secDic = new { sec = section.name, newSec = newName };

                result.Add(secDic);

            }

            return result;

        }


        public static List<object> reNamePipeTaperedsections2(List<object> pipeTaperedSectionsDimentions)
        {
            List<object> result = new List<object>();
            //var obj = new { name = profile, st_diameter = t3, st_thick = tw, end_diameter = t32, end_thick = tw2, mat = MatProp };

            double largeDepth = 0;
            double smallDepth = 0;
            double constDepth = 0;
            bool isConstDiameter = false;
            int ret = 0;

            String newName = null;

            foreach (dynamic section in pipeTaperedSectionsDimentions)
            {
                if (section.st_diameter == section.end_diameter)
                {
                    isConstDiameter = true;
                    constDepth = section.st_diameter;
                }

                else if (section.st_diameter > section.end_diameter)
                {
                    largeDepth = section.st_diameter;
                    smallDepth = section.end_diameter;

                }

                else
                {
                    largeDepth = section.st_diameter;
                    smallDepth = section.end_diameter;
                }



                if (isConstDiameter)
                {
                    newName = "PIPE " + constDepth + "X" + section.st_thick;
                }

                else
                {
                    newName = "PIPE " + largeDepth + "/" + smallDepth + "X" + section.st_thick;

                }


                object secDic = new { sec = section.name, newSec = newName };

                result.Add(secDic);

            }

            return result;




        }

        public static List<object> reNameSolidPipesections2(List<object> solidPipeSectionsDimentions)
        {
            List<object> result = new List<object>();

            //var obj = new { name = profile, diameter = t3, mat = MatProp };
            int ret = 0;

            foreach (dynamic section in solidPipeSectionsDimentions)
            {
                if (section.diameter > 200)
                {
                    string newName = "BM Cir " + section.diameter;
                    object secDic = new { sec = section.name, newSec = newName };
                    result.Add(secDic);
                }

                else
                {
                    string newName = "ROD " + section.diameter;
                    object secDic = new { sec = section.name, newSec = newName };

                    result.Add(secDic);

                }

            }
            return result;

        }

        public static List<object> reNameSolidBoxsections2(List<object> soldBOXSectionsDimentions)
        {
            List<object> result = new List<object>();

            //var obj = new { name = profile, diameter = t3, mat = MatProp };
            int ret = 0;

            foreach (dynamic section in soldBOXSectionsDimentions)
            {


                string newName = "SolidRectangle " + section.width + "X" + section.depth;
                object secDic = new { sec = section.name, newSec = newName };
                result.Add(secDic);


            }
            return result;

        }

        public static List<object> reNameChannelsections2(List<object> channelSectionsDimentions)
        {
            List<object> result = new List<object>();
            // var obj = new { name = profile, depth = t3, width = t2, flg_thick = tf, web_thick = tw, rf = rf, mat = MatProp };

            return result;
        }

        public static List<object> reNameTsections2(List<object> tSectionsDimentions)
        {
            List<object> result = new List<object>();
            // var obj = new { name = profile, depth = t3, width = t2, flg_thick = tf, web_thick = tw, rf = rf, mat = MatProp };

            return result;
        }

        public static List<object> reNameSingleAngle2(List<object> singleAngleSectionsDimentions)
        {
            List<object> result = new List<object>();
            // var obj = new { name = profile, depth = t3, width = t2, flg_thick = tf, web_thick = tw, rf = rf, mat = MatProp };

            return result;
        }



        // temporary

        public static void ChangeNamesDirectlyInSap()
        {
            List<Column> c = new List<Column>();
            List<Beam> b = new List<Beam>();


            cSapModel mySapModel = SapUtils.Get_sap_mode();
            List<Frame> frames = SapUtils.getSapFrames();


            List<object> secDic = new List<object>();

            //List<string> SapSectionS = SapUtils.Sections.getSections(mySapModel);
            List<string> SapSectionS = SapUtils.Sections.GetAllNamesAssignedToFramesInSap(mySapModel);
            Dictionary<string, Section_type> mydic = SapUtils.Sections.getSectionTypes(SapSectionS);



            List<string> iSection = mydic.Where(x => x.Value == Section_type.I_SECTION_B).ToDictionary(x => x.Key, x => x.Value).Keys.ToList();
            List<Object> IDimentions = SapUtils.Sections.Get_i_section_dimentions(mySapModel, iSection) as List<Object>;
            secDic = secDic.Union(reNameIsection2(IDimentions)).ToList();



            List<string> boxSection = mydic.Where(x => x.Value == Section_type.Hollow_box_B || x.Value == Section_type.Hollow_box_H).ToDictionary(x => x.Key, x => x.Value).Keys.ToList();
            List<Object> boxDimentions = SapUtils.Sections.Get_Box_section_dimentions(mySapModel, boxSection) as List<Object>;
            secDic = secDic.Union(reNameBoxsections2(boxDimentions)).ToList();



            List<string> pipeSection = mydic.Where(x => x.Value == Section_type.Hollow_pipe_B || x.Value == Section_type.Hollow_pipe_H).ToDictionary(x => x.Key, x => x.Value).Keys.ToList();
            List<Object> pipeDimentions = SapUtils.Sections.Get_pipe_section_dimentions(mySapModel, pipeSection) as List<Object>;
            secDic = secDic.Union(reNamePipesections2(pipeDimentions)).ToList();


            List<string> channelSection = mydic.Where(x => x.Value == Section_type.Channel_B).ToDictionary(x => x.Key, x => x.Value).Keys.ToList();
            List<Object> channelDimentions = SapUtils.Sections.Get_channel_section_dimentions(mySapModel, channelSection) as List<Object>;
            secDic = secDic.Union(reNameChannelsections2(channelDimentions)).ToList();


            List<string> iTaperedSection = mydic.Where(x => x.Value == Section_type.I_Tapered).ToDictionary(x => x.Key, x => x.Value).Keys.ToList();
            List<Object> iTaperedDimentions = SapUtils.Sections.Get_Tapered_i_section_dimentions(mySapModel, iTaperedSection) as List<Object>;
            secDic = secDic.Union(reNameITaperedsections2(iTaperedDimentions)).ToList();



            List<string> boxTaperedSection = mydic.Where(x => x.Value == Section_type.Box_Tapered).ToDictionary(x => x.Key, x => x.Value).Keys.ToList();
            List<Object> boxTaperedDimentions = SapUtils.Sections.Get_Tapered_BOX_section_dimentions(mySapModel, boxTaperedSection) as List<Object>;
            secDic = secDic.Union(reNameboxTaperedsections2(boxTaperedDimentions)).ToList();



            List<string> pipeTaperedSection = mydic.Where(x => x.Value == Section_type.Pipe_Tapered).ToDictionary(x => x.Key, x => x.Value).Keys.ToList();
            List<Object> pipeTaperedDimentions = SapUtils.Sections.Get_Tapered_PIPE_section_dimentions(mySapModel, pipeTaperedSection) as List<Object>;
            secDic = secDic.Union(reNamePipeTaperedsections2(pipeTaperedDimentions)).ToList();



            List<string> solidPipeSection = mydic.Where(x => x.Value == Section_type.Solid_pipe).ToDictionary(x => x.Key, x => x.Value).Keys.ToList();
            List<Object> sloidPipeDimentions = SapUtils.Sections.Get_solid_pipe_section_dimentions(mySapModel, solidPipeSection) as List<Object>;
            secDic = secDic.Union(reNameSolidPipesections2(sloidPipeDimentions)).ToList();



            List<string> solidBoxSection = mydic.Where(x => x.Value == Section_type.Solid_box).ToDictionary(x => x.Key, x => x.Value).Keys.ToList();
            List<Object> sloidBoxDimentions = SapUtils.Sections.Get_solid_rect_section_dimentions(mySapModel, solidBoxSection) as List<Object>;
            secDic = secDic.Union(reNameSolidBoxsections2(sloidBoxDimentions)).ToList();



            List<string> teeSection = mydic.Where(x => x.Value == Section_type.T_Section_B).ToDictionary(x => x.Key, x => x.Value).Keys.ToList();
            List<Object> teeDimentions = SapUtils.Sections.Get_Tee_section_dimentions(mySapModel, teeSection) as List<Object>;
            secDic = secDic.Union(reNameTsections2(teeDimentions)).ToList();



            List<string> singleAngleSection = mydic.Where(x => x.Value == Section_type.S_Angle_B).ToDictionary(x => x.Key, x => x.Value).Keys.ToList();
            List<Object> singleAngleDimentions = SapUtils.Sections.Get_Sig_angle_section_dimentions(mySapModel, singleAngleSection) as List<Object>;
            secDic = secDic.Union(reNameSingleAngle2(singleAngleDimentions)).ToList();



            List<string> concrBeamRectSection = mydic.Where(x => x.Value == Section_type.Concrete_rect).ToDictionary(x => x.Key, x => x.Value).Keys.ToList();
            List<Object> ConcrBeamRectDimentions = SapUtils.Sections.Get_solid_rect_section_dimentions(mySapModel, concrBeamRectSection) as List<Object>;
            secDic = secDic.Union(reNameSolidBoxsections2(ConcrBeamRectDimentions)).ToList();



            List<string> concrBeamPipeSection = mydic.Where(x => x.Value == Section_type.Concrete_circ).ToDictionary(x => x.Key, x => x.Value).Keys.ToList();
            List<Object> ConcrBeamPipeDimentions = SapUtils.Sections.Get_solid_pipe_section_dimentions(mySapModel, concrBeamPipeSection) as List<Object>;
            secDic = secDic.Union(reNameSolidPipesections2(ConcrBeamPipeDimentions)).ToList();


            foreach (string sec in SapSectionS)
            {

                object currentSec = secDic.Cast<dynamic>().FirstOrDefault(x => x.sec == sec);

                if (currentSec != null)

                {
                    string newSec = secDic.Cast<dynamic>().FirstOrDefault(x => x.sec == sec).newSec;

                    mySapModel.PropFrame.ChangeName(sec, newSec);

                }




            }


        }

    }
}
