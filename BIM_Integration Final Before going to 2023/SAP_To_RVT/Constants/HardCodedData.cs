using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Analysis;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace BIM_Integration
{
    internal class HardCodedData
{






public static double AngleValueInDegree = 0;


        public static string LengthParameter = "Length";
        public static string CutLengthParameter = "Cut Length";

        public static string VolumeParameter = "Volume";
        public static double GamaSteel = 7.85;
        public static string filePathForSerilization = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\Bim.txt";

        public static string revitSymbol = "Revit_Profile";
        public static string revitCount = "R_No_of_Members";
        public static string revitLength = "R_Total_Length_Meter";
        public static string revitTonnage = "R_Total_Tonnage_Ton";

        public static string sapSection = "Sap_Prfile";
        public static string sapCount = "S_No_of_Members";
        public static string sapLength = "S_Total_Length_Meter";
        public static string sapTonnage = "S_Total_Tonnage_Ton";
        public static string sapXSecArea = "S_X_sec_Area_MM";


        public static string RevitMainGroup = "ACM_SE_SubZone_Level_Name"; //(NewName)
        public static string RevitSubGroup = "ACM_SE_SubZone_Level_Name";



        public static string RevitSubGroupUsage = "ACM_SE_Member Usage";

        public static string SapIDParameter = "ACM_SE_SBM_SapID";
        public static string EtabsIDParameter = "ACM_SE_SBM_EtabsID";
        public static string StaadIDParameter = "ACM_SE_SBM_StaadID";
        public static string MaterialParameter = "ACM_SE_SBM_AnalysisMaterial";

        public static string StaadAnalysisModel = "Staad";
        public static string SAPAnalysisModel = "Sap";
        public static string ETABSAnalysisModel = "Etabs";

        #region Realeases

        public static string StartM3Release = "ACM_SE_StartMajorM_Release";
        public static string StartM2Release = "ACM_SE_StartMinorM_Release";
        public static string EndM3Release = "ACM_SE_EndMajorM_Release";
        public static string EndM2Release = "ACM_SE_EndMinorM_Release";


        public static string ConnectionType = "ACM_SE_ConnectionType";

        public static string BasePlateID = "ACM_SE_BasePlate_ID";

        #endregion


        public static Dictionary<string, string> MemberSubGroupDic = new Dictionary<string, string>
        {
            { "SB", "SECONDARY BEAM" },
            { "MB", "STEEL MAIN BEAM" },
            { "SC", "STEEL COLUMN" },
            { "SV", "VERTICAL MEMBER" },
            { "SD", "DIAGONAL MEMBER" },
            { "TC", "TOP CHORD" },
            { "BC", "BOTTOM CHORD" },
            { "ST", "STRUTS" },
            { "HB", "HORIZONTAL BRACING" },
            { "VB", "VERTICAL BRACING" },
            { "SG", "STAIR STRINGER" }
        };


        public static Dictionary<string, string> ColumnMarkDic = new Dictionary<string, string>
        {
            { "Mark", "ACM_SE_Main Type ID" },
            { "SubMark", "ACM_SE_Design Section" },
            { "Depth", "ACM_SE_Overall Depth" },
            { "Width", "ACM_SE_Overall Width" },
            { "Diameter", "DIAGONAL MEMBER" },
            { "Description", "TOP CHORD" },
            { "Designation", "ACM_SE_Column RFT Detail Number" },
            { "Count", "STRUTS" },
            { "ColumnType", "STRUTS" },


        };


        public static List<string> allowedParams = new List<string>
                    {
                        SapIDParameter,
                        RevitMainGroup,
                        RevitSubGroup,
                        RevitSubGroupUsage,
                        SapIDParameter,
                        EtabsIDParameter,
                        StaadIDParameter,
                        MaterialParameter,
                        StartM3Release,
                        StartM2Release,
                        EndM3Release,
                        EndM2Release,
                        ConnectionType,
                        BasePlateID,


                    };

    }
}
