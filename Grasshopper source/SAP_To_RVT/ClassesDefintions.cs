using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using BIM_Integration.STAAD;
using ETABSv1;

//using Eto.Forms;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;
using Rhino;
using Rhino.Geometry;
using Rhino.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Numerics;
using System.Reflection;
using Arc = Autodesk.Revit.DB.Arc;
using Curve = Autodesk.Revit.DB.Curve;
using Document = Autodesk.Revit.DB.Document;
using Ellipse = Autodesk.Revit.DB.Ellipse;
using Excel = Microsoft.Office.Interop.Excel;
using Line = Autodesk.Revit.DB.Line;
using Parameter = Autodesk.Revit.DB.Parameter;
using Plane = Rhino.Geometry.Plane;
using Quaternion = System.Numerics.Quaternion;

namespace BIM_Integration
{

    #region Classes_declerations

    public class Frame
    {
        public static int chopSize = 1000;
      

        #region  props
        public string ID { get; set; }
        public string SapID { get; set; }

        public string EtabsID { get; set; }


        public string ParentID { get; set; }

        public string RvtID { get; set; }

        public string RhinoID { get; set; }
        public Point StartPoint { get; set; }
        public Point EndPoint { get; set; }
        public Point MidPoint { get; set; }

        public int SerialNumber { get; set; }
        public ColumnRebarDesignData ColumnsRebarDesign {  get; set; }  

        public double Beta { get; set; }
        public string Shape { get; set; }
        #region Assignments
        public string distLoadAssignments { get; set; } 


        #endregion

        public Section_type SecType { get; set; }

        public SectType FamilyType { get; set; }
        #region Section Parameters
        public string Sec { get; set; }
        public string ExistSec { get; set; }
        public int RefSecNo { get; set; }
        public double secArea { get; set; }

        public double SecDepth { get; set; }

        public double SecWidth { get; set; }

        #endregion

        public ElementId revitEleId { get; set; }
        public InsertionPoint InsPt = InsertionPoint.Centroid;
        public double StartZOffset { get; set; }
        public double EndZOoffset { get; set; }
        public bool StartReleased { get; set; }
        public bool EndReleased { get; set; }

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


        #region Groups parameters
        public string Group { get; set; }


        public List<string> Groups { get; set; }
        public List<string> Groups2 { get; set; }
        public string LayerName { get; set; }

        public string MainGroup { get; set; }
        public string SubGroup { get; set; }
        public string SubGroupUsage { get; set; }
        public string Mark { get; set; }
        public string Designation { get; set; }

        public double Depth { get; set; }



        public double Width{ get; set; }
        public double Diameter { get; set; }

        public string Story { get; set; }
        public string EtabsTopStory { get; set; }
        public string EtabsBotStory { get; set; }


        public string RevitTopLevel { get; set; }
        public string RevitBotLevel { get; set; }


        public double RevitTopLevelOffset { get; set; }
        public double RevitBotLevelOffset { get; set; }


        #endregion
        public double Length { get; set; }
        public double EndLengthOffset { get; set; }
        public double StartLengthOffset { get; set; }

        public MemberType MemberType { get; set; }

        public int rowNo { get; set; }


        public string sectionTypeName { get; set; }

        public Autodesk.Revit.DB.Curve revitCurve { get; set; }

        public Rhino.Geometry.Curve RhinoCurve { get; set; }
        public bool IsNonPresmatic { get; set; }

        public string BasePlateID { get; set; }


        #endregion

        public Frame()
        {
            StartPoint = new Point();
            EndPoint = new Point();
            MidPoint = new Point();
            ColumnsRebarDesign = new ColumnRebarDesignData();

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


        public static void setDataInExcelBulk(List<Frame> analysisElements, List<FamilyInstance> revitElements, ForgeTypeId dis, Microsoft.Office.Interop.Excel.Worksheet ws, ref object[,] dataToExcel, string action)
        {

            #region get_the_cells
            Excel.Range Sap_ID = ws.Range["A:A"];
            Excel.Range RevitName = ws.Range["Y1"];
            Excel.Range RevitFamily = ws.Range["Z1"];
            Excel.Range RevitID = ws.Range["AA1"];
            Excel.Range StartX = ws.Range["AB1"];

            int count = analysisElements.Count;



            Document  doc =null;

            if (revitElements.Count>0)
            {
                doc = revitElements[0].Document;
            }

            //object[,] arr = new object[analysisElements.Count, 9];

            for (int i = 0; i < analysisElements.Count; i++)
            {

                FamilyInstance instance = revitElements[i];
                Frame bm = analysisElements[i];

                if (!instance.IsValidObject)
                {
                    dataToExcel[bm.rowNo, 0] = "DUPLICATED";
                    dataToExcel[bm.rowNo, 1] = "DUPLICATED";
                    dataToExcel[bm.rowNo, 2] = "DUPLICATED";
                    dataToExcel[bm.rowNo, 3] = "DUPLICATED";
                    dataToExcel[bm.rowNo, 4] = "DUPLICATED";
                    dataToExcel[bm.rowNo, 5] = "DUPLICATED";
                    dataToExcel[bm.rowNo, 6] = "DUPLICATED";
                    dataToExcel[bm.rowNo, 7] = "DUPLICATED";
                    dataToExcel[bm.rowNo, 8] = "DUPLICATED";
                    continue;
                }

                AnalyticalToPhysicalAssociationManager analyManger = AnalyticalToPhysicalAssociationManager.GetAnalyticalToPhysicalAssociationManager(doc);
                AnalyticalElement analytInstance = doc.GetElement(analyManger.GetAssociatedElementId(instance.Id)) as AnalyticalElement;

                XYZ startPoint = analytInstance.GetCurve().GetEndPoint(0);
                XYZ endPoint = analytInstance.GetCurve().GetEndPoint(1);


                dataToExcel[bm.rowNo - 8, 0] = instance.Symbol.Name;
                dataToExcel[bm.rowNo - 8, 1] = instance.Symbol.FamilyName;
                dataToExcel[bm.rowNo - 8, 2] = instance.Id.ToString();
                dataToExcel[bm.rowNo - 8, 3] = Convert.ToString(UnitUtils.ConvertFromInternalUnits(startPoint.X, dis));
                dataToExcel[bm.rowNo - 8, 4] = Convert.ToString(UnitUtils.ConvertFromInternalUnits(startPoint.Y, dis));
                dataToExcel[bm.rowNo - 8, 5] = Convert.ToString(UnitUtils.ConvertFromInternalUnits(startPoint.Z, dis));
                dataToExcel[bm.rowNo - 8, 6] = Convert.ToString(UnitUtils.ConvertFromInternalUnits(endPoint.X, dis));
                dataToExcel[bm.rowNo - 8, 7] = Convert.ToString(UnitUtils.ConvertFromInternalUnits(endPoint.Y, dis));
                dataToExcel[bm.rowNo - 8, 8] = Convert.ToString(UnitUtils.ConvertFromInternalUnits(endPoint.Z, dis));

            }


            if (action == "Do")
            {
                Microsoft.Office.Interop.Excel.Range startCell = (Range)ws.Cells[7, 25];
                Microsoft.Office.Interop.Excel.Range endCell = (Range)ws.Cells[dataToExcel.GetLength(0) + 7 - 1, 33];
                Microsoft.Office.Interop.Excel.Range range = (Range)ws.Range[startCell, endCell];
                range.Value2 = dataToExcel;
            }



            #endregion

        }

        // new
        public static void UpdateAnlaysisDataInExcel(List<Frame> analysisElements, List<Frame> revitElements, Microsoft.Office.Interop.Excel.Worksheet ws, ref object[,] dataToExcel, string action, string typeOfUpdate)
        {



            #region get_the_cells
            Excel.Range Sap_ID = ws.Range["A:A"];
            Excel.Range RevitName = ws.Range["Y1"];
            Excel.Range RevitFamily = ws.Range["Z1"];
            Excel.Range RevitID = ws.Range["AA1"];
            Excel.Range StartX = ws.Range["AB1"];



            for (int i = 0; i < revitElements.Count; i++)
            {

                try
                {
                    Frame analysisElement = analysisElements[i];
                    Frame revitElement = revitElements[i];

                    if (typeOfUpdate == "All")
                    {
                        dataToExcel[revitElement.rowNo - 7, 0] = analysisElement.ID;
                        dataToExcel[revitElement.rowNo - 7, 1] = analysisElement.StartPoint.Id;
                        dataToExcel[revitElement.rowNo - 7, 2] = analysisElement.EndPoint.Id;
                        dataToExcel[revitElement.rowNo - 7, 3] = analysisElement.StartPoint.X;
                        dataToExcel[revitElement.rowNo - 7, 4] = analysisElement.StartPoint.Y;
                        dataToExcel[revitElement.rowNo - 7, 5] = analysisElement.StartPoint.Z;
                        dataToExcel[revitElement.rowNo - 7, 6] = analysisElement.EndPoint.X;
                        dataToExcel[revitElement.rowNo - 7, 7] = analysisElement.EndPoint.Y;
                        dataToExcel[revitElement.rowNo - 7, 8] = analysisElement.EndPoint.Z;
                        dataToExcel[revitElement.rowNo - 7, 9] = analysisElement.Sec;
                        dataToExcel[revitElement.rowNo - 7, 10] = analysisElement.Lib;
                        dataToExcel[revitElement.rowNo - 7, 11] = analysisElement.Shape;
                        dataToExcel[revitElement.rowNo - 7, 12] = analysisElement.IsBUS;
                        dataToExcel[revitElement.rowNo - 7, 13] = analysisElement.IsCable;
                        dataToExcel[revitElement.rowNo - 7, 14] = analysisElement.IsLink;
                        dataToExcel[revitElement.rowNo - 7, 15] = analysisElement.IsConc;
                        dataToExcel[revitElement.rowNo - 7, 16] = analysisElement.Group;
                        dataToExcel[revitElement.rowNo - 7, 17] = analysisElement.Material;
                        dataToExcel[revitElement.rowNo - 7, 18] = analysisElement.Length;
                        dataToExcel[revitElement.rowNo - 7, 19] = analysisElement.Beta;
                    }
                    else if (typeOfUpdate == "Sec")
                    {
                        dataToExcel[revitElement.rowNo - 7, 0] = analysisElement.ID;
                        dataToExcel[revitElement.rowNo - 7, 9] = analysisElement.Sec;
                        dataToExcel[revitElement.rowNo - 7, 10] = analysisElement.Lib;
                        dataToExcel[revitElement.rowNo - 7, 11] = analysisElement.Shape;
                        dataToExcel[revitElement.rowNo - 7, 12] = analysisElement.IsBUS;
                        dataToExcel[revitElement.rowNo - 7, 13] = analysisElement.IsCable;
                        dataToExcel[revitElement.rowNo - 7, 14] = analysisElement.IsLink;
                        dataToExcel[revitElement.rowNo - 7, 15] = analysisElement.IsConc;
                        dataToExcel[revitElement.rowNo - 7, 16] = analysisElement.Group;
                        dataToExcel[revitElement.rowNo - 7, 17] = analysisElement.Material;
                        dataToExcel[revitElement.rowNo - 7, 18] = analysisElement.Length;
                        dataToExcel[revitElement.rowNo - 7, 19] = analysisElement.Beta;
                    }

                    else if (typeOfUpdate == "Geo")
                    {
                        dataToExcel[revitElement.rowNo - 7, 0] = analysisElement.ID;
                        dataToExcel[revitElement.rowNo - 7, 1] = analysisElement.StartPoint.Id;
                        dataToExcel[revitElement.rowNo - 7, 2] = analysisElement.EndPoint.Id;
                        dataToExcel[revitElement.rowNo - 7, 3] = analysisElement.StartPoint.X;
                        dataToExcel[revitElement.rowNo - 7, 4] = analysisElement.StartPoint.Y;
                        dataToExcel[revitElement.rowNo - 7, 5] = analysisElement.StartPoint.Z;
                        dataToExcel[revitElement.rowNo - 7, 6] = analysisElement.EndPoint.X;
                        dataToExcel[revitElement.rowNo - 7, 7] = analysisElement.EndPoint.Y;
                        dataToExcel[revitElement.rowNo - 7, 8] = analysisElement.EndPoint.Z;
                    }
                }

                catch (Exception ex)
                {

                    continue;
                }



            }


            if (action == "Do")
            {
                try
                {
                    Microsoft.Office.Interop.Excel.Range startCell = (Range)ws.Cells[7, 1];
                    Microsoft.Office.Interop.Excel.Range endCell = (Range)ws.Cells[dataToExcel.GetLength(0) + 7 - 1, 20];
                    Microsoft.Office.Interop.Excel.Range range = (Range)ws.Range[startCell, endCell];
                    range.Value2 = dataToExcel;
                }
                catch (Exception ex)
                {
                    return;

                }

            }



            #endregion

        }

        public static bool IsSameMemberType(Frame sapFrame, Frame revitFrame)
        {
            if (sapFrame.MemberType == revitFrame.MemberType)
            {
                return true;
            }

            else if (sapFrame.MemberType == MemberType.SlanttedColumn && revitFrame.MemberType == MemberType.Column)
            {
                sapFrame.Beta = sapFrame.Beta * -1;
                sapFrame.Beta += 90;

                return true;
            }

            else if (sapFrame.MemberType == MemberType.Column && revitFrame.MemberType == MemberType.SlanttedColumn)
            {
                sapFrame.Beta = sapFrame.Beta - 90;
                sapFrame.Beta = sapFrame.Beta * -1;


                return true;
            }

            return false;


        }

        #region assignments

        public static void ParseDistrLoadAssignmentsString(string dataString, out int ret, out int NumberItems2, out string[] FrameName, out string[] LoadPat, out int[] MyType2, out string[] CSys, out int[] Dir, out double[] RD1, out double[] RD2, out double[] Dist1, out double[] Dist2, out double[] Val1, out double[] Val2)
        {
            string[] dataArray = dataString.Split(',');

            ret = int.Parse(dataArray[0]);
            NumberItems2 = int.Parse(dataArray[1]);
            FrameName = dataArray[2].Split(',');
            LoadPat = dataArray[3].Split(',');
            MyType2 = Array.ConvertAll(dataArray[4].Split(','), int.Parse);
            CSys = dataArray[5].Split(',');
            Dir = Array.ConvertAll(dataArray[6].Split(','), int.Parse);
            RD1 = Array.ConvertAll(dataArray[7].Split(','), double.Parse);
            RD2 = Array.ConvertAll(dataArray[8].Split(','), double.Parse);
            Dist1 = Array.ConvertAll(dataArray[9].Split(','), double.Parse);
            Dist2 = Array.ConvertAll(dataArray[10].Split(','), double.Parse);
            Val1 = Array.ConvertAll(dataArray[11].Split(','), double.Parse);
            Val2 = Array.ConvertAll(dataArray[12].Split(','), double.Parse);
        }

        public static void ParseRealeasesAssignmentsString(string resultString, out bool[] ii, out bool[] jj, out double[] StartValue, out double[] EndValue)
        {
            string[] parts = resultString.Split(';');

            ii = parts[0].Split(',').Select(bool.Parse).ToArray();
            jj = parts[1].Split(',').Select(bool.Parse).ToArray();
            StartValue = parts[2].Split(',').Select(double.Parse).ToArray();
            EndValue = parts[3].Split(',').Select(double.Parse).ToArray();
        }



        #endregion


        #region UnUsed
        public static void setDataInExcel(Frame bm, FamilyInstance instance, ForgeTypeId dis, Microsoft.Office.Interop.Excel.Worksheet ws)
        {



            #region get_the_cells
            Excel.Range Sap_ID = ws.Range["A:A"];
            Excel.Range RevitName = ws.Range["Y1"];
            Excel.Range RevitFamily = ws.Range["Z1"];
            Excel.Range RevitID = ws.Range["AA1"];
            Excel.Range StartX = ws.Range["AB1"];


            if (!instance.IsValidObject)
            {
                RevitName.Offset[bm.rowNo - 1, 0].Value = "DUPLICATED";
                return;
            }

            Document doc = instance.Document;
            AnalyticalToPhysicalAssociationManager analyManger = AnalyticalToPhysicalAssociationManager.GetAnalyticalToPhysicalAssociationManager(doc);
            AnalyticalElement analytInstance = doc.GetElement(analyManger.GetAssociatedElementId(instance.Id)) as AnalyticalElement;

            XYZ startPoint = analytInstance.GetCurve().GetEndPoint(0);
            XYZ endPoint = analytInstance.GetCurve().GetEndPoint(1);



            RevitName.Offset[bm.rowNo - 1, 0].Value = instance.Symbol.Name;
            RevitFamily.Offset[bm.rowNo - 1, 0].Value = instance.Symbol.FamilyName;
            RevitID.Offset[bm.rowNo - 1, 0].Value = instance.Id.ToString();

            StartX.Offset[bm.rowNo - 1, 0].Value = Convert.ToString(UnitUtils.ConvertFromInternalUnits(startPoint.X, dis));
            StartX.Offset[bm.rowNo - 1, 1].Value = Convert.ToString(UnitUtils.ConvertFromInternalUnits(startPoint.Y, dis));
            StartX.Offset[bm.rowNo - 1, 2].Value = Convert.ToString(UnitUtils.ConvertFromInternalUnits(startPoint.Z, dis));

            StartX.Offset[bm.rowNo - 1, 3].Value = Convert.ToString(UnitUtils.ConvertFromInternalUnits(endPoint.X, dis));
            StartX.Offset[bm.rowNo - 1, 4].Value = Convert.ToString(UnitUtils.ConvertFromInternalUnits(endPoint.Y, dis));
            StartX.Offset[bm.rowNo - 1, 5].Value = Convert.ToString(UnitUtils.ConvertFromInternalUnits(endPoint.Z, dis));



            #endregion

        }
        public static void seDataInExceDummy(Frame bm, FamilyInstance instance, ForgeTypeId dis, Microsoft.Office.Interop.Excel.Worksheet ws)
        {
            Excel.Range Sap_ID = ws.Range["A:A"];
            Excel.Range RevitName = ws.Range["Y1"];
            Excel.Range RevitFamily = ws.Range["Z1"];
            Excel.Range RevitID = ws.Range["AA1"];
            Excel.Range StartX = ws.Range["AB1"];


            if (!instance.IsValidObject)
            {
                RevitName.Offset[bm.rowNo - 1, 0].Value = "DUPLICATED";
                return;
            }
            Document doc = instance.Document;
            AnalyticalToPhysicalAssociationManager analyManger = AnalyticalToPhysicalAssociationManager.GetAnalyticalToPhysicalAssociationManager(doc);
            AnalyticalElement analytInstance = doc.GetElement(analyManger.GetAssociatedElementId(instance.Id)) as AnalyticalElement;

            XYZ startPoint = analytInstance.GetCurve().GetEndPoint(0);
            XYZ endPoint = analytInstance.GetCurve().GetEndPoint(1);



            RevitName.Offset[bm.rowNo, 0].Value = instance.Symbol.Name;
            RevitFamily.Offset[bm.rowNo, 0].Value = instance.Symbol.FamilyName;
            RevitID.Offset[bm.rowNo, 0].Value = instance.Id.ToString();

            StartX.Offset[bm.rowNo, 0].Value = Convert.ToString(UnitUtils.ConvertFromInternalUnits(startPoint.X, dis));
            StartX.Offset[bm.rowNo, 1].Value = Convert.ToString(UnitUtils.ConvertFromInternalUnits(startPoint.Y, dis));
            StartX.Offset[bm.rowNo, 2].Value = Convert.ToString(UnitUtils.ConvertFromInternalUnits(startPoint.Z, dis));

            StartX.Offset[bm.rowNo, 3].Value = Convert.ToString(UnitUtils.ConvertFromInternalUnits(endPoint.X, dis));
            StartX.Offset[bm.rowNo, 4].Value = Convert.ToString(UnitUtils.ConvertFromInternalUnits(endPoint.Y, dis));
            StartX.Offset[bm.rowNo, 5].Value = Convert.ToString(UnitUtils.ConvertFromInternalUnits(endPoint.Z, dis));


        }

        #endregion


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


        public Beam(FamilyInstance instance)
        {
            if (instance == null)
            {

            }
            else
            {
                Document doc = instance.Document;
                AnalyticalToPhysicalAssociationManager analyManger = AnalyticalToPhysicalAssociationManager.GetAnalyticalToPhysicalAssociationManager(doc);
                AnalyticalElement analytInstance = doc.GetElement(analyManger.GetAssociatedElementId(instance.Id)) as AnalyticalElement;

                this.RvtID = instance.Id.ToString();
                this.RvtName = instance.Symbol.Name;
                this.RvtFamily = instance.Symbol.FamilyName;
                this.StartPoint = new Point(analytInstance.GetCurve().GetEndPoint(0));
                this.EndPoint = new Point(analytInstance.GetCurve().GetEndPoint(1));
            }

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


        public static List<ElementId> CreateAllBeams3(Document doc, List<Frame> beams, ForgeTypeId dis, Dictionary<string, FamilySymbol> dic, Microsoft.Office.Interop.Excel.Worksheet ws, string analysisModel, object[,] dataToExcel)
        {

            List<ElementId> result = new List<ElementId>();
            List<FamilyInstance> revitInstances = new List<FamilyInstance>();

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

                                revitSection = dic[bm.Sec];

                                if (!revitSection.IsActive)
                                    revitSection.Activate();

                                FamilyInstance instance = doc.Create.NewFamilyInstance(line, revitSection, lvl0, StructuralType.Beam);
                                instance.LookupParameter("Cross-Section Rotation").Set(angle);
                                Parameter sss = instance.LookupParameter("z Justification");
                                int zjustf = (int)ZJustification.Center;
                                sss.Set(zjustf);
                                Autodesk.Revit.DB.Structure.StructuralFramingUtils.DisallowJoinAtEnd(instance, 0);
                                Autodesk.Revit.DB.Structure.StructuralFramingUtils.DisallowJoinAtEnd(instance, 1);



                                if (analysisModel == "Staad")
                                {
                                    RevitUtils.setParametersSTAAD(instance, bm);
                                }

                                else if (analysisModel == "Sap")
                                {
                                    RevitUtils.setParametersSAP(instance, bm);

                                }

                                else if (analysisModel == "Etabs")
                                {
                                    RevitUtils.setParametersEtabs(instance, bm);

                                }



                                revitInstances.Add(instance);
                                result.Add(instance.Id);


                            }


                            catch (Exception ex2)

                            {
                                System.Diagnostics.StackTrace trace = new System.Diagnostics.StackTrace(ex2, true);
                                RevitUtils.Log.Add("Error in CreateAllBeams3 Method... beam no" + bm.ID + ex2.Message + trace.GetFrame(trace.FrameCount - 1).GetFileLineNumber());

                                //TaskDialog.Show("Error", "error in beam no" + bm.ID + ex2.Message + trace.GetFrame(trace.FrameCount - 1).GetFileLineNumber());

                                continue;

                            }


                        }

                        trns.Commit();


                    }

                }

                if (ws != null)
                {
                    Frame.setDataInExcelBulk(beams, revitInstances, dis, ws, ref dataToExcel, "not");
                }

                tg.Assimilate();


            }

            return result;
        }

        public static void updateRevitSectionsForBeams2(Document doc, List<SapRevitFrame> MappedBeams, Dictionary<string, FamilySymbol> dic, ForgeTypeId dis, Microsoft.Office.Interop.Excel.Worksheet ws, string analysisModel, object[,] dataToExcel)
        {

            List<FamilyInstance> revitInstances = new List<FamilyInstance>();

            RevitServices.Transactions.TransactionManager.Instance.ForceCloseTransaction();

            using (TransactionGroup tg = new TransactionGroup(doc, "aaa"))
            {
                tg.Start();

                var chopSize = Frame.chopSize;
                var MappedBeamsChuncks = MappedBeams.Chunk(chopSize).ToList();


                foreach (var mappedBeamChunck in MappedBeamsChuncks)
                {

                    using (Transaction trns = new Transaction(doc, "trns"))
                    {
                        trns.Start();
                        FailureHandlingOptions failOpt = trns.GetFailureHandlingOptions();
                        failOpt.SetFailuresPreprocessor(new RoomWarningSwallower());
                        trns.SetFailureHandlingOptions(failOpt);


                        foreach (var mappedBeam in mappedBeamChunck.AsParallel())
                        {


                            try
                            {

                                FamilySymbol revitSection = dic[mappedBeam.sapFrame.Sec];

                                if (revitSection == null)
                                {
                                    continue;
                                }

                                // NO update
                                else if (revitSection.Name == mappedBeam.revitFrame.RvtName)
                                {
                                    FamilyInstance instance = (FamilyInstance)doc.GetElement(mappedBeam.revitFrame.revitEleId);
                                    revitInstances.Add(instance);

                                    double angle = (mappedBeam.sapFrame.Beta) * Math.PI / 180;
                                    instance.LookupParameter("Cross-Section Rotation").Set(angle);


                                    if (analysisModel == "Staad")
                                    {
                                        RevitUtils.setParametersSTAAD(instance, mappedBeam.sapFrame);
                                    }

                                    else if (analysisModel == "Sap")
                                    {
                                        RevitUtils.setParametersSAP(instance, mappedBeam.sapFrame);

                                    }


                                    else if (analysisModel == "Etabs")
                                    {
                                        RevitUtils.setParametersEtabs(instance, mappedBeam.sapFrame);

                                    }


                                }


                                // update
                                else
                                {
                                    FamilyInstance instance = (FamilyInstance)doc.GetElement(mappedBeam.revitFrame.revitEleId);
                                    revitInstances.Add(instance);

                                    double angle = (mappedBeam.sapFrame.Beta) * Math.PI / 180;
                                    instance.LookupParameter("Cross-Section Rotation").Set(angle);

                                    instance.Symbol = revitSection;

                                    if (analysisModel == "Staad")
                                    {
                                        RevitUtils.setParametersSTAAD(instance, mappedBeam.sapFrame);
                                    }

                                    else if (analysisModel == "Sap")
                                    {
                                        RevitUtils.setParametersSAP(instance, mappedBeam.sapFrame);

                                    }
                                    else if (analysisModel == "Etabs")
                                    {
                                        RevitUtils.setParametersEtabs(instance, mappedBeam.sapFrame);

                                    }
                                }

                            }

                            catch (Exception ex2)
                            {
                                System.Diagnostics.StackTrace trace = new System.Diagnostics.StackTrace(ex2, true);
                                RevitUtils.Log.Add("Error in updateRevitSectionsForBeams2 Method  sapBeam no " + mappedBeam.sapFrame.ID + ex2.Message + trace.GetFrame(trace.FrameCount - 1).GetFileLineNumber());
                                continue;
                            }

                            trns.Commit();




                        }


                    }




                }

                tg.Assimilate();
            }

            if (ws != null)
            {
                List<Frame> sapMappedBeams = MappedBeams.Select(x => x.sapFrame).ToList();
                Frame.setDataInExcelBulk(sapMappedBeams, revitInstances, dis, (Microsoft.Office.Interop.Excel.Worksheet)ws, ref dataToExcel, "NO");

            }
        }

        // should be modified to be like sap 
        public static void updateEtabsSections(List<SapRevitFrame> MappedBeams, ETABSv1.cSapModel mysapmodel, Document doc, string type, Dictionary<string, string> secDicForBeams, Dictionary<string, string> secDicForColumns)
        {
            string materialName = null;
            string sAuto = null;
            String sapSection = null;

            foreach (var mappedBeam in MappedBeams.AsParallel())
            {

                //mysapmodel.View.RefreshView();
                if (type == "BM")
                {
                    sapSection = secDicForBeams[mappedBeam.revitFrame.RvtName];

                }
                else if (type == "COL")
                {
                    sapSection = secDicForColumns[mappedBeam.revitFrame.RvtName];
                }




                if (sapSection == null)
                {
                    continue;
                }


                else if (sapSection == mappedBeam.sapFrame.Sec)
                {

                    // Update rotation 
                    mysapmodel.FrameObj.SetLocalAxes(mappedBeam.sapFrame.ID, mappedBeam.revitFrame.Beta);


                    // set Parameters in revit    
                    mysapmodel.FrameObj.GetSection(mappedBeam.sapFrame.ID, ref sapSection, ref sAuto);
                    mysapmodel.PropFrame.GetMaterial(sapSection, ref materialName);


                    FamilyInstance instance = (FamilyInstance)doc.GetElement(mappedBeam.revitFrame.revitEleId);
                    using (Transaction trns = new Transaction(doc, "aa"))
                    {
                        trns.Start();
                        FailureHandlingOptions failOpt = trns.GetFailureHandlingOptions();

                        RevitUtils.setParametersEtabs(instance, mappedBeam.sapFrame);



                        trns.Commit();
                    }
                    continue;

                }


                else
                {
                    // updateSapSection
                    int ret = mysapmodel.FrameObj.SetSection(mappedBeam.sapFrame.ID, sapSection);
                    if (ret == 1)
                    {
                        mysapmodel.FrameObj.SetSection(mappedBeam.sapFrame.ID, "dummySection");
                    }

                    // Update rotation 
                    mysapmodel.FrameObj.SetLocalAxes(mappedBeam.sapFrame.ID, mappedBeam.revitFrame.Beta);


                    // set Parameters in revit    
                    mysapmodel.FrameObj.GetSection(mappedBeam.sapFrame.ID, ref sapSection, ref sAuto);
                    mysapmodel.PropFrame.GetMaterial(sapSection, ref materialName);


                    FamilyInstance instance = (FamilyInstance)doc.GetElement(mappedBeam.revitFrame.revitEleId);
                    using (Transaction trns = new Transaction(doc, "aa"))
                    {
                        trns.Start();
                        FailureHandlingOptions failOpt = trns.GetFailureHandlingOptions();
                        RevitUtils.setParametersEtabs(instance, mappedBeam.sapFrame);

                        trns.Commit();
                    }
                    continue;

                }

            }


        }
        public static void updateStaadSections(List<SapRevitFrame> MappedBeams, Document doc, string type)
        {


            Dictionary<string, int> revitSTaadSectionDic = new Dictionary<string, int>();
            List<Frame> FilteredRevitFrameByName = MappedBeams.Select(x => x.revitFrame).ToList().GroupBy(x => x.RvtName).Select(x => x.FirstOrDefault()).ToList();


            if (type == "BM")
            {
                revitSTaadSectionDic = StaadUtils.createStaadSectionForBeams(FilteredRevitFrameByName);

            }

            else if (type == "COL")
            {
                revitSTaadSectionDic = StaadUtils.createStaadSectionForColumns(FilteredRevitFrameByName);

            }

            foreach (var mappedBeam in MappedBeams.AsParallel())
            {

                // Update section 
                StaadUtils.property.AssignBeamProperty(mappedBeam.sapFrame.ID, revitSTaadSectionDic[mappedBeam.revitFrame.RvtName]);
                mappedBeam.sapFrame.Sec = StaadUtils.property.GetBeamSectionName(mappedBeam.sapFrame.ID);


                // update rotation
                StaadUtils.property.AssignBetaAngle(mappedBeam.sapFrame.ID, -mappedBeam.revitFrame.Beta);


                FamilyInstance instance = (FamilyInstance)doc.GetElement(mappedBeam.revitFrame.revitEleId);
                using (Transaction trns = new Transaction(doc, "aa"))
                {
                    trns.Start();
                    FailureHandlingOptions failOpt = trns.GetFailureHandlingOptions();

                    RevitUtils.setParametersSTAAD(instance, mappedBeam.sapFrame);
                    trns.Commit();
                }
                continue;


            }


        }


        public static List<FamilyInstance> updateRevitGeomertyForBeams2(Document doc, List<SapRevitFrame> MappedBeams, ForgeTypeId dis, Microsoft.Office.Interop.Excel.Worksheet ws, double margin, string analysisModel, object[,] dataToExcel)
        {

            XYZ StartPoint = null;
            XYZ EndPoint = null;
            double x, y, z, x1, y1, z1;
            double margin2 = 0.0;

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


            List<FamilyInstance> revitElements = new List<FamilyInstance>();

            List<FamilyInstance> newlyCreatedElements = new List<FamilyInstance>();


            RevitServices.Transactions.TransactionManager.Instance.ForceCloseTransaction();

            using (TransactionGroup tg = new TransactionGroup(doc, "aaa"))
            {

                tg.Start();
                var chopSize = Frame.chopSize;
                var MappedbeamChunks = MappedBeams.Chunk(chopSize);


                foreach (var MappedBeamChunk in MappedbeamChunks)
                {


                    using (Transaction trns = new Transaction(doc, "trns"))
                    {
                        trns.Start();
                        FailureHandlingOptions failOpt = trns.GetFailureHandlingOptions();
                        failOpt.SetFailuresPreprocessor(new RoomWarningSwallower());
                        trns.SetFailureHandlingOptions(failOpt);


                        foreach (var mappedBeam in MappedBeamChunk)
                        {

                            try
                            {

                                margin2 = margin * mappedBeam.revitFrame.Length;
                                double dist1 = Point.getDistance(mappedBeam.revitFrame.StartPoint, mappedBeam.sapFrame.StartPoint);
                                double dist2 = Point.getDistance(mappedBeam.revitFrame.EndPoint, mappedBeam.sapFrame.EndPoint);
                                double angle = (mappedBeam.sapFrame.Beta) * Math.PI / 180;

                                if (dist1 < margin2 && dist2 < margin2)
                                {

                                    FamilyInstance instance = (FamilyInstance)doc.GetElement(mappedBeam.revitFrame.revitEleId);

                                    instance.LookupParameter("Cross-Section Rotation").Set(angle);

                                    if (analysisModel == "Staad")
                                    {
                                        RevitUtils.setParametersSTAAD(instance, mappedBeam.sapFrame);
                                    }

                                    else if (analysisModel == "Sap")
                                    {
                                        RevitUtils.setParametersSAP(instance, mappedBeam.sapFrame);

                                    }

                                    else if (analysisModel == "Etabs")
                                    {
                                        RevitUtils.setParametersEtabs(instance, mappedBeam.sapFrame);

                                    }

                                    revitElements.Add(instance);

                                    continue;
                                }

                                else
                                {
                                    x = UnitUtils.ConvertToInternalUnits(mappedBeam.sapFrame.StartPoint.X, dis);
                                    y = UnitUtils.ConvertToInternalUnits(mappedBeam.sapFrame.StartPoint.Y, dis);
                                    z = UnitUtils.ConvertToInternalUnits(mappedBeam.sapFrame.StartPoint.Z, dis);
                                    StartPoint = new XYZ(x, y, z);

                                    x1 = UnitUtils.ConvertToInternalUnits(mappedBeam.sapFrame.EndPoint.X, dis);
                                    y1 = UnitUtils.ConvertToInternalUnits(mappedBeam.sapFrame.EndPoint.Y, dis);
                                    z1 = UnitUtils.ConvertToInternalUnits(mappedBeam.sapFrame.EndPoint.Z, dis);
                                    EndPoint = new XYZ(x1, y1, z1);

                                    Line line = Line.CreateBound(StartPoint, EndPoint);


                                    FamilyInstance instance = (FamilyInstance)doc.GetElement(mappedBeam.revitFrame.revitEleId);

                                    Location loc = instance.Location;


                                    #region No change in type
                                    if ((BuiltInCategory)instance.Category.Id.IntegerValue == BuiltInCategory.OST_StructuralFraming)
                                    {
                                        // Create a new line with the modified end points
                                        Line newLine = Line.CreateBound(StartPoint, EndPoint);

                                        LocationCurve lc = instance.Location as LocationCurve;
                                        lc.Curve = newLine;

                                        //Set Parameters
                                        instance.LookupParameter("Cross-Section Rotation").Set(angle);

                                        if (analysisModel == "Staad")
                                        {
                                            RevitUtils.setParametersSTAAD(instance, mappedBeam.sapFrame);
                                        }

                                        else if (analysisModel == "Sap")
                                        {
                                            RevitUtils.setParametersSAP(instance, mappedBeam.sapFrame);

                                        }

                                        revitElements.Add(instance);

                                        continue;
                                    }


                                    #endregion

                                    #region Change in type (from Column to beam)
                                    else
                                    {

                                        // handle error in case revit section of new beam not found
                                        Dictionary<string, FamilySymbol> dic = new Dictionary<string, FamilySymbol>();
                                        List<Frame> Sapframes = new List<Frame>() { mappedBeam.sapFrame };
                                        dic = RevitUtils.createRvtFamilySectionForAllBeam(Sapframes, doc, dis);
                                        FamilySymbol revitSection = dic[mappedBeam.sapFrame.Sec];

                                        if (!revitSection.IsActive)
                                            revitSection.Activate();


                                        FamilyInstance newInstance = doc.Create.NewFamilyInstance(line, revitSection, lvl0, StructuralType.Beam);

                                        newInstance.LookupParameter("Cross-Section Rotation").Set(angle);
                                        Parameter sss = newInstance.LookupParameter("z Justification");
                                        int zjustf = (int)ZJustification.Center;

                                        Autodesk.Revit.DB.Structure.StructuralFramingUtils.DisallowJoinAtEnd(newInstance, 0);
                                        Autodesk.Revit.DB.Structure.StructuralFramingUtils.DisallowJoinAtEnd(newInstance, 1);

                                        //Set Parameters

                                        if (analysisModel == "Staad")
                                        {
                                            RevitUtils.setParametersSTAAD(newInstance, mappedBeam.sapFrame);
                                        }

                                        else if (analysisModel == "Sap")
                                        {
                                            RevitUtils.setParametersSAP(newInstance, mappedBeam.sapFrame);

                                        }

                                        else if (analysisModel == "Etabs")
                                        {
                                            RevitUtils.setParametersEtabs(newInstance, mappedBeam.sapFrame);

                                        }

                                        doc.Delete(instance.Id);

                                        revitElements.Add(newInstance);
                                        newlyCreatedElements.Add(newInstance);
                                        continue;
                                    }




                                    #endregion



                                }
                            }
                            catch (Exception ex2)
                            {
                                System.Diagnostics.StackTrace trace = new System.Diagnostics.StackTrace(ex2, true);
                                RevitUtils.Log.Add("Error in updateRevitGeomertyForBeams2 Method  sapBeam no " + mappedBeam.sapFrame.ID + ex2.Message + trace.GetFrame(trace.FrameCount - 1).GetFileLineNumber());
                                continue;
                            }
                        }
                        trns.Commit();

                    }
                }

                if (ws != null)
                {
                    List<Frame> sapMappedBeams = MappedBeams.Select(a => a.sapFrame).ToList();
                    Frame.setDataInExcelBulk(sapMappedBeams, revitElements, dis, ws, ref dataToExcel, "not");
                }

                tg.Assimilate();
            }


            return newlyCreatedElements;



        }

        // should be modified to be likse sap
        public static void updateEtabsGeomerty(List<SapRevitFrame> MappedBeams, ETABSv1.cSapModel myEtabsmodel, Document doc, double margin)
        {
            double margin2 = 0.0;
            string propName = null;
            string Sauto = null;
            string materialName = null;
            int ret = 0;
            string stPoint = null;
            string endPoint = null;
            foreach (var mappedBeam in MappedBeams.AsParallel())
            {

                //myEtabsmodel.View.RefreshView();
                margin2 = margin * mappedBeam.revitFrame.Length;
                double dist1 = Point.getDistance(mappedBeam.revitFrame.StartPoint, mappedBeam.sapFrame.StartPoint);
                double dist2 = Point.getDistance(mappedBeam.revitFrame.EndPoint, mappedBeam.sapFrame.EndPoint);


                if (dist1 < margin2 && dist2 < margin2)
                {
                    // set Parameters in revit    
                    myEtabsmodel.FrameObj.GetSection(mappedBeam.sapFrame.ID, ref propName, ref Sauto);
                    myEtabsmodel.PropFrame.GetMaterial(propName, ref materialName);


                    // Update rotation 
                    myEtabsmodel.FrameObj.SetLocalAxes(mappedBeam.sapFrame.ID, mappedBeam.revitFrame.Beta);


                    FamilyInstance instance = (FamilyInstance)doc.GetElement(mappedBeam.revitFrame.revitEleId);
                    using (Transaction trns = new Transaction(doc, "aa"))
                    {
                        trns.Start();
                        FailureHandlingOptions failOpt = trns.GetFailureHandlingOptions();
                        RevitUtils.setParametersEtabs(instance, mappedBeam.sapFrame);
                        trns.Commit();
                    }
                    continue;
                }

                else
                {

                    ret = myEtabsmodel.PointObj.AddCartesian(mappedBeam.revitFrame.StartPoint.X, mappedBeam.revitFrame.StartPoint.Y, mappedBeam.revitFrame.StartPoint.Z, ref stPoint);
                    ret = myEtabsmodel.PointObj.AddCartesian(mappedBeam.revitFrame.EndPoint.X, mappedBeam.revitFrame.EndPoint.Y, mappedBeam.revitFrame.EndPoint.Z, ref endPoint);

                    ret = myEtabsmodel.EditFrame.ChangeConnectivity(mappedBeam.sapFrame.ID, stPoint, endPoint);


                    // Update rotation 
                    myEtabsmodel.FrameObj.SetLocalAxes(mappedBeam.sapFrame.ID, mappedBeam.revitFrame.Beta);


                    // set Parameters in revit    
                    myEtabsmodel.FrameObj.GetSection(mappedBeam.sapFrame.ID, ref propName, ref Sauto);
                    myEtabsmodel.PropFrame.GetMaterial(propName, ref materialName);

                    FamilyInstance instance = (FamilyInstance)doc.GetElement(mappedBeam.revitFrame.revitEleId);
                    using (Transaction trns = new Transaction(doc, "aa"))
                    {
                        trns.Start();
                        FailureHandlingOptions failOpt = trns.GetFailureHandlingOptions();
                        RevitUtils.setParametersEtabs(instance, mappedBeam.sapFrame);
                        trns.Commit();
                    }
                    continue;
                }


            }


        }
        public static void updateStaadGeomerty(List<SapRevitFrame> MappedBeams, Document doc, double margin)
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


            foreach (var mappedBeam in MappedBeams.AsParallel())
            {


                margin2 = margin * mappedBeam.revitFrame.Length;
                double dist1 = Point.getDistance(mappedBeam.revitFrame.StartPoint, mappedBeam.sapFrame.StartPoint);
                double dist2 = Point.getDistance(mappedBeam.revitFrame.EndPoint, mappedBeam.sapFrame.EndPoint);


                if (dist1 < margin2 && dist2 < margin2)
                {

                    // Update rotation 
                    StaadUtils.property.AssignBetaAngle(mappedBeam.sapFrame.ID, -mappedBeam.revitFrame.Beta);

                    FamilyInstance instance = (FamilyInstance)doc.GetElement(mappedBeam.revitFrame.revitEleId);
                    using (Transaction trns = new Transaction(doc, "aa"))
                    {
                        trns.Start();
                        FailureHandlingOptions failOpt = trns.GetFailureHandlingOptions();
                        RevitUtils.setParametersSTAAD(instance, mappedBeam.sapFrame);

                        trns.Commit();
                    }
                    continue;
                }

                else
                {

                    StaadUtils.getUnitFactorsFactorToStaad(ref forceFactor, ref momentFactor, ref lengthFactor);


                    int newStartNodeId = StaadUtils.gometry.AddNode(mappedBeam.revitFrame.StartPoint.X * lengthFactor, mappedBeam.revitFrame.StartPoint.Z * lengthFactor, mappedBeam.revitFrame.StartPoint.Y * lengthFactor * -1);
                    int newEndtNodeId = StaadUtils.gometry.AddNode(mappedBeam.revitFrame.EndPoint.X * lengthFactor, mappedBeam.revitFrame.EndPoint.Z * lengthFactor, mappedBeam.revitFrame.EndPoint.Y * lengthFactor * -1);

                    StaadUtils.gometry.SetNodeCoordinate(mappedBeam.sapFrame.StartPoint.Id, mappedBeam.revitFrame.StartPoint.X * lengthFactor, mappedBeam.revitFrame.StartPoint.Z * lengthFactor, mappedBeam.revitFrame.StartPoint.Y * lengthFactor * -1);
                    StaadUtils.gometry.SetNodeCoordinate(mappedBeam.sapFrame.EndPoint.Id, mappedBeam.revitFrame.EndPoint.X * lengthFactor, mappedBeam.revitFrame.EndPoint.Z * lengthFactor, mappedBeam.revitFrame.EndPoint.Y * lengthFactor * -1);


                    // Update rotation 
                    StaadUtils.property.AssignBetaAngle(mappedBeam.sapFrame.ID, -mappedBeam.revitFrame.Beta);



                    // set Parameters in revit    


                    FamilyInstance instance = (FamilyInstance)doc.GetElement(mappedBeam.revitFrame.revitEleId);
                    using (Transaction trns = new Transaction(doc, "aa"))
                    {
                        trns.Start();
                        FailureHandlingOptions failOpt = trns.GetFailureHandlingOptions();
                        RevitUtils.setParametersSTAAD(instance, mappedBeam.sapFrame);

                        trns.Commit();
                    }
                    continue;
                }


            }


        }





        #region New update for Sap, and should be implemented in etabs (to set in excel when go from revit to sap)

        public static void updateSapSections2(List<SapRevitFrame> MappedBeams, SAP2000v1.cSapModel mysapmodel, Document doc, string type, Dictionary<string, string> secDicForBeams, Dictionary<string, string> secDicForColumns, Microsoft.Office.Interop.Excel.Worksheet ws, ref object[,] dataToExcel)
        {
            string materialName = null;
            string sAuto = null;
            String sapSection = null;
            List<Frame> analysisElements = new List<Frame>();
            List<Frame> revitElements = MappedBeams.Select(x => x.revitFrame).ToList();


            foreach (var mappedBeam in MappedBeams.AsParallel())
            {

                //mysapmodel.View.RefreshView();
                if (type == "BM")
                {
                    sapSection = secDicForBeams[mappedBeam.revitFrame.RvtName];

                }
                else if (type == "COL")
                {
                    sapSection = secDicForColumns[mappedBeam.revitFrame.RvtName];
                }




                if (sapSection == null)
                {
                    continue;
                }


                else if (sapSection == mappedBeam.sapFrame.Sec)
                {

                    // Update rotation 
                    mysapmodel.FrameObj.SetLocalAxes(mappedBeam.sapFrame.ID, mappedBeam.revitFrame.Beta);
                    mappedBeam.sapFrame.Beta = mappedBeam.revitFrame.Beta;

                    // set Parameters in revit    
                    mysapmodel.FrameObj.GetSection(mappedBeam.sapFrame.ID, ref sapSection, ref sAuto);
                    mysapmodel.PropFrame.GetMaterial(sapSection, ref materialName);
                    mappedBeam.sapFrame.Material = materialName;

                    FamilyInstance instance = (FamilyInstance)doc.GetElement(mappedBeam.revitFrame.revitEleId);
                    using (Transaction trns = new Transaction(doc, "aa"))
                    {
                        trns.Start();
                        FailureHandlingOptions failOpt = trns.GetFailureHandlingOptions();

                        RevitUtils.setParametersSAP(instance, mappedBeam.sapFrame);



                        trns.Commit();
                    }

                    analysisElements.Add(mappedBeam.sapFrame);
                    continue;

                }


                else
                {
                    // updateSapSection
                    int ret = mysapmodel.FrameObj.SetSection(mappedBeam.sapFrame.ID, sapSection);
                    mappedBeam.sapFrame.Sec = sapSection;

                    if (ret == 1)
                    {
                        mysapmodel.FrameObj.SetSection(mappedBeam.sapFrame.ID, "dummySection");
                        mappedBeam.sapFrame.Sec = "dummySection";

                    }

                    // Update rotation 
                    mysapmodel.FrameObj.SetLocalAxes(mappedBeam.sapFrame.ID, mappedBeam.revitFrame.Beta);
                    mappedBeam.sapFrame.Beta = mappedBeam.revitFrame.Beta;

                    // set Parameters in revit    
                    mysapmodel.FrameObj.GetSection(mappedBeam.sapFrame.ID, ref sapSection, ref sAuto);
                    mysapmodel.PropFrame.GetMaterial(sapSection, ref materialName);
                    mappedBeam.sapFrame.Material = materialName;


                    FamilyInstance instance = (FamilyInstance)doc.GetElement(mappedBeam.revitFrame.revitEleId);

                    using (Transaction trns = new Transaction(doc, "aa"))
                    {
                        trns.Start();
                        FailureHandlingOptions failOpt = trns.GetFailureHandlingOptions();
                        RevitUtils.setParametersSAP(instance, mappedBeam.sapFrame);

                        trns.Commit();
                    }


                    analysisElements.Add(mappedBeam.sapFrame);
                    continue;

                }

            }

            Frame.UpdateAnlaysisDataInExcel(analysisElements, revitElements, ws, ref dataToExcel, "Not", "Sec");

        }

        public static void updateSapGeomerty2(List<SapRevitFrame> MappedBeams, SAP2000v1.cSapModel mysapmodel, Document doc, double margin, Microsoft.Office.Interop.Excel.Worksheet ws, ref object[,] dataToExcel)
        {
            double margin2 = 0.0;
            string propName = null;
            string Sauto = null;
            string materialName = null;
            int ret = 0;

            List<Frame> analysisElements = new List<Frame>();
            List<Frame> revitElements = MappedBeams.Select(x => x.revitFrame).ToList();

            foreach (var mappedBeam in MappedBeams.AsParallel())
            {

                //mysapmodel.View.RefreshView();
                margin2 = margin * mappedBeam.revitFrame.Length;
                double dist1 = Point.getDistance(mappedBeam.revitFrame.StartPoint, mappedBeam.sapFrame.StartPoint);
                double dist2 = Point.getDistance(mappedBeam.revitFrame.EndPoint, mappedBeam.sapFrame.EndPoint);


                if (dist1 < margin2 && dist2 < margin2)
                {

                    mysapmodel.FrameObj.GetSection(mappedBeam.sapFrame.ID, ref propName, ref Sauto);
                    mysapmodel.PropFrame.GetMaterial(propName, ref materialName);
                    // Update rotation 
                    mysapmodel.FrameObj.SetLocalAxes(mappedBeam.sapFrame.ID, mappedBeam.revitFrame.Beta);


                    // reflect updates on the sapElement
                    mappedBeam.sapFrame.Beta = mappedBeam.revitFrame.Beta;
                    mappedBeam.sapFrame.Material = materialName;


                    // set Parameters in revit 
                    FamilyInstance instance = (FamilyInstance)doc.GetElement(mappedBeam.revitFrame.revitEleId);
                    using (Transaction trns = new Transaction(doc, "aa"))
                    {
                        trns.Start();
                        FailureHandlingOptions failOpt = trns.GetFailureHandlingOptions();
                        RevitUtils.setParametersSAP(instance, mappedBeam.sapFrame);
                        trns.Commit();
                    }

                    analysisElements.Add(mappedBeam.sapFrame);
                    continue;
                }

                else
                {


                    // Update rotation 
                    mysapmodel.FrameObj.SetLocalAxes(mappedBeam.sapFrame.ID, mappedBeam.revitFrame.Beta);

                    ret = mysapmodel.EditPoint.ChangeCoordinates_1(mappedBeam.sapFrame.StartPoint.Id, mappedBeam.revitFrame.StartPoint.X, mappedBeam.revitFrame.StartPoint.Y, mappedBeam.revitFrame.StartPoint.Z);
                    ret = mysapmodel.EditPoint.ChangeCoordinates_1(mappedBeam.sapFrame.EndPoint.Id, mappedBeam.revitFrame.EndPoint.X, mappedBeam.revitFrame.EndPoint.Y, mappedBeam.revitFrame.EndPoint.Z);


                    // reflect updates on the sapElement
                    mappedBeam.sapFrame.Beta = mappedBeam.revitFrame.Beta;
                    mappedBeam.sapFrame.Material = materialName;
                    mappedBeam.sapFrame.StartPoint.X = mappedBeam.revitFrame.StartPoint.X;
                    mappedBeam.sapFrame.StartPoint.Y = mappedBeam.revitFrame.StartPoint.Y;
                    mappedBeam.sapFrame.StartPoint.Z = mappedBeam.revitFrame.StartPoint.Z;
                    mappedBeam.sapFrame.EndPoint.X = mappedBeam.revitFrame.EndPoint.X;
                    mappedBeam.sapFrame.EndPoint.Y = mappedBeam.revitFrame.EndPoint.Y;
                    mappedBeam.sapFrame.EndPoint.Z = mappedBeam.revitFrame.EndPoint.Z;




                    // set Parameters in revit    
                    mysapmodel.FrameObj.GetSection(mappedBeam.sapFrame.ID, ref propName, ref Sauto);
                    mysapmodel.PropFrame.GetMaterial(propName, ref materialName);

                    FamilyInstance instance = (FamilyInstance)doc.GetElement(mappedBeam.revitFrame.revitEleId);
                    using (Transaction trns = new Transaction(doc, "aa"))
                    {
                        trns.Start();
                        FailureHandlingOptions failOpt = trns.GetFailureHandlingOptions();
                        RevitUtils.setParametersSAP(instance, mappedBeam.sapFrame);
                        trns.Commit();
                    }

                    analysisElements.Add(mappedBeam.sapFrame);
                    continue;
                }


            }

            Frame.UpdateAnlaysisDataInExcel(analysisElements, revitElements, ws, ref dataToExcel, "Not", "Geo");

        }

        #endregion



        #region  unUsed

        public static void updateSapSections(List<SapRevitFrame> MappedBeams, SAP2000v1.cSapModel mysapmodel, Document doc, string type, Dictionary<string, string> secDicForBeams, Dictionary<string, string> secDicForColumns)
        {
            string materialName = null;
            string sAuto = null;
            String sapSection = null;

            foreach (var mappedBeam in MappedBeams.AsParallel())
            {

                //mysapmodel.View.RefreshView();
                if (type == "BM")
                {
                    sapSection = secDicForBeams[mappedBeam.revitFrame.RvtName];

                }
                else if (type == "COL")
                {
                    sapSection = secDicForColumns[mappedBeam.revitFrame.RvtName];
                }




                if (sapSection == null)
                {
                    continue;
                }


                else if (sapSection == mappedBeam.sapFrame.Sec)
                {

                    // Update rotation 
                    mysapmodel.FrameObj.SetLocalAxes(mappedBeam.sapFrame.ID, mappedBeam.revitFrame.Beta);


                    // set Parameters in revit    
                    mysapmodel.FrameObj.GetSection(mappedBeam.sapFrame.ID, ref sapSection, ref sAuto);
                    mysapmodel.PropFrame.GetMaterial(sapSection, ref materialName);


                    FamilyInstance instance = (FamilyInstance)doc.GetElement(mappedBeam.revitFrame.revitEleId);
                    using (Transaction trns = new Transaction(doc, "aa"))
                    {
                        trns.Start();
                        FailureHandlingOptions failOpt = trns.GetFailureHandlingOptions();

                        RevitUtils.setParametersSAP(instance, mappedBeam.sapFrame);



                        trns.Commit();
                    }
                    continue;

                }


                else
                {
                    // updateSapSection
                    int ret = mysapmodel.FrameObj.SetSection(mappedBeam.sapFrame.ID, sapSection);
                    if (ret == 1)
                    {
                        mysapmodel.FrameObj.SetSection(mappedBeam.sapFrame.ID, "dummySection");
                    }

                    // Update rotation 
                    mysapmodel.FrameObj.SetLocalAxes(mappedBeam.sapFrame.ID, mappedBeam.revitFrame.Beta);


                    // set Parameters in revit    
                    mysapmodel.FrameObj.GetSection(mappedBeam.sapFrame.ID, ref sapSection, ref sAuto);
                    mysapmodel.PropFrame.GetMaterial(sapSection, ref materialName);


                    FamilyInstance instance = (FamilyInstance)doc.GetElement(mappedBeam.revitFrame.revitEleId);
                    using (Transaction trns = new Transaction(doc, "aa"))
                    {
                        trns.Start();
                        FailureHandlingOptions failOpt = trns.GetFailureHandlingOptions();
                        RevitUtils.setParametersSAP(instance, mappedBeam.sapFrame);

                        trns.Commit();
                    }
                    continue;

                }

            }


        }

        public static void updateSapGeomerty(List<SapRevitFrame> MappedBeams, SAP2000v1.cSapModel mysapmodel, Document doc, double margin)
        {
            double margin2 = 0.0;
            string propName = null;
            string Sauto = null;
            string materialName = null;
            int ret = 0;
            string stPoint = null;
            string endPoint = null;
            foreach (var mappedBeam in MappedBeams.AsParallel())
            {

                //mysapmodel.View.RefreshView();
                margin2 = margin * mappedBeam.revitFrame.Length;
                double dist1 = Point.getDistance(mappedBeam.revitFrame.StartPoint, mappedBeam.sapFrame.StartPoint);
                double dist2 = Point.getDistance(mappedBeam.revitFrame.EndPoint, mappedBeam.sapFrame.EndPoint);


                if (dist1 < margin2 && dist2 < margin2)
                {
                    // set Parameters in revit    
                    mysapmodel.FrameObj.GetSection(mappedBeam.sapFrame.ID, ref propName, ref Sauto);
                    mysapmodel.PropFrame.GetMaterial(propName, ref materialName);


                    // Update rotation 
                    mysapmodel.FrameObj.SetLocalAxes(mappedBeam.sapFrame.ID, mappedBeam.revitFrame.Beta);


                    FamilyInstance instance = (FamilyInstance)doc.GetElement(mappedBeam.revitFrame.revitEleId);
                    using (Transaction trns = new Transaction(doc, "aa"))
                    {
                        trns.Start();
                        FailureHandlingOptions failOpt = trns.GetFailureHandlingOptions();
                        RevitUtils.setParametersSAP(instance, mappedBeam.sapFrame);
                        trns.Commit();
                    }
                    continue;
                }

                else
                {


                    // Update rotation 
                    mysapmodel.FrameObj.SetLocalAxes(mappedBeam.sapFrame.ID, mappedBeam.revitFrame.Beta);

                    ret = mysapmodel.EditPoint.ChangeCoordinates_1(mappedBeam.sapFrame.StartPoint.Id, mappedBeam.revitFrame.StartPoint.X, mappedBeam.revitFrame.StartPoint.Y, mappedBeam.revitFrame.StartPoint.Z);
                    ret = mysapmodel.EditPoint.ChangeCoordinates_1(mappedBeam.sapFrame.EndPoint.Id, mappedBeam.revitFrame.EndPoint.X, mappedBeam.revitFrame.EndPoint.Y, mappedBeam.revitFrame.EndPoint.Z);


                    // set Parameters in revit    
                    mysapmodel.FrameObj.GetSection(mappedBeam.sapFrame.ID, ref propName, ref Sauto);
                    mysapmodel.PropFrame.GetMaterial(propName, ref materialName);

                    FamilyInstance instance = (FamilyInstance)doc.GetElement(mappedBeam.revitFrame.revitEleId);
                    using (Transaction trns = new Transaction(doc, "aa"))
                    {
                        trns.Start();
                        FailureHandlingOptions failOpt = trns.GetFailureHandlingOptions();
                        RevitUtils.setParametersSAP(instance, mappedBeam.sapFrame);
                        trns.Commit();
                    }
                    continue;
                }


            }


        }


        public static ElementId CreateBeam(Document doc, Frame bm, ForgeTypeId dis, Dictionary<string, FamilySymbol> dic, Microsoft.Office.Interop.Excel.Worksheet ws, string analysisModel)
        {


            #region defintions for creation of beams and columns 
            FamilySymbol revitSection = null;
            double x, y, z, x1, y1, z1;
            XYZ StartPoint = null;
            XYZ EndPoint = null;

        #endregion

        Beammodeling:
            using (Transaction trns = new Transaction(doc, "aa"))

            {

                trns.Start();
                FailureHandlingOptions failOpt = trns.GetFailureHandlingOptions();
                failOpt.SetFailuresPreprocessor(new RoomWarningSwallower());
                trns.SetFailureHandlingOptions(failOpt);


                x = UnitUtils.ConvertToInternalUnits(bm.StartPoint.X, dis);
                y = UnitUtils.ConvertToInternalUnits(bm.StartPoint.Y, dis);
                z = UnitUtils.ConvertToInternalUnits(bm.StartPoint.Z, dis);
                StartPoint = new XYZ(x, y, z);

                x1 = UnitUtils.ConvertToInternalUnits(bm.EndPoint.X, dis);
                y1 = UnitUtils.ConvertToInternalUnits(bm.EndPoint.Y, dis);
                z1 = UnitUtils.ConvertToInternalUnits(bm.EndPoint.Z, dis);
                EndPoint = new XYZ(x1, y1, z1);


                Curve line = Line.CreateBound(StartPoint, EndPoint);
                Level lvl0 = new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().ToList().First();
                double angle = (bm.Beta) * Math.PI / 180;

                try
                {
                modeling:

                    revitSection = dic[bm.Sec];

                    if (!revitSection.IsActive)
                        revitSection.Activate();

                    FamilyInstance instance = doc.Create.NewFamilyInstance(line, revitSection, lvl0, StructuralType.Beam);
                    instance.LookupParameter("Cross-Section Rotation").Set(angle);
                    Parameter sss = instance.LookupParameter("z Justification");
                    int zjustf = (int)ZJustification.Center;
                    sss.Set(zjustf);
                    Autodesk.Revit.DB.Structure.StructuralFramingUtils.DisallowJoinAtEnd(instance, 0);
                    Autodesk.Revit.DB.Structure.StructuralFramingUtils.DisallowJoinAtEnd(instance, 1);

                    if (analysisModel == "Staad")
                    {
                        RevitUtils.setParametersSTAAD(instance, bm);
                    }

                    else if (analysisModel == "Sap")
                    {
                        RevitUtils.setParametersSAP(instance, bm);

                    }

                    else if (analysisModel == "Etabs")
                    {

                        RevitUtils.setParametersEtabs(instance, bm);


                    }

                    trns.Commit();

                    if (ws != null)
                    {

                        Frame.setDataInExcel(bm, instance, dis, (Microsoft.Office.Interop.Excel.Worksheet)ws);



                    }

                    else
                    {

                    }


                    try
                    {
                        return instance.Id;

                    }

                    catch (Exception)
                    {

                        return null;
                    }




                }

                catch (Exception ex2)

                {
                    System.Diagnostics.StackTrace trace = new System.Diagnostics.StackTrace(ex2, true);
                    TaskDialog.Show("Error", "error in beam no" + bm.ID + ex2.Message + trace.GetFrame(trace.FrameCount - 1).GetFileLineNumber());

                    return null;

                }

            }

        }

        public static List<ElementId> CreateAllBeams(Document doc, List<Frame> beams, ForgeTypeId dis, Dictionary<string, FamilySymbol> dic, Microsoft.Office.Interop.Excel.Worksheet ws, string analysisModel)
        {

            List<ElementId> result = new List<ElementId>();

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
            foreach (var bm in beams)
            {
                var x = UnitUtils.ConvertToInternalUnits(bm.StartPoint.X, dis);
                var y = UnitUtils.ConvertToInternalUnits(bm.StartPoint.Y, dis);
                var z = UnitUtils.ConvertToInternalUnits(bm.StartPoint.Z, dis);
                var StartPoint = new XYZ(x, y, z);

                var x1 = UnitUtils.ConvertToInternalUnits(bm.EndPoint.X, dis);
                var y1 = UnitUtils.ConvertToInternalUnits(bm.EndPoint.Y, dis);
                var z1 = UnitUtils.ConvertToInternalUnits(bm.EndPoint.Z, dis);
                var EndPoint = new XYZ(x1, y1, z1);


                using (Transaction trns = new Transaction(doc, "aa"))
                {
                    try
                    {
                        Curve line = Line.CreateBound(StartPoint, EndPoint);
                        angle = (bm.Beta) * Math.PI / 180;

                        revitSection = dic[bm.Sec];

                        trns.Start();
                        FailureHandlingOptions failOpt = trns.GetFailureHandlingOptions();
                        failOpt.SetFailuresPreprocessor(new RoomWarningSwallower());
                        trns.SetFailureHandlingOptions(failOpt);


                        if (!revitSection.IsActive)
                            revitSection.Activate();

                        FamilyInstance instance = doc.Create.NewFamilyInstance(line, revitSection, lvl0, StructuralType.Beam);
                        instance.LookupParameter("Cross-Section Rotation").Set(angle);
                        Parameter sss = instance.LookupParameter("z Justification");
                        int zjustf = (int)ZJustification.Center;
                        sss.Set(zjustf);
                        Autodesk.Revit.DB.Structure.StructuralFramingUtils.DisallowJoinAtEnd(instance, 0);
                        Autodesk.Revit.DB.Structure.StructuralFramingUtils.DisallowJoinAtEnd(instance, 1);


                        if (analysisModel == "Staad")
                        {
                            RevitUtils.setParametersSTAAD(instance, bm);
                        }

                        else if (analysisModel == "Sap")
                        {
                            RevitUtils.setParametersSAP(instance, bm);

                        }

                        trns.Commit();


                        if (ws != null)
                        {
                            Frame.setDataInExcel(bm, instance, dis, (Microsoft.Office.Interop.Excel.Worksheet)ws);

                        }



                        result.Add(instance.Id);


                    }


                    catch (Exception ex2)

                    {
                        System.Diagnostics.StackTrace trace = new System.Diagnostics.StackTrace(ex2, true);
                        TaskDialog.Show("Error", "error in beam no" + bm.ID + ex2.Message + trace.GetFrame(trace.FrameCount - 1).GetFileLineNumber());

                        continue;

                    }


                }

            }

            return result;
        }

        public static void updateRevitSectionsForBeams(Document doc, List<SapRevitFrame> MappedBeams, Dictionary<string, FamilySymbol> dic, ForgeTypeId dis, Microsoft.Office.Interop.Excel.Worksheet ws, string analysisModel)
        {

            foreach (var mappedBeam in MappedBeams.AsParallel())
            {


                FamilySymbol revitSection = dic[mappedBeam.sapFrame.Sec];

                if (revitSection == null)
                {
                    continue;
                }

                // NO update
                else if (revitSection.Name == mappedBeam.revitFrame.RvtName)
                {
                    FamilyInstance instance = (FamilyInstance)doc.GetElement(mappedBeam.revitFrame.revitEleId);

                    using (Transaction trns = new Transaction(doc, "aa"))
                    {
                        trns.Start();
                        FailureHandlingOptions failOpt = trns.GetFailureHandlingOptions();
                        failOpt.SetFailuresPreprocessor(new RoomWarningSwallower());
                        trns.SetFailureHandlingOptions(failOpt);


                        double angle = (mappedBeam.sapFrame.Beta) * Math.PI / 180;
                        instance.LookupParameter("Cross-Section Rotation").Set(angle);

                        if (analysisModel == "Staad")
                        {
                            RevitUtils.setParametersSTAAD(instance, mappedBeam.sapFrame);
                        }

                        else if (analysisModel == "Sap")
                        {
                            RevitUtils.setParametersSAP(instance, mappedBeam.sapFrame);

                        }

                        trns.Commit();
                    }

                }


                // update
                else
                {
                    FamilyInstance instance = (FamilyInstance)doc.GetElement(mappedBeam.revitFrame.revitEleId);

                    using (Transaction trns = new Transaction(doc, "aa"))
                    {
                        trns.Start();
                        FailureHandlingOptions failOpt = trns.GetFailureHandlingOptions();
                        failOpt.SetFailuresPreprocessor(new RoomWarningSwallower());
                        trns.SetFailureHandlingOptions(failOpt);
                        instance.Symbol = revitSection;

                        if (analysisModel == "Staad")
                        {
                            RevitUtils.setParametersSTAAD(instance, mappedBeam.sapFrame);
                        }

                        else if (analysisModel == "Sap")
                        {
                            RevitUtils.setParametersSAP(instance, mappedBeam.sapFrame);

                        }


                        trns.Commit();

                    }

                    if (ws != null)
                    {
                        Frame.setDataInExcel(mappedBeam.sapFrame, instance, dis, (Microsoft.Office.Interop.Excel.Worksheet)ws);

                    }

                }

            }


        }


        public static List<FamilyInstance> updateRevitGeomertyForBeams(Document doc, List<SapRevitFrame> MappedBeams, ForgeTypeId dis, Microsoft.Office.Interop.Excel.Worksheet ws, double margin, string analysisModel)
        {

            XYZ StartPoint = null;
            XYZ EndPoint = null;
            double x, y, z, x1, y1, z1;
            double margin2 = 0.0;
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

            List<FamilyInstance> newlyCreatedElements = new List<FamilyInstance>();

            foreach (var mappedBeam in MappedBeams.AsParallel())
            {
                try
                {
                    margin2 = margin * mappedBeam.revitFrame.Length;
                    double dist1 = Point.getDistance(mappedBeam.revitFrame.StartPoint, mappedBeam.sapFrame.StartPoint);
                    double dist2 = Point.getDistance(mappedBeam.revitFrame.EndPoint, mappedBeam.sapFrame.EndPoint);
                    double angle = (mappedBeam.sapFrame.Beta) * Math.PI / 180;


                    if (dist1 < margin2 && dist2 < margin2)
                    {

                        FamilyInstance instance = (FamilyInstance)doc.GetElement(mappedBeam.revitFrame.revitEleId);
                        using (Transaction trns = new Transaction(doc, "aa"))
                        {
                            trns.Start();
                            FailureHandlingOptions failOpt = trns.GetFailureHandlingOptions();
                            instance.LookupParameter("Cross-Section Rotation").Set(angle);

                            if (analysisModel == "Staad")
                            {
                                RevitUtils.setParametersSTAAD(instance, mappedBeam.sapFrame);
                            }

                            else if (analysisModel == "Sap")
                            {
                                RevitUtils.setParametersSAP(instance, mappedBeam.sapFrame);

                            }


                            trns.Commit();
                        }


                        if (ws != null)
                        {
                            Frame.setDataInExcel(mappedBeam.sapFrame, instance, dis, (Microsoft.Office.Interop.Excel.Worksheet)ws);

                        }

                        continue;
                    }


                    else
                    {
                        x = UnitUtils.ConvertToInternalUnits(mappedBeam.sapFrame.StartPoint.X, dis);
                        y = UnitUtils.ConvertToInternalUnits(mappedBeam.sapFrame.StartPoint.Y, dis);
                        z = UnitUtils.ConvertToInternalUnits(mappedBeam.sapFrame.StartPoint.Z, dis);
                        StartPoint = new XYZ(x, y, z);

                        x1 = UnitUtils.ConvertToInternalUnits(mappedBeam.sapFrame.EndPoint.X, dis);
                        y1 = UnitUtils.ConvertToInternalUnits(mappedBeam.sapFrame.EndPoint.Y, dis);
                        z1 = UnitUtils.ConvertToInternalUnits(mappedBeam.sapFrame.EndPoint.Z, dis);
                        EndPoint = new XYZ(x1, y1, z1);

                        Line line = Line.CreateBound(StartPoint, EndPoint);




                        FamilyInstance instance = (FamilyInstance)doc.GetElement(mappedBeam.revitFrame.revitEleId);

                        Location loc = instance.Location;


                        try
                        {
                            #region No change in type
                            if ((BuiltInCategory)instance.Category.Id.IntegerValue == BuiltInCategory.OST_StructuralFraming)
                            {
                                using (Transaction trns = new Transaction(doc, "aa"))
                                {
                                    trns.Start();
                                    FailureHandlingOptions failOpt = trns.GetFailureHandlingOptions();
                                    failOpt.SetFailuresPreprocessor(new RoomWarningSwallower());
                                    trns.SetFailureHandlingOptions(failOpt);

                                    // Create a new line with the modified end points
                                    Line newLine = Line.CreateBound(StartPoint, EndPoint);

                                    LocationCurve lc = instance.Location as LocationCurve;
                                    lc.Curve = newLine;

                                    //Set Parameters
                                    instance.LookupParameter("Cross-Section Rotation").Set(angle);

                                    if (analysisModel == "Staad")
                                    {
                                        RevitUtils.setParametersSTAAD(instance, mappedBeam.sapFrame);
                                    }

                                    else if (analysisModel == "Sap")
                                    {
                                        RevitUtils.setParametersSAP(instance, mappedBeam.sapFrame);

                                    }


                                    trns.Commit();

                                    if (ws != null)
                                    {
                                        Frame.setDataInExcel(mappedBeam.sapFrame, instance, dis, (Microsoft.Office.Interop.Excel.Worksheet)ws);

                                    }

                                    continue;
                                }
                            }

                            #endregion

                            #region Change in type (from Column to beam)
                            else
                            {
                                using (Transaction trns = new Transaction(doc, "aa"))
                                {
                                    FamilySymbol revitSection = RevitUtils.createRvtFamilySectionForBeam(mappedBeam.sapFrame, doc, dis);
                                    trns.Start();
                                    FailureHandlingOptions failOpt = trns.GetFailureHandlingOptions();
                                    failOpt.SetFailuresPreprocessor(new RoomWarningSwallower());
                                    trns.SetFailureHandlingOptions(failOpt);

                                    if (!revitSection.IsActive)
                                        revitSection.Activate();


                                    FamilyInstance newInstance = doc.Create.NewFamilyInstance(line, revitSection, lvl0, StructuralType.Beam);

                                    newInstance.LookupParameter("Cross-Section Rotation").Set(angle);
                                    Parameter sss = newInstance.LookupParameter("z Justification");
                                    int zjustf = (int)ZJustification.Center;

                                    Autodesk.Revit.DB.Structure.StructuralFramingUtils.DisallowJoinAtEnd(newInstance, 0);
                                    Autodesk.Revit.DB.Structure.StructuralFramingUtils.DisallowJoinAtEnd(newInstance, 1);

                                    //Set Parameters

                                    if (analysisModel == "Staad")
                                    {
                                        RevitUtils.setParametersSTAAD(newInstance, mappedBeam.sapFrame);
                                    }

                                    else if (analysisModel == "Sap")
                                    {
                                        RevitUtils.setParametersSAP(newInstance, mappedBeam.sapFrame);

                                    }


                                    doc.Delete(instance.Id);
                                    trns.Commit();

                                    if (ws != null)
                                    {
                                        Frame.setDataInExcel(mappedBeam.sapFrame, instance, dis, (Microsoft.Office.Interop.Excel.Worksheet)ws);

                                    }
                                    newlyCreatedElements.Add(newInstance);
                                    continue;
                                }


                            }

                            #endregion
                        }


                        catch (Exception ex2)

                        {
                            System.Diagnostics.StackTrace trace = new System.Diagnostics.StackTrace(ex2, true);
                            TaskDialog.Show("Error", "error in beam no" + mappedBeam.sapFrame.ID + ex2.Message + trace.GetFrame(trace.FrameCount - 1).GetFileLineNumber());

                            continue;

                        }




                    }
                }

                catch (Exception eeex)
                {
                    System.Diagnostics.StackTrace trace = new System.Diagnostics.StackTrace(eeex, true);
                    TaskDialog.Show("Error", "error in beam no" + mappedBeam.sapFrame.ID + eeex.Message + trace.GetFrame(trace.FrameCount - 1).GetFileLineNumber());

                    continue;
                }



            }

            return newlyCreatedElements;
        }


        #endregion
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


        public Column(FamilyInstance instance)
        {
            if (instance == null)
            {

            }
            else
            {
                Document doc = instance.Document;
                AnalyticalToPhysicalAssociationManager analyManger = AnalyticalToPhysicalAssociationManager.GetAnalyticalToPhysicalAssociationManager(doc);
                AnalyticalElement analytInstance = doc.GetElement(analyManger.GetAssociatedElementId(instance.Id)) as AnalyticalElement;


                this.RvtID = instance.Id.ToString();
                this.RvtName = instance.Symbol.Name;
                this.RvtFamily = instance.Symbol.FamilyName;
                this.StartPoint = new Point(analytInstance.GetCurve().GetEndPoint(0));
                this.EndPoint = new Point(analytInstance.GetCurve().GetEndPoint(1));
            }


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


        public static List<ElementId> CreateAllColumns3(Document doc, List<Frame> columns, ForgeTypeId dis, Dictionary<string, FamilySymbol> dic, Microsoft.Office.Interop.Excel.Worksheet ws, string analysisModel, ref object[,] dataToExcel)
        {

            List<ElementId> result = new List<ElementId>();
            List<FamilyInstance> revitEelemtns = new List<FamilyInstance>();

            #region defintions for creation of beams and columns 
            FamilySymbol revitSection = null;
            double x, y, z, x1, y1, z1;
            XYZ StartPoint = null;
            XYZ EndPoint = null;
            Curve line;
            Curve line3;

            double angle = 0;
            Level lvl0;
            Level lvl1;
            Parameter param;
            LocationCurve instanceCurve;
            Parameter pbase;
            Parameter ptop;
            Level baselvl;
            Level toplvl;
            double basePointElevation;
            double topPointElevation;
            double baseLevelElevation;
            double topLevelElevation;
            double baseOffset;
            double topOffset;
            double base_offset;
            double top_offset;
            Line line2;
            double col_angle;


            lvl0 = new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().FirstOrDefault();
            lvl1 = new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().ElementAtOrDefault(1);

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

            if (lvl1 == null)
            {
                using (Transaction t = new Transaction(doc, "Create Levels"))
                {
                    t.Start();
                    lvl1 = Level.Create(doc, 10.0);
                    lvl1.Name = "Hydra-level-2";

                    t.Commit();

                }
            }


            if (lvl0.Elevation > lvl1.Elevation)
            {
                lvl0 = new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().ToList()[1];
                lvl1 = new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().ToList().First();
            }

            #endregion

            RevitServices.Transactions.TransactionManager.Instance.ForceCloseTransaction();

            using (TransactionGroup tg = new TransactionGroup(doc, "aaa"))

            {
                tg.Start();

                var chopSize = Frame.chopSize;
                var columnChuncks = columns.Chunk(chopSize).ToList();

                foreach (var columnChunck in columnChuncks)
                {

                    using (Transaction trns = new Transaction(doc, "trns"))
                    {
                        trns.Start();
                        FailureHandlingOptions failOpt = trns.GetFailureHandlingOptions();
                        failOpt.SetFailuresPreprocessor(new RoomWarningSwallower());
                        trns.SetFailureHandlingOptions(failOpt);


                        foreach (var cl in columnChunck)
                        {

                            x = UnitUtils.ConvertToInternalUnits(cl.StartPoint.X, dis);
                            y = UnitUtils.ConvertToInternalUnits(cl.StartPoint.Y, dis);
                            z = UnitUtils.ConvertToInternalUnits(cl.StartPoint.Z, dis);
                            StartPoint = new XYZ(x, y, z);

                            x1 = UnitUtils.ConvertToInternalUnits(cl.EndPoint.X, dis);
                            y1 = UnitUtils.ConvertToInternalUnits(cl.EndPoint.Y, dis);
                            z1 = UnitUtils.ConvertToInternalUnits(cl.EndPoint.Z, dis);
                            EndPoint = new XYZ(x1, y1, z1);


                            try
                            {
                                line = Line.CreateBound(StartPoint, EndPoint);
                                line3 = Line.CreateBound(EndPoint, StartPoint);
                                revitSection = dic[cl.Sec];


                                if (!revitSection.IsActive)
                                    revitSection.Activate();

                                FamilyInstance instance = doc.Create.NewFamilyInstance(StartPoint, revitSection, lvl0, StructuralType.Column);

                                if (cl.MemberType == MemberType.SlanttedColumn)
                                {
                                    param = instance.LookupParameter("Column Style");
                                    param.Set((int)SlantedOrVerticalColumnType.CT_EndPoint);

                                    instanceCurve = (LocationCurve)instance.Location;
                                    if (instanceCurve != null)
                                    {
                                        if (z > z1)
                                        {
                                            instanceCurve.Curve = line3;
                                        }
                                        else
                                        {
                                            instanceCurve.Curve = line;

                                        }

                                    }

                                }

                                else
                                {
                                    #region handling levels elevations
                                    pbase = instance.LookupParameter("Base Level");
                                    ptop = instance.LookupParameter("Top Level");

                                    //
                                    pbase.Set(lvl0.Id);
                                    ptop.Set(lvl1.Id);


                                    baselvl = doc.GetElement(pbase.AsElementId()) as Level;
                                    toplvl = doc.GetElement(ptop.AsElementId()) as Level;


                                    basePointElevation = Convert.ToDouble(cl.StartPoint.Z);
                                    topPointElevation = Convert.ToDouble(cl.EndPoint.Z);

                                    if (basePointElevation > topPointElevation)
                                    {
                                        basePointElevation = Convert.ToDouble(cl.EndPoint.Z);
                                        topPointElevation = Convert.ToDouble(cl.StartPoint.Z);

                                    }


                                    baseLevelElevation = UnitUtils.ConvertFromInternalUnits(baselvl.Elevation, dis);
                                    topLevelElevation = UnitUtils.ConvertFromInternalUnits(toplvl.Elevation, dis);


                                    baseOffset = basePointElevation - baseLevelElevation;
                                    topOffset = topPointElevation - topLevelElevation;

                                    base_offset = UnitUtils.ConvertToInternalUnits(basePointElevation - baseLevelElevation, dis);
                                    top_offset = UnitUtils.ConvertToInternalUnits(topPointElevation - topLevelElevation, dis);




                                    if ((topLevelElevation + topOffset) == baseLevelElevation)
                                    {
                                        // Replaced
                                        instance.LookupParameter("Base Offset").Set(base_offset);
                                        instance.LookupParameter("Top Offset").Set(top_offset);

                                    }

                                    else
                                    {
                                        instance.LookupParameter("Top Offset").Set(top_offset);
                                        instance.LookupParameter("Base Offset").Set(base_offset);
                                    }


                                    #endregion

                                }


                                line2 = Line.CreateBound(StartPoint, EndPoint);
                                col_angle = (cl.Beta) * Math.PI / 180;
                                ElementTransformUtils.RotateElement(doc, instance.Id, line2, col_angle);

                                // set Parameters 

                                if (analysisModel == "Staad")
                                {
                                    RevitUtils.setParametersSTAAD(instance, cl);
                                }

                                else if (analysisModel == "Sap")
                                {
                                    RevitUtils.setParametersSAP(instance, cl);

                                }


                                else if (analysisModel == "Etabs")
                                {
                                    RevitUtils.setParametersEtabs(instance, cl);

                                }



                                //if (ws != null)
                                //{
                                //    Frame.setDataInExcel(cl, instance, dis, (Microsoft.Office.Interop.Excel.Worksheet)ws);

                                //}


                                try
                                {
                                    revitEelemtns.Add(instance);
                                    result.Add(instance.Id);

                                }
                                catch (Exception)
                                {
                                    continue;
                                }

                            }


                            catch (Exception ex2)

                            {
                                System.Diagnostics.StackTrace trace = new System.Diagnostics.StackTrace(ex2, true);
                                RevitUtils.Log.Add("Error in CreateAllColumns3 Method... Col no" + cl.ID + ex2.Message + trace.GetFrame(trace.FrameCount - 1).GetFileLineNumber());

                                //TaskDialog.Show("Error", ex2.Message + " ,line of the error is " + trace.GetFrame(trace.FrameCount - 1).GetFileLineNumber());
                                continue;
                            }



                        }

                        trns.Commit();


                    }

                }

                if (ws != null)
                {
                    Frame.setDataInExcelBulk(columns, revitEelemtns, dis, ws, ref dataToExcel, "Do");

                }

                tg.Assimilate();
            }

            return result;

        }
        public static void updateRevitSectionsForColumns2(Document doc, List<SapRevitFrame> MappedColumns, Dictionary<string, FamilySymbol> dic, ForgeTypeId dis, Microsoft.Office.Interop.Excel.Worksheet ws, string analysisModel, object[,] dataToExcel)
        {
            List<FamilyInstance> revitInstances = new List<FamilyInstance>();

            RevitServices.Transactions.TransactionManager.Instance.ForceCloseTransaction();

            using (TransactionGroup tg = new TransactionGroup(doc, "aaa"))
            {
                tg.Start();

                var chopSize = Frame.chopSize;
                var MappedColumnsChuncks = MappedColumns.Chunk(chopSize).ToList();

                foreach (var mappedColumnChunck in MappedColumnsChuncks)
                {


                    using (Transaction trns = new Transaction(doc, "trns"))
                    {
                        trns.Start();
                        FailureHandlingOptions failOpt = trns.GetFailureHandlingOptions();
                        failOpt.SetFailuresPreprocessor(new RoomWarningSwallower());
                        trns.SetFailureHandlingOptions(failOpt);



                        foreach (var mappedColumn in mappedColumnChunck.AsParallel())
                        {

                            try
                            {

                                FamilySymbol revitSection = dic[mappedColumn.sapFrame.Sec];

                                if (revitSection == null)
                                {
                                    continue;
                                }


                                else if (revitSection.Name == mappedColumn.revitFrame.RvtName)
                                {

                                    // set Parameters in revit    
                                    FamilyInstance instance = (FamilyInstance)doc.GetElement(mappedColumn.revitFrame.revitEleId);
                                    revitInstances.Add(instance);

                                    #region SetRotationAngle 

                                    double x = UnitUtils.ConvertToInternalUnits(mappedColumn.sapFrame.StartPoint.X, dis);
                                    double y = UnitUtils.ConvertToInternalUnits(mappedColumn.sapFrame.StartPoint.Y, dis);
                                    double z = UnitUtils.ConvertToInternalUnits(mappedColumn.sapFrame.StartPoint.Z, dis);
                                    XYZ StartPoint = new XYZ(x, y, z);

                                    double x1 = UnitUtils.ConvertToInternalUnits(mappedColumn.sapFrame.EndPoint.X, dis);
                                    double y1 = UnitUtils.ConvertToInternalUnits(mappedColumn.sapFrame.EndPoint.Y, dis);
                                    double z1 = UnitUtils.ConvertToInternalUnits(mappedColumn.sapFrame.EndPoint.Z, dis);
                                    XYZ EndPoint = new XYZ(x1, y1, z1);

                                    double currentRotation = 0;
                                    Parameter colStyleParam = instance.LookupParameter("Column Style");
                                    int ColumnStyle = instance.LookupParameter("Column Style").AsInteger();


                                    if (ColumnStyle == (int)SlantedOrVerticalColumnType.CT_Vertical)
                                    {
                                        LocationPoint pt = instance.Location as LocationPoint;
                                        currentRotation = pt.Rotation;
                                        Line line2 = Line.CreateBound(StartPoint, EndPoint);
                                        double col_angle = (mappedColumn.sapFrame.Beta) * Math.PI / 180;


                                        if (currentRotation != col_angle)
                                        {
                                            ElementTransformUtils.RotateElement(doc, instance.Id, line2, col_angle - currentRotation);


                                        }

                                    }



                                    else
                                    {
                                        instance.LookupParameter("Cross-Section Rotation").Set(mappedColumn.sapFrame.Beta * Math.PI / 180);
                                    }









                                    #endregion

                                    if (analysisModel == "Staad")
                                    {
                                        RevitUtils.setParametersSTAAD(instance, mappedColumn.sapFrame);
                                    }

                                    else if (analysisModel == "Sap")
                                    {
                                        RevitUtils.setParametersSAP(instance, mappedColumn.sapFrame);

                                    }

                                    else if (analysisModel == "Etabs")
                                    {
                                        RevitUtils.setParametersEtabs(instance, mappedColumn.sapFrame);

                                    }

                                }


                                else
                                {
                                    // set Parameters in revit    
                                    FamilyInstance instance = (FamilyInstance)doc.GetElement(mappedColumn.revitFrame.revitEleId);
                                    revitInstances.Add(instance);

                                    instance.Symbol = revitSection;

                                    if (analysisModel == "Staad")
                                    {
                                        RevitUtils.setParametersSTAAD(instance, mappedColumn.sapFrame);
                                    }

                                    else if (analysisModel == "Sap")
                                    {
                                        RevitUtils.setParametersSAP(instance, mappedColumn.sapFrame);

                                    }

                                    else if (analysisModel == "Etabs")
                                    {
                                        RevitUtils.setParametersEtabs(instance, mappedColumn.sapFrame);

                                    }
                                }

                            }
                            catch (Exception ex2)
                            {
                                System.Diagnostics.StackTrace trace = new System.Diagnostics.StackTrace(ex2, true);
                                RevitUtils.Log.Add("Error in updateRevitSectionsForColumns2 Method  sapColumn no " + mappedColumn.sapFrame.ID + ex2.Message + trace.GetFrame(trace.FrameCount - 1).GetFileLineNumber());
                                continue;
                            }


                        }


                        trns.Commit();
                    }


                }


                tg.Assimilate();

                if (ws != null)
                {
                    List<Frame> sapMappedColumns = MappedColumns.Select(a => a.sapFrame).ToList();
                    Frame.setDataInExcelBulk(sapMappedColumns, revitInstances, dis, ws, ref dataToExcel, "not");
                }

            }
        }
        public static List<FamilyInstance> updateRevitGeomertyForColumns2(Document doc, List<SapRevitFrame> MappedColumns, ForgeTypeId dis, Microsoft.Office.Interop.Excel.Worksheet ws, double margin, string analysisModel, object[,] dataToExcel)
        {

            XYZ StartPoint = null;
            XYZ EndPoint = null;
            double x, y, z, x1, y1, z1;
            double margin2 = 0.0;
            Level lvl0;
            Level lvl1;

            lvl0 = new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().FirstOrDefault();
            lvl1 = new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().ElementAtOrDefault(1);

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

            if (lvl1 == null)
            {
                using (Transaction t = new Transaction(doc, "Create Levels"))
                {
                    t.Start();
                    lvl1 = Level.Create(doc, 10.0);
                    lvl1.Name = "Hydra-level-2";

                    t.Commit();

                }
            }


            List<FamilyInstance> revitElements = new List<FamilyInstance>();
            List<FamilyInstance> newlyCreatedElements = new List<FamilyInstance>();



            RevitServices.Transactions.TransactionManager.Instance.ForceCloseTransaction();

            using (TransactionGroup tg = new TransactionGroup(doc, "aaa"))
            {

                tg.Start();
                var chopSize = Frame.chopSize;
                var MappedColumnsChunks = MappedColumns.Chunk(chopSize);

                foreach (var MappedColumnChunk in MappedColumnsChunks)
                {

                    using (Transaction trns = new Transaction(doc, "trns"))
                    {
                        trns.Start();
                        FailureHandlingOptions failOpt = trns.GetFailureHandlingOptions();
                        failOpt.SetFailuresPreprocessor(new RoomWarningSwallower());
                        trns.SetFailureHandlingOptions(failOpt);


                        foreach (var mappedColumn in MappedColumnChunk.AsParallel())
                        {

                            try
                            {


                                margin2 = margin * mappedColumn.revitFrame.Length;
                                double dist1 = Point.getDistance(mappedColumn.revitFrame.StartPoint, mappedColumn.sapFrame.StartPoint);
                                double dist2 = Point.getDistance(mappedColumn.revitFrame.EndPoint, mappedColumn.sapFrame.EndPoint);
                                bool isNotMatched = true;


                                if (dist1 < margin2 && dist2 < margin2)
                                {

                                    FamilyInstance instance = (FamilyInstance)doc.GetElement(mappedColumn.revitFrame.revitEleId);
                                    revitElements.Add(instance);

                                    if (analysisModel == "Staad")
                                    {
                                        RevitUtils.setParametersSTAAD(instance, mappedColumn.sapFrame);
                                    }

                                    else if (analysisModel == "Sap")
                                    {
                                        RevitUtils.setParametersSAP(instance, mappedColumn.sapFrame);

                                    }

                                    else if (analysisModel == "Etabs")
                                    {
                                        RevitUtils.setParametersEtabs(instance, mappedColumn.sapFrame);

                                    }


                                    #region SetRotationAngle 

                                    x = UnitUtils.ConvertToInternalUnits(mappedColumn.sapFrame.StartPoint.X, dis);
                                    y = UnitUtils.ConvertToInternalUnits(mappedColumn.sapFrame.StartPoint.Y, dis);
                                    z = UnitUtils.ConvertToInternalUnits(mappedColumn.sapFrame.StartPoint.Z, dis);
                                    StartPoint = new XYZ(x, y, z);

                                    x1 = UnitUtils.ConvertToInternalUnits(mappedColumn.sapFrame.EndPoint.X, dis);
                                    y1 = UnitUtils.ConvertToInternalUnits(mappedColumn.sapFrame.EndPoint.Y, dis);
                                    z1 = UnitUtils.ConvertToInternalUnits(mappedColumn.sapFrame.EndPoint.Z, dis);
                                    EndPoint = new XYZ(x1, y1, z1);

                                    double currentRotation = 0;
                                    Parameter colStyleParam = instance.LookupParameter("Column Style");
                                    int ColumnStyle = instance.LookupParameter("Column Style").AsInteger();


                                    if (ColumnStyle == (int)SlantedOrVerticalColumnType.CT_Vertical)
                                    {
                                        LocationPoint pt = instance.Location as LocationPoint;
                                        currentRotation = pt.Rotation;
                                        Line line2 = Line.CreateBound(StartPoint, EndPoint);
                                        double col_angle = (mappedColumn.sapFrame.Beta) * Math.PI / 180;


                                        if (currentRotation != col_angle)
                                        {
                                            ElementTransformUtils.RotateElement(doc, instance.Id, line2, col_angle - currentRotation);


                                        }

                                    }



                                    else
                                    {
                                        instance.LookupParameter("Cross-Section Rotation").Set(mappedColumn.sapFrame.Beta * Math.PI / 180);
                                    }

                                    #endregion



                                    continue;
                                }


                                else
                                {
                                    x = UnitUtils.ConvertToInternalUnits(mappedColumn.sapFrame.StartPoint.X, dis);
                                    y = UnitUtils.ConvertToInternalUnits(mappedColumn.sapFrame.StartPoint.Y, dis);
                                    z = UnitUtils.ConvertToInternalUnits(mappedColumn.sapFrame.StartPoint.Z, dis);
                                    StartPoint = new XYZ(x, y, z);

                                    x1 = UnitUtils.ConvertToInternalUnits(mappedColumn.sapFrame.EndPoint.X, dis);
                                    y1 = UnitUtils.ConvertToInternalUnits(mappedColumn.sapFrame.EndPoint.Y, dis);
                                    z1 = UnitUtils.ConvertToInternalUnits(mappedColumn.sapFrame.EndPoint.Z, dis);
                                    EndPoint = new XYZ(x1, y1, z1);


                                    Curve line = Line.CreateBound(StartPoint, EndPoint);
                                    Curve line3 = Line.CreateBound(EndPoint, StartPoint);


                                    FamilyInstance instance = (FamilyInstance)doc.GetElement(mappedColumn.revitFrame.revitEleId);



                                    #region No change in the type 
                                    if ((BuiltInCategory)instance.Category.Id.IntegerValue == BuiltInCategory.OST_StructuralColumns)
                                    {
                                        revitElements.Add(instance);

                                        Parameter colStyleParam = instance.LookupParameter("Column Style");

                                        int ColumnStyle = instance.LookupParameter("Column Style").AsInteger();


                                        if (ColumnStyle == (int)SlantedOrVerticalColumnType.CT_Vertical && mappedColumn.sapFrame.MemberType == MemberType.Column)
                                        {

                                            isNotMatched = false;

                                            #region handling levels elevations

                                            (instance.Location as LocationPoint).Point = StartPoint;
                                            Level baseLevel = (Level)doc.GetElement(instance.LookupParameter("Base Level").AsElementId());
                                            Level topLevel = (Level)doc.GetElement(instance.LookupParameter("Top Level").AsElementId());

                                            double basePointElevation = Convert.ToDouble(mappedColumn.sapFrame.StartPoint.Z);
                                            double topPointElevation = Convert.ToDouble(mappedColumn.sapFrame.EndPoint.Z);

                                            if (basePointElevation > topPointElevation)
                                            {
                                                basePointElevation = Convert.ToDouble(mappedColumn.sapFrame.EndPoint.Z);
                                                topPointElevation = Convert.ToDouble(mappedColumn.sapFrame.StartPoint.Z);
                                            }


                                            double baseLevelElevation = UnitUtils.ConvertFromInternalUnits(baseLevel.Elevation, dis);
                                            double topLevelElevation = UnitUtils.ConvertFromInternalUnits(topLevel.Elevation, dis);


                                            double baseOffset = basePointElevation - baseLevelElevation;
                                            double topOffset = topPointElevation - topLevelElevation;


                                            double base_offset = UnitUtils.ConvertToInternalUnits(basePointElevation - baseLevelElevation, dis);
                                            double top_offset = UnitUtils.ConvertToInternalUnits(topPointElevation - topLevelElevation, dis);


                                            if ((topLevelElevation + topOffset) == baseLevelElevation)
                                            {
                                                instance.LookupParameter("Base Offset").Set(base_offset);
                                                instance.LookupParameter("Top Offset").Set(top_offset);

                                            }
                                            else
                                            {
                                                instance.LookupParameter("Top Offset").Set(top_offset);
                                                instance.LookupParameter("Base Offset").Set(base_offset);
                                            }



                                            #endregion

                                            #region SetRotationAngle 
                                            LocationPoint pt = instance.Location as LocationPoint;
                                            double currentRotation = pt.Rotation;



                                            Line line2 = Line.CreateBound(StartPoint, EndPoint);
                                            double col_angle = (mappedColumn.sapFrame.Beta) * Math.PI / 180;


                                            if (currentRotation != col_angle)
                                            {
                                                // to be checked 30/5/2023
                                                ElementTransformUtils.RotateElement(doc, instance.Id, line2, col_angle - currentRotation);

                                            }



                                            #endregion


                                            if (analysisModel == "Staad")
                                            {
                                                RevitUtils.setParametersSTAAD(instance, mappedColumn.sapFrame);
                                            }

                                            else if (analysisModel == "Sap")
                                            {
                                                RevitUtils.setParametersSAP(instance, mappedColumn.sapFrame);

                                            }

                                        }


                                        else if (ColumnStyle == (int)SlantedOrVerticalColumnType.CT_EndPoint && mappedColumn.sapFrame.MemberType == MemberType.SlanttedColumn)
                                        {

                                            isNotMatched = false;

                                            if (z > z1)
                                            {
                                                (instance.Location as LocationCurve).Curve = line3;
                                            }
                                            else
                                            {
                                                (instance.Location as LocationCurve).Curve = line;
                                            }


                                            if (analysisModel == "Staad")
                                            {
                                                RevitUtils.setParametersSTAAD(instance, mappedColumn.sapFrame);
                                            }

                                            else if (analysisModel == "Sap")
                                            {
                                                RevitUtils.setParametersSAP(instance, mappedColumn.sapFrame);

                                            }


                                            Line line2 = Line.CreateBound(StartPoint, EndPoint);
                                            double col_angle = (mappedColumn.sapFrame.Beta) * Math.PI / 180;
                                            ElementTransformUtils.RotateElement(doc, instance.Id, line2, col_angle);

                                        }


                                    }

                                    #endregion



                                    #region change in the type

                                    if (isNotMatched)
                                    {
                                        // case of VrColumn (to vr column)
                                        if (mappedColumn.sapFrame.MemberType == MemberType.Column)
                                        {

                                            Dictionary<string, FamilySymbol> dic = new Dictionary<string, FamilySymbol>();
                                            List<Frame> Sapframes = new List<Frame>() { mappedColumn.sapFrame };
                                            dic = RevitUtils.createRvtFamilySectionForAllColumn(Sapframes, doc, dis);
                                            FamilySymbol revitSection = dic[mappedColumn.sapFrame.Sec];
                                            if (!revitSection.IsActive)
                                                revitSection.Activate();


                                            FamilyInstance newInstance = doc.Create.NewFamilyInstance(StartPoint, revitSection, lvl0, StructuralType.Column);

                                            newlyCreatedElements.Add(newInstance);
                                            revitElements.Add(instance);

                                            #region handling levels elevations
                                            Parameter pbase = newInstance.LookupParameter("Base Level");
                                            Parameter ptop = newInstance.LookupParameter("Top Level");

                                            //
                                            pbase.Set(lvl0.Id);
                                            ptop.Set(lvl1.Id);

                                            Level baselvl = doc.GetElement(pbase.AsElementId()) as Level;
                                            Level toplvl = doc.GetElement(ptop.AsElementId()) as Level;


                                            double base_elevation = UnitUtils.ConvertFromInternalUnits(baselvl.Elevation, dis);
                                            double top_elevation = UnitUtils.ConvertFromInternalUnits(toplvl.Elevation, dis);

                                            double base_elev = Convert.ToDouble(mappedColumn.sapFrame.StartPoint.Z);
                                            double top_elv = Convert.ToDouble(mappedColumn.sapFrame.EndPoint.Z);


                                            double base_offset = UnitUtils.ConvertToInternalUnits(base_elev - base_elevation, dis);
                                            double top_offset = UnitUtils.ConvertToInternalUnits(top_elv - top_elevation, dis);

                                            if (base_elev - base_elevation > top_elv - top_elevation)

                                            {
                                                base_elev = Convert.ToDouble(mappedColumn.sapFrame.EndPoint.Z);
                                                top_elv = Convert.ToDouble(mappedColumn.sapFrame.StartPoint.Z);

                                                base_offset = UnitUtils.ConvertToInternalUnits(base_elev - base_elevation, dis);
                                                top_offset = UnitUtils.ConvertToInternalUnits(top_elv - top_elevation, dis);
                                            }

                                            newInstance.LookupParameter("Top Offset").Set(top_offset);
                                            newInstance.LookupParameter("Base Offset").Set(base_offset);



                                            #endregion


                                            if (analysisModel == "Staad")
                                            {
                                                RevitUtils.setParametersSTAAD(newInstance, mappedColumn.sapFrame);
                                            }

                                            else if (analysisModel == "Sap")
                                            {
                                                RevitUtils.setParametersSAP(newInstance, mappedColumn.sapFrame);

                                            }


                                            #region Rotation angle
                                            Line line2 = Line.CreateBound(StartPoint, EndPoint);
                                            double col_angle = (mappedColumn.sapFrame.Beta) * Math.PI / 180;

                                            ElementTransformUtils.RotateElement(doc, newInstance.Id, line2, col_angle);


                                            #endregion



                                            doc.Delete(instance.Id);




                                        }


                                        // case of SlanttedColumn (to slantted column)
                                        if (mappedColumn.sapFrame.MemberType == MemberType.SlanttedColumn)
                                        {
                                            FamilySymbol revitSection = RevitUtils.createRvtFamilySectionForColumn(mappedColumn.sapFrame, doc, dis);

                                            if (!revitSection.IsActive)
                                                revitSection.Activate();


                                            FamilyInstance newInstance = doc.Create.NewFamilyInstance(StartPoint, revitSection, lvl0, StructuralType.Column);
                                            newlyCreatedElements.Add(newInstance);
                                            revitElements.Add(instance);


                                            Parameter param = newInstance.LookupParameter("Column Style");
                                            param.Set((int)SlantedOrVerticalColumnType.CT_EndPoint);
                                            LocationCurve instanceCurve = (LocationCurve)newInstance.Location;

                                            if (instanceCurve != null)
                                            {
                                                if (z > z1)
                                                {
                                                    instanceCurve.Curve = line3;
                                                }
                                                else
                                                {
                                                    instanceCurve.Curve = line;

                                                }

                                            }

                                            if (analysisModel == "Staad")
                                            {
                                                RevitUtils.setParametersSTAAD(newInstance, mappedColumn.sapFrame);
                                            }

                                            else if (analysisModel == "Sap")
                                            {
                                                RevitUtils.setParametersSAP(newInstance, mappedColumn.sapFrame);

                                            }


                                            #region Rotation angle

                                            double col_angle = (mappedColumn.sapFrame.Beta) * Math.PI / 180;
                                            newInstance.LookupParameter("Cross-Section Rotation").Set(col_angle);


                                            #endregion



                                            doc.Delete(instance.Id);





                                        }
                                    }




                                    #endregion






                                }

                            }


                            catch (Exception ex2)
                            {
                                System.Diagnostics.StackTrace trace = new System.Diagnostics.StackTrace(ex2, true);
                                RevitUtils.Log.Add("Error in updateRevitGeomertyForColumns2 Method  sapCol no " + mappedColumn.sapFrame.ID + ex2.Message + trace.GetFrame(trace.FrameCount - 1).GetFileLineNumber());
                                continue;
                            }


                        }


                        trns.Commit();

                    }
                }

                if (ws != null)
                {
                    List<Frame> sapMappedColumns = MappedColumns.Select(a => a.sapFrame).ToList();
                    Frame.setDataInExcelBulk(sapMappedColumns, revitElements, dis, ws, ref dataToExcel, "not");
                }

                tg.Assimilate();
            }

            return newlyCreatedElements;
        }

        #region  unUsed
        public static ElementId CreateColumn(Document doc, Frame cl, ForgeTypeId dis, Dictionary<string, FamilySymbol> dic, Microsoft.Office.Interop.Excel.Worksheet ws, String analysisModel)
        {

            #region defintions for creation of beams and columns 
            FamilySymbol revitSection = null;
            double x, y, z, x1, y1, z1;
            XYZ StartPoint = null;
            XYZ EndPoint = null;

        #endregion

        ColumnModeling:
            using (Transaction trns = new Transaction(doc, "aa"))

            {
                trns.Start();
                FailureHandlingOptions failOpt = trns.GetFailureHandlingOptions();
                failOpt.SetFailuresPreprocessor(new RoomWarningSwallower());
                trns.SetFailureHandlingOptions(failOpt);

                x = UnitUtils.ConvertToInternalUnits(cl.StartPoint.X, dis);
                y = UnitUtils.ConvertToInternalUnits(cl.StartPoint.Y, dis);
                z = UnitUtils.ConvertToInternalUnits(cl.StartPoint.Z, dis);
                StartPoint = new XYZ(x, y, z);

                x1 = UnitUtils.ConvertToInternalUnits(cl.EndPoint.X, dis);
                y1 = UnitUtils.ConvertToInternalUnits(cl.EndPoint.Y, dis);
                z1 = UnitUtils.ConvertToInternalUnits(cl.EndPoint.Z, dis);
                EndPoint = new XYZ(x1, y1, z1);

                Curve line = Line.CreateBound(StartPoint, EndPoint);
                Curve line3 = Line.CreateBound(EndPoint, StartPoint);



                Level lvl0 = new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().ToList().First();
                Level lvl1 = new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().ToList()[1];

                if (lvl0.Elevation > lvl1.Elevation)
                {
                    lvl0 = new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().ToList()[1];
                    lvl1 = new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().ToList().First();
                }


                try
                {

                    revitSection = dic[cl.Sec];

                    if (!revitSection.IsActive)
                        revitSection.Activate();

                    FamilyInstance instance = doc.Create.NewFamilyInstance(StartPoint, revitSection, lvl0, StructuralType.Column);


                    if (cl.MemberType == MemberType.SlanttedColumn)
                    {

                        //instance.get_Parameter(BuiltInParameter.SLANTED_COLUMN_GEOMETRY_TREATMENT_TOP).Set(0);


                        Parameter param = instance.LookupParameter("Column Style");

                        param.Set((int)SlantedOrVerticalColumnType.CT_EndPoint);



                        LocationCurve instanceCurve = (LocationCurve)instance.Location;
                        if (instanceCurve != null)
                        {
                            if (z > z1)
                            {
                                instanceCurve.Curve = line3;
                            }
                            else
                            {
                                instanceCurve.Curve = line;

                            }

                        }

                    }

                    else
                    {
                        #region handling levels elevations
                        Parameter pbase = instance.LookupParameter("Base Level");
                        Parameter ptop = instance.LookupParameter("Top Level");

                        //
                        pbase.Set(lvl0.Id);
                        ptop.Set(lvl1.Id);


                        Level baselvl = doc.GetElement(pbase.AsElementId()) as Level;
                        Level toplvl = doc.GetElement(ptop.AsElementId()) as Level;


                        double basePointElevation = Convert.ToDouble(cl.StartPoint.Z);
                        double topPointElevation = Convert.ToDouble(cl.EndPoint.Z);

                        if (basePointElevation > topPointElevation)
                        {
                            basePointElevation = Convert.ToDouble(cl.EndPoint.Z);
                            topPointElevation = Convert.ToDouble(cl.StartPoint.Z);

                        }


                        double baseLevelElevation = UnitUtils.ConvertFromInternalUnits(baselvl.Elevation, dis);
                        double topLevelElevation = UnitUtils.ConvertFromInternalUnits(toplvl.Elevation, dis);


                        double baseOffset = basePointElevation - baseLevelElevation;
                        double topOffset = topPointElevation - topLevelElevation;

                        double base_offset = UnitUtils.ConvertToInternalUnits(basePointElevation - baseLevelElevation, dis);
                        double top_offset = UnitUtils.ConvertToInternalUnits(topPointElevation - topLevelElevation, dis);




                        if ((topLevelElevation + topOffset) == baseLevelElevation)
                        {
                            instance.LookupParameter("Base Offset").Set(base_offset);
                            instance.LookupParameter("Top Offset").Set(top_offset);

                        }
                        else
                        {
                            instance.LookupParameter("Top Offset").Set(top_offset);
                            instance.LookupParameter("Base Offset").Set(base_offset);
                        }


                        #endregion

                    }




                    // set orientation angle for columns
                    Line line2 = Line.CreateBound(StartPoint, EndPoint);
                    double col_angle = (cl.Beta) * Math.PI / 180;
                    ElementTransformUtils.RotateElement(doc, instance.Id, line2, col_angle);

                    // set Parameters 

                    if (analysisModel == "Staad")
                    {
                        RevitUtils.setParametersSTAAD(instance, cl);
                    }

                    else if (analysisModel == "Sap")
                    {
                        RevitUtils.setParametersSAP(instance, cl);

                    }




                    trns.Commit();
                    if (ws != null)
                    {
                        Frame.setDataInExcel(cl, instance, dis, (Microsoft.Office.Interop.Excel.Worksheet)ws);

                    }

                    try
                    {
                        return instance.Id;

                    }
                    catch (Exception)
                    {

                        return null;
                    }
                }

                catch (Exception ex2)

                {
                    System.Diagnostics.StackTrace trace = new System.Diagnostics.StackTrace(ex2, true);
                    TaskDialog.Show("Error", ex2.Message + " ,line of the error is " + trace.GetFrame(trace.FrameCount - 1).GetFileLineNumber());
                    return null;
                }

            }





        }

        public static List<ElementId> CreateAllColumns(Document doc, List<Frame> columns, ForgeTypeId dis, Dictionary<string, FamilySymbol> dic, Microsoft.Office.Interop.Excel.Worksheet ws, string analysisModel)
        {

            List<ElementId> result = new List<ElementId>();

            #region defintions for creation of beams and columns 
            FamilySymbol revitSection = null;
            double x, y, z, x1, y1, z1;
            XYZ StartPoint = null;
            XYZ EndPoint = null;
            Curve line;
            Curve line3;

            double angle = 0;
            Level lvl0;
            Level lvl1;
            Parameter param;
            LocationCurve instanceCurve;
            Parameter pbase;
            Parameter ptop;
            Level baselvl;
            Level toplvl;
            double basePointElevation;
            double topPointElevation;
            double baseLevelElevation;
            double topLevelElevation;
            double baseOffset;
            double topOffset;
            double base_offset;
            double top_offset;
            Line line2;
            double col_angle;


            lvl0 = new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().FirstOrDefault();
            lvl1 = new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().ElementAtOrDefault(1);

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

            if (lvl1 == null)
            {
                using (Transaction t = new Transaction(doc, "Create Levels"))
                {
                    t.Start();
                    lvl1 = Level.Create(doc, 10.0);
                    lvl1.Name = "Hydra-level-2";

                    t.Commit();

                }
            }


            if (lvl0.Elevation > lvl1.Elevation)
            {
                lvl0 = new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().ToList()[1];
                lvl1 = new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().ToList().First();
            }

            #endregion


            foreach (var cl in columns.AsParallel())
            {

                x = UnitUtils.ConvertToInternalUnits(cl.StartPoint.X, dis);
                y = UnitUtils.ConvertToInternalUnits(cl.StartPoint.Y, dis);
                z = UnitUtils.ConvertToInternalUnits(cl.StartPoint.Z, dis);
                StartPoint = new XYZ(x, y, z);

                x1 = UnitUtils.ConvertToInternalUnits(cl.EndPoint.X, dis);
                y1 = UnitUtils.ConvertToInternalUnits(cl.EndPoint.Y, dis);
                z1 = UnitUtils.ConvertToInternalUnits(cl.EndPoint.Z, dis);
                EndPoint = new XYZ(x1, y1, z1);




                using (Transaction trns = new Transaction(doc, "aa"))

                {
                    try
                    {
                        line = Line.CreateBound(StartPoint, EndPoint);
                        line3 = Line.CreateBound(EndPoint, StartPoint);
                        revitSection = dic[cl.Sec];

                        trns.Start();
                        FailureHandlingOptions failOpt = trns.GetFailureHandlingOptions();
                        failOpt.SetFailuresPreprocessor(new RoomWarningSwallower());
                        trns.SetFailureHandlingOptions(failOpt);

                        if (!revitSection.IsActive)
                            revitSection.Activate();

                        FamilyInstance instance = doc.Create.NewFamilyInstance(StartPoint, revitSection, lvl0, StructuralType.Column);

                        if (cl.MemberType == MemberType.SlanttedColumn)
                        {
                            param = instance.LookupParameter("Column Style");
                            param.Set((int)SlantedOrVerticalColumnType.CT_EndPoint);

                            instanceCurve = (LocationCurve)instance.Location;
                            if (instanceCurve != null)
                            {
                                if (z > z1)
                                {
                                    instanceCurve.Curve = line3;
                                }
                                else
                                {
                                    instanceCurve.Curve = line;

                                }

                            }

                        }

                        else
                        {
                            #region handling levels elevations
                            pbase = instance.LookupParameter("Base Level");
                            ptop = instance.LookupParameter("Top Level");

                            //
                            pbase.Set(lvl0.Id);
                            ptop.Set(lvl1.Id);


                            baselvl = doc.GetElement(pbase.AsElementId()) as Level;
                            toplvl = doc.GetElement(ptop.AsElementId()) as Level;


                            basePointElevation = Convert.ToDouble(cl.StartPoint.Z);
                            topPointElevation = Convert.ToDouble(cl.EndPoint.Z);

                            if (basePointElevation > topPointElevation)
                            {
                                basePointElevation = Convert.ToDouble(cl.EndPoint.Z);
                                topPointElevation = Convert.ToDouble(cl.StartPoint.Z);

                            }


                            baseLevelElevation = UnitUtils.ConvertFromInternalUnits(baselvl.Elevation, dis);
                            topLevelElevation = UnitUtils.ConvertFromInternalUnits(toplvl.Elevation, dis);


                            baseOffset = basePointElevation - baseLevelElevation;
                            topOffset = topPointElevation - topLevelElevation;

                            base_offset = UnitUtils.ConvertToInternalUnits(basePointElevation - baseLevelElevation, dis);
                            top_offset = UnitUtils.ConvertToInternalUnits(topPointElevation - topLevelElevation, dis);




                            if ((topLevelElevation + topOffset) == baseLevelElevation)
                            {
                                instance.LookupParameter("Base Offset").Set(base_offset);
                                instance.LookupParameter("Top Offset").Set(top_offset);

                            }
                            else
                            {
                                instance.LookupParameter("Top Offset").Set(top_offset);
                                instance.LookupParameter("Base Offset").Set(base_offset);
                            }


                            #endregion

                        }


                        line2 = Line.CreateBound(StartPoint, EndPoint);
                        col_angle = (cl.Beta) * Math.PI / 180;
                        ElementTransformUtils.RotateElement(doc, instance.Id, line2, col_angle);

                        // set Parameters 

                        if (analysisModel == "Staad")
                        {
                            RevitUtils.setParametersSTAAD(instance, cl);
                        }

                        else if (analysisModel == "Sap")
                        {
                            RevitUtils.setParametersSAP(instance, cl);

                        }
                        else if (analysisModel == "Etabs")
                        {
                            RevitUtils.setParametersEtabs(instance, cl);
                        }

                        trns.Commit();



                        if (ws != null)
                        {
                            Frame.setDataInExcel(cl, instance, dis, (Microsoft.Office.Interop.Excel.Worksheet)ws);

                        }


                        try
                        {
                            result.Add(instance.Id);

                        }
                        catch (Exception)
                        {
                            continue;
                        }

                    }


                    catch (Exception ex2)

                    {
                        System.Diagnostics.StackTrace trace = new System.Diagnostics.StackTrace(ex2, true);
                        TaskDialog.Show("Error", ex2.Message + " ,line of the error is " + trace.GetFrame(trace.FrameCount - 1).GetFileLineNumber());
                        continue;
                    }

                }

            }

            return result;

        }

        public static void updateRevitSectionsForColumns(Document doc, List<SapRevitFrame> MappedColumns, Dictionary<string, FamilySymbol> dic, ForgeTypeId dis, Microsoft.Office.Interop.Excel.Worksheet ws, string analysisModel)
        {

            foreach (var mappedColumn in MappedColumns.AsParallel())
            {

                FamilySymbol revitSection = dic[mappedColumn.sapFrame.Sec];

                if (revitSection == null)
                {
                    continue;
                }


                else if (revitSection.Name == mappedColumn.revitFrame.RvtName)
                {

                    // set Parameters in revit    
                    FamilyInstance instance = (FamilyInstance)doc.GetElement(mappedColumn.revitFrame.revitEleId);
                    using (Transaction trns = new Transaction(doc, "aa"))
                    {
                        trns.Start();
                        FailureHandlingOptions failOpt = trns.GetFailureHandlingOptions();

                        #region SetRotationAngle 

                        double x = UnitUtils.ConvertToInternalUnits(mappedColumn.sapFrame.StartPoint.X, dis);
                        double y = UnitUtils.ConvertToInternalUnits(mappedColumn.sapFrame.StartPoint.Y, dis);
                        double z = UnitUtils.ConvertToInternalUnits(mappedColumn.sapFrame.StartPoint.Z, dis);
                        XYZ StartPoint = new XYZ(x, y, z);

                        double x1 = UnitUtils.ConvertToInternalUnits(mappedColumn.sapFrame.EndPoint.X, dis);
                        double y1 = UnitUtils.ConvertToInternalUnits(mappedColumn.sapFrame.EndPoint.Y, dis);
                        double z1 = UnitUtils.ConvertToInternalUnits(mappedColumn.sapFrame.EndPoint.Z, dis);
                        XYZ EndPoint = new XYZ(x1, y1, z1);

                        double currentRotation = 0;
                        Parameter colStyleParam = instance.LookupParameter("Column Style");
                        int ColumnStyle = instance.LookupParameter("Column Style").AsInteger();


                        if (ColumnStyle == (int)SlantedOrVerticalColumnType.CT_Vertical)
                        {
                            LocationPoint pt = instance.Location as LocationPoint;
                            currentRotation = pt.Rotation;
                            Line line2 = Line.CreateBound(StartPoint, EndPoint);
                            double col_angle = (mappedColumn.sapFrame.Beta) * Math.PI / 180;


                            if (currentRotation != col_angle)
                            {
                                ElementTransformUtils.RotateElement(doc, instance.Id, line2, col_angle - currentRotation);


                            }

                        }



                        else
                        {
                            instance.LookupParameter("Cross-Section Rotation").Set(mappedColumn.sapFrame.Beta * Math.PI / 180);
                        }









                        #endregion

                        if (analysisModel == "Staad")
                        {
                            RevitUtils.setParametersSTAAD(instance, mappedColumn.sapFrame);
                        }

                        else if (analysisModel == "Sap")
                        {
                            RevitUtils.setParametersSAP(instance, mappedColumn.sapFrame);

                        }


                        trns.Commit();
                    }


                    if (ws != null)
                    {
                        Frame.setDataInExcel(mappedColumn.sapFrame, instance, dis, (Microsoft.Office.Interop.Excel.Worksheet)ws);

                    }

                }


                else
                {
                    // set Parameters in revit    
                    FamilyInstance instance = (FamilyInstance)doc.GetElement(mappedColumn.revitFrame.revitEleId);
                    using (Transaction trns = new Transaction(doc, "aa"))
                    {
                        trns.Start();
                        FailureHandlingOptions failOpt = trns.GetFailureHandlingOptions();
                        instance.Symbol = revitSection;

                        if (analysisModel == "Staad")
                        {
                            RevitUtils.setParametersSTAAD(instance, mappedColumn.sapFrame);
                        }

                        else if (analysisModel == "Sap")
                        {
                            RevitUtils.setParametersSAP(instance, mappedColumn.sapFrame);

                        }

                        trns.Commit();
                    }

                    if (ws != null)
                    {
                        Frame.setDataInExcel(mappedColumn.sapFrame, instance, dis, (Microsoft.Office.Interop.Excel.Worksheet)ws);

                    }


                }

            }


        }

        public static List<FamilyInstance> updateRevitGeomertyForColumns(Document doc, List<SapRevitFrame> MappedColumns, ForgeTypeId dis, Microsoft.Office.Interop.Excel.Worksheet ws, double margin, string analysisModel)
        {

            XYZ StartPoint = null;
            XYZ EndPoint = null;
            double x, y, z, x1, y1, z1;
            double margin2 = 0.0;
            List<FamilyInstance> newlyCreatedElements = new List<FamilyInstance>();
            Level lvl0;
            Level lvl1;

            lvl0 = new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().FirstOrDefault();
            lvl1 = new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().ElementAtOrDefault(1);

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

            if (lvl1 == null)
            {
                using (Transaction t = new Transaction(doc, "Create Levels"))
                {
                    t.Start();
                    lvl1 = Level.Create(doc, 10.0);
                    lvl1.Name = "Hydra-level-2";

                    t.Commit();

                }
            }



            foreach (var mappedColumn in MappedColumns.AsParallel())
            {
                try
                {
                    margin2 = margin * mappedColumn.revitFrame.Length;
                    double dist1 = Point.getDistance(mappedColumn.revitFrame.StartPoint, mappedColumn.sapFrame.StartPoint);
                    double dist2 = Point.getDistance(mappedColumn.revitFrame.EndPoint, mappedColumn.sapFrame.EndPoint);
                    bool isNotMatched = true;


                    if (dist1 < margin2 && dist2 < margin2)
                    {

                        FamilyInstance instance = (FamilyInstance)doc.GetElement(mappedColumn.revitFrame.revitEleId);
                        using (Transaction trns = new Transaction(doc, "aa"))
                        {
                            trns.Start();
                            FailureHandlingOptions failOpt = trns.GetFailureHandlingOptions();


                            if (analysisModel == "Staad")
                            {
                                RevitUtils.setParametersSTAAD(instance, mappedColumn.sapFrame);
                            }

                            else if (analysisModel == "Sap")
                            {
                                RevitUtils.setParametersSAP(instance, mappedColumn.sapFrame);

                            }



                            #region SetRotationAngle 

                            x = UnitUtils.ConvertToInternalUnits(mappedColumn.sapFrame.StartPoint.X, dis);
                            y = UnitUtils.ConvertToInternalUnits(mappedColumn.sapFrame.StartPoint.Y, dis);
                            z = UnitUtils.ConvertToInternalUnits(mappedColumn.sapFrame.StartPoint.Z, dis);
                            StartPoint = new XYZ(x, y, z);

                            x1 = UnitUtils.ConvertToInternalUnits(mappedColumn.sapFrame.EndPoint.X, dis);
                            y1 = UnitUtils.ConvertToInternalUnits(mappedColumn.sapFrame.EndPoint.Y, dis);
                            z1 = UnitUtils.ConvertToInternalUnits(mappedColumn.sapFrame.EndPoint.Z, dis);
                            EndPoint = new XYZ(x1, y1, z1);

                            double currentRotation = 0;
                            Parameter colStyleParam = instance.LookupParameter("Column Style");
                            int ColumnStyle = instance.LookupParameter("Column Style").AsInteger();


                            if (ColumnStyle == (int)SlantedOrVerticalColumnType.CT_Vertical)
                            {
                                LocationPoint pt = instance.Location as LocationPoint;
                                currentRotation = pt.Rotation;
                                Line line2 = Line.CreateBound(StartPoint, EndPoint);
                                double col_angle = (mappedColumn.sapFrame.Beta) * Math.PI / 180;


                                if (currentRotation != col_angle)
                                {
                                    ElementTransformUtils.RotateElement(doc, instance.Id, line2, col_angle - currentRotation);


                                }

                            }



                            else
                            {
                                instance.LookupParameter("Cross-Section Rotation").Set(mappedColumn.sapFrame.Beta * Math.PI / 180);
                            }









                            #endregion


                            trns.Commit();
                        }

                        if (ws != null)
                        {
                            Frame.setDataInExcel(mappedColumn.sapFrame, instance, dis, (Microsoft.Office.Interop.Excel.Worksheet)ws);

                        }


                        continue;
                    }


                    else
                    {
                        x = UnitUtils.ConvertToInternalUnits(mappedColumn.sapFrame.StartPoint.X, dis);
                        y = UnitUtils.ConvertToInternalUnits(mappedColumn.sapFrame.StartPoint.Y, dis);
                        z = UnitUtils.ConvertToInternalUnits(mappedColumn.sapFrame.StartPoint.Z, dis);
                        StartPoint = new XYZ(x, y, z);

                        x1 = UnitUtils.ConvertToInternalUnits(mappedColumn.sapFrame.EndPoint.X, dis);
                        y1 = UnitUtils.ConvertToInternalUnits(mappedColumn.sapFrame.EndPoint.Y, dis);
                        z1 = UnitUtils.ConvertToInternalUnits(mappedColumn.sapFrame.EndPoint.Z, dis);
                        EndPoint = new XYZ(x1, y1, z1);


                        Curve line = Line.CreateBound(StartPoint, EndPoint);
                        Curve line3 = Line.CreateBound(EndPoint, StartPoint);




                        FamilyInstance instance = (FamilyInstance)doc.GetElement(mappedColumn.revitFrame.revitEleId);

                        try
                        {

                            #region No change in the type 
                            if ((BuiltInCategory)instance.Category.Id.IntegerValue == BuiltInCategory.OST_StructuralColumns)
                            {

                                Parameter colStyleParam = instance.LookupParameter("Column Style");

                                int ColumnStyle = instance.LookupParameter("Column Style").AsInteger();


                                if (ColumnStyle == (int)SlantedOrVerticalColumnType.CT_Vertical && mappedColumn.sapFrame.MemberType == MemberType.Column)
                                {
                                    using (Transaction trns = new Transaction(doc, "aa"))
                                    {
                                        isNotMatched = false;
                                        #region handling levels elevations

                                        trns.Start();

                                        FailureHandlingOptions failOpt = trns.GetFailureHandlingOptions();
                                        failOpt.SetFailuresPreprocessor(new RoomWarningSwallower());
                                        trns.SetFailureHandlingOptions(failOpt);

                                        (instance.Location as LocationPoint).Point = StartPoint;
                                        Level baseLevel = (Level)doc.GetElement(instance.LookupParameter("Base Level").AsElementId());
                                        Level topLevel = (Level)doc.GetElement(instance.LookupParameter("Top Level").AsElementId());


                                        double base_elevation = UnitUtils.ConvertFromInternalUnits(baseLevel.Elevation, dis);
                                        double top_elevation = UnitUtils.ConvertFromInternalUnits(topLevel.Elevation, dis);

                                        double base_elev = Convert.ToDouble(mappedColumn.sapFrame.StartPoint.Z);
                                        double top_elv = Convert.ToDouble(mappedColumn.sapFrame.EndPoint.Z);

                                        double base_offset = UnitUtils.ConvertToInternalUnits(base_elev - base_elevation, dis);
                                        double top_offset = UnitUtils.ConvertToInternalUnits(top_elv - top_elevation, dis);

                                        if (base_elev - base_elevation > top_elv - top_elevation)

                                        {
                                            base_elev = Convert.ToDouble(mappedColumn.sapFrame.EndPoint.Z);
                                            top_elv = Convert.ToDouble(mappedColumn.sapFrame.StartPoint.Z);

                                            base_offset = UnitUtils.ConvertToInternalUnits(base_elev - base_elevation, dis);
                                            top_offset = UnitUtils.ConvertToInternalUnits(top_elv - top_elevation, dis);
                                        }

                                        instance.LookupParameter("Top Offset").Set(top_offset);
                                        instance.LookupParameter("Base Offset").Set(base_offset);


                                        #endregion

                                        #region SetRotationAngle 
                                        LocationPoint pt = instance.Location as LocationPoint;
                                        double currentRotation = pt.Rotation;



                                        Line line2 = Line.CreateBound(StartPoint, EndPoint);
                                        double col_angle = (mappedColumn.sapFrame.Beta) * Math.PI / 180;


                                        if (currentRotation != col_angle)
                                        {
                                            ElementTransformUtils.RotateElement(doc, instance.Id, line2, col_angle - currentRotation);

                                        }



                                        #endregion


                                        if (analysisModel == "Staad")
                                        {
                                            RevitUtils.setParametersSTAAD(instance, mappedColumn.sapFrame);
                                        }

                                        else if (analysisModel == "Sap")
                                        {
                                            RevitUtils.setParametersSAP(instance, mappedColumn.sapFrame);

                                        }



                                        trns.Commit();

                                        if (ws != null)
                                        {
                                            Frame.setDataInExcel(mappedColumn.sapFrame, instance, dis, (Microsoft.Office.Interop.Excel.Worksheet)ws);
                                        }


                                    }

                                }


                                else if (ColumnStyle == (int)SlantedOrVerticalColumnType.CT_EndPoint && mappedColumn.sapFrame.MemberType == MemberType.SlanttedColumn)
                                {
                                    using (Transaction trns = new Transaction(doc, "aaa"))
                                    {
                                        isNotMatched = false;
                                        trns.Start();


                                        if (z > z1)
                                        {
                                            (instance.Location as LocationCurve).Curve = line3;
                                        }
                                        else
                                        {
                                            (instance.Location as LocationCurve).Curve = line;
                                        }


                                        if (analysisModel == "Staad")
                                        {
                                            RevitUtils.setParametersSTAAD(instance, mappedColumn.sapFrame);
                                        }

                                        else if (analysisModel == "Sap")
                                        {
                                            RevitUtils.setParametersSAP(instance, mappedColumn.sapFrame);

                                        }


                                        Line line2 = Line.CreateBound(StartPoint, EndPoint);
                                        double col_angle = (mappedColumn.sapFrame.Beta) * Math.PI / 180;
                                        ElementTransformUtils.RotateElement(doc, instance.Id, line2, col_angle);

                                        trns.Commit();


                                        if (ws != null)
                                        {
                                            Frame.setDataInExcel(mappedColumn.sapFrame, instance, dis, (Microsoft.Office.Interop.Excel.Worksheet)ws);
                                        }


                                    }



                                }


                            }

                            #endregion

                            #region change in the type

                            if (isNotMatched)
                            {
                                // case of VrColumn (to vr column)
                                if (mappedColumn.sapFrame.MemberType == MemberType.Column)
                                {
                                    using (Transaction trns = new Transaction(doc, "changeType to vertical"))
                                    {

                                        trns.Start();
                                        FailureHandlingOptions failOpt = trns.GetFailureHandlingOptions();
                                        failOpt.SetFailuresPreprocessor(new RoomWarningSwallower());
                                        trns.SetFailureHandlingOptions(failOpt);
                                        FamilySymbol revitSection = RevitUtils.createRvtFamilySectionForColumn(mappedColumn.sapFrame, doc, dis);

                                        if (!revitSection.IsActive)
                                            revitSection.Activate();



                                        FamilyInstance newInstance = doc.Create.NewFamilyInstance(StartPoint, revitSection, lvl0, StructuralType.Column);

                                        #region handling levels elevations
                                        Parameter pbase = newInstance.LookupParameter("Base Level");
                                        Parameter ptop = newInstance.LookupParameter("Top Level");

                                        //
                                        pbase.Set(lvl0.Id);
                                        ptop.Set(lvl1.Id);

                                        Level baselvl = doc.GetElement(pbase.AsElementId()) as Level;
                                        Level toplvl = doc.GetElement(ptop.AsElementId()) as Level;


                                        double base_elevation = UnitUtils.ConvertFromInternalUnits(baselvl.Elevation, dis);
                                        double top_elevation = UnitUtils.ConvertFromInternalUnits(toplvl.Elevation, dis);

                                        double base_elev = Convert.ToDouble(mappedColumn.sapFrame.StartPoint.Z);
                                        double top_elv = Convert.ToDouble(mappedColumn.sapFrame.EndPoint.Z);


                                        double base_offset = UnitUtils.ConvertToInternalUnits(base_elev - base_elevation, dis);
                                        double top_offset = UnitUtils.ConvertToInternalUnits(top_elv - top_elevation, dis);

                                        if (base_elev - base_elevation > top_elv - top_elevation)

                                        {
                                            base_elev = Convert.ToDouble(mappedColumn.sapFrame.EndPoint.Z);
                                            top_elv = Convert.ToDouble(mappedColumn.sapFrame.StartPoint.Z);

                                            base_offset = UnitUtils.ConvertToInternalUnits(base_elev - base_elevation, dis);
                                            top_offset = UnitUtils.ConvertToInternalUnits(top_elv - top_elevation, dis);
                                        }

                                        newInstance.LookupParameter("Top Offset").Set(top_offset);
                                        newInstance.LookupParameter("Base Offset").Set(base_offset);



                                        #endregion
                                        if (analysisModel == "Staad")
                                        {
                                            RevitUtils.setParametersSTAAD(newInstance, mappedColumn.sapFrame);
                                        }

                                        else if (analysisModel == "Sap")
                                        {
                                            RevitUtils.setParametersSAP(newInstance, mappedColumn.sapFrame);

                                        }


                                        #region Rotation angle
                                        Line line2 = Line.CreateBound(StartPoint, EndPoint);
                                        double col_angle = (mappedColumn.sapFrame.Beta) * Math.PI / 180;

                                        ElementTransformUtils.RotateElement(doc, newInstance.Id, line2, col_angle);


                                        #endregion

                                        //Line line2 = Line.CreateBound(StartPoint, EndPoint);
                                        //double col_angle = (mappedColumn.sapFrame.Beta - 90) * Math.PI / 180;
                                        //ElementTransformUtils.RotateElement(doc, newInstance.Id, line2, col_angle);


                                        doc.Delete(instance.Id);
                                        trns.Commit();


                                        if (ws != null)
                                        {

                                            Frame.setDataInExcel(mappedColumn.sapFrame, newInstance, dis, (Microsoft.Office.Interop.Excel.Worksheet)ws);
                                        }

                                        newlyCreatedElements.Add(newInstance);
                                    }

                                }

                                // case of SlanttedColumn (to slantted column)
                                if (mappedColumn.sapFrame.MemberType == MemberType.SlanttedColumn)
                                {

                                    using (Transaction trns = new Transaction(doc, "changeType to slantted"))
                                    {
                                        trns.Start();
                                        FailureHandlingOptions failOpt = trns.GetFailureHandlingOptions();
                                        failOpt.SetFailuresPreprocessor(new RoomWarningSwallower());
                                        trns.SetFailureHandlingOptions(failOpt);
                                        FamilySymbol revitSection = RevitUtils.createRvtFamilySectionForColumn(mappedColumn.sapFrame, doc, dis);


                                        if (!revitSection.IsActive)
                                            revitSection.Activate();


                                        FamilyInstance newInstance = doc.Create.NewFamilyInstance(StartPoint, revitSection, lvl0, StructuralType.Column);

                                        Parameter param = newInstance.LookupParameter("Column Style");
                                        param.Set((int)SlantedOrVerticalColumnType.CT_EndPoint);
                                        LocationCurve instanceCurve = (LocationCurve)newInstance.Location;

                                        if (instanceCurve != null)
                                        {
                                            if (z > z1)
                                            {
                                                instanceCurve.Curve = line3;
                                            }
                                            else
                                            {
                                                instanceCurve.Curve = line;

                                            }

                                        }

                                        if (analysisModel == "Staad")
                                        {
                                            RevitUtils.setParametersSTAAD(newInstance, mappedColumn.sapFrame);
                                        }

                                        else if (analysisModel == "Sap")
                                        {
                                            RevitUtils.setParametersSAP(newInstance, mappedColumn.sapFrame);

                                        }


                                        #region Rotation angle

                                        double col_angle = (mappedColumn.sapFrame.Beta) * Math.PI / 180;
                                        newInstance.LookupParameter("Cross-Section Rotation").Set(col_angle);


                                        #endregion


                                        //Line line2 = Line.CreateBound(StartPoint, EndPoint);
                                        //double col_angle = (mappedColumn.sapFrame.Beta - 90) * Math.PI / 180;
                                        //ElementTransformUtils.RotateElement(doc, newInstance.Id, line2, col_angle);


                                        doc.Delete(instance.Id);
                                        trns.Commit();
                                        if (ws != null)
                                        {
                                            Frame.setDataInExcel(mappedColumn.sapFrame, newInstance, dis, (Microsoft.Office.Interop.Excel.Worksheet)ws);

                                        }
                                        newlyCreatedElements.Add(newInstance);

                                    }

                                }
                            }




                            #endregion

                        }

                        catch (Exception ex2)

                        {
                            System.Diagnostics.StackTrace trace = new System.Diagnostics.StackTrace(ex2, true);
                            TaskDialog.Show("Error", "error in beam no " + mappedColumn.sapFrame.ID + ex2.Message + trace.GetFrame(trace.FrameCount - 1).GetFileLineNumber());

                            continue;

                        }





                    }
                }

                catch (Exception eeex)
                {
                    System.Diagnostics.StackTrace trace = new System.Diagnostics.StackTrace(eeex, true);
                    TaskDialog.Show("Error", "error in beam no" + mappedColumn.sapFrame.ID + eeex.Message + trace.GetFrame(trace.FrameCount - 1).GetFileLineNumber());

                    continue;
                }



            }









            return newlyCreatedElements;
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

        public bool [] Restrains { get; set; }

        public int SerialNumber { get; set; }
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


        public Point(XYZ pt)
        {
            X = pt.X;
            Y = pt.Y;
            Z = pt.Z;
        }

        public Point(Rhino.Geometry.Point3d pt)
        {
            X = pt.X;
            Y = pt.Y;
            Z = pt.Z;
        }

        public static Vector3 ToVector(Point p)
        {
            return new Vector3((float)p.X, (float)p.Y, (float)p.Z);
        }

        public static XYZ toXYZ(Point p)
        {
            return new XYZ(p.X, p.Y, p.Z);
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

        public static void getMemberType(Frame F, bool isColumn = false)
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
                if (isColumn)
                {
                    F.MemberType = MemberType.Column; // vertical column

                }
                else
                {
                    F.MemberType = MemberType.Beam; // vertical column

                }
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
                        if (isColumn)
                        {
                            F.MemberType = MemberType.Column; // vrcolumn

                        }
                        else
                        {
                            F.MemberType = MemberType.Beam; // vrcolumn

                        }
                    }

                    //SlanttedColumn
                    else
                    {
                        if (isColumn)
                        {
                            F.MemberType = MemberType.SlanttedColumn; // slanted column

                        }
                        else
                        {
                            F.MemberType = MemberType.Beam; // slanted column

                        }

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

        public static List<Point> OrderPointsCounterClockwise(List<Point> points)
        {
            // Calculate the centroid of all points
            double centroidX = points.Average(p => p.X);
            double centroidY = points.Average(p => p.Y);


            // Sort the points counterclockwise using the angle between the centroid and each point
            List<Point> orderedPoints = points
                .OrderBy(p => Math.Atan2(p.Y - centroidY, p.X - centroidX)) // Atan2 handles the angle computation
                .ToList();

            return orderedPoints;
        }

        public static List<Point> RemoveDuplicatePoints(List<Point> points, double tolerance)
        {
            var uniquePoints = new List<Point>();

            foreach (Point pt in points)
            {
                bool isDuplicate = false;
                foreach (Point uniquePt in uniquePoints)
                {
                    if  (getDistance(pt, uniquePt) < tolerance)
                    {
                        isDuplicate = true;
                        break;
                    }
                }

                if (!isDuplicate)
                    uniquePoints.Add(pt);
            }

            return uniquePoints;
        }
        public static List<Point> OrderPointsCounterClockwise2(List<Point> points)
        {
            if (points.Count < 3) return points; // At least 3 points needed to define a plane

            // Step 1: Calculate normal vector of the plane using the first three points
            Vector3 v1 = new Vector3((float)(points[1].X - points[0].X), (float)(points[1].Y - points[0].Y), (float)(points[1].Z - points[0].Z));
            Vector3 v2 = new Vector3((float)(points[2].X - points[0].X), (float)(points[2].Y - points[0].Y), (float)(points[2].Z - points[0].Z));
            Vector3 normal = Vector3.Cross(v1, v2); // Cross product to get the normal of the plane

            // Step 2: Calculate centroid of the points
            double centroidX = points.Average(p => p.X);
            double centroidY = points.Average(p => p.Y);
            double centroidZ = points.Average(p => p.Z);
            Point centroid = new Point("centroid", centroidX, centroidY, centroidZ);

            // Step 3: Define a local 2D coordinate system on the plane
            Vector3 axisX = Vector3.Normalize(v1); // Take v1 as the X-axis of the local coordinate system
            Vector3 axisY = Vector3.Normalize(Vector3.Cross(normal, axisX)); // Y-axis is perpendicular to both normal and X-axis

            // Step 4: Project points onto the plane's local 2D coordinate system and compute angles
            List<Tuple<Point, double>> pointsWithAngles = new List<Tuple<Point, double>>();

            foreach (var point in points)
            {
                // Vector from centroid to the current point
                Vector3 vectorFromCentroid = new Vector3((float)(point.X - centroid.X), (float)(point.Y - centroid.Y), (float)(point.Z - centroid.Z));

                // Project onto the local 2D plane
                double projectedX = Vector3.Dot(vectorFromCentroid, axisX); // X coordinate in the local 2D system
                double projectedY = Vector3.Dot(vectorFromCentroid, axisY); // Y coordinate in the local 2D system

                // Compute the angle with respect to the positive X-axis of the local 2D system
                double angle = Math.Atan2(projectedY, projectedX);

                pointsWithAngles.Add(new Tuple<Point, double>(point, angle));
            }

            // Step 5: Sort the points based on the angle
            var orderedPoints = pointsWithAngles.OrderBy(tuple => tuple.Item2).Select(tuple => tuple.Item1).ToList();

            return orderedPoints;
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

        public static bool IsSameType(Frame sapElement, Frame revitElement)
        {
            if ((sapElement.MemberType == MemberType.Beam || sapElement.MemberType == MemberType.VLBracing) && (revitElement.MemberType == MemberType.Beam || revitElement.MemberType == MemberType.VLBracing))
            {
                return true;
            }

            else if ((sapElement.MemberType == MemberType.Column || sapElement.MemberType == MemberType.SlanttedColumn) && (revitElement.MemberType == MemberType.Column || revitElement.MemberType == MemberType.SlanttedColumn))
            {
                return true;
            }
            else
            {
                return false;
            }

        }



        public static List<SapRevitFrame> getMappedSapRevit(List<Frame> sapFrames, List<Frame> revitFrames, double margin)
        {
            List<SapRevitFrame> allResults = new List<SapRevitFrame>();


            List<string> sapIDS = sapFrames.Select(x => x.ID).AsParallel().ToList();
            List<string> sapIDParameter = revitFrames.Select(x => x.ID).AsParallel().ToList();


            List<Frame> sapElementsMappedByIDS = sapFrames.Where(x => sapIDParameter.Contains(x.ID)).AsParallel().OrderBy(x => x.ID).ToList();
            List<Frame> revitElementsMappedByIDS = revitFrames.Where(x => sapIDS.Contains(x.ID)).AsParallel().OrderBy(x => x.ID).ToList();



            for (int i = sapElementsMappedByIDS.Count - 1; i >= 0; i--)
            {



                if (sapElementsMappedByIDS[i].MemberType == revitElementsMappedByIDS[i].MemberType)
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

                else
                {
                    sapElementsMappedByIDS.RemoveAt(i);
                    revitElementsMappedByIDS.RemoveAt(i);
                }


            }


            List<Frame> sapElementsNotMappedByIDS = sapFrames.Except(sapElementsMappedByIDS).ToList();
            List<Frame> revitElementsNotMappedByIDS = revitFrames.Except(revitElementsMappedByIDS).ToList();


            foreach (var sapFrame in sapElementsNotMappedByIDS.AsParallel())
            {
                SapRevitFrame result = new SapRevitFrame();
                Dictionary<Frame, double> dictionary = new Dictionary<Frame, double>();
                List<double> lengthDiffernce = new List<double>();
                result.sapFrame = sapFrame;
                result.revitFrame = null;



                foreach (Frame revitFrame in revitElementsNotMappedByIDS.AsParallel())
                {
                    double lengthDifferencePersentage = Math.Abs(revitFrame.Length - sapFrame.Length) / sapFrame.Length;
                    if (lengthDifferencePersentage < margin)
                    {
                        dictionary.Add(revitFrame, Point.getDistance(sapFrame.MidPoint, revitFrame.MidPoint) + Point.getDistance(sapFrame.StartPoint, revitFrame.StartPoint));
                    }

                }


                #region getMinValueFromDictionary
                double mainValue = double.MaxValue;
                if (dictionary.Count != 0)
                {
                    mainValue = dictionary.Values.AsParallel().ToList().Min();
                }


                #endregion


                if (mainValue < margin * 2 * sapFrame.Length)
                {
                    int index = dictionary.Values.AsParallel().ToList().IndexOf(dictionary.Values.Min());
                    result.revitFrame = dictionary.Keys.AsParallel().ToList()[index];
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


            //foreach (var sapFrame in sapFrames.AsParallel())
            //{
            //    SapRevitFrame result = new SapRevitFrame();
            //    Dictionary<Frame, double> dictionary = new Dictionary<Frame, double>();
            //    List<double> lengthDiffernce = new List<double>();
            //    result.sapFrame = sapFrame;
            //    result.revitFrame = null;



            //    foreach (Frame revitFrame in revitFrames.AsParallel())
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
            //        mainValue = dictionary.Values.AsParallel().ToList().Min();
            //    }


            //    #endregion


            //    if (mainValue < margin * sapFrame.Length)
            //    {
            //        int index = dictionary.Values.AsParallel().ToList().IndexOf(dictionary.Values.Min());
            //        result.revitFrame = dictionary.Keys.AsParallel().ToList()[index];
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

        public static List<SapRevitFrame> getMappedSapRevit2(List<Frame> sapFrames, List<Frame> revitFrames, double margin)
        {
            List<SapRevitFrame> allResults = new List<SapRevitFrame>();


            List<string> sapIDS = sapFrames.Select(x => x.ID).AsParallel().ToList();
            List<string> sapIDParameter = revitFrames.Select(x => x.ID).AsParallel().ToList();


            List<Frame> sapElementsMappedByIDS = sapFrames.Where(x => sapIDParameter.Contains(x.ID)).AsParallel().OrderBy(x => x.ID).ToList();
            List<Frame> revitElementsMappedByIDS = revitFrames.Where(x => sapIDS.Contains(x.ID)).AsParallel().OrderBy(x => x.ID).ToList();



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


            foreach (var sapFrame in sapElementsNotMappedByIDS.AsParallel())
            {
                SapRevitFrame result = new SapRevitFrame();
                Dictionary<Frame, double> dictionary = new Dictionary<Frame, double>();
                List<double> lengthDiffernce = new List<double>();
                result.sapFrame = sapFrame;
                result.revitFrame = null;



                foreach (Frame revitFrame in revitElementsNotMappedByIDS.AsParallel())
                {
                    double lengthDifferencePersentage = Math.Abs(revitFrame.Length - sapFrame.Length) / sapFrame.Length;
                    if (lengthDifferencePersentage < margin)
                    {
                        dictionary.Add(revitFrame, Point.getDistance(sapFrame.MidPoint, revitFrame.MidPoint) + Point.getDistance(sapFrame.StartPoint, revitFrame.StartPoint));
                    }

                }


                #region getMinValueFromDictionary
                double mainValue = double.MaxValue;
                if (dictionary.Count != 0)
                {
                    mainValue = dictionary.Values.AsParallel().ToList().Min();
                }


                #endregion


                if (mainValue < margin * 2 * sapFrame.Length)
                {
                    int index = dictionary.Values.AsParallel().ToList().IndexOf(dictionary.Values.Min());
                    result.revitFrame = dictionary.Keys.AsParallel().ToList()[index];
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


            //foreach (var sapFrame in sapFrames.AsParallel())
            //{
            //    SapRevitFrame result = new SapRevitFrame();
            //    Dictionary<Frame, double> dictionary = new Dictionary<Frame, double>();
            //    List<double> lengthDiffernce = new List<double>();
            //    result.sapFrame = sapFrame;
            //    result.revitFrame = null;



            //    foreach (Frame revitFrame in revitFrames.AsParallel())
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
            //        mainValue = dictionary.Values.AsParallel().ToList().Min();
            //    }


            //    #endregion


            //    if (mainValue < margin * sapFrame.Length)
            //    {
            //        int index = dictionary.Values.AsParallel().ToList().IndexOf(dictionary.Values.Min());
            //        result.revitFrame = dictionary.Keys.AsParallel().ToList()[index];
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

            foreach (var sapFrame in sapFrames.AsParallel())
            {
                SapRevitFrame result = new SapRevitFrame();
                Dictionary<Frame, double> dictionary = new Dictionary<Frame, double>();
                List<double> lengthDiffernce = new List<double>();
                result.sapFrame = sapFrame;
                result.revitFrame = null;

                foreach (Frame revitFrame in revitFrames.AsParallel())
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

            unMappedSapFrames = sapRevitFrames.AsParallel().Where(x => x.revitFrame == null).ToList().Select(x => x.sapFrame).ToList();
            UnMappedRevitFrames = revitFrames.Except(sapRevitFrames.Select(x => x.revitFrame).ToList()).ToList();
        }

        public static void setMarkInRevit(List<SapRevitFrame> mappedElements, Document doc)
        {
            foreach (var item in mappedElements)
            {
                if (item.revitFrame != null)
                {
                    FamilyInstance instance = doc.GetElement(item.revitFrame.revitEleId) as FamilyInstance;
                    Parameter mark = instance.LookupParameter("Mark");


                    using (Transaction trns = new Transaction(doc, "aa"))
                    {

                        trns.Start();
                        FailureHandlingOptions failOpt = trns.GetFailureHandlingOptions();
                        failOpt.SetFailuresPreprocessor(new RoomWarningSwallower());
                        trns.SetFailureHandlingOptions(failOpt);

                        mark.Set(item.sapFrame.ID);
                        trns.Commit();

                    }
                }


            }


        }
        public static View3D createView(List<ElementId> ids, ExternalCommandData commandData, string groupName)
        {
            Document doc = commandData.Application.ActiveUIDocument.Document;

            using (Transaction trns = new Transaction(doc, "aa"))
            {

                trns.Start();
                ViewFamilyType viewType = new FilteredElementCollector(doc).OfClass(typeof(ViewFamilyType)).Cast<ViewFamilyType>().First(x => x.ViewFamily == ViewFamily.ThreeDimensional);
                View3D unMappedView = new FilteredElementCollector(doc).OfClass(typeof(View3D)).Cast<View3D>().FirstOrDefault(x => x.Name == groupName);
                if (unMappedView == null)
                {
                    unMappedView = View3D.CreateIsometric(doc, viewType.Id);
                    unMappedView.Name = groupName;
                }

                TemporaryViewMode tempView = TemporaryViewMode.TemporaryHideIsolate;
                unMappedView.DisableTemporaryViewMode(tempView);

                UIDocument uiDoc = commandData.Application.ActiveUIDocument;
                unMappedView.IsolateElementsTemporary(ids);

                trns.Commit();
                //commandData.Application.ActiveUIDocument.ActiveView = unMappedView;
                //commandData.Application.ActiveUIDocument.RequestViewChange(unMappedView);
                return unMappedView;
            }



        }



        public static View3D createView(List<ElementId> ids, Document doc, UIDocument uiDoc, string groupName)
        {


            using (Transaction trns = new Transaction(doc, "aa"))
            {

                trns.Start();
                ViewFamilyType viewType = new FilteredElementCollector(doc).OfClass(typeof(ViewFamilyType)).Cast<ViewFamilyType>().First(x => x.ViewFamily == ViewFamily.ThreeDimensional);
                View3D unMappedView = new FilteredElementCollector(doc).OfClass(typeof(View3D)).Cast<View3D>().FirstOrDefault(x => x.Name == groupName);
                if (unMappedView == null)
                {
                    unMappedView = View3D.CreateIsometric(doc, viewType.Id);
                    unMappedView.Name = groupName;
                }

                TemporaryViewMode tempView = TemporaryViewMode.TemporaryHideIsolate;
                unMappedView.DisableTemporaryViewMode(tempView);


                unMappedView.IsolateElementsTemporary(ids);

                trns.Commit();
                //commandData.Application.ActiveUIDocument.ActiveView = unMappedView;
                //commandData.Application.ActiveUIDocument.RequestViewChange(unMappedView);
                return unMappedView;
            }



        }


    }



    public class Area
    {
        #region  props
        public string ID { get; set; }

        public List<Point> vertices { get; set; }

        public List<string> Groups { get; set; }

        public string Group { get; set; } // Not Used
        public string MainGroup { get; set; } // Not Used
        public string SubGroup { get; set; } // Not Used

        public Guid guid { get; set; }

        public string RvtID { get; set; }

        public double Beta { get; set; }

        public ElementId revitEleId { get; set; }

        public string Sec { get; set; }

        public string RvtName { get; set; }

        public string aa { get; set; }
        public int rowNo { get; set; }

        public int NoOfPts { get; set; }

        public string type { get; set; }

        public List<Rhino.Geometry.Curve> crvList { get; set; }

        public CurveArrArray crvArray { get; set; }

        public bool IsCounterClockWise { get; set; }
        #endregion

        public Area()
        {
            //vertices = new List<Point>();
            //crvArray = new CurveArrArray();
           // crvList = new List<Rhino.Geometry.Curve>();
        }

        public static bool IsAreaVertical(Area area, double tolerance)
        {
            // Check if the area has at least 2 points
            if (area.vertices.Count < 2)
            {
                return false; // Not enough points to determine verticality
            }

            for (int i = 0; i < area.vertices.Count - 1; i++)
            {
                for (int j = i + 1; j < area.vertices.Count; j++)
                {
                    Point p1 = area.vertices[i];
                    Point p2 = area.vertices[j];

                    // Check if the X and Y coordinates are within the tolerance range
                    if (Math.Abs(p1.X - p2.X) <= tolerance && Math.Abs(p1.Y - p2.Y) <= tolerance)
                    {
                        return true; // Area is vertical
                    }
                }
            }

            if (area.vertices.Count == 3)
            {
                for (int i = 0; i < area.vertices.Count - 1; i++)
                {
                    for (int j = i + 1; j < area.vertices.Count; j++)
                    {
                        Point p1 = area.vertices[i];
                        Point p2 = area.vertices[j];

                        // Check if the difference in Z values is more than the tolerance
                        if (Math.Abs(p1.Z - p2.Z) > 1200)
                        {
                            return true; // Area is vertical
                        }
                    }
                }

            }


            return false; // Area is not vertical
        }

        // Method to calculate if vertices are counterclockwise around a center point
        public static bool AreVerticesCounterClockwise(List<Point3d> vertices, Point3d center, Vector3d normal)
        {
            double sum = 0;
            for (int i = 0; i < vertices.Count; i++)
            {
                Point3d current = vertices[i];
                Point3d next = vertices[(i + 1) % vertices.Count];

                // Calculate vectors from center to current and next vertices
                Vector3d v1 = current - center;
                Vector3d v2 = next - center;

                // Cross product to determine the orientation
                double crossProduct = Vector3d.CrossProduct(v1, v2) * normal;
                sum += crossProduct;
            }

            // Positive sum indicates counterclockwise order, negative indicates clockwise
            return sum > 0;
        }

        // Method to get vertices in counterclockwise order
        public static List<Point3d> GetCounterClockwiseVertices(BrepFace face)
        {
            // Calculate the centroid of the face
            AreaMassProperties areaProps = AreaMassProperties.Compute(face);
            Point3d center = areaProps.Centroid;

            // Get the normal of the face to use as reference for orientation
            Rhino.Geometry.Surface surface = face.UnderlyingSurface();
            Vector3d normal = surface.NormalAt(surface.Domain(0).Mid, surface.Domain(1).Mid);

            // Retrieve unique vertices of the face's Brep
            List<Point3d> vertices = face.Brep.Vertices.Select(v => v.Location).Distinct().ToList();

            // Project the points onto the plane of the face (defined by the center and normal)
            Plane facePlane = new Plane(center, normal);
            List<Point2d> projectedPoints = new List<Point2d>();
            foreach (var vertex in vertices)
            {
                double u, v;
                facePlane.ClosestParameter(vertex, out u, out v);
                projectedPoints.Add(new Point2d(u, v));
            }

            // Sort the points in counterclockwise order around the centroid
            List<Point3d> sortedVertices = vertices
                .Select((point, index) => new { Point = point, Angle = Math.Atan2(projectedPoints[index].Y, projectedPoints[index].X) })
                .OrderBy(p => p.Angle)
                .Select(p => p.Point)
                .ToList();

            // Define the length of the X and Y axis lines for visualization
            double axisLength = 5.0; // Adjust length as needed

            // Create lines along the X and Y axes of the plane
           Rhino.Geometry. Line xAxisLine = new Rhino.Geometry.Line(center, center + facePlane.XAxis * axisLength);
            Rhino.Geometry.Line yAxisLine = new Rhino.Geometry.Line(center, center + facePlane.YAxis * axisLength);


            Rhino.RhinoDoc doc = RhinoDoc.ActiveDoc;
            Rhino.DocObjects.Tables.ObjectTable ot = Rhino.RhinoDoc.ActiveDoc.Objects;

            //// Add the X and Y axis lines to the document
            //ot.AddLine(xAxisLine);
            //ot.AddLine(yAxisLine);

            //// Optionally, add a point at the center for clarity
            //ot.AddPoint(center);

            return sortedVertices;
        }
    }


    public class Area1
    {
        #region  props
        public string ID { get; set; }

        public List<Point> vertices { get; set; }

        public List<string> Groups { get; set; }

        public string Group { get; set; } // Not Used

        public string RvtID { get; set; }

        public double Beta { get; set; }

        public ElementId revitEleId { get; set; }

        public string Sec { get; set; }

        public string RvtName { get; set; }

        public string aa { get; set; }
        public int rowNo { get; set; }

        public int NoOfPts { get; set; }

        public string type { get; set; }

        public List<Rhino.Geometry.Curve> crvList { get; set; }

        public CurveArrArray crvArray { get; set; }
        #endregion

        public Area1()
        {
            vertices = new List<Point>();
            crvArray = new CurveArrArray();
            crvList = new List<Rhino.Geometry.Curve>();
        }

    }


    public class ColumnRebarDesignData
    {
        public string section { get; set; }
        public string MatPropLong { get; set; }
        public string MatPropConfine { get; set; }
        public string Pattern { get; set; }
        public string ConfineType { get; set; }
        public double Cover { get; set; }
        public int NumberCBars { get; set; }
        public int NumberOfWidthBars { get; set; }
        public int NumberOfDepthBars { get; set; }
        public string RebarSize { get; set; }
        public string TieSize { get; set; }
        public double TieSpacingLongit { get; set; }
        public int Number2DirTieBars { get; set; }
        public int Number3DirTieBars { get; set; }
        public bool ToBeDesigned { get; set; }


        public static void RetrieveDataFromETABS(ref List<Frame> etabsFrames)
        {
            cSapModel myEtabsModel = EtabsUtils.Get_sap_mode();

            List<string> sections = etabsFrames.Where(x => x.MemberType == MemberType.Column).Select(x => x.Sec).ToList();
            sections = sections.Distinct().ToList();


            #region definitions
            int ret = -1;
            string RebarName = "RebarName";
            string MatPropLong = "MatPropLong";
            string MatPropConfine = "MatPropConfine";
            int Pattern = 2;
            int ConfineType = 2;
            double Cover = 2.0;
            int NumberCBars = 10;
            int NumberR3Bars = 0;
            int NumberR2Bars = 0;
            string RebarSize = "#10";
            string TieSize = "#5";
            double TieSpacingLongit = 4.0;
            int Number2DirTieBars = 0;
            int Number3DirTieBars = 0;
            bool ToBeDesigned = false;

            eFramePropType secType = eFramePropType.Rectangular;
            #endregion


            foreach (var sec in sections)
            {
                ColumnRebarDesignData designData = new ColumnRebarDesignData();
                ret = myEtabsModel.PropFrame.GetRebarColumn(sec, ref MatPropLong, ref MatPropConfine, ref Pattern, ref ConfineType, ref Cover, ref NumberCBars, ref NumberR3Bars, ref NumberR2Bars, ref RebarSize, ref TieSize, ref TieSpacingLongit, ref Number2DirTieBars, ref Number3DirTieBars, ref ToBeDesigned);
                ret = myEtabsModel.PropFrame.GetTypeOAPI(sec, ref secType);
                #region Get ColumnsType
                if (Pattern == 1 && secType == eFramePropType.Rectangular)
                {
                    //Rectangle
                    designData.Pattern = "R";

                }

                else if (Pattern == 2 && secType == eFramePropType.Circle)
                {
                    //Circle
                    designData.Pattern = "C";
                }

                //else
                //{
                //    continue;
                //}

                #endregion

                #region get COnfinType
                if (ConfineType == 1)
                {
                    designData.ConfineType = "Ties";
                }
                else
                {
                    designData.ConfineType = "Spiral";

                }
                #endregion

                designData.section = sec;

                designData.Cover = Cover;

                designData.NumberCBars = NumberCBars;

                designData.NumberOfWidthBars = NumberR3Bars;

                designData.NumberOfDepthBars = NumberR2Bars;

                designData.RebarSize = RebarSize;

                designData.TieSize = TieSize;

                designData.ToBeDesigned  = ToBeDesigned;

                foreach (var item in etabsFrames.Where(x => x.Sec == sec))
                {
                    item.ColumnsRebarDesign = designData;
                }
            }


        }

        public static string getDesignationMark (Frame item)
        {
            string result = null;
            if (item.ColumnsRebarDesign.Pattern =="R")
            {
                int noOfBars = (item.ColumnsRebarDesign.NumberOfDepthBars * 2 + item.ColumnsRebarDesign.NumberOfWidthBars * 2) - 4;
                result = item.Width+"X"+item.Depth+"-"+ noOfBars+"T"+ item.ColumnsRebarDesign.RebarSize;


            }

            else if (item.ColumnsRebarDesign.Pattern == "C")
            {
                result = "D"+item.Diameter+"-"+ item.ColumnsRebarDesign.NumberCBars+"T"+ item.ColumnsRebarDesign.RebarSize;

            }

            else
            {
                result = "N.A";
            }


            return result;
        }
    }


    public class Record
    {
        public string Mark { get; set; }
        public string SubMark { get; set; }
        public string RvtBotLvl { get; set; }
        public string RvtTopLvl { get; set; }
        public string EtabsBotStory { get; set; }
        public string EtabsTopStory { get; set; }
        public string ColumnType { get; set; }
        public string Designation { get; set; }
        public int Count { get; set; }
        public string Description { get; set; }
        public string Depth { get; set; }
        public string Width { get; set; }
        public string Diameter { get; set; }

        public int numberOfBarsInDepthDir  { get; set; }
        public int numberOfBarsInWidthDir { get; set; }
        public int BarDiameter { get; set; }

        public int CircularBars { get; set; }



    }


    #region DownGrade

    public class RevitBeam : Frame
    {


        // Additional properties specific to RevitBeam
        public string ReferenceLevel { get; set; }
        public string EndLevel { get; set; }
        public double StartLevelOffset { get; set; }
        public double EndLevelOffset { get; set; }
        public double StartExtension { get; set; }
        public double EndExtension { get; set; }
        public ZJustification ZJustification { get; set; }
        public YJustification YJustification { get; set; }
        public double ZOffsetValue { get; set; }
        public double YOffsetValue { get; set; }

        //public Curve crv { get; set; }

        //public Arc arc { get; set; }    

        // public Line line { get; set; }

        // public Level refLevelObject { get; set; }

        public List<Point> ArcVerticies { get; set; }

        public CurveType CrvType { get; set; }

        public static void getRevitBeams(Document doc, UIDocument uidoc, ref List<RevitBeam> Beams, ForgeTypeId dis)
        {

            using (Transaction allTrns = new Transaction(doc, "aaa"))
            {
                allTrns.Start();
                FailureHandlingOptions failOpt = allTrns.GetFailureHandlingOptions();
                failOpt.SetFailuresPreprocessor(new RoomWarningSwallower());
                allTrns.SetFailureHandlingOptions(failOpt);

                #region get Selected Elements 
                ICollection<ElementId> selectedIds = uidoc.Selection.GetElementIds();

                // Create a filter for Structural Framing and Structural Beams categories
                List<BuiltInCategory> categories = new List<BuiltInCategory>();
                categories.Add(BuiltInCategory.OST_StructuralFraming);
                categories.Add(BuiltInCategory.OST_StructuralColumns);


                // Create an ElementCategoryFilter for each category and combine them with a LogicalOrFilter
                List<ElementFilter> categoryFilters = new List<ElementFilter>();
                foreach (BuiltInCategory category in categories)
                {
                    categoryFilters.Add(new ElementCategoryFilter(category));
                }
                //
                LogicalOrFilter orFilter = new LogicalOrFilter(categoryFilters);

                // Get the selected elements that match the filter
                List<Element> selectedElements = new List<Element>();
                foreach (ElementId id in selectedIds)
                {
                    Element element = uidoc.Document.GetElement(id);


                    FilteredElementCollector tagCollector = new FilteredElementCollector(doc)
    .OfCategory(BuiltInCategory.OST_Tags);



                    if (orFilter.PassesFilter(element))
                    {
                        selectedElements.Add(element);
                    }
                }


                List<FamilyInstance> selectedStrucralFraming = selectedElements.Where(x => x.Category.Id.IntegerValue == (int)BuiltInCategory.OST_StructuralFraming).Cast<FamilyInstance>().ToList();


                #endregion



                List<Frame> beamFarames = new List<Frame>();
                List<Frame> columnFarames = new List<Frame>();
                List<Frame> frames = new List<Frame>();

                List<FamilyInstance> Allbeams = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralFraming).WhereElementIsNotElementType().Cast<FamilyInstance>().ToList();

                if (selectedStrucralFraming.Count > 0)
                {
                    Allbeams = selectedStrucralFraming;
                }

                #region get data of beams

                #region handling curvedBeams (not implemented yet)
                List<FamilyInstance> strightBeams = new List<FamilyInstance>();
                List<FamilyInstance> curvedBeams = new List<FamilyInstance>();
                List<FamilyInstance> StrightElementsSplittedFromCurved = new List<FamilyInstance>();

                foreach (var instance in Allbeams)
                {
                    AnalyticalToPhysicalAssociationManager analyManger = AnalyticalToPhysicalAssociationManager.GetAnalyticalToPhysicalAssociationManager(doc);
                    AnalyticalElement analytInstance = doc.GetElement(analyManger.GetAssociatedElementId(instance.Id)) as AnalyticalElement;

                    if (analytInstance.GetCurve() != null)
                    {
                        Point startPoint = new Point(analytInstance.GetCurve().GetEndPoint(0));
                        Point endPoint = new Point(analytInstance.GetCurve().GetEndPoint(1));

                        Vector3 vec1 = new Vector3((float)(endPoint.X - startPoint.X), (float)(endPoint.Y - startPoint.Y), (float)(endPoint.Z - startPoint.Z));
                        vec1 = Vector3.Normalize(vec1);

                        LocationCurve locCurve = instance.Location as LocationCurve;
                        Curve curve = locCurve.Curve;
                        double totalLength = curve.Length;
                        Point midPoint = new Point(curve.Evaluate(totalLength / 2, false));


                        Vector3 vec2 = new Vector3((float)(midPoint.X - startPoint.X), (float)(midPoint.Y - startPoint.Y), (float)(midPoint.Z - startPoint.Z));
                        vec2 = Vector3.Normalize(vec2);


                        if (vec1.X == vec2.X && vec1.Y == vec2.Y && vec1.Z == vec2.Z)
                        {
                            strightBeams.Add(instance);

                        }

                        else
                        {
                            if (UnitUtils.ConvertFromInternalUnits((instance.Location as LocationCurve).Curve.Length, dis) > 6000)
                            {
                                curvedBeams.Add(instance);
                            }
                            else
                            {
                                strightBeams.Add(instance);
                            }

                        }

                    }



                }

                // get stright and curved beams


                // split curved beams to straight beams 
                foreach (var instance in curvedBeams)
                {


                    //SapRevitFrame.createView(curvedBeams.Select(x => x.Id).ToList(), doc, uidoc, "curvedBeams");

                    //return null;




                    StrightElementsSplittedFromCurved.Add(instance);


                    // uncomment from here
                    //LocationCurve locCurve = instance.Location as LocationCurve;
                    //Curve curve = locCurve.Curve;
                    //double totalLength = curve.Length;

                    //double segmentLength = UnitUtils.ConvertToInternalUnits(5000.0, dis);

                    //using (Transaction trns = new Transaction(doc, "aa"))

                    //{
                    //    trns.Start();

                    //    while ((instance.Location as LocationCurve).Curve.Length > segmentLength)
                    //    {
                    //        totalLength = (instance.Location as LocationCurve).Curve.Length;
                    //        StrightElementsSplittedFromCurved.Add(doc.GetElement(instance.Split(1 - segmentLength / totalLength)) as FamilyInstance);

                    //    }

                    //aa:
                    //    trns.Commit();

                    //}

                }



                // union between  straightBeams and StraightBeamsFromSplittedCurvedBeams
                strightBeams = strightBeams.Union(StrightElementsSplittedFromCurved).ToList();
                #endregion

                #region get Data Of Beams after handling curvedBeams

                foreach (var instance in strightBeams.AsParallel())
                {
                    try
                    {
                        RevitBeam frame = new RevitBeam();
                        AnalyticalToPhysicalAssociationManager analyManger = AnalyticalToPhysicalAssociationManager.GetAnalyticalToPhysicalAssociationManager(doc);
                        AnalyticalElement analytInstance = doc.GetElement(analyManger.GetAssociatedElementId(instance.Id)) as AnalyticalElement;

                        if (analytInstance.GetCurve() != null)
                        {
                            Point startPoint = new Point(analytInstance.GetCurve().GetEndPoint(0));
                            Point endPoint = new Point(analytInstance.GetCurve().GetEndPoint(1));

                            //startPoint.X = UnitUtils.ConvertFromInternalUnits(startPoint.X, dis);
                            //startPoint.Y = UnitUtils.ConvertFromInternalUnits(startPoint.Y, dis);
                            //startPoint.Z = UnitUtils.ConvertFromInternalUnits(startPoint.Z, dis);

                            //endPoint.X = UnitUtils.ConvertFromInternalUnits(endPoint.X, dis);
                            //endPoint.Y = UnitUtils.ConvertFromInternalUnits(endPoint.Y, dis);
                            //endPoint.Z = UnitUtils.ConvertFromInternalUnits(endPoint.Z, dis);


                            Point midPoint = Point.getMidPoint(startPoint, endPoint);
                            frame.RvtID = instance.Id.ToString();
                            frame.revitEleId = instance.Id;
                            frame.StartPoint = startPoint;
                            frame.EndPoint = endPoint;
                            frame.MidPoint = midPoint;
                            frame.Length = Point.getDistance(startPoint, endPoint);
                            frame.Material = instance.LookupParameter("Material").AsString();



                            frame.Beta = instance.LookupParameter("Cross-Section Rotation").AsDouble() * 180 / (Math.PI);
                            frame.RvtFamily = instance.Symbol.FamilyName;
                            frame.RvtName = instance.Symbol.Name;

                            #region getLevelOfssetsParamaters
                            frame.ReferenceLevel = (doc.GetElement(instance.get_Parameter(BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM).AsElementId()) as Level).Name;
                            frame.StartLevelOffset = instance.LookupParameter("Start Level Offset").AsDouble();
                            frame.EndLevelOffset = instance.LookupParameter("End Level Offset").AsDouble();

                            frame.StartExtension = instance.LookupParameter("Start Extension").AsDouble();
                            frame.EndExtension = instance.LookupParameter("End Extension").AsDouble();

                            frame.ZJustification = (ZJustification)instance.get_Parameter(BuiltInParameter.Z_JUSTIFICATION).AsInteger();
                            frame.YJustification = (YJustification)instance.get_Parameter(BuiltInParameter.Y_JUSTIFICATION).AsInteger();


                            frame.ZOffsetValue = instance.get_Parameter(BuiltInParameter.Z_OFFSET_VALUE).AsDouble();
                            frame.YOffsetValue = instance.get_Parameter(BuiltInParameter.Y_OFFSET_VALUE).AsDouble();

                            #endregion



                            Beams.Add(frame);
                        }

                        else
                        {
                            continue;
                        }
                    }


                    catch (Exception ex2)
                    {

                        System.Diagnostics.StackTrace trace = new System.Diagnostics.StackTrace(ex2, true);

                        RevitUtils.Log.Add("Error in getRevitFrames Method in the Beams part... BeamInstance no" + instance.Id + ex2.Message + trace.GetFrame(trace.FrameCount - 1).GetFileLineNumber());

                        // TaskDialog.Show("ErrorInGetRevitFramesMethod", "ErrorInRevitBeamID= "+ instance.Id + ex2.Message + trace.GetFrame(trace.FrameCount - 1).GetFileLineNumber());
                        continue;
                    }

                }
                #endregion

                #endregion

                allTrns.Commit();

            }




        }

        public static void getRevitBeams2(Document doc, UIDocument uidoc, ref List<RevitBeam> Beams, ForgeTypeId dis)
        {

            using (Transaction allTrns = new Transaction(doc, "aaa"))
            {
                allTrns.Start();
                FailureHandlingOptions failOpt = allTrns.GetFailureHandlingOptions();
                failOpt.SetFailuresPreprocessor(new RoomWarningSwallower());
                allTrns.SetFailureHandlingOptions(failOpt);

                #region get Selected Elements 
                ICollection<ElementId> selectedIds = uidoc.Selection.GetElementIds();

                // Create a filter for Structural Framing and Structural Beams categories
                List<BuiltInCategory> categories = new List<BuiltInCategory>();
                categories.Add(BuiltInCategory.OST_StructuralFraming);
                categories.Add(BuiltInCategory.OST_StructuralColumns);


                // Create an ElementCategoryFilter for each category and combine them with a LogicalOrFilter
                List<ElementFilter> categoryFilters = new List<ElementFilter>();
                foreach (BuiltInCategory category in categories)
                {
                    categoryFilters.Add(new ElementCategoryFilter(category));
                }
                //
                LogicalOrFilter orFilter = new LogicalOrFilter(categoryFilters);

                // Get the selected elements that match the filter
                List<Element> selectedElements = new List<Element>();
                foreach (ElementId id in selectedIds)
                {
                    Element element = uidoc.Document.GetElement(id);


                    FilteredElementCollector tagCollector = new FilteredElementCollector(doc)
    .OfCategory(BuiltInCategory.OST_Tags);



                    if (orFilter.PassesFilter(element))
                    {
                        selectedElements.Add(element);
                    }
                }


                List<FamilyInstance> selectedStrucralFraming = selectedElements.Where(x => x.Category.Id.IntegerValue == (int)BuiltInCategory.OST_StructuralFraming).Cast<FamilyInstance>().ToList();


                #endregion



                List<Frame> beamFarames = new List<Frame>();
                List<Frame> columnFarames = new List<Frame>();
                List<Frame> frames = new List<Frame>();

                List<FamilyInstance> Allbeams = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralFraming).WhereElementIsNotElementType().Cast<FamilyInstance>().ToList();

                if (selectedStrucralFraming.Count > 0)
                {
                    Allbeams = selectedStrucralFraming;
                }


                foreach (var instance in Allbeams.AsParallel())
                {
                    try
                    {
                        AnalyticalToPhysicalAssociationManager analyManger = AnalyticalToPhysicalAssociationManager.GetAnalyticalToPhysicalAssociationManager(doc);
                        AnalyticalElement analytInstance = doc.GetElement(analyManger.GetAssociatedElementId(instance.Id)) as AnalyticalElement;

                        RevitBeam frame = new RevitBeam();

                        if (analytInstance.GetCurve() != null)
                        {

                            Point startPoint = new Point(analytInstance.GetCurve().GetEndPoint(0));
                            Point endPoint = new Point(analytInstance.GetCurve().GetEndPoint(1));


                            Point midPoint = Point.getMidPoint(startPoint, endPoint);
                            frame.RvtID = instance.Id.ToString();
                            frame.revitEleId = instance.Id;
                            frame.StartPoint = startPoint;
                            frame.EndPoint = endPoint;
                            frame.MidPoint = midPoint;
                            frame.Length = Point.getDistance(startPoint, endPoint);
                            frame.Material = instance.LookupParameter("Material").AsString();


                            #region Sections
                            // sections 
                            frame.Beta = instance.LookupParameter("Cross-Section Rotation").AsDouble() * 180 / (Math.PI);
                            frame.RvtFamily = instance.Symbol.FamilyName;
                            frame.RvtName = instance.Symbol.Name;
                            #endregion

                            #region getLevelOfssetsParamaters
                            // frame.refLevelObject = (doc.GetElement(instance.get_Parameter(BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM).AsElementId()) as Level);
                            frame.ReferenceLevel = (doc.GetElement(instance.get_Parameter(BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM).AsElementId()) as Level).Name;
                            frame.StartLevelOffset = instance.LookupParameter("Start Level Offset").AsDouble();
                            frame.EndLevelOffset = instance.LookupParameter("End Level Offset").AsDouble();

                            frame.StartExtension = instance.LookupParameter("Start Extension").AsDouble();
                            frame.EndExtension = instance.LookupParameter("End Extension").AsDouble();

                            frame.ZJustification = (ZJustification)instance.get_Parameter(BuiltInParameter.Z_JUSTIFICATION).AsInteger();
                            frame.YJustification = (YJustification)instance.get_Parameter(BuiltInParameter.Y_JUSTIFICATION).AsInteger();


                            frame.ZOffsetValue = instance.get_Parameter(BuiltInParameter.Z_OFFSET_VALUE).AsDouble();
                            frame.YOffsetValue = instance.get_Parameter(BuiltInParameter.Y_OFFSET_VALUE).AsDouble();



                            Autodesk.Revit.DB.Curve crv = (instance.Location as LocationCurve).Curve;

                            if (crv is Line)
                            {
                                frame.CrvType = CurveType.Line;
                            }

                            else if (crv is Arc)
                            {
                                frame.CrvType = CurveType.Arc;

                                XYZ stPoint1 = crv.Evaluate(0, true);
                                XYZ endPoint2 = crv.Evaluate(1, true);
                                XYZ MidPoint = crv.Evaluate(0.5, true);

                                frame.ArcVerticies.Add(new Point(stPoint1));
                                frame.ArcVerticies.Add(new Point(MidPoint));
                                frame.ArcVerticies.Add(new Point(endPoint2));
                            }

                            else if (crv is Ellipse)
                            {
                                frame.CrvType = CurveType.Ellipse;

                            }

                            else if (crv is HermiteSpline)
                            {
                                frame.CrvType = CurveType.HermiteSpline;
                            }

                            else
                            {
                                frame.CrvType = CurveType.Other;

                            }


                            // frame.crv = (instance.Location as LocationCurve).Curve;





                            #endregion



                            Beams.Add(frame);
                        }

                        else
                        {
                            continue;
                        }
                    }


                    catch (Exception ex2)
                    {

                        System.Diagnostics.StackTrace trace = new System.Diagnostics.StackTrace(ex2, true);

                        RevitUtils.Log.Add("Error in getRevitFrames Method in the Beams part... BeamInstance no" + instance.Id + ex2.Message + trace.GetFrame(trace.FrameCount - 1).GetFileLineNumber());

                        // TaskDialog.Show("ErrorInGetRevitFramesMethod", "ErrorInRevitBeamID= "+ instance.Id + ex2.Message + trace.GetFrame(trace.FrameCount - 1).GetFileLineNumber());
                        continue;
                    }

                }




                allTrns.Commit();

            }




        }

        public static void setRevitBeams(Document doc, UIDocument uidoc, ref List<RevitBeam> Beams, ForgeTypeId dis)
        {

            RevitServices.Transactions.TransactionManager.Instance.ForceCloseTransaction();

            using (TransactionGroup tg = new TransactionGroup(doc, "aaa"))
            {

                tg.Start();

                var chopSize = Frame.chopSize;
                var beamChunks = Beams.Chunk(chopSize).ToList();

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
                            Curve crv = null;


                            if (bm.CrvType == CurveType.Line)
                            {
                                var StartPoint = new XYZ(bm.StartPoint.X, bm.StartPoint.Y, bm.StartPoint.Z);

                                var EndPoint = new XYZ(bm.EndPoint.X, bm.EndPoint.Y, bm.EndPoint.Z);
                                crv = Line.CreateBound(StartPoint, EndPoint);
                            }

                            else
                            {
                                XYZ stPoint = new XYZ(bm.ArcVerticies[0].X, bm.ArcVerticies[0].Y, bm.ArcVerticies[0].Z);
                                XYZ midPoint = new XYZ(bm.ArcVerticies[1].X, bm.ArcVerticies[1].Y, bm.ArcVerticies[1].Z);
                                XYZ endPoint = new XYZ(bm.ArcVerticies[2].X, bm.ArcVerticies[2].Y, bm.ArcVerticies[2].Z);

                                crv = Arc.Create(stPoint, endPoint, midPoint);
                            }


                            try
                            {
                                double angle = (bm.Beta) * Math.PI / 180;

                                FamilySymbol revitSection = (FamilySymbol)new FilteredElementCollector(doc).WhereElementIsElementType().OfCategory(BuiltInCategory.OST_StructuralFraming).FirstOrDefault(x => x.Name == bm.RvtName);

                                if (revitSection == null)
                                {
                                    revitSection = (FamilySymbol)new FilteredElementCollector(doc).WhereElementIsElementType().OfCategory(BuiltInCategory.OST_StructuralFraming).FirstOrDefault();


                                }


                                Level lvl0 = (Level)new FilteredElementCollector(doc).OfClass(typeof(Level)).FirstOrDefault(x => x.Name == bm.ReferenceLevel);

                                if (!revitSection.IsActive)
                                    revitSection.Activate();

                                FamilyInstance instance = doc.Create.NewFamilyInstance(crv, revitSection, lvl0, StructuralType.Beam);

                                instance.LookupParameter("Mark").Set(bm.RvtName);

                                instance.LookupParameter("Cross-Section Rotation").Set(bm.Beta);
                                instance.LookupParameter("z Justification").Set((int)bm.ZJustification);
                                instance.LookupParameter("y Justification").Set((int)bm.YJustification);
                                instance.LookupParameter("Start Level Offset").Set((double)bm.StartLevelOffset);
                                instance.LookupParameter("End Level Offset").Set((double)bm.EndLevelOffset);
                                instance.LookupParameter("Start Extension").Set((double)bm.StartExtension);
                                instance.LookupParameter("End Extension").Set((double)bm.EndExtension);


                                instance.get_Parameter(BuiltInParameter.Z_OFFSET_VALUE).Set((double)bm.ZOffsetValue);
                                instance.get_Parameter(BuiltInParameter.Y_OFFSET_VALUE).Set((double)bm.YOffsetValue);

                                Autodesk.Revit.DB.Structure.StructuralFramingUtils.DisallowJoinAtEnd(instance, 0);
                                Autodesk.Revit.DB.Structure.StructuralFramingUtils.DisallowJoinAtEnd(instance, 1);

                            }


                            catch (Exception ex2)

                            {
                                System.Diagnostics.StackTrace trace = new System.Diagnostics.StackTrace(ex2, true);
                                RevitUtils.Log.Add("Error in CreateAllBeams3 Method... beam no" + bm.ID + ex2.Message + trace.GetFrame(trace.FrameCount - 1).GetFileLineNumber());

                                //TaskDialog.Show("Error", "error in beam no" + bm.ID + ex2.Message + trace.GetFrame(trace.FrameCount - 1).GetFileLineNumber());

                                continue;

                            }


                        }

                        trns.Commit();


                    }

                }



                tg.Assimilate();


            }



        }

        public RevitBeam()
        {
            ArcVerticies = new List<Point>();
        }

        #region Serializeation

        // generic one
        public static List<T> DeserializeFromJson<T>(string filePath)
        {
            List<T> items = new List<T>();
            if (File.Exists(filePath))
            {
                string json = File.ReadAllText(filePath);
                items = JsonConvert.DeserializeObject<List<T>>(json);
            }
            return items;
        }

        public static List<RevitBeam> DeserializeRevitBeamsFromJson(string filePath)
        {
            List<RevitBeam> beams = new List<RevitBeam>();
            if (File.Exists(filePath))
            {
                string json = File.ReadAllText(filePath);
                beams = JsonConvert.DeserializeObject<List<RevitBeam>>(json);
            }
            return beams;
        }

        public static List<RevitLevel> DeserializeRevitLevelsFromJson(string filePath)
        {
            List<RevitLevel> Levels = new List<RevitLevel>();
            if (File.Exists(filePath))
            {
                string json = File.ReadAllText(filePath);
                Levels = JsonConvert.DeserializeObject<List<RevitLevel>>(json);
            }
            return Levels;
        }



        //short helper class to ignore some properties from serialization
        public class IgnorePropertiesResolver : DefaultContractResolver
        {
            private readonly HashSet<string> ignoreProps;
            public IgnorePropertiesResolver(IEnumerable<string> propNamesToIgnore)
            {
                this.ignoreProps = new HashSet<string>(propNamesToIgnore);
            }

            protected override JsonProperty CreateProperty(MemberInfo member, MemberSerialization memberSerialization)
            {
                JsonProperty property = base.CreateProperty(member, memberSerialization);
                if (this.ignoreProps.Contains(property.PropertyName))
                {
                    property.ShouldSerialize = _ => false;
                }
                return property;
            }
        }
        public static void SerializeRevitBeamsToJsonWithSkip(List<RevitBeam> beams, string filePath)
        {

            string json = JsonConvert.SerializeObject(beams, new JsonSerializerSettings()
            { ContractResolver = new IgnorePropertiesResolver(new[] { "" }) });

            File.WriteAllText(filePath, json);
        }
        public static void SerializeLevelsToJsonWithSkip(List<RevitLevel> levels, string filePath)
        {

            string json = JsonConvert.SerializeObject(levels, new JsonSerializerSettings()
            { ContractResolver = new IgnorePropertiesResolver(new[] { "" }) });


            // string json = JsonConvert.SerializeObject(beams, Newtonsoft.Json.Formatting.Indented);

            File.WriteAllText(filePath, json);
        }


        // short helper Calss to include only a certain propetries for srialization
        public class IgnoreParentPropertiesResolver : DefaultContractResolver
        {
            private readonly List<string> _includedProperties;

            public IgnoreParentPropertiesResolver(List<string> includedProperties)
            {
                _includedProperties = includedProperties;
            }

            protected override IList<JsonProperty> CreateProperties(Type type, MemberSerialization memberSerialization)
            {
                var allProps = base.CreateProperties(type, memberSerialization);

                // Filter properties based on the includedProperties list
                return allProps.Where(p => _includedProperties.Contains(p.PropertyName)).ToList();
            }
        }
        public static void SerializeRevitBeamsToJsonWithSpecificProps(List<RevitBeam> beams, string filePath)
        {

            // Example: Specify the properties you want to include
            var includedProperties = new List<string> { "PropertyName1", "PropertyName2", "PropertyName3" };

            // Create an instance of the resolver with the included properties
            var resolver = new IgnoreParentPropertiesResolver(includedProperties);

            var settings = new JsonSerializerSettings
            {
                ContractResolver = resolver,
            };


            // Serialize your object using the settings
            var json = JsonConvert.SerializeObject(beams, settings);

            File.WriteAllText(filePath, json);
        }
        #endregion

    }

    public class RevitColumn : Frame
    {

        // Additional properties specific to RevitBeam
        public string BaseLevel { get; set; }
        public string TopLevel { get; set; }
        public double BaseLevelOffset { get; set; }
        public double TopLevelOffset { get; set; }

        public static void getRevitColumns(Document doc, UIDocument uidoc, ref List<RevitColumn> Columns, ForgeTypeId dis)
        {

            using (Transaction allTrns = new Transaction(doc, "aaa"))
            {
                allTrns.Start();
                FailureHandlingOptions failOpt = allTrns.GetFailureHandlingOptions();
                failOpt.SetFailuresPreprocessor(new RoomWarningSwallower());
                allTrns.SetFailureHandlingOptions(failOpt);

                #region get Selected Elements 
                ICollection<ElementId> selectedIds = uidoc.Selection.GetElementIds();

                // Create a filter for Structural Framing and Structural Beams categories
                List<BuiltInCategory> categories = new List<BuiltInCategory>();
                categories.Add(BuiltInCategory.OST_StructuralFraming);
                categories.Add(BuiltInCategory.OST_StructuralColumns);


                // Create an ElementCategoryFilter for each category and combine them with a LogicalOrFilter
                List<ElementFilter> categoryFilters = new List<ElementFilter>();
                foreach (BuiltInCategory category in categories)
                {
                    categoryFilters.Add(new ElementCategoryFilter(category));
                }
                //
                LogicalOrFilter orFilter = new LogicalOrFilter(categoryFilters);

                // Get the selected elements that match the filter
                List<Element> selectedElements = new List<Element>();
                foreach (ElementId id in selectedIds)
                {
                    Element element = uidoc.Document.GetElement(id);


                    FilteredElementCollector tagCollector = new FilteredElementCollector(doc)
    .OfCategory(BuiltInCategory.OST_Tags);



                    if (orFilter.PassesFilter(element))
                    {
                        selectedElements.Add(element);
                    }
                }


                List<FamilyInstance> selectedStrucralColumns = selectedElements.Where(x => x.Category.Id.IntegerValue == (int)BuiltInCategory.OST_StructuralColumns).Cast<FamilyInstance>().ToList();


                #endregion


                List<FamilyInstance> Allcolumns = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralColumns).WhereElementIsNotElementType().Cast<FamilyInstance>().ToList();

                if (selectedStrucralColumns.Count > 0)
                {
                    Allcolumns = selectedStrucralColumns;
                }

                #region get Data of columns
                foreach (var instance in Allcolumns.AsParallel())
                {

                    try
                    {
                        AnalyticalToPhysicalAssociationManager analyManger = AnalyticalToPhysicalAssociationManager.GetAnalyticalToPhysicalAssociationManager(doc);
                        AnalyticalElement analytInstance = doc.GetElement(analyManger.GetAssociatedElementId(instance.Id)) as AnalyticalElement;


                        RevitColumn frame = new RevitColumn();

                        if (analytInstance.GetCurve() != null)
                        {
                            Point startPoint = new Point(analytInstance.GetCurve().GetEndPoint(0));
                            Point endPoint = new Point(analytInstance.GetCurve().GetEndPoint(1));

                            startPoint.X = UnitUtils.ConvertFromInternalUnits(startPoint.X, dis);
                            startPoint.Y = UnitUtils.ConvertFromInternalUnits(startPoint.Y, dis);
                            startPoint.Z = UnitUtils.ConvertFromInternalUnits(startPoint.Z, dis);

                            endPoint.X = UnitUtils.ConvertFromInternalUnits(endPoint.X, dis);
                            endPoint.Y = UnitUtils.ConvertFromInternalUnits(endPoint.Y, dis);
                            endPoint.Z = UnitUtils.ConvertFromInternalUnits(endPoint.Z, dis);

                            Point midPoint = Point.getMidPoint(startPoint, endPoint);
                            frame.RvtID = instance.Id.ToString();
                            frame.revitEleId = instance.Id;
                            frame.StartPoint = startPoint;
                            frame.EndPoint = endPoint;
                            frame.MidPoint = midPoint;
                            frame.Length = Point.getDistance(startPoint, endPoint);


                            frame.Material = instance.LookupParameter("Material").AsString();
                            frame.RvtFamily = instance.Symbol.FamilyName;
                            frame.RvtName = instance.Symbol.Name;
                            frame.MemberType = MemberType.Column;

                            #region getRotationAngle
                            LocationPoint locationPoint = instance.Location as LocationPoint;
                            double angle = 0.0;
                            if (locationPoint == null)
                            {

                                frame.Beta = instance.LookupParameter("Cross-Section Rotation").AsDouble() * 180 / (Math.PI);

                            }

                            else
                            {
                                angle = locationPoint.Rotation;
                                frame.Beta = (angle * 180 / Math.PI) + 90;

                            }

                            #endregion


                            frame.BaseLevel = ((Level)doc.GetElement(instance.LookupParameter("Base Level").AsElementId())).Name;
                            frame.TopLevel = ((Level)doc.GetElement(instance.LookupParameter("Top Level").AsElementId())).Name;


                            frame.BaseLevelOffset = instance.LookupParameter("Base Offset").AsDouble();
                            frame.TopLevelOffset = instance.LookupParameter("Top Offset").AsDouble();




                            Columns.Add(frame);
                        }

                        else
                        {
                            continue;
                        }

                    }

                    catch (Exception ex2)
                    {

                        System.Diagnostics.StackTrace trace = new System.Diagnostics.StackTrace(ex2, true);
                        RevitUtils.Log.Add("Error in getRevitFrames Method in the Columns part... Columninstance no" + instance.Id + ex2.Message + trace.GetFrame(trace.FrameCount - 1).GetFileLineNumber());

                        // TaskDialog.Show("ErrorInGetRevitFramesMethod", "ErrorInRevitBeamID= "+ instance.Id + ex2.Message + trace.GetFrame(trace.FrameCount - 1).GetFileLineNumber());
                        continue;
                    }

                }





                #endregion



                allTrns.Commit();

            }




        }


    }



    public class RevitLevel
    {
        public string Name { get; set; }
        public double Elevation { get; set; }

        public static void getRevitLevels(Document doc, UIDocument uidoc, ref List<RevitLevel> levels, ForgeTypeId dis)
        {

            List<Level> revitLevels = new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().ToList();

            foreach (var item in revitLevels)
            {
                RevitLevel lvl = new RevitLevel();

                lvl.Name = item.Name;
                lvl.Elevation = item.Elevation;
                levels.Add(lvl);
            }

        }


        public static void setRevitLevels(Document doc, UIDocument uidoc, List<RevitLevel> levels, ForgeTypeId dis)
        {


            List<Level> revitLevels = new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().ToList();
            List<string> revitLevelsNames = revitLevels.Select(x => x.Name).ToList();

            List<Level> existingLevels = revitLevels.Where(L1 => levels.Any(L2 => L2.Name == L1.Name)).ToList();


            using (Transaction t = new Transaction(doc, "Modify Levels"))
            {
                t.Start();
                // add transaction
                existingLevels.ForEach(L1 => L1.Elevation = levels.FirstOrDefault(L2 => L2.Name == L1.Name).Elevation);

                t.Commit();

            }


            List<RevitLevel> newlevels = levels.Where(l1 => !revitLevels.Any(l2 => l2.Name == l1.Name)).ToList();

            foreach (var lvl in newlevels)
            {

                using (Transaction t = new Transaction(doc, "Create Levels"))
                {
                    t.Start();
                    Level l = Level.Create(doc, lvl.Elevation);
                    l.Name = lvl.Name;

                    t.Commit();

                }


            }

        }


    }


    #endregion

    #region RevitFamilies
    public class RevitFamily
    {

        #region MethodsOfDimentions

        public static void getPipeDimentionsFromAnalysisSectionNames(Frame analysisMember, out string dimentionswithX, out string[] dimentions)
        {
            char cc = analysisMember.Sec.SkipWhile(c => !char.IsDigit(c)).ToList()[0];
            int lastSpacecharacterIndex = analysisMember.Sec.IndexOf(cc);
            dimentionswithX = analysisMember.Sec.Substring(lastSpacecharacterIndex);
            dimentions = dimentionswithX.Split(("X".ToCharArray()));
        }

        public static void getTubeDimnsionsFromAnalysisSectionNames(Frame analysisMember, out string dimentionswithX, out string[] dimentions, out int noOfX)
        {
            char cc = analysisMember.Sec.SkipWhile(c => !char.IsDigit(c)).ToList()[0];
            int lastSpacecharacterIndex = analysisMember.Sec.IndexOf(cc);
            dimentionswithX = analysisMember.Sec.Substring(lastSpacecharacterIndex);
            dimentions = dimentionswithX.Split(("X".ToCharArray()));
            noOfX = analysisMember.Sec.Count(a => a == 'X');
        }

        public static void getBUS_IDimensionsFromAnalysisSectionNames(Frame analysisMember, out string webThick, out string flgThick, out string consWebDepth, out string consFlgDepth)
        {
            #region getDimentions 
            int seperatorIndex = analysisMember.Sec.LastIndexOf("+WEB");

            string webString = analysisMember.Sec.Substring(seperatorIndex + 4);
            string flgString = analysisMember.Sec.Substring(0, seperatorIndex - 1);


            string[] webstring2 = webString.Split(("X".ToCharArray()));
            webThick = webstring2[1];

            string[] flgString2 = flgString.Split(("X".ToCharArray()));
            flgThick = flgString2[1];

            consWebDepth = null;
            consFlgDepth = null;

            consWebDepth = (Convert.ToDouble(new string((webstring2[0].SkipWhile(c => !char.IsDigit(c))
                    .TakeWhile(c => char.IsDigit(c)).ToArray()
                   ))) + 2 * Convert.ToDouble(flgThick)).ToString();



            consFlgDepth = new string((flgString2[0].Substring(2).SkipWhile(c => !char.IsDigit(c))
                   .TakeWhile(c => char.IsDigit(c)).ToArray()
                   ));


            #endregion
        }

        public static string getChannelDimnsionsFromAnalysisSectionNames(Frame analysisMember)
        {
            char cc = analysisMember.Sec.SkipWhile(c => !char.IsDigit(c)).ToList()[0];
            int lastSpacecharacterIndex = analysisMember.Sec.IndexOf(cc);
            string dimentionswithX = analysisMember.Sec.Substring(lastSpacecharacterIndex);
            return dimentionswithX;
        }

        public static void getTapered_IDimensionsFromAnalysisSectionNames(Frame analysisMember, out string webThick, out string flgThick, out string startWebDepth, out string endWebDepth, out string consFlgDepth)
        {
            #region getDimentions 

            int seperatorIndex = analysisMember.Sec.LastIndexOf("+WEB");

            string webString = analysisMember.Sec.Substring(seperatorIndex + 4);
            string flgString = analysisMember.Sec.Substring(0, seperatorIndex - 1);


            string[] webstring2 = webString.Split(("X".ToCharArray()));
            webThick = webstring2[1];

            string[] flgString2 = flgString.Split(("X".ToCharArray()));
            flgThick = flgString2[1];
            startWebDepth = null;
            endWebDepth = null;

            consFlgDepth = null;


            // get largeDepth ,smalldepth of the web and constWidth of the flange
            string[] webstring3 = webstring2[0].Split(("/".ToCharArray()));




            startWebDepth = (Convert.ToDouble(new string((webstring3[0].SkipWhile(c => !char.IsDigit(c))
                 .TakeWhile(c => char.IsDigit(c)).ToArray()
                )))).ToString();



            endWebDepth = (Convert.ToDouble(new string((webstring3[1].SkipWhile(c => !char.IsDigit(c))
                   .TakeWhile(c => char.IsDigit(c)).ToArray()
                  )))).ToString();



            consFlgDepth = new string((flgString2[0].Substring(2).SkipWhile(c => !char.IsDigit(c))
                   .TakeWhile(c => char.IsDigit(c)).ToArray()
                   ));

            #endregion
        }

        public static void getTaperedPipeDimensionsFromAnalysisSectionNames(Frame analysisMember, out string startDiamaterValue, out string endDiamaterValue, out string PipeThicknessValue)
        {
            #region Dimensions
            int lastXIndex = analysisMember.Sec.LastIndexOf("X");
            string Diameters = analysisMember.Sec.Substring(0, lastXIndex);

            int lastseperetorIndex = Diameters.LastIndexOf("/");



            startDiamaterValue = new string((Diameters.SkipWhile(c => !char.IsDigit(c))
                 .TakeWhile(c => char.IsDigit(c)).ToArray()
                ));
            endDiamaterValue = Diameters.Substring(lastseperetorIndex + 1);
            PipeThicknessValue = analysisMember.Sec.Substring(lastXIndex + 1);

            #endregion
        }

        public static void getTaperedTubeDimensionsFromAnalysisSectionNames(Frame analysisMember, out string startWebDepth, out string endWebDepth, out string webThick, out string startFlgWidth, out string endFlgWidth, out string flgThick)
        {
            #region GetDimentions 

            startWebDepth = null;
            endWebDepth = null;
            webThick = null;

            startFlgWidth = null;
            endFlgWidth = null;
            flgThick = null;
            bool isFlgTapered = false;



            int no_of_x = Frame.CountCharsUsingIndex(analysisMember.Sec, 'X');
            int no_of_slash = Frame.CountCharsUsingIndex(analysisMember.Sec, '/');


            //  800/500X20
            if (no_of_x == 1 && no_of_slash == 1)
            {

                startWebDepth = startFlgWidth = new string(analysisMember.Sec.SkipWhile(c => !char.IsDigit(c)).TakeWhile(c => char.IsDigit(c)).ToArray());
                endWebDepth = endFlgWidth = new string(analysisMember.Sec.Substring(analysisMember.Sec.IndexOf("/")).SkipWhile(c => !char.IsDigit(c)).TakeWhile(c => char.IsDigit(c)).ToArray());



                webThick = flgThick = new string(analysisMember.Sec.Substring(analysisMember.Sec.IndexOf('X')).SkipWhile(c => !char.IsDigit(c)).TakeWhile(c => char.IsDigit(c)).ToArray());


            }

            //  800/500X300X20
            else if (no_of_x == 2 && no_of_slash == 1)
            {

                startWebDepth = new string(analysisMember.Sec.SkipWhile(c => !char.IsDigit(c)).TakeWhile(c => char.IsDigit(c)).ToArray());
                endWebDepth = new string(analysisMember.Sec.Substring(analysisMember.Sec.IndexOf("/")).SkipWhile(c => !char.IsDigit(c)).TakeWhile(c => char.IsDigit(c)).ToArray());


                startFlgWidth = endFlgWidth = new string(analysisMember.Sec.Substring(analysisMember.Sec.IndexOf('X')).SkipWhile(c => !char.IsDigit(c)).TakeWhile(c => char.IsDigit(c)).ToArray());
                webThick = flgThick = new string(analysisMember.Sec.Substring(analysisMember.Sec.LastIndexOf('X')).SkipWhile(c => !char.IsDigit(c)).TakeWhile(c => char.IsDigit(c)).ToArray());

            }

            //  800/500X400/300X20
            else if (no_of_x == 2 && no_of_slash == 2)
            {


                isFlgTapered = true;
                int smallFlgWidthIndex = Frame.IndexOfNth(analysisMember.Sec, "/", 1);



                startWebDepth = new string(analysisMember.Sec.SkipWhile(c => !char.IsDigit(c)).TakeWhile(c => char.IsDigit(c)).ToArray());
                endWebDepth = new string(analysisMember.Sec.Substring(analysisMember.Sec.IndexOf("/")).SkipWhile(c => !char.IsDigit(c)).TakeWhile(c => char.IsDigit(c)).ToArray());



                startFlgWidth = new string(analysisMember.Sec.Substring(analysisMember.Sec.IndexOf('X')).SkipWhile(c => !char.IsDigit(c)).TakeWhile(c => char.IsDigit(c)).ToArray());
                endFlgWidth = new string(analysisMember.Sec.Substring(smallFlgWidthIndex).SkipWhile(c => !char.IsDigit(c)).TakeWhile(c => char.IsDigit(c)).ToArray());


                webThick = flgThick = new string(analysisMember.Sec.Substring(analysisMember.Sec.LastIndexOf("X")).SkipWhile(c => !char.IsDigit(c)).TakeWhile(c => char.IsDigit(c)).ToArray());



            }



            #endregion
        }

        public static void getAnglesDimensionsFromAnalysisSectionNames(Frame analysisMember, out bool isEQSingleAngle, out bool isUNEQSingleAngle, out bool isEQDblAngle, out string dimentionswithX, out string firstDimention)
        {
            isEQSingleAngle = false;
            isUNEQSingleAngle = false;
            isEQDblAngle = false;



            int noOfX = Frame.CountCharsUsingIndex(analysisMember.Sec, 'X');

            dimentionswithX = new string(analysisMember.Sec.SkipWhile(c => !char.IsDigit(c)).ToArray());
            int indexOfSecondX = Frame.IndexOfNth(dimentionswithX, "X", 0);


            firstDimention = null;
            string secondDimention = null;



            if (noOfX == 2 && analysisMember.IsBUS != true && analysisMember.Shape == "Angle")
            {
                dimentionswithX = new string(analysisMember.Sec.SkipWhile(c => !char.IsDigit(c)).ToArray());
                firstDimention = new string(dimentionswithX.TakeWhile(c => char.IsDigit(c)).ToArray());
                secondDimention = new string(dimentionswithX.Substring(indexOfSecondX).SkipWhile(c => !Char.IsDigit(c)).TakeWhile(c => char.IsDigit(c)).ToArray());

                if (firstDimention == secondDimention)
                {
                    isEQSingleAngle = true;

                }

                else
                {
                    isUNEQSingleAngle = true;
                }

            }

            else if (noOfX == 1 && analysisMember.IsBUS != true && analysisMember.Shape == "Angle")
            {
                isEQSingleAngle = true;
                dimentionswithX = new string(analysisMember.Sec.SkipWhile(c => !char.IsDigit(c)).ToArray());
                firstDimention = new string(dimentionswithX.TakeWhile(c => char.IsDigit(c)).ToArray());

            }

            else if (analysisMember.Shape == "DblAngle")
            {

                isEQDblAngle = true;

                if (analysisMember.Sec[0] == '2')
                {
                    dimentionswithX = dimentionswithX.Substring(1, dimentionswithX.Length - 1);

                }



                int _1stindexofSlash = dimentionswithX.IndexOf("/");

                if (_1stindexofSlash != -1)
                {
                    dimentionswithX = dimentionswithX.Substring(1, _1stindexofSlash - 1);
                }

                dimentionswithX = new string(dimentionswithX.SkipWhile(c => !char.IsDigit(c)).ToArray());
                firstDimention = new string(dimentionswithX.TakeWhile(c => char.IsDigit(c)).ToArray());

            }
        }

        public static void getConcRectDimensionsFromAnalysisSectionNames(Frame analysisMember, out string dimentionswithX, out string[] dimentions)
        {
            char cc = analysisMember.Sec.SkipWhile(c => !char.IsDigit(c)).ToList()[0];
            int lastSpacecharacterIndex = analysisMember.Sec.IndexOf(cc);
            dimentionswithX = analysisMember.Sec.Substring(lastSpacecharacterIndex);
            dimentions = dimentionswithX.Split(("X".ToCharArray()));
        }

        #endregion

    }

    #region RVtFamilies

    public class BM_BUS_I : RevitFamily
    {
        public static string FamilyName = "BM_STL_BUS-I_DAR";
        public static string shape = "I";
        public static string library = "N.A";
        public static bool IsBus = true;
        public static string totalBeamDepth = "H";
        public static string flgWidth = "bf";
        public static string webThick = "tw";
        public static string flgThick = "tf";



    }

    public class BM_HEA : RevitFamily
    {
        public static string FamilyName = "BM_STL_HEA_DAR";
        public static string shape = "I";
        public static string library = "euro";
        public static bool IsBus = false;
        public static string totalDepth = "d";
        public static string flgWidth = "bf";
        public static string webThick = "tw";
        public static string flgThick = "tf";
        public static string flgRad = "ra";

    }

    public class BM_HEAA : RevitFamily
    {
        public static string FamilyName = "BM_STL_HEAA_DAR";
        public static string shape = "I";
        public static string library = "euro";
        public static bool IsBus = false;
        public static string totalDepth = "d";
        public static string flgWidth = "bf";
        public static string webThick = "tw";
        public static string flgThick = "tf";
        public static string flgRad = "ra";

    }
    public class BM_HEB : RevitFamily
    {
        public static string FamilyName = "BM_STL_HEB_DAR";
        public static string shape = "I";
        public static string library = "euro";
        public static bool IsBus = false;
        public static string totalDepth = "d";
        public static string flgWidth = "bf";
        public static string webThick = "tw";
        public static string flgThick = "tf";
        public static string flgRad = "ra";

    }

    public class BM_HEM : RevitFamily
    {
        public static string FamilyName = "BM_STL_HEM_DAR";
        public static string shape = "I";
        public static string library = "euro";
        public static bool IsBus = false;
        public static string totalDepth = "d";
        public static string flgWidth = "bf";
        public static string webThick = "tw";
        public static string flgThick = "tf";
        public static string flgRad = "ra";
    }
    public class BM_IPE : RevitFamily
    {
        public static string FamilyName = "BM_STL_IPE_DAR";
        public static string shape = "I";
        public static string library = "euro";
        public static bool IsBus = false;
        public static string totalDepth = "d";
        public static string flgWidth = "bf";
        public static string webThick = "tw";
        public static string flgThick = "tf";
        public static string flgRad = "ra";
    }
    public class BM_IPEA : RevitFamily
    {
        public static string FamilyName = "BM_STL_IPEA_DAR";
        public static string shape = "I";
        public static string library = "euro";
        public static bool IsBus = false;
        public static string totalDepth = "d";
        public static string flgWidth = "bf";
        public static string webThick = "tw";
        public static string flgThick = "tf";
        public static string flgRad = "ra";
    }

    public class BM_IPEO : RevitFamily
    {
        public static string FamilyName = "BM_STL_IPEO_DAR";
        public static string shape = "I";
        public static string library = "euro";
        public static bool IsBus = false;
        public static string totalDepth = "d";
        public static string flgWidth = "bf";
        public static string webThick = "tw";
        public static string flgThick = "tf";
        public static string flgRad = "ra";
    }

    public class BM_L_EQUAL : RevitFamily
    {
        public static string FamilyName = "BM_STL_L_EQUAL_DAR";
        public static string shape = "Angle";
        public static string library = "euro";
        public static bool IsBus = false;
        public static string hzLeg = "bf";
        public static string vrLeg = "d";
        public static string hzThick = "tf";
        public static string vrThick = "tw";
        public static string innerRad = "ri";
        public static string outerRad = "ra";
    }
    public class BM_2L_EQUALANGLE : RevitFamily
    {
        public static string FamilyName = "BM_STL_2L_EQUAL_DAR";
        public static string shape = "DblAngle";
        public static string library = "euro";
        public static bool IsBus = false;
        public static string hzLeg = "bf";
        public static string vrLeg = "d";
        public static string hzThick = "tf";
        public static string vrThick = "tw";
        public static string innerRad = "ri";
        public static string outerRad = "ra";
        public static string gap = "s";


    }
    public class BM_L_UNEQUAL : RevitFamily
    {
        public static string FamilyName = "BM_STL_L_UNEQUAL_DAR";
        public static string shape = "Angle";
        public static string library = "euro";
        public static bool IsBus = false;
        public static string totalHzLeg = "bf";
        public static string vrLeg = "d";
        public static string hzThick = "tf";
        public static string vrThick = "tw";
        public static string innerRad = "ri";
        public static string outerRad = "ra";

    }
    public class BM_PIPE : RevitFamily
    {
        public static string FamilyName = "BM_STL_PIPE_DAR";
        public static string shape = "Pipe";
        public static string library = "euro";
        public static bool IsBus = false;
        public static string outerDiameter = "d";
        public static string thick = "tw";


    }
    public class BM_TUBE : RevitFamily
    {
        public static string FamilyName = "BM_STL_TUBE_DAR";
        public static string shape = "Box";
        public static string library = "euro";
        public static bool IsBus = false;
        public static string webDepth = "d";
        public static string flgWidth = "bf";
        public static string webThick = "tw";
        public static string flgThick = "tf";
        public static string outerRad = "ri";
        public static string innerRad = "ra";


    }
    public class BM_UPN : RevitFamily
    {
        public static string FamilyName = "BM_STL_UPN_DAR";
        public static string shape = "Channel";
        public static string library = "euro";
        public static bool IsBus = false;
        public static string totalDepth = "d";
        public static string width = "bf";
        public static string flgThick = "tf";
        public static string webThick = "tw";
        public static string innerDiameter = "ra";
        public static string outerDiamater = "ri";

    }

    public class BM_BUS_I_Tapered : RevitFamily
    {
        public static string FamilyName = "BM_STL_BUS-I-Tapered_DAR";
        public static string shape = "VARI";
        public static string library = "N.A";
        public static bool IsBus = true;
        public static string startWebDepth = "H Web 2";
        public static string endWebDepth = "H Web 1";
        public static string flgWidth = "fl Width";
        public static string webThick = "Web Thickness";
        public static string flgThick = "fl Thickness";
    }
    public class BM_BUS_PIPE_Tapered : RevitFamily
    {
        public static string FamilyName = "BM_STL_BUS-PIPE-Tapered_DAR";
        public static string shape = "VARPipe";
        public static string library = "N.A";
        public static bool IsBus = true;
        public static string startDiametr = "Start d";
        public static string endDiametr = "End d";
        public static string startThick = "Start t";
        public static string endThick = "End t";


    }
    public class BM_BUS_TUBE_Tapered : RevitFamily
    {
        public static string FamilyName = "BM_STL_BUS-TUBE-Tapered_DAR";
        public static string shape = "VARBox";
        public static string library = "N.A";
        public static bool IsBus = true;
        public static string startWebDepth = "H1";
        public static string endWebDepth = "H2";
        public static string startFlgWidth = "B1";
        public static string endFlgWidth = "B2";
        public static string overAllThick = "thick";


    }
    public class BM_W : RevitFamily
    {
        public static string FamilyName = "BM_STL_W_DAR";
        public static string shape = "I";
        public static string library = "AISC";
        public static bool IsBus = false;
        public static string totalDepth = "d";
        public static string flgWidth = "bf";
        public static string webThick = "tw";
        public static string flgThick = "tf";
        public static string flgRad = "ra";
    }
    public class BM_UB : RevitFamily
    {

        public static string FamilyName = "BM_STL_UB_DAR";
        public static string shape = "I";
        public static string library = "BSShapes";
        public static bool IsBus = false;
        public static string totalDepth = "d";
        public static string flgWidth = "bf";
        public static string webThick = "tw";
        public static string flgThick = "tf";
        public static string flgRad = "ra";
    }
    public class BM_UC : RevitFamily
    {
        public static string FamilyName = "BM_STL_UC_DAR";
        public static string shape = "I";
        public static string library = "BSShapes";
        public static bool IsBus = false;
        public static string totalDepth = "d";
        public static string flgWidth = "bf";
        public static string webThick = "tw";
        public static string flgThick = "tf";
        public static string flgRad = "ra";
    }
    public class BM_ConcRect : RevitFamily
    {
        public static string FamilyName = "M_Concrete-Rectangular Beam";
        public static string shape = "Rectangular";
        public static string library = "N.A";
        public static bool IsBus = false;
        public static string depth = "H";
        public static string width = "B";

    }

    public class BM_Dummy : RevitFamily
    {
        public static string FamilyName = "BM_Dummy_DAR";

    }



    public class COL_BUS_I : RevitFamily
    {
        public static string FamilyName = "COL_STL_BUS-I_DAR";
        public static string shape = "I";
        public static string library = "N.A";
        public static bool IsBus = true;
        public static string totalBeamDepth = "H";
        public static string flgWidth = "B";
        public static string webThick = "tw";
        public static string flgThick = "tf";



    }

    public class COL_HEA : RevitFamily
    {
        public static string FamilyName = "COL_STL_HEA_DAR";
        public static string shape = "I";
        public static string library = "euro";
        public static bool IsBus = false;
        public static string totalDepth = "d";
        public static string flgWidth = "bf";
        public static string webThick = "tw";
        public static string flgThick = "tf";
        public static string flgRad = "ra";

    }

    public class COL_HEAA : RevitFamily
    {
        public static string FamilyName = "COL_STL_HEAA_DAR";
        public static string shape = "I";
        public static string library = "euro";
        public static bool IsBus = false;
        public static string totalDepth = "d";
        public static string flgWidth = "bf";
        public static string webThick = "tw";
        public static string flgThick = "tf";
        public static string flgRad = "ra";
    }
    public class COL_HEB : RevitFamily
    {
        public static string FamilyName = "COL_STL_HEB_DAR";
        public static string shape = "I";
        public static string library = "euro";
        public static bool IsBus = false;
        public static string totalDepth = "d";
        public static string flgWidth = "bf";
        public static string webThick = "tw";
        public static string flgThick = "tf";
        public static string flgRad = "ra";
    }

    public class COL_HEM : RevitFamily
    {
        public static string FamilyName = "COL_STL_HEM_DAR";
        public static string shape = "I";
        public static string library = "euro";
        public static bool IsBus = false;
        public static string totalDepth = "d";
        public static string flgWidth = "bf";
        public static string webThick = "tw";
        public static string flgThick = "tf";
        public static string flgRad = "ra";
    }
    public class COL_IPE : RevitFamily
    {
        public static string FamilyName = "COL_STL_IPE_DAR";
        public static string shape = "I";
        public static string library = "euro";
        public static bool IsBus = false;
        public static string totalDepth = "d";
        public static string flgWidth = "bf";
        public static string webThick = "tw";
        public static string flgThick = "tf";
        public static string flgRad = "ra";
    }
    public class COL_IPEA : RevitFamily
    {
        public static string FamilyName = "COL_STL_IPEA_DAR";
        public static string shape = "I";
        public static string library = "euro";
        public static bool IsBus = false;
        public static string totalDepth = "d";
        public static string flgWidth = "bf";
        public static string webThick = "tw";
        public static string flgThick = "tf";
        public static string flgRad = "ra";
    }

    public class COL_IPEO : RevitFamily
    {
        public static string FamilyName = "COL_STL_IPEO_DAR";
        public static string shape = "I";
        public static string library = "euro";
        public static bool IsBus = false;
        public static string totalDepth = "d";
        public static string flgWidth = "bf";
        public static string webThick = "tw";
        public static string flgThick = "tf";
        public static string flgRad = "ra";
    }
    public class COL_L_EQUAL : RevitFamily
    {
        public static string FamilyName = "COL_STL_L_EQUAL_DAR";
        public static string shape = "Angle";
        public static string library = "euro";
        public static bool IsBus = false;
        public static string hzLeg = "bf";
        public static string vrLeg = "d";
        public static string hzThick = "tf";
        public static string vrThick = "tw";
        public static string innerRad = "ri";
        public static string outerRad = "ra";
    }
    public class COL_2L_EQUALANGLE : RevitFamily
    {
        public static string FamilyName = "COL_STL_2L_EQUAL_DAR";
        public static string shape = "DblAngle";
        public static string library = "euro";
        public static bool IsBus = false;
        public static string hzLeg = "bf";
        public static string vrLeg = "d";
        public static string hzThick = "tf";
        public static string vrThick = "tw";
        public static string innerRad = "ri";
        public static string outerRad = "ra";
        public static string gap = "s";


    }
    public class COL_L_UNEQUAL : RevitFamily
    {
        public static string FamilyName = "COL_STL_L_UNEQUAL_DAR";
        public static string shape = "Angle";
        public static string library = "euro";
        public static bool IsBus = false;
        public static string totalHzLeg = "bf";
        public static string vrLeg = "d";
        public static string hzThick = "tf";
        public static string vrThick = "tw";
        public static string innerRad = "ri";
        public static string outerRad = "ra";

    }
    public class COL_PIPE : RevitFamily
    {
        public static string FamilyName = "COL_STL_PIPE_DAR";
        public static string shape = "Pipe";
        public static string library = "euro";
        public static bool IsBus = false;
        public static string outerDiameter = "d";
        public static string thick = "tw";

    }
    public class COL_TUBE : RevitFamily
    {
        public static string FamilyName = "COL_STL_TUBE_DAR";
        public static string shape = "Box";
        public static string library = "euro";
        public static bool IsBus = false;
        public static string webDepth = "d";
        public static string flgWidth = "bf";
        public static string flgThick = "tf";
        public static string webThick = "tw";
        public static string innerRad = "ra";
        public static string outerRad = "ri";


    }
    public class COL_UPN : RevitFamily
    {
        public static string FamilyName = "COL_STL_UPN_DAR";
        public static string shape = "Channel";
        public static string library = "euro";
        public static bool IsBus = false;
        public static string totalDepth = "d";
        public static string width = "bf";
        public static string flgThick = "tf";
        public static string webThick = "tw";
        public static string innerDiameter = "ra";
        public static string outerDiamater = "ri";

    }

    public class COL_BUS_I_Tapered : RevitFamily
    {
        public static string FamilyName = "COL_STL_BUS-I-Tapered_DAR";
        public static string shape = "VARI";
        public static string library = "N.A";
        public static bool IsBus = true;
        public static string startWebDepth = "H_Bott";
        public static string endWebDepth = "H_Top";
        public static string startFlgWidth = "Bf_Bott";
        public static string endFlgWidth = "Bf_Top";

        public static string webThick = "tw";
        public static string flgThick = "tf";
    }

    public class COL_BUS_PIPE_Tapered : RevitFamily
    {
        public static string FamilyName = "COL_STL_BUS-PIPE-Tapered_DAR";
        public static string shape = "VARPipe";
        public static string library = "N.A";
        public static bool IsBus = true;
        public static string startInnerDiam = "Bott Inner Dia";
        public static string startOuterDiam = "Bott Outer Dia";

        public static string endInnerDiam = "Top Inner Dia";
        public static string endOuterDiam = "Top Outer Dia";

    }
    public class COL_BUS_TUBE_Tapered : RevitFamily
    {
        public static string FamilyName = "COL_STL_BUS-TUBE-Tapered_DAR";
        public static string shape = "VARBox";
        public static string library = "N.A";
        public static bool IsBus = true;
        public static string endWebDepth = "Top h";
        public static string endWebThick = "Top tw";
        public static string endFlgWidth = "Top b";
        public static string endFlgThick = "Top tf";

        public static string startWebDepth = "Bottom h";
        public static string startWebThick = "Bottom tw";
        public static string startFlgWidth = "Bottom b";
        public static string startFlgThick = "Bottom tf";


    }
    public class COL_W : RevitFamily
    {
        public static string FamilyName = "COL_STL_W_DAR";
        public static string shape = "I";
        public static string library = "AISC";
        public static bool IsBus = false;
        public static string totalDepth = "d";
        public static string flgWidth = "bf";
        public static string webThick = "tw";
        public static string flgThick = "tf";
        public static string flgRad = "ra";
    }
    public class COL_UB : RevitFamily
    {
        public static string FamilyName = "COL_STL_UB_DAR";
        public static string shape = "I";
        public static string library = "BSShapes";
        public static bool IsBus = false;
        public static string totalDepth = "d";
        public static string flgWidth = "bf";
        public static string webThick = "tw";
        public static string flgThick = "tf";
        public static string flgRad = "ra";
    }
    public class COL_UC : RevitFamily
    {

        public static string FamilyName = "COL_STL_UC_DAR";
        public static string shape = "I";
        public static string library = "BSShapes";
        public static bool IsBus = false;
        public static string totalDepth = "d";
        public static string flgWidth = "bf";
        public static string webThick = "tw";
        public static string flgThick = "tf";
        public static string flgRad = "ra";
    }

    public class COL_ConcRect : RevitFamily

    {
        public static string FamilyName = "M_Concrete-Rectangular Column";
        public static string shape = "Rectangular";
        public static string library = "N.A";
        public static bool IsBus = false;
        public static string width = "b";
        public static string depth = "h";

    }

    public class COL_Dummy : RevitFamily
    {

        public static string FamilyName = "COL_Dummy_DAR";
    }


    #endregion


    #endregion
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
    public enum InsertionPoint
    {
        BottomLeft = 1,
        BottomCenter = 2,
        BottomRight = 3,
        MiddleLeft = 4,
        MiddleCenter = 5,
        MiddleRight = 6,
        TopLeft = 7,
        TopCenter = 8,
        TopRight = 9,
        Centroid = 10,
        ShearCenter = 11
    }

    public enum CurveType
    {
        Line,
        Arc,
        Ellipse,
        HermiteSpline,
        Other
    }




    #endregion

    #region handling all errors in revit
    public class RoomWarningSwallower : IFailuresPreprocessor
    {

        public FailureProcessingResult PreprocessFailures(FailuresAccessor a)
        {
            // inside event handler, get all warnings

            IList<FailureMessageAccessor> failures = a.GetFailureMessages();


            foreach (FailureMessageAccessor f in failures)
            {

                FailureDefinitionId id = f.GetFailureDefinitionId();
                string jj = f.GetDescriptionText();
                if (jj.Contains("Keep") && jj.Contains("Join"))
                {
                    jj = "";
                }

                #region old
                //if (id == BuiltInFailures.FamilyFailures.CannotMakeTypeCB)
                //{
                //    if (f.HasResolutions())
                //    {
                //        a.ResolveFailure(f);
                //        a.DeleteAllWarnings();
                //        a.DeleteWarning(f);
                //        return FailureProcessingResult.Continue;
                //    }
                //    a.DeleteWarning(f);
                //    return FailureProcessingResult.Continue;


                //}

                //if (id == BuiltInFailures.FamilyFailures.SlantedColumnEndPlanesHaveFlipped)
                //{
                //    List<ElementId> allelementsw = f.GetFailingElementIds().ToList();
                //    string sapId = a.GetDocument().GetElement(allelementsw[0]).LookupParameter("SapID").AsString();

                //    a.DeleteElements(allelementsw);
                //    //a.DeleteWarning(f);
                //    return FailureProcessingResult.ProceedWithCommit;
                //}
                //if (id == BuiltInFailures.CurveFailures.LineTooShortError)
                //{
                //    if (f.HasResolutions())
                //    {
                //        a.ResolveFailure(f);
                //        a.DeleteWarning(f);
                //        return FailureProcessingResult.Continue;
                //    }
                //    a.DeleteWarning(f);
                //    return FailureProcessingResult.Continue;


                //}

                //if (id == BuiltInFailures.GeneralFailures.ErrorInFamilyResolved || id == BuiltInFailures.GeneralFailures.ErrorInSymbolFamilyResolved || id == BuiltInFailures.GeneralFailures.ErrorResolved)
                //{
                //    if (f.HasResolutions())
                //    {
                //        a.ResolveFailure(f);
                //        a.DeleteWarning(f);
                //        return FailureProcessingResult.Continue;
                //    }
                //    a.DeleteWarning(f);
                //    return FailureProcessingResult.Continue;


                //}

                //if (id == BuiltInFailures.OverlapFailures.DuplicateInstances)
                //{
                //    List<ElementId> allelementsw = f.GetFailingElementIds().ToList();
                //    List<ElementId> deletedelements = new List<ElementId>();
                //    deletedelements.Add(allelementsw[1]);
                //    a.DeleteElements(deletedelements);
                //    a.DeleteWarning(f);
                //    return FailureProcessingResult.ProceedWithCommit;
                //}


                //if (f.HasResolutions())
                //{
                //    a.ResolveFailure(f);
                //    a.DeleteWarning(f);
                //    a.DeleteAllWarnings();

                //    return FailureProcessingResult.ProceedWithCommit;
                //}

                //else
                //{
                //    a.DeleteWarning(f);
                //    return FailureProcessingResult.ProceedWithCommit;
                //}


                #endregion

                #region new


                if (id == BuiltInFailures.FamilyFailures.CannotMakeTypeCB)
                {
                    if (f.HasResolutions())
                    {
                        a.ResolveFailure(f);
                        a.DeleteAllWarnings();
                        a.DeleteWarning(f);
                        return FailureProcessingResult.Continue;
                    }
                    a.DeleteWarning(f);
                    return FailureProcessingResult.Continue;


                }

                if (id == BuiltInFailures.FamilyFailures.SlantedColumnEndPlanesHaveFlipped)
                {
                    List<ElementId> allelementsw = f.GetFailingElementIds().ToList();
                    string sapId = a.GetDocument().GetElement(allelementsw[0]).LookupParameter("SapID").AsString();

                    a.DeleteElements(allelementsw);
                    //a.DeleteWarning(f);
                    return FailureProcessingResult.ProceedWithCommit;
                }
                if (id == BuiltInFailures.CurveFailures.LineTooShortError)
                {
                    if (f.HasResolutions())
                    {
                        a.ResolveFailure(f);
                        a.DeleteWarning(f);
                        return FailureProcessingResult.Continue;
                    }
                    a.DeleteWarning(f);
                    return FailureProcessingResult.Continue;


                }

                if (id == BuiltInFailures.GeneralFailures.ErrorInFamilyResolved || id == BuiltInFailures.GeneralFailures.ErrorInSymbolFamilyResolved || id == BuiltInFailures.GeneralFailures.ErrorResolved)
                {
                    if (f.HasResolutions())
                    {
                        a.ResolveFailure(f);
                        a.DeleteWarning(f);
                        return FailureProcessingResult.Continue;
                    }
                    a.DeleteWarning(f);
                    return FailureProcessingResult.Continue;


                }

                if (id == BuiltInFailures.OverlapFailures.DuplicateInstances)
                {
                    List<ElementId> allelementsw = f.GetFailingElementIds().ToList();
                    List<ElementId> deletedelements = new List<ElementId>();
                    deletedelements.Add(allelementsw[1]);
                    a.DeleteElements(deletedelements);
                    a.DeleteWarning(f);
                    return FailureProcessingResult.ProceedWithCommit;
                }


                if (f.GetSeverity() == FailureSeverity.Warning)
                {
                    if (f.HasResolutions())
                    {
                        a.ResolveFailure(f);
                        a.DeleteWarning(f);
                        a.DeleteAllWarnings();

                        return FailureProcessingResult.ProceedWithCommit;
                    }

                    else
                    {
                        a.DeleteWarning(f);
                        return FailureProcessingResult.ProceedWithCommit;
                    }
                }

                else if (f.GetSeverity() == FailureSeverity.Error)
                {
                    List<ElementId> allelementsw = f.GetFailingElementIds().ToList();
                    List<ElementId> deletedelements = new List<ElementId>();
                    // deletedelements.Add(allelementsw[1]);
                    a.DeleteElements(allelementsw);
                    // a.DeleteWarning(f);
                    return FailureProcessingResult.ProceedWithCommit;
                }



                #endregion

                if (BuiltInFailures.InaccurateFailures.InaccurateBeamOrBrace == id || BuiltInFailures.InaccurateFailures.InaccurateAreaLine == id || BuiltInFailures.InaccurateFailures.InaccurateSketchLine == id || BuiltInFailures.InaccurateFailures.InaccurateCurveBasedFamily == id || BuiltInFailures.InaccurateFailures.InaccurateDriveCurve == id || BuiltInFailures.InaccurateFailures.InaccurateMassingSketchLine == id || BuiltInFailures.InaccurateFailures.InaccurateGrid == id || BuiltInFailures.InaccurateFailures.InaccurateLevel == id || BuiltInFailures.InaccurateFailures.InaccurateLine == id || BuiltInFailures.InaccurateFailures.InaccurateRefPlane == id || BuiltInFailures.InaccurateFailures.InaccurateRoomSeparation == id)
                {

                    a.DeleteWarning(f);
                    return FailureProcessingResult.ProceedWithCommit;


                }


                if (id == BuiltInFailures.JoinElementsFailures.CannotJoinElements || id == BuiltInFailures.JoinElementsFailures.CannotJoinElementsError || id == BuiltInFailures.JoinElementsFailures.CannotJoinElementsStructural || id == BuiltInFailures.JoinElementsFailures.CannotJoinElementsWarn || id == BuiltInFailures.JoinElementsFailures.CannotKeepJoined || id == BuiltInFailures.JoinElementsFailures.CannotCutJoinedGeometry || id == BuiltInFailures.JoinElementsFailures.CannotJoinDesignOptions || id == BuiltInFailures.JoinElementsFailures.CannotKeepReferenceOfJoinedElement || id == BuiltInFailures.JoinElementsFailures.CannotKeepReferenceOfJoinedElementWarning || id == BuiltInFailures.JoinElementsFailures.ElementWillBeDetachedFromPlane || id == BuiltInFailures.JoinElementsFailures.CannotCope)
                {

                    bool isResolved = false;

                    if (f.HasResolutions())
                    {
                        a.ResolveFailure(f);
                        isResolved = true;
                    }


                    if (isResolved)
                    {

                        return FailureProcessingResult.ProceedWithCommit;
                    }

                    return FailureProcessingResult.Continue;

                }


                return FailureProcessingResult.Continue;
            }
            return FailureProcessingResult.Continue;
        }
    }
    #endregion


}