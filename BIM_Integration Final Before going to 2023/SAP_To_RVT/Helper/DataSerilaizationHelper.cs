
using System;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json;

namespace BIM_Integration
{
    internal class DataSerilaizationHelper
    {
        public static List<T> LoadFromJsonFileOnDeskTop<T>(string fileName)
        {

            List<T> data = new List<T>();

            try
            {
                string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                string filePath = Path.Combine(desktopPath, fileName);

                if (File.Exists(filePath))
                {
                    using (StreamReader file = File.OpenText(filePath))
                    {
                        JsonSerializer serializer = new JsonSerializer();
                        data = (List<T>)serializer.Deserialize(file, typeof(List<T>));
                    }
                }

            }
            catch (Exception ex)
            {
            }

            return data;
        }

        public static void SaveToJsonFileOnDeskTop<T>(List<T> data, string fileName)
        {
            try
            {
                string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                string filePath = Path.Combine(desktopPath, fileName);

                using (StreamWriter file = File.CreateText(filePath))
                {
                    JsonSerializer serializer = new JsonSerializer();
                    serializer.Serialize(file, data);
                }

                Console.WriteLine($"Data saved to {filePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving data: {ex.Message}");
            }
        }


        public static List<dynamic> LoadFromJsonFile(string fileName)
        {
            List<dynamic> data = new List<dynamic>();

            try
            {
                string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                string filePath = Path.Combine(desktopPath, fileName);

                if (File.Exists(filePath))
                {
                    using (StreamReader file = File.OpenText(filePath))
                    {
                        JsonSerializer serializer = new JsonSerializer();
                        using (JsonReader reader = new JsonTextReader(file))
                        {
                            while (reader.Read())
                            {
                                if (reader.TokenType == JsonToken.StartObject)
                                {
                                    // Load object from JSON
                                    dynamic obj = serializer.Deserialize(reader);
                                    // Add only the desired properties to the list
                                    data.Add(obj);
                                }
                            }
                        }
                    }
                }
                else
                {
                    Console.WriteLine($"File '{fileName}' does not exist.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading data: {ex.Message}");
            }

            return data;
        }

        public static dynamic GetRequiredProperties<T>(dynamic obj)
        {
            // Define the properties you want to keep
            // Adjust this according to your object structure
            dynamic newObj = new System.Dynamic.ExpandoObject();
            newObj.ID = obj.StdID;
            newObj.StartNode = obj.StartNode;
            newObj.EndNode = obj.EndNode;
            newObj.startX = obj.startX;
            newObj.startY = obj.startY;
            newObj.startZ = obj.startZ;
            newObj.endX = obj.endX;
            newObj.endY = obj.endY;
            newObj.endZ = obj.endZ;
            newObj.Sec = obj.stdSec;
            newObj.Lib = obj.stdLib;
            newObj.Shape = obj.stdShape;
            newObj.IsBUS = obj.IsBUS;
            newObj.IsCable = obj.IsCable;
            newObj.IsLink = obj.IsLink;
            newObj.IsConc = obj.IsConc;
            newObj.Group = obj.Group;
            newObj.Material = obj.Material;
            newObj.Length = obj.Length;
            newObj.Beta = obj.Beta;

            return newObj;
        }

        public static List<Frame> getAllStaadFrames(string jsonFileName)
        {


            List<Frame> allFrames = new List<Frame>();

            List<dynamic> objs = LoadFromJsonFile(jsonFileName);

            foreach (var item in objs)
            {
                Frame frame = new Frame();
                frame.ID = item.ID;
                frame.StartPoint.Id = item.StartNode;
                frame.EndPoint.Id = item.EndNode;
                frame.StartPoint.X = item.startX;
                frame.StartPoint.Y = item.startY;
                frame.StartPoint.Z = item.startZ;

                frame.EndPoint.X = item.endX;
                frame.EndPoint.Y = item.endY;
                frame.EndPoint.Z = item.endZ;


                frame.Sec = item.Sec;
                frame.Lib = item.Lib;
                frame.Shape = item.Shape;
                frame.IsBUS = item.IsBUS;
                frame.IsCable = item.IsCable; ;
                frame.IsLink = item.IsLink; ;
                frame.IsConc = item.IsConc; ;
                frame.Group = item.Group;
                frame.Material = item.Material;
                frame.Length = item.Length;



                Point.getMemberType(frame);

                if (frame.MemberType == MemberType.SlanttedColumn || frame.MemberType == MemberType.Beam || frame.MemberType == MemberType.VLBracing)
                {
                    frame.Beta = -(Convert.ToDouble(item.Beta));
                }
                else
                {

                    frame.Beta = (Convert.ToDouble(item.Beta) + 90.00);
                }


                frame.MidPoint.X = (frame.StartPoint.X + frame.EndPoint.X) / 2;
                frame.MidPoint.Y = (frame.StartPoint.Y + frame.EndPoint.Y) / 2;
                frame.MidPoint.Z = (frame.StartPoint.Z + frame.EndPoint.Z) / 2;

                allFrames.Add(frame);
            }



            return allFrames;

        }


        public static List<Frame> getAllStaadFrames2(string jsonFileName)
        {


            List<Frame> allFrames = new List<Frame>();


            allFrames = LoadFromJsonFil2(jsonFileName);


            return allFrames;

        }


        public static List<Frame> LoadFromJsonFil2(string jsonFileName)
        {
            List<Frame> allFrames = new List<Frame>();

            try
            {
                string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                string filePath = Path.Combine(desktopPath, jsonFileName);

                if (File.Exists(filePath))
                {
                    using (StreamReader file = File.OpenText(filePath))
                    {
                        JsonSerializer serializer = new JsonSerializer();
                        allFrames = (List<Frame>)serializer.Deserialize(file, typeof(List<Frame>));
                    }
                }
                else
                {
                    Console.WriteLine($"File '{jsonFileName}' does not exist.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading data from JSON: {ex.Message}");
            }

            return allFrames;
        }
    }
}