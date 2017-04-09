using Microsoft.Office.Interop.Excel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using salesfusion;
using System.Diagnostics;

namespace salesfusion
{
    public class Program
    {
        public static Entity GetExcelData(Entity entityObj)
        {

            entityObj.eventWeightsObj = new Hashtable();
            entityObj.eventWeightsObj.Add("web", 1.0);
            entityObj.eventWeightsObj.Add("email", 1.2);
            entityObj.eventWeightsObj.Add("social", 1.5);
            entityObj.eventWeightsObj.Add("webinar", 2.0);
            try
            {
                entityObj.applicationObj = new Application();
                entityObj.scoresObj = new SortedDictionary<int, double>();
                if (entityObj.applicationObj.Workbooks.Open(entityObj.path) != null)
                {
                    entityObj.workBookobj = entityObj.applicationObj.Workbooks.Open(entityObj.path);
                }
                else
                {
                    throw new Exception("Enter valid path!");
                }

                foreach (Worksheet worksheetObj in entityObj.workBookobj.Worksheets)
                {
                    foreach (Range rangeObj in worksheetObj.UsedRange)
                    {
                        if (!entityObj.scoresObj.ContainsKey(Convert.ToInt32(rangeObj.Cells[1, 1].Value2.ToString().Split(',')[0])))
                        {
                            entityObj.scoresObj.Add(Convert.ToInt32(rangeObj.Cells[1, 1].Value2.ToString().Split(',')[0]), Convert.ToDouble(rangeObj.Cells[1, 1].Value2.ToString().Split(',')[2]) * entityObj.eventWeightsObj[rangeObj.Cells[1, 1].Value2.ToString().Split(',')[1]]);
                        }
                        else
                        {
                            entityObj.scoresObj[Convert.ToInt32(rangeObj.Cells[1, 1].Value2.ToString().Split(',')[0])] += Convert.ToDouble(rangeObj.Cells[1, 1].Value2.ToString().Split(',')[2]) * entityObj.eventWeightsObj[rangeObj.Cells[1, 1].Value2.ToString().Split(',')[1]];
                            entityObj.scoresObj[Convert.ToInt32(rangeObj.Cells[1, 1].Value2.ToString().Split(',')[0])] = Math.Round(entityObj.scoresObj[Convert.ToInt32(rangeObj.Cells[1, 1].Value2.ToString().Split(',')[0])]);
                        }
                    }
                }

                return entityObj;
            }
            catch (Exception e)
            {
                if (e.Message == "Cannot perform runtime binding on a null reference")
                    Console.WriteLine("Check Excel Input!");
                else
                    Console.WriteLine(e.Message);
                return null;
            }
            finally
            {
                if (entityObj.applicationObj != null)
                {
                    entityObj.applicationObj = null;
                    entityObj.workBookobj = null;
                }
            }
        }

        static void Main(string[] args)
        {
            Entity entityObj = new Entity();
            //entityObj.path = @"V:\Input.xlsx";
            Console.WriteLine("Enter the path and press Enter - \n");
            entityObj.path = Console.ReadLine();
            entityObj = entityObj.path != null ? GetExcelData(entityObj) : null;
            if (entityObj != null)
            {
                List<int> keysObj = new List<int>(entityObj.scoresObj.Keys);
                double min = entityObj.scoresObj.Select(score => score.Value).Min();
                if (min < 0)
                {
                    Console.WriteLine("Invalid Data");
                }
                double max = entityObj.scoresObj.Select(score => score.Value).Max();
                foreach (int key in keysObj)
                {
                    entityObj.scoresObj[key] = ((entityObj.scoresObj[key] - min) / (max - min)) * 100;
                    string quartileLabel = entityObj.scoresObj[key] > 75 ? "platinum" : entityObj.scoresObj[key] > 50 && entityObj.scoresObj[key] <= 75 ? "gold" : entityObj.scoresObj[key] > 25 && entityObj.scoresObj[key] <= 50 ? "silver" : "bronze";
                    Console.WriteLine(key + ", " + quartileLabel + ", " + Convert.ToInt32(Math.Round(entityObj.scoresObj[key])));
                }
            }
            Console.WriteLine("Press Enter to exit!");
            Console.Read();
        }
    }
}
