using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace DaysheetGenerator_2023
{
    internal class DataExtractor
    {

        public List<List<string>> extractSchedule(string filePath) 
        {
            List<List<string>> data = new List<List<string>>();
            HSSFWorkbook workbook;

            using (FileStream file = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                workbook = new HSSFWorkbook(file);
                file.Close();
                file.Dispose();
            }

            ISheet sheet = workbook.GetSheetAt(0);

            for (int row = 1; row < sheet.PhysicalNumberOfRows - 1; row++) 
            {
                List<string> sheetRow = new List<string>();
                sheetRow.Add(sheet.GetRow(row).GetCell(0).ToString());
                sheetRow.Add(sheet.GetRow(row).GetCell(1).ToString());
                sheetRow.Add(sheet.GetRow(row).GetCell(2).ToString().Replace(",", ""));
                sheetRow.Add(sheet.GetRow(row).GetCell(3).ToString());
                sheetRow.Add(sheet.GetRow(row).GetCell(4).ToString());
                sheetRow.Add(sheet.GetRow(row).GetCell(5).ToString());
                //Console.WriteLine(sheet.GetRow(row).GetCell(2).ToString());
                
                string dateString = sheet.GetRow(row).GetCell(6).ToString();
                DateTime dateTimeDay = DateTime.Parse(dateString);
                string day = dateTimeDay.ToString("ddd");
                sheetRow.Add(day);

                data.Add(sheetRow);
            }
            workbook.Close();
            return data;
        }

        public Dictionary<string, string> extractConditions(int index) 
        { 
            Dictionary<string, string> keyValuePairs = new Dictionary<string, string>();
            string conditionsFilePath = "DaysheetConditions.xlsx";
            XSSFWorkbook workbook;

            using (FileStream file = new FileStream(conditionsFilePath, FileMode.Open, FileAccess.Read))
            {
                workbook = new XSSFWorkbook(file);
                file.Close();
                file.Dispose();
            }

            ISheet sheet = workbook.GetSheetAt(index);
            //Console.WriteLine(index);

            for (int row = 0; row < sheet.PhysicalNumberOfRows; row++) 
            {
               
                if (sheet.GetRow(row) != null)
                {
                    if (sheet.GetRow(row).GetCell(0) == null)
                    {
                        break;
                    }
                    if (sheet.GetRow(row).GetCell(0).StringCellValue == "" ) 
                    {
                        break;
                    }
                    else if (sheet.GetRow(row).GetCell(1) != null)
                    {
                        //Console.WriteLine(sheet.GetRow(row).GetCell(0).ToString() + " " + sheet.GetRow(row).GetCell(1).ToString());
                        keyValuePairs.Add(sheet.GetRow(row).GetCell(0).ToString(), sheet.GetRow(row).GetCell(1).ToString());
                    }
                    else
                    {
                        keyValuePairs.Add(sheet.GetRow(row).GetCell(0).ToString(), "");
                    }
                    
                }
            }
            workbook.Close();
            return keyValuePairs;
        }

        public List<string> extractResidents(int index)
        {
            List<string> residents = new List<string>();
            string conditionsFilePath = "DaysheetConditions.xlsx";
            XSSFWorkbook workbook;

            using (FileStream file = new FileStream(conditionsFilePath, FileMode.Open, FileAccess.Read))
            {
                workbook = new XSSFWorkbook(file);
                file.Close();
                file.Dispose();
            }

            ISheet sheet = workbook.GetSheetAt(index);
            

            for (int row = 0; row < sheet.PhysicalNumberOfRows; row++)
            {

                if (sheet.GetRow(row) != null)
                {
                    Console.WriteLine("Hello");

                    if (sheet.GetRow(row).GetCell(0) == null)
                    {
                        Console.WriteLine("Hello1");
                        break;
                    }
                    if (sheet.GetRow(row).GetCell(0).StringCellValue == "")
                    {
                        Console.WriteLine("Hello2");
                        break;
                    }
                    else if (sheet.GetRow(row).GetCell(0) != null)
                    {
                        Console.WriteLine("Hello3");
                        residents.Add(sheet.GetRow(row).GetCell(0).ToString());
                    }
                    

                }
            }
            workbook.Close();
            return residents;
        }

    }
}
