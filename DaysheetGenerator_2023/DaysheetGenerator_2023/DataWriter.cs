using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Org.BouncyCastle.Asn1.Cmp;

namespace DaysheetGenerator_2023
{
    internal class DataWriter
    {
        private int _attending;
        private int _associates;
        private int _assistants;
        private int _fellows;
        private int _residents;
        private string _outputFilePath;

        public DataWriter(Dictionary<string, string> staffPositiong, string outputFilePath) 
        {
            _attending = Int32.Parse(staffPositiong["Staff"]);
            _associates = Int32.Parse(staffPositiong["Associates"]);
            _assistants = Int32.Parse(staffPositiong["AAs"]);
            _fellows = Int32.Parse(staffPositiong["Fellows"]);
            _residents = Int32.Parse(staffPositiong["Residents"]);
            _outputFilePath = outputFilePath;
        }

        public void WriteSchedule(StaffType data, DateExtractor dateExtractor, Dictionary<string, string> ohipFellows)
        {
            string outputPath = _outputFilePath;
            XSSFWorkbook workbook;

            using (FileStream file = new FileStream(outputPath, FileMode.Open, FileAccess.Read))
            {
                workbook = new XSSFWorkbook(file);

                file.Close();
                file.Dispose();
            }

            int sheetIndex = getSheetIndex(dateExtractor);
            if (sheetIndex != 5) 
            {
                ISheet sheet = workbook.GetSheetAt(sheetIndex);

                writeAttendings(sheet, data.Attendings);
                writeAssociates(sheet, data.Associates);
                writeAssistants(sheet, data.Assistants);
                writeFellows(sheet, data.Fellows, ohipFellows);
                writeResidents(sheet, data.Residents);
                writeDate(sheet, dateExtractor);

                using (FileStream file = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
                {

                    workbook.Write(file);
                    file.Close();
                    file.Dispose();


                }
                //Console.WriteLine("Finished");
            }
            


        }

        private void writeDate(ISheet sheet, DateExtractor dateExtractor) 
        {
            int[] rows = {0, 72, 73, 141};
            foreach (int rowIndex in rows)
            {
                IRow row = sheet.GetRow(rowIndex);

                ICell cell = row.GetCell(2);
                cell.SetCellValue(dateExtractor.DayOfWeek);

                cell = row.GetCell(5);
                cell.SetCellValue(dateExtractor.Month);

                cell = row.GetCell(6);
                cell.SetCellValue(dateExtractor.DayDate);

                cell = row.GetCell(7);
                cell.SetCellValue(dateExtractor.Year);
            }
        }

        private void writeAttendings(ISheet sheet, List<List<string>> attendings) 
        {

            foreach (List<string> rowData in attendings)
            {
                IRow row = sheet.GetRow(_attending);
                ICell cell = row.GetCell(0);
                cell.SetCellValue(rowData[0]);

                cell = row.GetCell(2);
                cell.SetCellValue(rowData[1]);

                cell = row.GetCell(5);
                cell.SetCellValue(formatName(rowData[2]));

                cell = row.GetCell(6);
                cell.SetCellValue(rowData[3]);

                cell = row.GetCell(7);
                cell.SetCellValue(rowData[4]);

                _attending++;
            }

        }

        private void writeAssociates(ISheet sheet, List<List<string>> associates)
        {

            foreach (List<string> rowData in associates)
            {
                IRow row = sheet.GetRow(_associates);
                ICell cell = row.GetCell(0);
                cell.SetCellValue(rowData[0]);

                cell = row.GetCell(2);
                cell.SetCellValue(rowData[1]);

                cell = row.GetCell(5);
                cell.SetCellValue(formatName(rowData[2]));

                cell = row.GetCell(6);
                cell.SetCellValue(rowData[3]);

                cell = row.GetCell(7);
                cell.SetCellValue(rowData[4]);

                _associates++;
            }

        }

        private void writeAssistants(ISheet sheet, List<List<string>> assistants)
        {

            foreach (List<string> rowData in assistants)
            {
                IRow row = sheet.GetRow(_assistants);
                ICell cell = row.GetCell(0);
                cell.SetCellValue(rowData[0]);

                cell = row.GetCell(2);
                cell.SetCellValue(rowData[1]);

                cell = row.GetCell(5);
                cell.SetCellValue(formatName(rowData[2]));

                cell = row.GetCell(6);
                cell.SetCellValue(rowData[3]);

                cell = row.GetCell(7);
                cell.SetCellValue(rowData[4]);

                _assistants++;
            }

        }

        private void writeFellows(ISheet sheet, List<List<string>> fellows, Dictionary<string, string> ohipFellows)
        {

            foreach (List<string> rowData in fellows)
            {
                IRow row = sheet.GetRow(_fellows);
                //ICell cell = row.GetCell(0);
                //cell.SetCellValue(rowData[0]);

                ICell cell = row.GetCell(2);
                cell.SetCellValue(rowData[0] + rowData[1]);

                if (rowData[1] == "FC")
                {
                    rowData[3] = "PreCall";
                }

                Dictionary<string, string>.KeyCollection keys = ohipFellows.Keys;
                foreach (string key in keys)
                {
                    //Console.WriteLine("key: {0}", key);
                }
                //Console.WriteLine(ohipFellows.);

                if (ohipFellows.ContainsKey(rowData[2]))
                {
                    //Console.WriteLine(rowData[2]);
                    cell = row.GetCell(4);
                    cell.SetCellValue(ohipFellows[rowData[2]]);
                }

                cell = row.GetCell(5);
                cell.SetCellValue(formatName(rowData[2]));

                cell = row.GetCell(6);
                cell.SetCellValue(rowData[3]);

                cell = row.GetCell(7);
                cell.SetCellValue(rowData[4]);

                _fellows++;
            }

        }

        private void writeResidents(ISheet sheet, List<List<string>> residents)
        {

            foreach (List<string> rowData in residents)
            {
                //Console.WriteLine(rowData[2]);

                IRow row = sheet.GetRow(_residents);
                ICell cell = row.GetCell(0);
                cell.SetCellValue(rowData[0]);

                cell = row.GetCell(2);
                cell.SetCellValue(rowData[1]);

                cell = row.GetCell(5);
                cell.SetCellValue(formatName(rowData[2]));

                cell = row.GetCell(6);
                cell.SetCellValue(rowData[3]);

                cell = row.GetCell(7);
                cell.SetCellValue(rowData[4]);

                _residents++;
            }

        }

        private int getSheetIndex(DateExtractor dateExtractor) 
        { 
            switch (dateExtractor.DayOfWeek) 
            {
                case "Mon":
                    return 0;

                case "Tue":
                    return 1;

                case "Wed":
                    return 2;

                case "Thu":
                    return 3;

                case "Fri":
                    return 4;

                default:
                    return 5;

            }
        }

        private string formatName(string name) 
        {
            string finalizedName = "";
            string[] tokenizedName = name.Split(' ');

            for (int i = 0; i < tokenizedName.Length; i++) 
            { 
                if (tokenizedName[i].All(char.IsUpper) || tokenizedName[i].Contains("-") || tokenizedName[i].Contains("'"))
                {
                    finalizedName += char.ToUpper(tokenizedName[i][0]) + tokenizedName[i].Substring(1).ToLower() + " ";
                }
                else 
                {
                    finalizedName = finalizedName.TrimEnd();
                    finalizedName += ", " + char.ToUpper(tokenizedName[i][0]);
                    break;
                }
            }

            return finalizedName;
        }
    }
}
