using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using Org.BouncyCastle.Security;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DaysheetGenerator_2023
{
    internal class DateExtractor
    {
        private string _dayOfWeek;
        private string _dayDate;
        private string _month;
        private string _year;
        public DateExtractor(string filePath) 
        {
            HSSFWorkbook workbook;

            using (FileStream file = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                workbook = new HSSFWorkbook(file);
                file.Close();
                file.Dispose();
            }

            ISheet sheet = workbook.GetSheetAt(0);
            string sheetDate = sheet.GetRow(2).GetCell(6).ToString();
            
            DateTime dateValue = DateTime.Parse(sheetDate);
            _dayOfWeek = dateValue.ToString("ddd");
            _dayDate = dateValue.ToString("dd");
            _month = dateValue.ToString("MMM");
            _year = dateValue.ToString("yyyy");
        }

        public string DayOfWeek { get { return _dayOfWeek; } }
        public string DayDate { get { return _dayDate; } }
        public string Month { get { return _month; } }
        public string Year { get { return _year; } }
    }
}
