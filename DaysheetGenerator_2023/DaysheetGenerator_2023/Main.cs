using System;
using System.IO;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;

/*
 Main function of the program.
 */

namespace DaysheetGenerator_2023
{
    class Main
    {
        public Main() { }

        public void RunProgram(ItemCollection filePaths, string outputFolderPath) 
        { 
            DataExtractor extractor = new DataExtractor();
            Dictionary<string, string> assignmentWording = extractor.extractConditions(0);
            Dictionary<string, string> assignmentsToNotes = extractor.extractConditions(1);
            Dictionary<string, string> AddNotes = extractor.extractConditions(2);
            Dictionary<string, string> fellowNotes= extractor.extractConditions(3);
            Dictionary<string, string> outputPositioning = extractor.extractConditions(4);
            List<string> residentOrdering = extractor.extractResidents(5);
            Dictionary<string, string> ohipFellows = extractor.extractConditions(7);


            PostCallExtractor postCallExtractor = new PostCallExtractor();
            List<List<List<string>>> finalSchedule = new List<List<List<string>>>();
            List<DateExtractor> dateList = new List<DateExtractor>();

           
            for (int i = 0; i < filePaths.Count; i++)
            {
                DateExtractor dateExtractor = new DateExtractor(filePaths[i].ToString());
                dateList.Add(dateExtractor);

                List<List<string>> sheetData = extractor.extractSchedule(filePaths[i].ToString());
                DataManipulator manipulator = new DataManipulator(sheetData, assignmentWording, assignmentsToNotes, AddNotes, fellowNotes);

                manipulator.CollateDuplicates();
                manipulator.RemoveDuplicates();
                manipulator.AdditionalNotes();
                manipulator.ChangeWording();
                manipulator.MoveToNotes();
                manipulator.FellowNotes();
                manipulator.AddSlash();

                List<List<string>> modifiedData = manipulator.getData();
                finalSchedule.Add(modifiedData);     
            }

            finalSchedule = OrganizeScheduleByDay(finalSchedule);
            dateList = OrganizeDatesByDay(dateList);
            finalSchedule = postCallExtractor.insertPostCalls(finalSchedule);

            string outputFilePath = createOutputFile(outputFolderPath, dateList[1].Month, dateList[1].DayDate);

            int dateIndex = 0;

            foreach (List<List<string>> daySchedule in finalSchedule)
            {
                StaffType staff = new StaffType();
                staff.StaffTypeSeparator(daySchedule);
                staff.orderResidents(residentOrdering);

                DataWriter writer = new DataWriter(outputPositioning, outputFilePath);
                writer.WriteSchedule(staff, dateList[dateIndex], ohipFellows);
                dateIndex++;
            }
        }

        public List<List<List<string>>> OrganizeScheduleByDay(List<List<List<string>>> unorderedSchedules)
        {
            List <List<List<string>>> orderedSchedules = new List<List<List<string>>>();
            string[] days = {"Sun", "Mon", "Tue", "Wed", "Thu", "Fri" };

            foreach (string day in days)
            {
                foreach (List<List<string>> daySchedule in unorderedSchedules)
                {
                    if (daySchedule[0][6] == day)
                    {  
                        orderedSchedules.Add(daySchedule);
                        break;
                    }
                }
            }
            return orderedSchedules;
        }

        public List<DateExtractor> OrganizeDatesByDay(List<DateExtractor> unorderedDates)
        {
            List<DateExtractor> orderedDate = new List<DateExtractor>();
            string[] days = { "Sun", "Mon", "Tue", "Wed", "Thu", "Fri" };

            foreach (string day in days)
            {
                foreach (DateExtractor unorderedDate in unorderedDates)
                {
                    if (unorderedDate.DayOfWeek == day)
                    {
                        
                        orderedDate.Add(unorderedDate);
                        break;
                    }
                }
            }
            return orderedDate;
        }

        public string createOutputFile(string outputFolderPath, string fileMonth, string fileDay)
        {
            string outputFilePath = outputFolderPath + "\\Daysheet_" + fileMonth + "_" + fileDay + ".xlsx";

            if (File.Exists(outputFilePath))
            {
                File.Delete(outputFilePath);
            }

            File.Copy("Template.xlsx", outputFilePath);
            return outputFilePath;
        }
    }
}
