using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DaysheetGenerator_2023
{
    internal class PostCallExtractor
    {

        public PostCallExtractor() { }


        public List<List<List<string>>> insertPostCalls(List<List<List<string>>> weekSchedule)
        {
            string thirdCall = "";
            string fourthCall = "";
            for (int i = 0; i < weekSchedule.Count; i++) 
            { 
                for (int j = 0; j < weekSchedule[i].Count; j++)
                {
                    if (weekSchedule[i][j][2] == thirdCall)
                    {
                        weekSchedule[i][j][0] = "Post 3rd Call";
                    }
                    if (weekSchedule[i][j][2] == fourthCall)
                    {
                        weekSchedule[i][j][0] = "Post 4th Call";
                    }
                }

                for (int j = 0; j < weekSchedule[i].Count; j++)
                {
                    if (weekSchedule[i][j][1] == "3")
                    {
                        thirdCall = weekSchedule[i][j][2];
                    }
                    if (weekSchedule[i][j][1] == "4")
                    {
                        fourthCall = weekSchedule[i][j][2];
                    }

                }
                
                
            }

            return weekSchedule;
        }
    }
}
