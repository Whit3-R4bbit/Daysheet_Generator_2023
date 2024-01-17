using NPOI.SS.Formula.Functions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DaysheetGenerator_2023
{
    internal class DataManipulator
    {
        private List<List<string>> _data;
        private Dictionary<string,string> _assignmentWording;
        private Dictionary<string, string> _assignmentsToNotes;
        private Dictionary<string, string> _addNotes;
        private Dictionary<string, string> _fellowNotes;

        public DataManipulator(List<List<string>> data, Dictionary<string, string> assignmentWording, Dictionary<string, string> assignmentsToNotes, 
            Dictionary<string, string> addNotes, Dictionary<string, string> fellowNotes) 
        {
            this._data = data;
            this._assignmentWording = assignmentWording;
            this._assignmentsToNotes = assignmentsToNotes;
            this._addNotes = addNotes;
            this._fellowNotes = fellowNotes;
        }

        public void CollateDuplicates() 
        {
            List<string> row = _data[0];
            for (int i =1; i < _data.Count - 1; i++)
            { 
                if (row[2] == _data[i][2])
                {
                    row[0] += "|" + _data[i][0];
                    row[1] += "|" + _data[i][1];
                    row[3] += "|" + _data[i][3];
                    row[4] += "|" + _data[i][4];
                }
                else 
                {
                    row = _data[i];
                }
                row[0] = row[0].Trim('|');
                row[1] = row[1].Trim('|');
                row[3] = row[3].Trim('|');
                row[4] = row[4].Trim('|');
            }
            
        }

        public void RemoveDuplicates() 
        {   
            List<List<string>> duplicatesRemoved = new List<List<string>>();
            List<string> row = _data[0];
            duplicatesRemoved.Add(row);
            for (int i = 1; i < _data.Count - 1; i++) 
            { 
                if (row[2] != _data[i][2]) 
                {
                    row = _data[i];
                    duplicatesRemoved.Add(row);
                }
            }
            this._data = duplicatesRemoved;
            
        }

        public void ChangeWording() 
        { 
            for (int row = 0; row < _data.Count; row++) 
            {
                string outputAssignments = "";
                List<string> inputAssignments = _data[row][3].Split('|').ToList();

                foreach (string inputAssignment in inputAssignments) 
                {
                    if (_assignmentWording.ContainsKey(inputAssignment)) 
                    {
                        outputAssignments += "|" + _assignmentWording[inputAssignment];
                    }
                    else 
                    {
                        outputAssignments += "|" + inputAssignment;
                    }
                }
                outputAssignments = outputAssignments.Trim('|');
                _data[row][3] = outputAssignments;
            }


        }

        public void MoveToNotes() 
        { 
            for (int row = 0; row < _data.Count; row++) 
            {
                string outputAssignments = "";
                string outputNotes = "";
                List<string> inputAssignments = _data[row][3].Split('|').ToList();

                foreach (string inputAssignment in inputAssignments) 
                { 
                    if (_assignmentsToNotes.ContainsKey(inputAssignment)) 
                    {
                        
                        outputNotes += "|" + _assignmentsToNotes[inputAssignment];
                    }
                    else 
                    {
                        outputAssignments += "|" + inputAssignment;
                    }
                }
                _data[row][3] = outputAssignments.Trim('|');
                _data[row][4] += outputNotes.Trim('|');
                _data[row][4] = _data[row][4].Replace("\n", "");
                
            }
           
        }

        public void AdditionalNotes() 
        {
            for (int row = 0; row < _data.Count; row++) 
            {
                string outputNotes = "";
                List<string> inputAssignments = _data[row][3].Split('|').ToList();

                foreach (string inputAssignment in inputAssignments) 
                {
                    if (_addNotes.ContainsKey(inputAssignment))
                    {
                        outputNotes += "|" + _addNotes[inputAssignment];
                        
                    }
                }
                _data[row][4] += outputNotes;
            }
            
        }

        public void FellowNotes() 
        { 
            for (int row = 0; row < _data.Count; row++) 
            { 
                if (_fellowNotes.ContainsKey(_data[row][2].Trim()))
                {
                    
                    _data[row][4] = _fellowNotes[_data[row][2].Trim()] + "|" + _data[row][4];
                    _data[row][4] = _data[row][4].Trim('|');
                    
                }
            }
           
        }

        public void AddSlash() 
        {
            for (int row = 0; row < _data.Count; row++) 
            {
                _data[row][3] = _data[row][3].Replace('|', '/');
                _data[row][4] = _data[row][4].Replace('|', '/');

                _data[row][3] = _data[row][3].Trim('/');
                _data[row][4] = _data[row][4].Trim('/');
            }
        }

        public List<List<string>> getData() 
        {
            return _data;
        }

        


    }
}
