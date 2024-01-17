using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DaysheetGenerator_2023
{
    internal class StaffType
    {
        private List<List<string>> _attendings;
        private List<List<string>> _associates;
        private List<List<string>> _assistants;
        private List<List<string>> _fellows;
        private List<List<string>> _residents;

        public StaffType()
        {
            this._attendings = new List<List<string>>();
            this._associates = new List<List<string>>();
            this._assistants = new List<List<string>>();
            this._fellows = new List<List<string>>();
            this._residents = new List<List<string>>();
        }

        public List<List<string>> Attendings { get { return _attendings; } }
        public List<List<string>> Associates { get { return _associates; } }
        public List<List<string>> Assistants { get { return _assistants; } }
        public List<List<string>> Fellows { get { return _fellows; } }
        public List<List<string>> Residents { get { return _residents; } }



        public void StaffTypeSeparator(List<List<string>> data)
        {
            foreach (List<string> row in data)
            {
                if (row[5] == "Staff")
                {
                    _attendings.Add(row);
                }
                else if (row[5] == "Associate")
                {
                    _associates.Add(row);
                }
                else if (row[5] == "Anesthesia Assistant")
                {
                    _assistants.Add(row);
                }
                else if (row[5] == "Fellow")
                {
                    _fellows.Add(row);
                }
                else if (row[5] == "Resident")
                {
                    _residents.Add(row);
                }
            }
        }

        public void orderResidents(List<string> residentOrdering)
        {
            List<List<string>> orderedResidents = new List<List<string>>();

            foreach (string resident in  residentOrdering)
            { 
                foreach (List<string> row in _residents)
                {
                    if (resident == row[2])
                    {
                        orderedResidents.Add(row);
                    }
                }
            }
            _residents = orderedResidents;
        }

    }
}
