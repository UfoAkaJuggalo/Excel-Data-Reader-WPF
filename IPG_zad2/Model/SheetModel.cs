using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IPG_zad2.Model
{
    public class SheetModel
    {
        public List<int> Id { get; set; }
        public List<string> Name { get; set; }
        public List<int> Price { get; set; }
        public List<int> Position { get; set; }
        public List<string> Level { get; set; }
        public List<string> Description { get; set; }
        public List<string> Order { get; set; }
        public List<DateRangeColumn> EmissionDatesList { get; set; }

        public SheetModel()
        {
            EmissionDatesList = new List<DateRangeColumn>();
        }
    }
}
