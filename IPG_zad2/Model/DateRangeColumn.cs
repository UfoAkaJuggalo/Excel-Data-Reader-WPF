using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IPG_zad2.Model
{
    public class DateRangeColumn
    {
        public DateTime DtFrom { get; set; }
        public DateTime DtTo { get; set; }
        public string Title { get; set; }
        public List<bool> EmissionsList { get; set; }
    }
}
