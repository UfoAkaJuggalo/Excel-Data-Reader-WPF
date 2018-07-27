using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IPG_zad2.Model
{
    public class FileDataModel
    {
        public List<SheetModel> SheetList { get; set; }

        public FileDataModel()
        {
            SheetList = new List<SheetModel>();
        }
    }
}
