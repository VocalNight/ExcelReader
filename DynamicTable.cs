using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReader
{
    public class DynamicTable
    {
        public int Id {  get; set; }

        [NotMapped]
        public Dictionary<string, object> DynamicProperties { get; set; }
    }
}
