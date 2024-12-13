using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDbSchemaTool.Models
{
    public class SchemaModel
    {
        public string Name { get; set; }
        public List<Table> Tables { get; set; } = new();
    }
}
