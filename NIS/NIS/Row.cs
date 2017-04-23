using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NIS
{
    class Row
    {
        public int Id { get; set; }
        public string Manager { get; set; }
        public string Name { get; set; }
        public List<string> Criteria { get; set; }
        public double Sum { get; set; }
    }
}
