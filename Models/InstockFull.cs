using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InstockAutoMailService.Models
{
    public class InstockFull
    {
        public string Site { get; set; }
        public string Project { get; set; }
        public string Group { get; set; }
        public string ModelCode { get; set; }
        public int QTY { get; set; }

    }
}
