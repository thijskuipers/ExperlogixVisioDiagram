using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Broes.Experlogix.DAL.Entities
{
    public class Rule
    {
        public string RuleID { get; set; }
        public string Type { get; set; }
        public string Conclusion { get; set; }
        public string Premise { get; set; }
    }
}
