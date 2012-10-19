using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Broes.Experlogix.DAL.Entities
{
    public class CategoryAttribute
    {
        public string AttributeName { get; set; }
        public string Source { get; set; }
        public string VariableOne { get; set; }
        public string HideFormula { get; set; }
        public string ErrorFormula { get; set; }
        public string TypeData { get; set; }
        public int? AttSeqNo { get; set; }

        public List<CategoryAttLookup> CategoryAttLookups { get; set; }
    }
}
