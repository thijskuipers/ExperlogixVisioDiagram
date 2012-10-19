using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Broes.Experlogix.DAL
{
    public static class Bootstrapper
    {
        public static void Bootstrap()
        {
            AutoMapper.Mapper.CreateMap<Jet.ExperlogixDataSet.SeriesRow, Entities.Series>();
            AutoMapper.Mapper.CreateMap<Jet.ExperlogixDataSet.ModelRow, Entities.Model>();
            AutoMapper.Mapper.CreateMap<Jet.ExperlogixDataSet.CategoryRow, Entities.Category>();
            AutoMapper.Mapper.CreateMap<Jet.ExperlogixDataSet.ListRow, Entities.List>();
            AutoMapper.Mapper.CreateMap<Jet.ExperlogixDataSet.LookupRow, Entities.Lookup>();
            AutoMapper.Mapper.CreateMap<Jet.ExperlogixDataSet.CategoryAttLookupRow, Entities.CategoryAttLookup>();
            AutoMapper.Mapper.CreateMap<Jet.ExperlogixDataSet.FormulaRow, Entities.Formula>()
                .ForMember(f => f.Formula1, opt => opt.MapFrom(f => f.Formula));
            AutoMapper.Mapper.CreateMap<Jet.ExperlogixDataSet.RuleRow, Entities.Rule>();
            AutoMapper.Mapper.CreateMap<Jet.ExperlogixDataSet.CategoryAttributeRow, Entities.CategoryAttribute>();
        }
    }
}
