using Broes.Experlogix.DAL.Entities;
using Broes.Experlogix.DAL.Jet;
using Broes.Experlogix.DAL.Jet.ExperlogixDataSetTableAdapters;
using System.Collections.Generic;
using System.Linq;

namespace Broes.Experlogix.DAL
{
    public class ExperlogixRepository
    {
        private SeriesTableAdapter _seriesAdapter;
        private ModelTableAdapter _modelAdapter;
        private CategoryTableAdapter _categoryAdapter;
        private ListTableAdapter _listAdapter;
        private LookupTableAdapter _lookupAdapter;
        private RuleTableAdapter _ruleAdapter;
        private CategoryAttributeTableAdapter _attributeAdapter;
        private FormulaTableAdapter _formulaAdapter;
        private CategoryAttLookupTableAdapter _attributeLookupAdapter;

        public ExperlogixRepository()
        {
            _seriesAdapter = new SeriesTableAdapter();
            _modelAdapter = new ModelTableAdapter();
            _categoryAdapter = new CategoryTableAdapter();
            _listAdapter = new ListTableAdapter();
            _lookupAdapter = new LookupTableAdapter();
            _ruleAdapter = new RuleTableAdapter();
            _attributeAdapter = new CategoryAttributeTableAdapter();
            _formulaAdapter = new FormulaTableAdapter();
            _attributeLookupAdapter = new CategoryAttLookupTableAdapter();
        }

        public List<Series> RetrieveSeries()
        {
            return AutoMapper.Mapper.Map<List<Series>>(_seriesAdapter.GetData());
        }

        public List<Model> RetrieveModelsBySeriesID(string seriesID)
        {
            return AutoMapper.Mapper.Map<List<Model>>(_modelAdapter.GetModelsBySeriesID(seriesID));
        }

        public List<Category> RetrieveCategoriesBySeriesID(string seriesID)
        {
            return AutoMapper.Mapper.Map<List<Category>>(_categoryAdapter.GetCategoriesBySeriesID(seriesID));
        }

        public List<List> RetrieveLists()
        {
            return AutoMapper.Mapper.Map<List<List>>(_listAdapter.GetData());
        }

        public List<Lookup> RetrieveLookupTables()
        {
            return AutoMapper.Mapper.Map<List<Lookup>>(_lookupAdapter.GetData());
        }

        public List<Rule> RetrieveRulesByModelID(string modelID)
        {
            return AutoMapper.Mapper.Map<List<Rule>>(_ruleAdapter.GetRulesByModelID(modelID));
        }

        public List<CategoryAttribute> RetrieveAttributesByCategoryID(string categoryID)
        {
            ExperlogixDataSet.CategoryAttributeDataTable attributeTable = _attributeAdapter.GetCategoryAttributesByCategoryID(categoryID);

            foreach (var attributeRow in attributeTable)
            {
                attributeRow.CategoryAttLookups = _attributeLookupAdapter
                    .GetCategoryAttLookupsByCategoryAttributeID(attributeRow.CatID, attributeRow.AttributeName)
                    .Select().Cast<ExperlogixDataSet.CategoryAttLookupRow>().ToArray();
            }

            return AutoMapper.Mapper.Map<List<CategoryAttribute>>(attributeTable);
        }

        public List<Formula> RetrieveFormulas()
        {
            return AutoMapper.Mapper.Map<List<Formula>>(_formulaAdapter.GetData());
        }
    }
}
