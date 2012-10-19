using Broes.Experlogix.DAL;
using Broes.Experlogix.DAL.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Visio = Microsoft.Office.Interop.Visio;

namespace Broes.Experlogix.VisioDiagram
{
    class Program
    {
        private const double X_POS1 = 0;
        private const double Y_POS1 = 0;
        private const double X_POS2 = 1.75;
        private const double Y_POS2 = 0.6;
        const short VISIO_SECTION_OJBECT_INDEX = 1;
        const short NAME_CHARACTER_SIZE = 10; // 10pt

        const double SHDW_PATTERN = 0; // None
        const double BEGIN_ARROW = 4; // Filled arrow
        const double BEGIN_ARROW_NONE = 0; // None
        const double END_ARROW = 4; // Filled arrow
        const double LINE_COLOR_MANY = 10;
        const double LINE_COLOR = 8; // Black
        const double LINE_PATTERN = 1; // ______ solid line
        const double LINE_PATTERN_ERROR = 2; // _ _ _ _ dashed lined
        const double LINE_PATTERN_HIDE = 3; // . . . . dotted line
        const double LINE_PATTERN_MATCH = 4; // _ . _ . _ dash/dot-(t)ed line
        const string LINE_WEIGHT = "2pt";
        const double ROUNDING = 0.0625;
        const double HEIGHT = 0.25;

        const string CATEGORY_PREFIX = "CAT_";
        const string FORMULA_PREFIX = "FOR_";
        const string RULE_PREFIX = "RUL_";
        const string LIST_PREFIX = "LIS_";
        const string LOOKUP_PREFIX = "LKP_";

        const string COLOR_GREEN_LIGHT = "RGB(204,255,204)";
        const string COLOR_ORANGE_LIGHT = "RGB(255,202,176)";
        const string COLOR_YELLOW_LIGHT = "RGB(255,255,176)";
        const string COLOR_BLUE_LIGHT = "RGB(204,204,255)";
        const string COLOR_PINK_LIGHT = "RGB(255,204,255)";
        const string COLOR_WHITE = "RGB(255,255,255)";
        const string COLOR_BLACK = "RGB(0,0,0)";

        /// <summary>
        /// Contains Tuple&lt;string, string> of to --> from shape names.
        /// </summary>
        private static readonly HashSet<Tuple<string, string, FormulaUse>> relations = new HashSet<Tuple<string, string, FormulaUse>>();
        private static readonly HashSet<string> usedFormulaNames = new HashSet<string>();
        private static readonly HashSet<string> usedListNames = new HashSet<string>();
        private static readonly HashSet<string> usedLookupNames = new HashSet<string>();

        private static readonly Regex matchCategoryAttributeRegex = new Regex(@"\[([a-z_]+)\.([a-z_]+)\]", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        private static readonly Regex matchoptionalCategoryAttributeRegex = new Regex(@"\[([a-z]+)?\.([a-z_]+)\]", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        private static readonly Regex rulePremiseRegex = new Regex(@"([crf])\:([a-z]+)", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

        private static readonly ExperlogixRepository repository = new ExperlogixRepository();

        static void Main(string[] args)
        {
            Visio.Application application = new Visio.Application();
            application.Visible = false; // Hiding the application speeds up rendering.
            application.AutoLayout = false; // Delay autolayout until after putting all shapes on page.

            Visio.Document document = application.Documents.Add(string.Empty);

            try
            {
                Bootstrapper.Bootstrap();

                List<Category> categories;
                List<Formula> formulae;
                List<Rule> rules;
                List<List> lists;
                List<Lookup> lookups;

                var series = repository.RetrieveSeries();

                int i = 0;
                for (; i < series.Count; i++)
                {
                    Console.WriteLine("[{0}] {1}", i + 1, series[i].Description);
                }
                Console.Write("Choose the series [1 - {0}]: ", i);

                int seriesIndex = 0;
                string seriesInput = Console.ReadLine();
                Console.WriteLine();
                Console.WriteLine();

                if (int.TryParse(seriesInput, out seriesIndex) && seriesIndex >= 0 && seriesIndex <= i)
                {
                    Series chosenSerie = series[seriesIndex - 1];

                    var models = repository.RetrieveModelsBySeriesID(chosenSerie.SeriesID);
                    int j = 0;
                    for (; j < models.Count; j++)
                    {
                        Console.WriteLine("[{0}] {1}", j + 1, models[j].Description);
                    }
                    Console.Write("Choose the model [1 - {0}]: ", j);

                    int modelIndex = 0;
                    string modelInput = Console.ReadLine();
                    Console.WriteLine();
                    Console.WriteLine();

                    if (int.TryParse(modelInput, out modelIndex) && modelIndex >= 0 && modelIndex <= j)
                    {
                        Model chosenModel = models[modelIndex - 1];
                        categories = repository.RetrieveCategoriesBySeriesID(chosenSerie.SeriesID);

                        formulae = repository.RetrieveFormulas();

                        rules = repository.RetrieveRulesByModelID(chosenModel.ModelID);

                        lists = repository.RetrieveLists();
                        lookups = repository.RetrieveLookupTables();

                        // Get the default page of our new document
                        Visio.Page page = document.Pages[1];
                        page.Name = "Experlogix";
                        
                        // Set Placement Spacing to "25mm"
                        page.PageSheet.get_CellsU("AvenueSizeY").FormulaU = "25 mm";
                        page.PageSheet.get_CellsU("AvenueSizeX").FormulaU = "25 mm";
                        
                        // Set Connector Appearance to "Curved"
                        page.PageSheet.get_CellsU("LineRouteExt").ResultIU = 2;

                        BuildDiagram(page, categories, formulae, rules, lists, lookups);
                    }
                    else
                    {
                        Console.WriteLine("Unknown model.");
                    }
                }
                else
                {
                    Console.WriteLine("Unknown series.");
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("An Exception occurred: {0}", e);
            }
            // The following is in a finally block to make sure the application is shown,
            // even when an exception is thrown.
            finally
            {
                // Set autolayout to true to update the layout.
                application.AutoLayout = true;

                // Show the application.
                application.Visible = true;
            }

            Console.WriteLine();
            Console.WriteLine("Press enter to exit...");
            Console.ReadLine();
        }


        private static void BuildDiagram(Visio.Page page, List<Category> categories, List<Formula> formulae, List<Rule> rules, List<List> lists, List<Lookup> lookups)
        {
            Console.WriteLine("Drawing {0} categories...", categories.Count);
            DrawCategories(page, categories);

            Console.WriteLine("Drawing {0} rules...", rules.Count);
            DrawRules(page, rules);

            var listsToDraw = lists.Where(l => usedListNames.Contains(l.ListName.ToUpperInvariant()));
            Console.WriteLine("Drawing {0} lists...", listsToDraw.Count());
            DrawLists(page, listsToDraw);

            var lookupsToDraw = lookups.Where(l => usedLookupNames.Contains(l.TableName.ToUpperInvariant()));
            Console.WriteLine("Drawing {0} lookup tables...", lookupsToDraw.Count());
            DrawLookups(page, lookupsToDraw);

            var formulaeToDraw = formulae.Where(f => usedFormulaNames.Contains(f.FormulaName.ToUpperInvariant()));
            Console.WriteLine("Drawing {0} formulas...", formulaeToDraw.Count());
            DrawFormulae(page, formulaeToDraw);

            Console.WriteLine("Drawing {0} relations...", relations.Count);
            DrawRelations(page);

            Console.WriteLine();
            Console.WriteLine("Laying out the page...");
            page.Layout();

            Console.WriteLine("Resizing to fit to contents...");
            page.ResizeToFitContents();
        }

        private static void DrawRules(Visio.Page page, IEnumerable<Rule> rules)
        {
            foreach (Rule rule in rules)
            {
                // Create a Visio rectangle shape.
                Visio.Shape rect;

                try
                {
                    // There is no "Get Try", so we have to rely on an exception to tell us it does not exists
                    rect = page.Shapes.get_ItemU(RULE_PREFIX + rule.RuleID);
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    rect = DrawRule(page, rule);
                }
            }
        }

        private static void DrawCategories(Visio.Page page, IEnumerable<Category> categories)
        {
            foreach (Category category in categories)
            {
                // Create a Visio rectangle shape.
                Visio.Shape rect;

                try
                {
                    // There is no "Get Try", so we have to rely on an exception to tell us it does not exists
                    rect = page.Shapes.get_ItemU(CATEGORY_PREFIX + category.CatID);
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    rect = DrawCategory(page, category);
                }
            }
        }

        private static void DrawFormulae(Visio.Page page, IEnumerable<Formula> formulae)
        {
            // Get the metadata for each passed-in entity, draw it, and draw its relationships.
            foreach (Formula formula in formulae)
            {
                // Create a Visio rectangle shape.
                Visio.Shape rect;

                try
                {
                    // There is no "Get Try", so we have to rely on an exception to tell us it does not exists
                    rect = page.Shapes.get_ItemU(FORMULA_PREFIX + formula.FormulaName);
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    rect = DrawFormula(page, formula);
                }
            }
        }

        private static void DrawLists(Visio.Page page, IEnumerable<List> lists)
        {
            foreach (List list in lists)
            {
                try
                {
                    page.Shapes.get_ItemU(LIST_PREFIX + list.ListName);
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    DrawList(page, list);
                }
            }
        }

        private static void DrawLookups(Visio.Page page, IEnumerable<Lookup> lookups)
        {
            foreach (Lookup lookup in lookups)
            {
                try
                {
                    page.Shapes.get_ItemU(LOOKUP_PREFIX + lookup.TableName);
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    DrawLookup(page, lookup);
                }
            }
        }

        private static Visio.Shape DrawFormula(Visio.Page page, Formula formula)
        {
            MatchCollection matches = matchoptionalCategoryAttributeRegex.Matches(formula.Formula1);

            string text = string.Format("{0} ({1})", formula.FormulaName, formula.FormulaType);
            StringBuilder sb = new StringBuilder(text);
            sb.AppendLine();

            SortedSet<string> matchSet = new SortedSet<string>();

            foreach (Match match in matches)
            {
                if (match.Groups[1].Length > 0)
                {
                    relations.Add(Tuple.Create(FORMULA_PREFIX + formula.FormulaName, CATEGORY_PREFIX + match.Groups[1].Value, FormulaUse.Value));
                }
                matchSet.Add(match.Value);
            }

            foreach (string attribute in matchSet)
            {
                sb.AppendLine(attribute);
            }
            
            string fullText = sb.ToString().TrimEnd();

            string fillColor = COLOR_ORANGE_LIGHT;

            Visio.Shape rect = DrawRectangle(page, FORMULA_PREFIX + formula.FormulaName, text.Length, fullText, fillColor);

            return rect;
        }

        private static Visio.Shape DrawCategory(Visio.Page page, Category category)
        {
            string name = CATEGORY_PREFIX + category.CatID;
            string text = (category.Description ?? category.CatID);
            StringBuilder sb = new StringBuilder(text);
            sb.AppendLine();

            foreach (CategoryAttribute attribute in repository.RetrieveAttributesByCategoryID(category.CatID))
            {
                sb.AppendFormat("{0} ({1}) {2}", attribute.AttributeName, attribute.TypeData, attribute.Source);
                sb.AppendLine();

                if (attribute.Source == "Formula" && !string.IsNullOrEmpty(attribute.VariableOne))
                {
                    relations.Add(Tuple.Create(name, FORMULA_PREFIX + attribute.VariableOne, FormulaUse.Value));
                    usedFormulaNames.Add(attribute.VariableOne.ToUpperInvariant());
                }

                if (attribute.Source == "Match" && !string.IsNullOrEmpty(attribute.VariableOne))
                {
                    var match = matchCategoryAttributeRegex.Match(attribute.VariableOne);
                    if (match.Success)
                    {
                        string matchingCategory = match.Groups[1].Value;
                        if (matchingCategory != category.CatID)
                        {
                            relations.Add(Tuple.Create(name, CATEGORY_PREFIX + matchingCategory, FormulaUse.Match));                            
                        }
                    }
                }

                if (attribute.Source == "List" && !string.IsNullOrEmpty(attribute.VariableOne))
                {
                    relations.Add(Tuple.Create(name, LIST_PREFIX + attribute.VariableOne, FormulaUse.Value));
                    usedListNames.Add(attribute.VariableOne.ToUpperInvariant());
                }

                if (attribute.Source == "Lookup" && !string.IsNullOrEmpty(attribute.VariableOne))
                {
                    var lookupRelations = attribute.CategoryAttLookups.ToList();

                    foreach (CategoryAttLookup lookupRelation in lookupRelations)
                    {
                        Match match = matchCategoryAttributeRegex.Match(lookupRelation.IndexAttribute);
                        if (match.Success)
                        {
                            relations.Add(Tuple.Create(name, CATEGORY_PREFIX + match.Groups[1].Value, FormulaUse.Value));
                        }
                    }

                    relations.Add(Tuple.Create(name, LOOKUP_PREFIX + attribute.VariableOne, FormulaUse.Value));
                    usedLookupNames.Add(attribute.VariableOne.ToUpperInvariant());
                }


                if (!string.IsNullOrEmpty(attribute.ErrorFormula))
                {
                    relations.Add(Tuple.Create(name, FORMULA_PREFIX + attribute.ErrorFormula, FormulaUse.Error));
                    usedFormulaNames.Add(attribute.ErrorFormula.ToUpperInvariant());
                }

                if (!string.IsNullOrEmpty(attribute.HideFormula))
                {
                    relations.Add(Tuple.Create(name, FORMULA_PREFIX + attribute.HideFormula, FormulaUse.Hide));
                    usedFormulaNames.Add(attribute.HideFormula.ToUpperInvariant());
                }
            }

            string fullText = sb.ToString().TrimEnd();

            // Determine the shape fill color based on category linetype.
            string fillColor = string.Empty;
            switch (category.MRPCatType)
            {
                case "L":
                    fillColor = COLOR_YELLOW_LIGHT; // Light yellow
                    break;
                default:
                    fillColor = COLOR_WHITE; // White
                    break;
            }

            Visio.Shape rect = DrawRectangle(page, CATEGORY_PREFIX + category.CatID, text.Length, fullText, fillColor);

            return rect;
        }

        private static Visio.Shape DrawRule(Visio.Page page, Rule rule)
        {
            string ruleId = RULE_PREFIX + rule.RuleID;

            if (null != rule.Premise)
            {
                MatchCollection premiseMatches = rulePremiseRegex.Matches(rule.Premise);

                foreach (Match premiseMatch in premiseMatches)
                {
                    switch (premiseMatch.Groups[1].Value.ToUpperInvariant())
                    {
                        case "C": // Category
                        case "R": // Ruleflag (points to category)
                            relations.Add(Tuple.Create(ruleId, CATEGORY_PREFIX + premiseMatch.Groups[2].Value, FormulaUse.Value));
                            break;
                        case "F": // Formula
                            relations.Add(Tuple.Create(ruleId, FORMULA_PREFIX + premiseMatch.Groups[2].Value, FormulaUse.Value));
                            usedFormulaNames.Add(premiseMatch.Groups[2].Value.ToUpperInvariant());
                            break;
                    }
                }
            }

            if (null != rule.Conclusion)
            {
                MatchCollection conclusionMatches = rulePremiseRegex.Matches(rule.Conclusion);

                foreach (Match conclusionMatch in conclusionMatches)
                {
                    switch (conclusionMatch.Groups[1].Value.ToUpperInvariant())
                    {
                        case "C": // Category
                        case "R": // Ruleflag (points to category)
                            relations.Add(Tuple.Create(ruleId, CATEGORY_PREFIX + conclusionMatch.Groups[2].Value, FormulaUse.Value));
                            break;
                        case "F": // Function
                            relations.Add(Tuple.Create(ruleId, FORMULA_PREFIX + conclusionMatch.Groups[2].Value, FormulaUse.Value));
                            break;
                    }
                }
            }

            string title = string.Format("{0} ({1})", rule.RuleID, rule.Type);

            return DrawRectangle(page, ruleId, title.Length, title, COLOR_GREEN_LIGHT); // Light green
        }

        private static Visio.Shape DrawList(Visio.Page page, List list)
        {
            string name = LIST_PREFIX + list.ListName;

            if (!string.IsNullOrEmpty(list.DependsOn) && list.DependsOn.IndexOf("No Dependency", StringComparison.OrdinalIgnoreCase) < 0)
            {
                if (list.DependsOn.IndexOf('[') >= 0) // Reference to a category
                {
                    Match match = matchCategoryAttributeRegex.Match(list.DependsOn);

                    if (match.Success)
                    {
                        relations.Add(Tuple.Create(name, CATEGORY_PREFIX + match.Groups[1].Value, FormulaUse.Value));
                    }
                }
                else // Reference to another list
                {
                    relations.Add(Tuple.Create(name, LIST_PREFIX + list.DependsOn, FormulaUse.Value));
                }
            }

            return DrawRectangle(page, name, list.ListName.Length, list.ListName, COLOR_BLUE_LIGHT);
        }

        private static Visio.Shape DrawLookup(Visio.Page page, Lookup lookup)
        {
            string name = LOOKUP_PREFIX + lookup.TableName;
            return DrawRectangle(page, name, lookup.TableName.Length, lookup.TableName, COLOR_PINK_LIGHT);
        }

        private static Visio.Shape DrawRectangle(Visio.Page page, string name, int titleLength, string text, string fillColor)
        {
            Visio.Shape rect = page.DrawRectangle(X_POS1, Y_POS1, X_POS2, Y_POS2);
            rect.Name = name.ToUpperInvariant();
            rect.Text = text;

            // Set the fill color, placement properties, and line weight of the shape.
            rect.get_CellsU("ObjType").FormulaU = ((int)Visio.VisCellVals.visLOFlagsPlacable).ToString();
            rect.get_CellsU("FillForegnd").FormulaU = fillColor;
            rect.get_CellsU("Width").FormulaU = "TEXTWIDTH(TheText,100mm)";
            rect.get_CellsU("Height").FormulaU = "TEXTHEIGHT(TheText,100mm)";

            // Update the style of the entity name
            Visio.Characters characters = rect.Characters;
            characters.set_CharProps((short)Visio.VisCellIndices.visCharacterSize, NAME_CHARACTER_SIZE);
            characters.set_ParaProps((short)Visio.VisCellIndices.visHorzAlign, (short)Visio.VisCellVals.visHorzLeft);

            Visio.Characters titleChars = rect.Characters;
            titleChars.End = titleLength;
            titleChars.set_CharProps((short)Visio.VisCellIndices.visCharacterStyle, (short)Visio.VisCellVals.visBold);
            titleChars.set_CharProps((short)Visio.VisCellIndices.visCharacterColor, (short)Visio.VisDefaultColors.visDarkBlue);
            titleChars.set_ParaProps((short)Visio.VisCellIndices.visHorzAlign, (short)Visio.VisCellVals.visHorzCenter);
            
            return rect;
        }

        private static Visio.Shape DrawDirectionalDynamicConnector(Visio.Shape shapeFrom, Visio.Shape shapeTo, FormulaUse formulaType)
        {
            // Add a dynamic connector to the page.
            Visio.Shape connectorShape = shapeFrom.ContainingPage.Drop(shapeFrom.Application.ConnectorToolDataObject, 0.0, 0.0);

            // Set the connector properties, using different arrows, colors, and patterns for many-to-many relationships.
            connectorShape.get_CellsU("ShdwPattern").ResultIU = SHDW_PATTERN;
            connectorShape.get_CellsU("BeginArrow").ResultIU = BEGIN_ARROW_NONE;
            connectorShape.get_CellsU("EndArrow").ResultIU = END_ARROW;
            connectorShape.get_CellsU("LineColor").FormulaU = COLOR_BLACK;
            var linePatternCell = connectorShape.get_CellsU("LinePattern");
            switch (formulaType)
            {
                case FormulaUse.Error:
                    linePatternCell.ResultIU = LINE_PATTERN_ERROR;
                    break;
                case FormulaUse.Hide:
                    linePatternCell.ResultIU = LINE_PATTERN_HIDE;
                    break;
                case FormulaUse.Match:
                    linePatternCell.ResultIU = LINE_PATTERN_MATCH;
                    break;
                case FormulaUse.Value:
                default:
                    linePatternCell.ResultIU = LINE_PATTERN;
                    break;
            }
            connectorShape.get_CellsU("Rounding").ResultIU = ROUNDING;

            // Connect the starting point.
            Visio.Cell cellBeginX = connectorShape.get_CellsU("BeginX");
            cellBeginX.GlueTo(shapeFrom.get_CellsU("PinX"));

            // Connect the ending point.
            Visio.Cell cellEndX = connectorShape.get_CellsU("EndX");
            cellEndX.GlueTo(shapeTo.get_CellsU("PinX"));

            return connectorShape;
        }

        private static void DrawRelations(Visio.Page page)
        {
            foreach (var relation in relations)
            {
                Visio.Shape fromShape;
                Visio.Shape toShape;

                try
                {
                    // There is no "Get Try", so we have to rely on an exception to tell us it does not exists
                    fromShape = page.Shapes.get_ItemU(relation.Item1.ToUpperInvariant());
                    toShape = page.Shapes.get_ItemU(relation.Item2.ToUpperInvariant());
                    DrawDirectionalDynamicConnector(fromShape, toShape, relation.Item3);
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    Console.WriteLine("* Unable to draw from {0} to {1}", relation.Item1, relation.Item2);
                }
            }
        }

        enum FormulaUse
        {
            Value,
            Error,
            Hide,
            Match
        }
    }
}
