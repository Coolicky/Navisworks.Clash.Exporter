using System.Linq;
using Autodesk.Navisworks.Api.Clash;

namespace Navisworks.Clash.Exporter.Methods
{

    public class CustomClashTest
    {
        public CustomClashTest(ClashTest test)
        {
            _clashTest = test;
        }

        public string DisplayName => _clashTest.DisplayName;

        private readonly ClashTest _clashTest;
        public ClashTest ClashTest => _clashTest;

        public string SelectionAName
        {
            get { return GetSelectedItem(_clashTest.SelectionA); }
        }

        public string SelectionBName
        {
            get { return GetSelectedItem(_clashTest.SelectionB); }
        }

        private string GetSelectedItem(ClashSelection selection)
        {
            string result;
            if (selection.Selection.HasSelectionSources)
            {
                result = selection.Selection.SelectionSources.FirstOrDefault()?.ToString();
                if (result.Contains("lcop_selection_set_tree\\"))
                {
                    result = result.Replace("lcop_selection_set_tree\\", "");
                }

                if (selection.Selection.SelectionSources.Count > 1)
                {
                    result += " (and other selection sets)";
                }

            }
            else if (selection.Selection.GetSelectedItems().Count == 0)
            {
                result = "No item have been selected.";
            }
            else if (selection.Selection.GetSelectedItems().Count == 1)
            {
                result = selection.Selection.GetSelectedItems().FirstOrDefault()?.DisplayName;
            }
            else
            {
                result = selection.Selection.GetSelectedItems().FirstOrDefault()?.DisplayName;
                foreach (var item in selection.Selection.GetSelectedItems().Skip(1))
                {
                    result = $"{result}; {item.DisplayName}";
                }
            }

            return result;
        }

    }
}