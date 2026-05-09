using Kevull.MultiHeader.Core.Columns;

namespace Kevull.MultiHeader.ClosedXml.Test
{
    public class ColumnInfoTest
    {
        [Fact]
        public void ColumnEnumeration_Dicitionary_WritesOneColumPerKey()
        {
            using var workbook = new ClosedXML.Excel.XLWorkbook();
            var sheet = workbook.Worksheets.Add("Enummeration");
            var writer = new ClosedXmlExcelWriter(workbook, sheet);

            var data = new RiskDict
            {
                Name = "TestRisk",
                Levels = new Dictionary<string, int>
                {
                    { "LOW", 10 },
                    { "MED", 20 },
                    { "HIGh", 30 }
                }
            };
            var properties = typeof(RiskDict).GetProperties().ToDictionary(x => x.Name, x => x);

            var column = new ColumnEnumeration<Dictionary<string, int>>("Levels", data.Levels.Keys);
            column.WriteCell(writer, 2, 2, properties, data);

            Assert.Equal(10, sheet.Cell(2, 2).GetValue<int>());
            Assert.Equal(20, sheet.Cell(2, 3).GetValue<int>());
            Assert.Equal(30, sheet.Cell(2, 4).GetValue<int>());
        }

        [Fact]
        public void ColumnEnumeration_Dicitionary_ThrowsWhenNotExpectedKey()
        {
            using var workbook = new ClosedXML.Excel.XLWorkbook();
            var sheet = workbook.Worksheets.Add("Enummeration");
            var writer = new ClosedXmlExcelWriter(workbook, sheet);

            var data = new RiskDict
            {
                Name = "TestRisk",
                Levels = new Dictionary<string, int>
                {
                    { "LOW", 10 },
                    { "MED", 20 },
                    { "HIGh", 30 }
                }
            };
            var properties = typeof(RiskDict).GetProperties().ToDictionary(x => x.Name, x => x);

            var column = new ColumnEnumeration<Dictionary<string, int>>("Levels", data.Levels.Keys.Take(2));
            Assert.Throws<KeyNotFoundException>(() => column.WriteCell(writer, 2, 2, properties, data));
        }

        [Fact]
        public void ColumnEnumeration_Enumberable_WritesOneColumPerKey()
        {
            using var workbook = new ClosedXML.Excel.XLWorkbook();
            var sheet = workbook.Worksheets.Add("Enummeration");
            var writer = new ClosedXmlExcelWriter(workbook, sheet);

            var data = new RiskList
            {
                Name = "TestRisk",
                Levels = new List<int> { 10, 20, 30 }
            };
            var properties = typeof(RiskList).GetProperties().ToDictionary(x => x.Name, x => x);

            var column = new ColumnEnumeration<List<int>>("Levels", data.Levels.ConvertAll(x => x.ToString()));
            column.WriteCell(writer, 2, 2, properties, data);

            Assert.Equal(10, sheet.Cell(2, 2).GetValue<int>());
            Assert.Equal(20, sheet.Cell(2, 3).GetValue<int>());
            Assert.Equal(30, sheet.Cell(2, 4).GetValue<int>());
        }

        [Fact]
        public void ColumnEnumeration_Enumerable_ThrowsWhenNotExpectedKey()
        {
            using var workbook = new ClosedXML.Excel.XLWorkbook();
            var sheet = workbook.Worksheets.Add("Enummeration");
            var writer = new ClosedXmlExcelWriter(workbook, sheet);

            var data = new RiskList
            {
                Name = "TestRisk",
                Levels = new List<int> { 10, 20, 30 }
            };
            var properties = typeof(RiskList).GetProperties().ToDictionary(x => x.Name, x => x);

            var column = new ColumnEnumeration<List<int>>("Levels", data.Levels.ConvertAll(x => x.ToString()).Take(2));
            Assert.Throws<KeyNotFoundException>(() => column.WriteCell(writer, 2, 2, properties, data));
        }
    }

    internal class RiskDict
    {
        public string Name { get; set; } = "";
        public Dictionary<string, int> Levels { get; set; } = new Dictionary<string, int>();
    }

    internal class RiskList
    {
        public string Name { get; set; } = "";
        public List<int> Levels { get; set; } = new List<int>();
    }
}
