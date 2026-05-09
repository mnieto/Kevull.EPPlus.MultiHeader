using Kevull.MultiHeader.Core.Columns;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Kevull.MultiHeader.EPPLus.Test
{
    public class ColumnInfoTest : BaseTest
    {
        [Fact]
        public void ColumnEnumeration_Dicitionary_WritesOneColumPerKey()
        {
            var xls = new ExcelPackage();
            xls.Workbook.Worksheets.Add("Enummeration");
            var sheet = xls.Workbook.Worksheets["Enummeration"];
            var writer = new EPPlusExcelWriter(xls, sheet);

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
            var properties = typeof(RiskDict).GetProperties()
                .ToDictionary(x => x.Name, x => x);

            var column = new ColumnEnumeration<Dictionary<string, int>>("Levels", data.Levels.Keys);
            column.WriteCell(writer, 2, 2, properties, data);

            Assert.Equal(10, sheet.GetValue<int>(2, 2));
            Assert.Equal(20, sheet.GetValue<int>(2, 3));
            Assert.Equal(30, sheet.GetValue<int>(2, 4));

        }

        [Fact]
        public void ColumnEnumeration_Dicitionary_ThrowsWhenNotExpectedKey()
        {
            var xls = new ExcelPackage();
            xls.Workbook.Worksheets.Add("Enummeration");
            var sheet = xls.Workbook.Worksheets["Enummeration"];
            var writer = new EPPlusExcelWriter(xls, sheet);

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
            var properties = typeof(RiskDict).GetProperties()
                .ToDictionary(x => x.Name, x => x);

            var column = new ColumnEnumeration<Dictionary<string, int>>("Levels", data.Levels.Keys.Take(2));
            Assert.Throws<KeyNotFoundException>(() => column.WriteCell(writer, 2, 2, properties, data));
        }

        [Fact]
        public void ColumnEnumeration_Enumberable_WritesOneColumPerKey()
        {
            var xls = new ExcelPackage();
            xls.Workbook.Worksheets.Add("Enummeration");
            var sheet = xls.Workbook.Worksheets["Enummeration"];
            var writer = new EPPlusExcelWriter(xls, sheet);

            var data = new RiskList
            {
                Name = "TestRisk",
                Levels = new List<int> { 10, 20, 30 }
            };
            var properties = typeof(RiskList).GetProperties()
                .ToDictionary(x => x.Name, x => x);

            var column = new ColumnEnumeration<List<int>>("Levels", data.Levels.ConvertAll(x => x.ToString()));
            column.WriteCell(writer, 2, 2, properties, data);

            Assert.Equal(10, sheet.GetValue<int>(2, 2));
            Assert.Equal(20, sheet.GetValue<int>(2, 3));
            Assert.Equal(30, sheet.GetValue<int>(2, 4));

        }

        [Fact]
        public void ColumnEnumeration_Enumerable_ThrowsWhenNotExpectedKey()
        {
            var xls = new ExcelPackage();
            xls.Workbook.Worksheets.Add("Enummeration");
            var sheet = xls.Workbook.Worksheets["Enummeration"];
            var writer = new EPPlusExcelWriter(xls, sheet);

            var data = new RiskList
            {
                Name = "TestRisk",
                Levels = new List<int> { 10, 20, 30 }
            };
            var properties = typeof(RiskList).GetProperties()
                .ToDictionary(x => x.Name, x => x);

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
