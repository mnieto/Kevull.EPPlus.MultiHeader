using EPPLus.MultiHeader.Columns;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace EPPLus.MultiHeader.Test
{
    public class ColumnInfoTest
    {

        public ColumnInfoTest()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        [Fact]
        public void Order_MustBeOneOrUpper()
        {
            var property = typeof(Person).GetProperties().First(x => x.Name == nameof(Person.Name));
            var sut = new ColumnInfo(nameof(Person.Name));
            Action act = () => sut.Order = 0;
            Assert.Throws<ArgumentOutOfRangeException>(act);
        }

        [Fact]
        public void DisplayName_IsName_IfNotAssigned()
        {
            var property = typeof(Person).GetProperties().First(x => x.Name == nameof(Person.BirthDate));
            var sut = new ColumnInfo(nameof(Person.BirthDate));
            Assert.Equal(sut.Name, sut.DisplayName);
        }

        [Fact]
        public void Deep_InDirectProprties_IsOne()
        {
            var sut = new ColumnInfo<RootLevelDictionary>(x => x.SimpleProperty);
            Assert.Equal(1, sut.Deep);
        }

        [Fact]
        public void Deep_InDirectChildProperties_IsTwo()
        {
            var sut = new ColumnInfo<RootLevelDictionary>(x => x.ComplexProperty.RightColumn);
            Assert.Equal(2, sut.Deep);
        }

        [Fact]
        public void ColumnEnumeration_Dicitionary_WritesOneColumPerKey()
        {
            var xls = new ExcelPackage();
            xls.Workbook.Worksheets.Add("Enummeration");
            var sheet = xls.Workbook.Worksheets["Enummeration"];
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

            var column = new ColumnEnumeration("Levels", data.Levels.Keys);
            column.WriteCell(sheet.Cells["B2"], properties, data);

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

            var column = new ColumnEnumeration("Levels", data.Levels.Keys.Take(2));
            Assert.Throws<KeyNotFoundException>(() => column.WriteCell(sheet.Cells["B2"], properties, data));
        }

        [Fact]
        public void ColumnEnumeration_Enumberable_WritesOneColumPerKey()
        {
            var xls = new ExcelPackage();
            xls.Workbook.Worksheets.Add("Enummeration");
            var sheet = xls.Workbook.Worksheets["Enummeration"];
            var data = new RiskList
            {
                Name = "TestRisk",
                Levels = new List<int> { 10, 20, 30 }
            };
            var properties = typeof(RiskList).GetProperties()
                .ToDictionary(x => x.Name, x => x);

            var column = new ColumnEnumeration("Levels", data.Levels.ConvertAll(x => x.ToString()));
            column.WriteCell(sheet.Cells["B2"], properties, data);

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
            var data = new RiskList
            {
                Name = "TestRisk",
                Levels = new List<int> { 10, 20, 30 }
            };
            var properties = typeof(RiskList).GetProperties()
                .ToDictionary(x => x.Name, x => x);

            var column = new ColumnEnumeration("Levels", data.Levels.ConvertAll(x => x.ToString()).Take(2));
            Assert.Throws<KeyNotFoundException>(() => column.WriteCell(sheet.Cells["B2"], properties, data));
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
