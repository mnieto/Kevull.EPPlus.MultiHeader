using System.Xml.Linq;
using System;
using OfficeOpenXml;
using NuGet.Frameworks;

namespace Kevull.EPPLus.MultiHeader.Test
{
    public class OneHeaderRenderTest
    {
        private int maxColumns;
        public OneHeaderRenderTest()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            maxColumns = typeof(Person).GetProperties().Length;
        }

        [Fact]
        public void Write2Rows()
        {
            var people = Person.BuildPeopleList();
            using var xls = new ExcelPackage();

            var report = new MultiHeaderReport<Person>(xls, "People");
            report.GenerateReport(people);

            var sheet = xls.Workbook.Worksheets["People"];
            //Size
            Assert.Equal(maxColumns, sheet.Dimension.End.Column);
            Assert.Equal(3, sheet.Dimension.End.Row);
            
            //Headers
            Assert.Equal(nameof(Person.NumOfComputers), sheet.Cells[1, 5].GetValue<string>());

            //Data
            Assert.Equal(Gender.Female.ToString(), sheet.GetValue<string>(3, 4));
            Assert.Null(sheet.GetValue(2, 5));
            Assert.Equal(2, sheet.GetValue<int>(3, 5));
            Assert.Equal("https://github.com/", sheet.GetValue(3, 6).ToString());
        }

        [Fact]
        public void Config_SetupOrder_ColumnsAreOrdered()
        {
            var people = Person.BuildPeopleList();
            using var xls = new ExcelPackage();
            var report = new MultiHeaderReport<Person>(xls, "People");
            report.Configure(options => options
                .AddColumn(x => x.NumOfComputers, 1)
            ).GenerateReport(people);

            var sheet = xls.Workbook.Worksheets["People"];
            Assert.Equal(nameof(Person.NumOfComputers), sheet.GetValue<string>(1, 1));
            Assert.Equal(nameof(Person.Name), sheet.GetValue<string>(1, 2));
        }

        [Fact]
        public void Config_IgnoredColumns_AreNotInTheList()
        {
            var people = Person.BuildPeopleList();
            using var xls = new ExcelPackage();
            var report = new MultiHeaderReport<Person>(xls, "People");
            report.Configure(options => options
                .AddColumn(x => x.Surname, 1)
                .IgnoreColumn(x => x.NumOfComputers)
            ).GenerateReport(people);

            var sheet = xls.Workbook.Worksheets["People"];
            Assert.Equal(maxColumns - 1, sheet.Dimension.End.Column);
        }

        [Fact]
        public void HiddenColumns_AreRendered_AsHidden()
        {
            var people = Person.BuildPeopleList();
            using var xls = new ExcelPackage();
            var report = new MultiHeaderReport<Person>(xls, "People");
            report.Configure(options => options
                .AddColumn(x => x.NumOfComputers, hidden: true)
            ).GenerateReport(people);

            var sheet = xls.Workbook.Worksheets["People"];
            Assert.True(sheet.Column(5).Hidden);
        }

        [Fact]
        public void HyperLinkColumns_UseAntherColumnTo_BuildTheLink()
        {
            var people = Person.BuildPeopleList();
            using var xls = new ExcelPackage();
            var report = new MultiHeaderReport<Person>(xls, "People");
            report.Configure(options => options
                .AddHyperLinkColumn(x => x.Name, x => x.Profile)
                .IgnoreColumn(x => x.Profile)
            ).GenerateReport(people);

            var sheet = xls.Workbook.Worksheets["People"];
            Assert.True(sheet.Cells[3, 1].Hyperlink != null);
        }

        [Fact]
        public void FormulaColumns_Write_Formulas()
        {
            var people = Person.BuildPeopleList();
            using var xls = new ExcelPackage();

            var report = new MultiHeaderReport<Person>(xls, "People");
            report.Configure(options => options
                .AddColumn(x => x.Name, 1)
                .AddColumn(x => x.Surname, 2)
                .AddFormula("CompleteName", "CONCATENATE(B2,\", \",A2)", 3)
            ).GenerateReport(people);

            var sheet = xls.Workbook.Worksheets["People"];
            Assert.Equal("Bateson, Aimée", sheet.GetValue<string>(3, 3));
        }

        [Fact]
        public void ExpressionColumns_Write_ExpressionResults()
        {
            var people = Person.BuildPeopleList();
            using var xls = new ExcelPackage();

            var report = new MultiHeaderReport<Person>(xls, "People");
            report.Configure(options => options
                .AddColumn(x => x.Name, 1)
                .AddColumn(x => x.Surname, 2)
                .AddExpression("Initials", x => string.Concat(x.Name[0], '.', x.Surname[0], '.'), 3)
            ).GenerateReport(people);

            var sheet = xls.Workbook.Worksheets["People"];
            Assert.Equal("A.B.", sheet.GetValue<string>(3, 3));
        }

    }
}