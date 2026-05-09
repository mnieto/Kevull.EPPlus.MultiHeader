using Kevull.MultiHeader.TestCommon;

namespace Kevull.MultiHeader.ClosedXml.Test
{
    public class OneHeaderRenderTest : BaseTest
    {
        private readonly int _maxColumns;

        public OneHeaderRenderTest()
        {
            _maxColumns = typeof(Person).GetProperties().Length;
        }

        [Fact]
        public void Write2Rows()
        {
            var people = Person.BuildPeopleList();
            using var workbook = new ClosedXML.Excel.XLWorkbook();

            var report = new MultiHeaderReport<Person>(workbook, "People");
            report.GenerateReport(people);

            var sheet = workbook.Worksheet("People");
            Assert.Equal(_maxColumns, sheet.LastColumnUsed()!.ColumnNumber());
            Assert.Equal(3, sheet.LastRowUsed()!.RowNumber());

            Assert.Equal(nameof(Person.NumOfComputers), sheet.Cell(1, 5).GetValue<string>());
            Assert.Equal(Gender.Female.ToString(), sheet.Cell(3, 4).GetValue<string>());
            Assert.True(sheet.Cell(2, 5).IsEmpty());
            Assert.Equal(2, sheet.Cell(3, 5).GetValue<int>());
            Assert.Equal("https://github.com/", sheet.Cell(3, 6).GetValue<string>());
        }

        [Fact]
        public void Config_SetupOrder_ColumnsAreOrdered()
        {
            var people = Person.BuildPeopleList();
            using var workbook = new ClosedXML.Excel.XLWorkbook();
            var report = new MultiHeaderReport<Person>(workbook, "People");

            report.Configure(options => options
                .AddColumn(x => x.NumOfComputers, 1)
            ).GenerateReport(people);

            var sheet = workbook.Worksheet("People");
            Assert.Equal(nameof(Person.NumOfComputers), sheet.Cell(1, 1).GetValue<string>());
            Assert.Equal(nameof(Person.Name), sheet.Cell(1, 2).GetValue<string>());
        }

        [Fact]
        public void Config_IgnoredColumns_AreNotInTheList()
        {
            var people = Person.BuildPeopleList();
            using var workbook = new ClosedXML.Excel.XLWorkbook();
            var report = new MultiHeaderReport<Person>(workbook, "People");

            report.Configure(options => options
                .AddColumn(x => x.Surname, 1)
                .IgnoreColumn(x => x.NumOfComputers)
            ).GenerateReport(people);

            var sheet = workbook.Worksheet("People");
            Assert.Equal(_maxColumns - 1, sheet.LastColumnUsed()!.ColumnNumber());
        }

        [Fact]
        public void HiddenColumns_AreRendered_AsHidden()
        {
            var people = Person.BuildPeopleList();
            using var workbook = new ClosedXML.Excel.XLWorkbook();
            var report = new MultiHeaderReport<Person>(workbook, "People");

            report.Configure(options => options
                .AddColumn(x => x.NumOfComputers, hidden: true)
            ).GenerateReport(people);

            var sheet = workbook.Worksheet("People");
            Assert.True(sheet.Column(5).IsHidden);
        }

        [Fact]
        public void HyperLinkColumns_UseAntherColumnTo_BuildTheLink()
        {
            var people = Person.BuildPeopleList();
            using var workbook = new ClosedXML.Excel.XLWorkbook();
            var report = new MultiHeaderReport<Person>(workbook, "People");

            report.Configure(options => options
                .AddHyperLinkColumn(x => x.Name, x => x.Profile)
                .IgnoreColumn(x => x.Profile)
            ).GenerateReport(people);

            var sheet = workbook.Worksheet("People");
            Assert.True(sheet.Cell(3, 1).HasHyperlink);
        }

        [Fact]
        public void FormulaColumns_Write_Formulas()
        {
            var people = Person.BuildPeopleList();
            using var workbook = new ClosedXML.Excel.XLWorkbook();

            var report = new MultiHeaderReport<Person>(workbook, "People");
            report.Configure(options => options
                .AddColumn(x => x.Name, 1)
                .AddColumn(x => x.Surname, 2)
                .AddFormula("CompleteName", "CONCATENATE(B2,\", \",A2)", 3)
            ).GenerateReport(people);

            var sheet = workbook.Worksheet("People");
            Assert.Contains("CONCATENATE", sheet.Cell(2, 3).FormulaA1);
        }

        [Fact]
        public void ExpressionColumns_Write_ExpressionResults()
        {
            var people = Person.BuildPeopleList();
            using var workbook = new ClosedXML.Excel.XLWorkbook();

            var report = new MultiHeaderReport<Person>(workbook, "People");
            report.Configure(options => options
                .AddColumn(x => x.Name, 1)
                .AddColumn(x => x.Surname, 2)
                .AddExpression("Initials", x => string.Concat(x.Name[0], '.', x.Surname[0], '.'), 3)
            ).GenerateReport(people);

            var sheet = workbook.Worksheet("People");
            Assert.Equal("A.B.", sheet.Cell(3, 3).GetValue<string>());
        }
    }
}
