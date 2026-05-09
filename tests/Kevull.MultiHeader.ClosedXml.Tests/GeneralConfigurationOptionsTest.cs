using Kevull.MultiHeader.TestCommon;

namespace Kevull.MultiHeader.ClosedXml.Test
{
    public class GeneralConfigurationOptionsTest : BaseTest
    {
        [Fact]
        public void ReportStatsAt_TopLeftStartingPoint()
        {
            var people = Person.BuildPeopleList();
            using var workbook = new ClosedXML.Excel.XLWorkbook();

            var report = new MultiHeaderReport<Person>(workbook, "People");
            report.Configure(config =>
                config.SetStartingAddres(3, 2)
            );
            report.GenerateReport(people);

            var sheet = workbook.Worksheet("People");
            Assert.Equal("Name", sheet.Cell(3, 2).GetValue<string>());
            Assert.Equal("Médiamass", sheet.Cell(4, 2).GetValue<string>());

            Assert.True(sheet.Cell(3, 8).Style.Font.Bold);
            Assert.NotNull(sheet.AutoFilter.Range);
            Assert.Equal(3, sheet.AutoFilter.Range.RangeAddress.FirstAddress.RowNumber);
            Assert.Equal(2, sheet.AutoFilter.Range.RangeAddress.FirstAddress.ColumnNumber);
            Assert.Equal(3, sheet.AutoFilter.Range.RangeAddress.LastAddress.RowNumber);
            Assert.Equal(8, sheet.AutoFilter.Range.RangeAddress.LastAddress.ColumnNumber);
        }

        [Fact]
        public void Report_WithAppendToExistingReport_AppendsNewRowsAtBottom()
        {
            var people = Person.BuildPeopleList();
            using var workbook = new ClosedXML.Excel.XLWorkbook();

            var report = new MultiHeaderReport<Person>(workbook, "People");
            report.GenerateReport(people);

            people = Person.BuildPeopleList(2, 3);
            report = new MultiHeaderReport<Person>(workbook, "People");
            report.Configure(config =>
                config.AppendToExistingReport = true
            );
            report.GenerateReport(people);

            var sheet = workbook.Worksheet("People");
            Assert.Equal("Michelle", sheet.Cell(4, 1).GetValue<string>());
        }

        [Fact]
        public void Report_WithAutoFreeze_ProperlyFreezes()
        {
            var people = Person.BuildPeopleList();
            using var workbook = new ClosedXML.Excel.XLWorkbook();
            var report = new MultiHeaderReport<Person>(workbook, "People");
            report.Configure(config =>
                config.AutoFreezePanes = true
            );
            report.GenerateReport(people);
            var sheet = workbook.Worksheet("People");

            Assert.Equal(1, sheet.SheetView.SplitRow);
            Assert.Equal(0, sheet.SheetView.SplitColumn);
        }

        [Fact]
        public void Report_WithoutAutoFreeze_DoNotHasFrozenPanes()
        {
            var people = Person.BuildPeopleList();
            using var workbook = new ClosedXML.Excel.XLWorkbook();
            var report = new MultiHeaderReport<Person>(workbook, "People");
            report.Configure(config =>
                config.AutoFreezePanes = false
            );
            report.GenerateReport(people);
            var sheet = workbook.Worksheet("People");

            Assert.Equal(0, sheet.SheetView.SplitRow);
            Assert.Equal(0, sheet.SheetView.SplitColumn);
        }
    }
}
