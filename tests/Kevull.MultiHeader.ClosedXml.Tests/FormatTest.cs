using Kevull.MultiHeader.Core;
using Kevull.MultiHeader.TestCommon;

namespace Kevull.MultiHeader.ClosedXml.Test
{
    public class FormatTest
    {
        [Fact]
        public void PropertiesWithoutChildren_HasVerticalMerge()
        {
            var complexObject = RootLevel.CreateTest();
            using var workbook = new ClosedXML.Excel.XLWorkbook();

            var report = new MultiHeaderReport<RootLevel>(workbook, "Object");
            report.GenerateReport(complexObject);
            var sheet = workbook.Worksheet("Object");

            Assert.True(sheet.Range("A1:A3").IsMerged());
        }

        [Fact]
        public void Configuration_WithHeaderStyle_HasOverridenDefaultStyle()
        {
            var complexObject = RootLevelDictionary.CreateTest();
            using var workbook = new ClosedXML.Excel.XLWorkbook();

            var report = new MultiHeaderReport<RootLevelDictionary>(workbook, "Object");
            report.Configure(options => options
                .AddEnumeration(x => x.ComplexProperty.RightColumn, complexObject.First().ComplexProperty.RightColumn.Keys)
                .AddHeaderStyle(x =>
                {
                    x.HorizontalAlignment = HorizontalAlignment.Center;
                    x.VerticalAlignment = VerticalAlignment.Center;
                    x.BackgroundColor = ExcelColor.Black;
                })
            );
            report.GenerateReport(complexObject);
            var sheet = workbook.Worksheet("Object");

            Assert.Equal(ClosedXML.Excel.XLAlignmentHorizontalValues.Center, sheet.Cell("A1").Style.Alignment.Horizontal);
            Assert.Equal(ExcelColor.Black.Argb, sheet.Cell("A1").Style.Fill.BackgroundColor.Color.ToArgb().ToString("X8"));
        }

        [Fact]
        public void Headers_WithAutoFilter_SetAutoFilterInLeafLevelHeader()
        {
            var complexObject = RootLevel.CreateTest();
            using var workbook = new ClosedXML.Excel.XLWorkbook();

            var report = new MultiHeaderReport<RootLevel>(workbook, "Object");
            report.GenerateReport(complexObject);
            var sheet = workbook.Worksheet("Object");

            Assert.NotNull(sheet.AutoFilter.Range);
            Assert.Equal(3, sheet.AutoFilter.Range.RangeAddress.FirstAddress.RowNumber);
            Assert.Equal(1, sheet.AutoFilter.Range.RangeAddress.FirstAddress.ColumnNumber);
            Assert.Equal(3, sheet.AutoFilter.Range.RangeAddress.LastAddress.RowNumber);
            Assert.Equal(5, sheet.AutoFilter.Range.RangeAddress.LastAddress.ColumnNumber);
        }

        [Fact]
        public void DateOrTimeColumns_HasByDefault_DateTimeNumberFormat()
        {
            var people = Person.BuildPeopleList();
            using var workbook = new ClosedXML.Excel.XLWorkbook();

            var report = new MultiHeaderReport<Person>(workbook, "People");
            report.GenerateReport(people);
            var sheet = workbook.Worksheet("People");

            int birthDateColumn = sheet.Row(1).CellsUsed().First(x => x.GetValue<string>() == nameof(Person.BirthDate)).Address.ColumnNumber;
            int alarmTimeColumn = sheet.Row(1).CellsUsed().First(x => x.GetValue<string>() == nameof(Person.AlarmTime)).Address.ColumnNumber;

            Assert.Equal(StyleNames.DateFormat, sheet.Cell(2, birthDateColumn).Style.NumberFormat.Format);
            Assert.Equal(StyleNames.TimeFormat, sheet.Cell(2, alarmTimeColumn).Style.NumberFormat.Format);
        }

        [Fact]
        public void DateColumns_WithAppliedStyle_HasSpecifiedFormat()
        {
            var people = Person.BuildPeopleList();
            using var workbook = new ClosedXML.Excel.XLWorkbook();

            var report = new MultiHeaderReport<Person>(workbook, "People");
            report.Configure(options => options
                .AddNamedStyle("BirthDay", s =>
                {
                    s.Italic = true;
                    s.NumberFormat = "dd/mm";
                })
                .AddColumn(x => x.BirthDate, styleName: "BirthDay")
            );
            report.GenerateReport(people);
            var sheet = workbook.Worksheet("People");

            Assert.Equal("dd/mm", sheet.Cell("C2").Style.NumberFormat.Format);
            Assert.True(sheet.Cell("C2").Style.Font.Italic);
        }

        [Fact]
        public void Columns_WithSpecifiedWidth_ApplyDefinedConfiguration()
        {
            var people = Person.BuildPeopleList();
            using var workbook = new ClosedXML.Excel.XLWorkbook();

            var report = new MultiHeaderReport<Person>(workbook, "People");
            report.Configure(options => options
                .AddColumn(x => x.Name, cfg =>
                    cfg.ColumnWidth.SetWidth(WidthType.Auto))
                .AddColumn(x => x.Surname, cfg =>
                    cfg.ColumnWidth.SetWidth(8.0))
                .AddColumn(x => x.BirthDate, cfg =>
                    cfg.ColumnWidth.SetWidth(WidthType.Hidden))
                .AddColumn(x => x.NumOfComputers, cfg =>
                    cfg.ColumnWidth.SetWidth(WidthType.Auto, 12.0, 20.0))
            );
            report.GenerateReport(people);
            var sheet = workbook.Worksheet("People");

            Assert.NotEqual(sheet.ColumnWidth, sheet.Column(1).Width);
            Assert.Equal(8.0, sheet.Column(2).Width);
            Assert.True(sheet.Column(3).IsHidden);
            Assert.InRange(sheet.Column(5).Width, 12.0, 20.0);
        }
    }
}
