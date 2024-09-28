using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace Kevull.EPPLus.MultiHeader.Test
{
    public class FormatTest
    {
        public FormatTest()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        [Fact]
        public void PropertiesWithoutChildren_HasVerticalMerge()
        {
            var complexObject = RootLevel.CreateTest();
            using var xls = new ExcelPackage();

            var report = new MultiHeaderReport<RootLevel>(xls, "Object");
            report.GenerateReport(complexObject);
            var sheet = xls.Workbook.Worksheets["Object"];

            Assert.True(sheet.Cells["A1:A3"].Merge);
        }

        [Fact]
        public void Configuration_WithHeaderStyle_HasOverridenDefaultStyle()
        {
            var complexObject = RootLevelDictionary.CreateTest();
            using var xls = new ExcelPackage();

            var report = new MultiHeaderReport<RootLevelDictionary>(xls, "Object");
            report.Configure(options => options
                .AddEnumeration(x => x.ComplexProperty.RightColumn, complexObject.First().ComplexProperty.RightColumn.Keys)
                .AddHeaderStyle(x =>
                {
                    x.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    x.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                })
            );
            report.GenerateReport(complexObject);
            var sheet = xls.Workbook.Worksheets["Object"];

            Assert.Equal(OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, sheet.Cells["A1"].Style.HorizontalAlignment);
            Assert.NotEqual(Color.LightGray.ToArgb().ToString("X"), sheet.Cells["A1"].Style.Fill.BackgroundColor.Rgb);
        }

        [Fact]
        public void Headers_WithAutoFilter_SetAutoFilterInLeafLevelHeader()
        {
            var complexObject = RootLevel.CreateTest();
            using var xls = new ExcelPackage();

            var report = new MultiHeaderReport<RootLevel>(xls, "Object");
            report.GenerateReport(complexObject);
            var sheet = xls.Workbook.Worksheets["Object"];

            Assert.True(sheet.Cells["A3:E3"].AutoFilter);
        }

        [Fact]
        public void DateOrTimeColumns_HasByDefault_DateTimeNumberFormat()
        {
            var people = Person.BuildPeopleList();
            using var xls = new ExcelPackage();

            var report = new MultiHeaderReport<Person>(xls, "People");
            report.GenerateReport(people);
            var sheet = xls.Workbook.Worksheets["People"];

            Assert.Equal(StyleNames.DateFormat, sheet.Cells["C2"].Style.Numberformat.Format);
            Assert.Equal(StyleNames.TimeFormat, sheet.Cells["G2"].Style.Numberformat.Format);
        }

        [Fact]
        public void DateColumns_WithAppliedStyle_HasSpecifiedFormat()
        {
            var people = Person.BuildPeopleList();
            using var xls = new ExcelPackage();

            var report = new MultiHeaderReport<Person>(xls, "People");
            report.Configure(options => options
                .AddNamedStyle("BirthDay", s =>
                {
                    s.Font.Italic = true;
                    s.Numberformat.Format = "dd/mm";
                })
                .AddColumn(x => x.BirthDate, styleName: "BirthDay")
            );
            report.GenerateReport(people);
            var sheet = xls.Workbook.Worksheets["People"];

            Assert.Equal("dd/mm", sheet.Cells["C2"].Style.Numberformat.Format);
            Assert.True(sheet.Cells["C2"].Style.Font.Italic);
        }

        [Fact]
        public void Columns_WithSpecifiedWidth_ApplyDefinedConfiguraiton()
        {

            var people = Person.BuildPeopleList();
            using var xls = new ExcelPackage();

            var report = new MultiHeaderReport<Person>(xls, "People");
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
            var sheet = xls.Workbook.Worksheets["People"];

            Assert.NotEqual(sheet.DefaultColWidth, sheet.Column(1).Width);
            Assert.Equal(8.0, sheet.Column(2).Width);
            Assert.True(sheet.Column(3).Hidden);
            Assert.Equal(sheet.DefaultColWidth, sheet.Column(4).Width);
            Assert.Equal(12.0, sheet.Column(5).Width);
        }
    }
}
