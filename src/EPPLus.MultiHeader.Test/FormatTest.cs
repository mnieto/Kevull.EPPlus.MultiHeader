using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace EPPLus.MultiHeader.Test
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
                .AddHeaderStyle(xls, x => {
                    x.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    x.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                })
            );
            report.GenerateReport(complexObject);
            var sheet = xls.Workbook.Worksheets["Object"];

            Assert.Equal(OfficeOpenXml.Style.ExcelHorizontalAlignment.Center, sheet.Cells["A1"].Style.HorizontalAlignment);
            Assert.NotEqual(Color.LightGray.ToArgb().ToString("X"), sheet.Cells["A1"].Style.Fill.BackgroundColor.Rgb);
        }
    }
}
