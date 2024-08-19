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
    public class TwoHeaderRenderTest
    {

        public TwoHeaderRenderTest() {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        [Fact]
        public void ComposedObjects_AreRendered_InSecondRow()
        {
            var complexObject = RootLevel.CreateTest();
            using var xls = new ExcelPackage();

            var report = new MultiHeaderReport<RootLevel>(xls, "Object");
            report.GenerateReport(complexObject);
            var sheet = xls.Workbook.Worksheets["Object"];

            //Headers1
            Assert.Equal(nameof(RootLevel.SimpleProperty), sheet.GetValue<string>(1, 1));
            Assert.Equal(nameof(RootLevel.ComplexProperty), sheet.GetValue<string>(1, 2));
            Assert.Null(sheet.GetValue(1, 3));
            //Headers2
            Assert.Null(sheet.GetValue(2, 1));
            Assert.Equal(nameof(RootLevel.ComplexProperty.LeftColumn), sheet.GetValue<string>(2, 2));
            Assert.Equal(nameof(RootLevel.ComplexProperty.RightColumn), sheet.GetValue<string>(2, 3));
            Assert.Null(sheet.GetValue(2, 4));
            //Headers3
            Assert.Null(sheet.GetValue(3, 1));
            Assert.Null(sheet.GetValue(3, 2));
            Assert.Equal(nameof(RootLevel.ComplexProperty.RightColumn.CatA), sheet.GetValue<string>(3, 3));

            //Data1
            Assert.Equal("String1", sheet.GetValue(4, 1));
            Assert.Equal("Left side 1", sheet.GetValue(4, 2));
            Assert.Equal(11, sheet.GetValue<int>(4, 3));
            Assert.Equal(12, sheet.GetValue<int>(4, 4));
            Assert.Equal(13, sheet.GetValue<int>(4, 5));

        }

        [Fact]
        public void ComposedObjects_WithEnumerables_NeedsToBeConfigured()
        {
            var complexObject = RootLevelDictionary.CreateTest();
            using var xls = new ExcelPackage();

            var report = new MultiHeaderReport<RootLevelDictionary>(xls, "Object");
            Assert.Throws<InvalidOperationException>(() => { report.GenerateReport(complexObject); });
        }

        [Fact]
        public void ComposedObjects_WithEnumerables_HasWithEqualsToCountOfKeys()
        {
            var complexObject = RootLevelDictionary.CreateTest();
            using var xls = new ExcelPackage();

            var report = new MultiHeaderReport<RootLevelDictionary>(xls, "Object");
            report.Configure(options =>
                options.AddEnumeration(x => x.ComplexProperty.RightColumn, complexObject.First().ComplexProperty.RightColumn.Keys)
            );
            report.GenerateReport(complexObject);
            var sheet = xls.Workbook.Worksheets["Object"];

            //Headers1
            Assert.Equal(nameof(RootLevel.SimpleProperty), sheet.GetValue<string>(1, 1));
            Assert.Equal(nameof(RootLevel.ComplexProperty), sheet.GetValue<string>(1, 2));
            Assert.Null(sheet.GetValue(1, 3));
            //Headers2
            Assert.Null(sheet.GetValue(2, 1));
            Assert.Equal(nameof(RootLevel.ComplexProperty.LeftColumn), sheet.GetValue<string>(2, 2));
            Assert.Equal(nameof(RootLevel.ComplexProperty.RightColumn), sheet.GetValue<string>(2, 3));
            Assert.Null(sheet.GetValue(2, 4));
            //Headers3
            Assert.Null(sheet.GetValue(3, 1));
            Assert.Null(sheet.GetValue(3, 2));
            Assert.Equal(nameof(RootLevel.ComplexProperty.RightColumn.CatA), sheet.GetValue<string>(3, 3));

            //Data1
            Assert.Equal("String1", sheet.GetValue(4, 1));
            Assert.Equal("Left side 1", sheet.GetValue(4, 2));
            Assert.Equal(11, sheet.GetValue<int>(4, 3));
            Assert.Equal(12, sheet.GetValue<int>(4, 4));
            Assert.Equal(13, sheet.GetValue<int>(4, 5));
        }
    }
}
