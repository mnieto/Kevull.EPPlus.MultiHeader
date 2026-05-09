using Kevull.MultiHeader.TestCommon;

namespace Kevull.MultiHeader.ClosedXml.Test
{
    public class TwoHeaderRenderTest
    {
        [Fact]
        public void ComposedObjects_AreRendered_InSecondRow()
        {
            var complexObject = RootLevel.CreateTest();
            using var workbook = new ClosedXML.Excel.XLWorkbook();

            var report = new MultiHeaderReport<RootLevel>(workbook, "Object");
            report.GenerateReport(complexObject);
            var sheet = workbook.Worksheet("Object");

            Assert.Equal(nameof(RootLevel.SimpleProperty), sheet.Cell(1, 1).GetValue<string>());
            Assert.Equal(nameof(RootLevel.ComplexProperty), sheet.Cell(1, 2).GetValue<string>());
            Assert.True(sheet.Cell(1, 3).IsEmpty());

            Assert.True(sheet.Cell(2, 1).IsEmpty());
            Assert.Equal(nameof(RootLevel.ComplexProperty.LeftColumn), sheet.Cell(2, 2).GetValue<string>());
            Assert.Equal(nameof(RootLevel.ComplexProperty.RightColumn), sheet.Cell(2, 3).GetValue<string>());
            Assert.True(sheet.Cell(2, 4).IsEmpty());

            Assert.True(sheet.Cell(3, 1).IsEmpty());
            Assert.True(sheet.Cell(3, 2).IsEmpty());
            Assert.Equal(nameof(RootLevel.ComplexProperty.RightColumn.CatA), sheet.Cell(3, 3).GetValue<string>());

            Assert.Equal("String1", sheet.Cell(4, 1).GetValue<string>());
            Assert.Equal("Left side 1", sheet.Cell(4, 2).GetValue<string>());
            Assert.Equal(11, sheet.Cell(4, 3).GetValue<int>());
            Assert.Equal(12, sheet.Cell(4, 4).GetValue<int>());
            Assert.Equal(13, sheet.Cell(4, 5).GetValue<int>());
        }

        [Fact]
        public void ComposedObjects_WithEnumerables_NeedsToBeConfigured()
        {
            var complexObject = RootLevelDictionary.CreateTest();
            using var workbook = new ClosedXML.Excel.XLWorkbook();

            var report = new MultiHeaderReport<RootLevelDictionary>(workbook, "Object");
            Assert.Throws<InvalidOperationException>(() => report.GenerateReport(complexObject));
        }

        [Fact]
        public void ComposedObjects_WithEnumerables_HasWithEqualsToCountOfKeys()
        {
            var complexObject = RootLevelDictionary.CreateTest();
            using var workbook = new ClosedXML.Excel.XLWorkbook();

            var report = new MultiHeaderReport<RootLevelDictionary>(workbook, "Object");
            report.Configure(options =>
                options.AddEnumeration(x => x.ComplexProperty.RightColumn, complexObject.First().ComplexProperty.RightColumn.Keys)
            );
            report.GenerateReport(complexObject);
            var sheet = workbook.Worksheet("Object");

            Assert.Equal(nameof(RootLevel.SimpleProperty), sheet.Cell(1, 1).GetValue<string>());
            Assert.Equal(nameof(RootLevel.ComplexProperty), sheet.Cell(1, 2).GetValue<string>());
            Assert.True(sheet.Cell(1, 3).IsEmpty());

            Assert.True(sheet.Cell(2, 1).IsEmpty());
            Assert.Equal(nameof(RootLevel.ComplexProperty.LeftColumn), sheet.Cell(2, 2).GetValue<string>());
            Assert.Equal(nameof(RootLevel.ComplexProperty.RightColumn), sheet.Cell(2, 3).GetValue<string>());
            Assert.True(sheet.Cell(2, 4).IsEmpty());

            Assert.True(sheet.Cell(3, 1).IsEmpty());
            Assert.True(sheet.Cell(3, 2).IsEmpty());
            Assert.Equal(nameof(RootLevel.ComplexProperty.RightColumn.CatA), sheet.Cell(3, 3).GetValue<string>());

            Assert.Equal("String1", sheet.Cell(4, 1).GetValue<string>());
            Assert.Equal("Left side 1", sheet.Cell(4, 2).GetValue<string>());
            Assert.Equal(11, sheet.Cell(4, 3).GetValue<int>());
            Assert.Equal(12, sheet.Cell(4, 4).GetValue<int>());
            Assert.Equal(13, sheet.Cell(4, 5).GetValue<int>());
        }
    }
}
