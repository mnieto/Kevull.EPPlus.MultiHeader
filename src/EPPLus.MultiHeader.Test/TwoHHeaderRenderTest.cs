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
    public class TwoHHeaderRenderTest
    {

        public TwoHHeaderRenderTest() {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        private List<RootLevel> BuildStructure()
        {
            return new List<RootLevel> { 
                new RootLevel {
                    SimpleProperty = "String1",
                    ComplexProperty = new SecondLevel
                    {
                        LeftColumn = "Left side 1",
                        RightColumn = new ThirdLevel
                        {
                            CatA = 11,
                            CatB = 12,
                            CatC = 13
                        }
                    }
                }, 
                new RootLevel {
                    SimpleProperty = "String2",
                    ComplexProperty = new SecondLevel
                    {
                        LeftColumn = "Left side 2",
                        RightColumn = new ThirdLevel
                        {
                            CatA = 21,
                            CatB = 22,
                            CatC = 23
                        }
                    }
                }
            };
        }

        [Fact]
        public void ComposedObjects_AreRendered_InSecondRow()
        {
            var complexObject = BuildStructure();
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

        }
    }
}
