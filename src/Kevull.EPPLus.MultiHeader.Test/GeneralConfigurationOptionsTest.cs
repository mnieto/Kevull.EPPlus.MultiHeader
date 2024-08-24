using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kevull.EPPLus.MultiHeader.Test
{
    public class GeneralConfigurationOptionsTest
    {
        public GeneralConfigurationOptionsTest()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        [Fact]
        public void ReporttStatsAt_TopLeftStartingPoint()
        {
            var people = Person.BuildPeopleList();
            using var xls = new ExcelPackage();

            var report = new MultiHeaderReport<Person>(xls, "People");
            report.Configure(config =>
                config.SetStartingAddres(3, 2)
            );
            report.GenerateReport(people);

            var sheet = xls.Workbook.Worksheets["People"];
            Assert.Equal("Name", sheet.GetValue<string>(3, 2));
            Assert.Equal("Médiamass", sheet.GetValue<string>(4, 2));
        }


    }
}
