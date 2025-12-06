using Kevull.MultiHeader.EPPlus;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kevull.EPPLus.MultiHeader.Test
{
    public class GeneralConfigurationOptionsTest : BaseTest
    {
        [Fact]
        public void ReportStatsAt_TopLeftStartingPoint()
        {
            var people = Person.BuildPeopleList();
            using var xls = new ExcelPackage();

            var report = new MultiHeaderReportEPPlus<Person>(xls, "People");
            report.Configure(config =>
                config.SetStartingAddres(3, 2)
            );
            report.GenerateReport(people);

            var sheet = xls.Workbook.Worksheets["People"];
            Assert.Equal("Name", sheet.GetValue<string>(3, 2));
            Assert.Equal("Médiamass", sheet.GetValue<string>(4, 2));

            Assert.True(sheet.Cells[3, 8].Style.Font.Bold);             //Last column is properly formatted as header
            Assert.Equal("B3:H3", sheet.AutoFilter.Address.Address);    //AutoFilter is applied to the correct range
        }

        [Fact]
        public void Report_WithAppendToExistingReport_AppendsNewRowsAtBottom()
        {
            var people = Person.BuildPeopleList();
            using var xls = new ExcelPackage();

            var report = new MultiHeaderReportEPPlus<Person>(xls, "People");
            report.GenerateReport(people);
            var sheet = xls.Workbook.Worksheets["People"];

            people = Person.BuildPeopleList(2, 3);
            report = new MultiHeaderReportEPPlus<Person>(xls, "People");
            report.Configure(config =>
                config.AppendToExistingReport = true
            );
            report.GenerateReport(people);

            sheet = xls.Workbook.Worksheets["People"];
            Assert.Equal("Michelle", sheet.GetValue<string>(4, 1));
        }

        [Fact]
        public void Report_WithAutoFreeze_ProperlyFreezes()
        {
            var people = Person.BuildPeopleList();
            using var xls = new ExcelPackage();
            var report = new MultiHeaderReportEPPlus<Person>(xls, "People");
            report.Configure(config =>
                config.AutoFreezePanes = true
            );
            report.GenerateReport(people);
            var sheet = xls.Workbook.Worksheets["People"];
            Assert.Equal("A2", sheet.View.PaneSettings.TopLeftCell);
        }

        [Fact]
        public void Report_WithoutAutoFreeze_DoNotHasFrozenPanes()
        {
            var people = Person.BuildPeopleList();
            using var xls = new ExcelPackage();
            var report = new MultiHeaderReportEPPlus<Person>(xls, "People");
            report.Configure(config =>
                config.AutoFreezePanes = false
            );
            report.GenerateReport(people);
            var sheet = xls.Workbook.Worksheets["People"];
            Assert.Null(sheet.View.PaneSettings);
        }
    }
}
