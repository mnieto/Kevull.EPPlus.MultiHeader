using OfficeOpenXml;
using System.Reflection;

namespace Kevull.EPPLus.MultiHeader.Test
{
    public class BaseTest
    {
        public BaseTest()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        protected void Save<T>(MultiHeaderReport<T> report, [System.Runtime.CompilerServices.CallerMemberName] string methodName = "")
        {
            report.Save(string.Concat(methodName, ".xlsx"));
        }

        protected string GetTestAssemblyFolder()
        {
            string fullPath = Assembly.GetExecutingAssembly().Location;
            string name = Assembly.GetExecutingAssembly().GetName().Name!;
            int index = fullPath.IndexOf(name);
            return Path.Combine(fullPath.Substring(0, index), name);
        }
    }
}