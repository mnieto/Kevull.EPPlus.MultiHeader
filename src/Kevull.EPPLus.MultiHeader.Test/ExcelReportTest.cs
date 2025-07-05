using System.Reflection;
using System.Runtime.CompilerServices;
using System.Xml.Linq;

namespace Kevull.EPPLus.MultiHeader.Test
{
    public static class TestUtils
    {
        public static void Save<T>(this MultiHeaderReport<T> report, [CallerMemberName]string methodName = "")
        {
            report.Save(methodName + ".xlsx");
        }

        public static string GetTestAssemblyFolder()
        {
            string fullPath = Assembly.GetExecutingAssembly().Location;
            string name = Assembly.GetExecutingAssembly().GetName().Name!;
            int index = fullPath.IndexOf(name);
            return Path.Combine(fullPath.Substring(0, index), name);
        }
    }

}