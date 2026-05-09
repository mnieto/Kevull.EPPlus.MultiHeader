using Kevull.MultiHeader.Core;
using System.IO;
using System.Reflection;

namespace Kevull.MultiHeader.TestCommon
{
    public class BaseTest
    {
        protected void Save<T>(IMultiHeaderReport<T> report, [System.Runtime.CompilerServices.CallerMemberName] string methodName = "")
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