using OfficeOpenXml;
using System.Reflection;

namespace Kevull.MultiHeader.EPPLus.Test
{
    public class EPPLusBaseTest : BaseTest
    {
        public EPPLusBaseTest()
        {
            ExcelPackage.License.SetNonCommercialPersonal("Kevull");
        }
    }
}