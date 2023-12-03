using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPLus.MultiHeader.Test
{
    internal class RootLevel
    {
        public string SimpleProperty { get; set; }
        public SecondLevel ComplexProperty { get; set; }
    }

    internal class SecondLevel
    {
        public string LeftColumn { get; set; }
        public ThirdLevel RightColumn { get; set; }
    }

    internal class ThirdLevel
    {
        public int CatA { get; set; }
        public int CatB { get; set; }
        public int CatC { get; set; }
    }


    internal class RootLevelDictionary
    {
        public string SimpleProperty { get; set; }
        public SecondLevelDictionary ComplexProperty { get; set; }
    }

    internal class SecondLevelDictionary
    {
        public string LeftColumn { get; set; }
        public Dictionary<string, int> RightColumn { get; set; }
    }


}
