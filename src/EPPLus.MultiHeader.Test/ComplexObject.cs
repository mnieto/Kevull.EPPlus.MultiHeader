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
        public static List<RootLevel> CreateTest()
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

        public static List<RootLevelDictionary> CreateTest()
        {
            return new List<RootLevelDictionary> {
                new RootLevelDictionary {
                    SimpleProperty = "String1",
                    ComplexProperty = new SecondLevelDictionary
                    {
                        LeftColumn = "Left side 1",
                        RightColumn = new Dictionary<string, int>
                        {
                            { "CatA", 11 },
                            { "CatB", 12 },
                            { "CatC", 13 }
                        }
                    }
                },
                new RootLevelDictionary {
                    SimpleProperty = "String2",
                    ComplexProperty = new SecondLevelDictionary
                    {
                        LeftColumn = "Left side 2",
                        RightColumn = new Dictionary<string, int>
                        {
                            { "CatA", 21 },
                            { "CatC", 23 }
                        }
                    }
                }
            };
        }
    }

    internal class SecondLevelDictionary
    {
        public string LeftColumn { get; set; }
        public Dictionary<string, int> RightColumn { get; set; }
    }


}
