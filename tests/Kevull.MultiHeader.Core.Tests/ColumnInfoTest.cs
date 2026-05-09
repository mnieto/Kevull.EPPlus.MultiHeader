using Kevull.MultiHeader.Core.Columns;
using Kevull.MultiHeader.TestCommon;
using System;
using System.Linq;

namespace Kevull.MultiHeader.Core.Tests
{
    public class ColumnInfoTest
    {
        [Fact]
        public void Order_MustBeOneOrUpper()
        {
            var sut = new ColumnInfo(nameof(Person.Name));
            Action act = () => sut.Order = 0;
            Assert.Throws<ArgumentOutOfRangeException>(act);
        }

        [Fact]
        public void DisplayName_IsName_IfNotAssigned()
        {
            var sut = new ColumnInfo(nameof(Person.BirthDate));
            Assert.Equal(sut.Name, sut.DisplayName);
        }

        [Fact]
        public void Deep_InDirectProprties_IsOne()
        {
            var sut = new ColumnInfo<RootLevelDictionary>(x => x.SimpleProperty);
            Assert.Equal(1, sut.Deep);
        }

        [Fact]
        public void Deep_InDirectChildProperties_IsTwo()
        {
            var sut = new ColumnInfo<RootLevelDictionary>(x => x.ComplexProperty.RightColumn);
            Assert.Equal(2, sut.Deep);
        }
    }
}
