using Kevull.EPPLus.MultiHeader.Columns;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kevull.EPPLus.MultiHeader.Test
{
    public class PropertyNameBuilderTest
    {
        [Fact]
        public void SimpleProperty_NameEqualstoFullName()
        {
            var sut = new PropertyNameBuilder<RootLevel>();
            var names = sut.Build(x => x.SimpleProperty);
            Assert.Equal(names.FullName, names.Name);
        }
        [Fact]
        public void SimpleProperty_MemberExpression()
        {
            var sut = new PropertyNameBuilder<RootLevel>();
            var names = sut.Build(x => x.SimpleProperty);
            Assert.Equal(nameof(RootLevel.SimpleProperty), names.Name);
        }

        [Fact]
        public void SimpleProperty_UnaryExpression()
        {
            var sut = new PropertyNameBuilder<Person>();
            var names = sut.Build(x => x.NumOfComputers);
            Assert.Equal(nameof(Person.NumOfComputers), names.Name);
        }

        [Fact]
        public void NestedProperty_WithTwoLevels()
        {
            var sut = new PropertyNameBuilder<RootLevel>();
            var names = sut.Build(x => x.ComplexProperty.LeftColumn);
            Assert.Equal("LeftColumn", names.Name);
            Assert.Equal("ComplexProperty.LeftColumn", names.FullName);
            Assert.Equal("ComplexProperty", names.ParentName);
            Assert.Equal(typeof(SecondLevel), names.ParentType);
        }

        [Fact]
        public void NestedProperty_WithThreeLevels()
        {
            var sut = new PropertyNameBuilder<RootLevel>();
            var names = sut.Build(x => x.ComplexProperty.RightColumn.CatA);
            Assert.Equal("CatA", names.Name);
            Assert.Equal("ComplexProperty.RightColumn.CatA", names.FullName);
            Assert.Equal("RightColumn", names.ParentName);
            Assert.Equal(typeof(ThirdLevel), names.ParentType);
        }
    }
}
