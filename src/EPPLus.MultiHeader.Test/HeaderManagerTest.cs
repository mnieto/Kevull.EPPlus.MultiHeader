using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection.PortableExecutable;
using System.Text;
using System.Threading.Tasks;

namespace EPPLus.MultiHeader.Test
{


    public class HeaderManagerTest
    {

        [Fact]
        public void SimpleProperty_Has_No_Children()
        {
            var header = new HeaderManager<RootLevel>();
            var col = header.Columns.First(x => x.Name == nameof(RootLevel.SimpleProperty));

            Assert.False(col.HasChildren);
        }

        [Fact]
        public void SimpleProperty_Has_Width1()
        {
            var header = new HeaderManager<RootLevel>();
            var col = header.Columns.First(x => x.Name == nameof(RootLevel.SimpleProperty));

            Assert.Equal(1, col.Width);
        }

        [Fact]
        public void ComplexObject_Width_IsTheSumOfLeafNodes()
        {
            var header = new HeaderManager<RootLevel>();
            var col = header.Columns.First(x => x.Name == nameof(RootLevel.ComplexProperty));
            Assert.Equal(4, col.Width);
        }

        [Fact]
        public void Height_IsTheMaxOfRows()
        {
            var header = new HeaderManager<RootLevel>();
            Assert.Equal(3, header.Height);
        }
}
}
