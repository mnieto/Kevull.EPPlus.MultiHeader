using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPLus.MultiHeader.Test
{
    public class ColumInfoTest
    {
        [Fact]
        public void Order_MustBeOneOrUpper()
        {
            var property = typeof(Person).GetProperties().First(x => x.Name == nameof(Person.Name));
            var sut = new ColumnInfo(property);
            Action act = () => sut.Order = 0;
            Assert.Throws<ArgumentOutOfRangeException>(act);
        }

        [Fact]
        public void DisplayName_IsName_IfNotAssigned()
        {
            var property = typeof(Person).GetProperties().First(x => x.Name == nameof(Person.BirthDate));
            var sut = new ColumnInfo(property);
            Assert.Equal(property.Name, sut.DisplayName);
        }
    }
}
