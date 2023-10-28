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
            var sut = new ColumnInfo(typeof(Person).GetProperty(nameof(Person.Name))!);
            Action act = () => sut.Order = 0;
            Assert.Throws<ArgumentOutOfRangeException>(act);
        }

        [Fact]
        public void DisplayName_IsName_IfNotAssigned()
        {
            var sut = new ColumnInfo(typeof(Person).GetProperty(nameof(Person.BirthDate))!);
            Assert.Equal(nameof(Person.BirthDate), sut.DisplayName);
        }
    }
}
