using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPLus.MultiHeader.Test
{
    public class ConfigurationBuilderTest
    {
        [Fact]
        public void Build_WithDefaultConfig_AllPropertiesAreOrdered()
        {
            var builder = new ConfigurationBuilder<Person>();

            var sut = builder.Build();

            Assert.Equal(1, sut.First(x => x.Name == nameof(Person.Name)).Order);
            Assert.Equal(2, sut.First(x => x.Name == nameof(Person.SurName)).Order);
            Assert.Equal(3, sut.First(x => x.Name == nameof(Person.BirthDate)).Order);
            Assert.Equal(4, sut.First(x => x.Name == nameof(Person.Age)).Order);

        }

        [Fact]
        public void Build_WithSingleColumnOrdered_RemainingPropertiesGoesBelow()
        {
            var builder = new ConfigurationBuilder<Person>(new ColumnConfig<Person>(x => x.BirthDate, 1));

            var sut = builder.Build();

            Assert.Equal(1, sut.First(x => x.Name == nameof(Person.BirthDate)).Order);
            Assert.Equal(2, sut.First(x => x.Name == nameof(Person.Name)).Order);
            Assert.Equal(3, sut.First(x => x.Name == nameof(Person.SurName)).Order);
            Assert.Equal(4, sut.First(x => x.Name == nameof(Person.Age)).Order);

        }

        [Fact]
        public void Build_WithRepeatedOrder_ThrowsError()
        {
            var builder = new ConfigurationBuilder<Person>(
                new ColumnConfig<Person>(x => x.BirthDate, 1),
                new ColumnConfig<Person>(x => x.Name, 1)
            );

            Action act = () => builder.Build();
            Assert.Throws<InvalidOperationException>(act);
        }

        [Fact]
        public void Build_IgnoredColumns_AreNotInTheList()
        {
            var builder = new ConfigurationBuilder<Person>(new ColumnConfig<Person>(x => x.Age, true));

            var sut = builder.Build();

            Assert.Collection(sut, 
                x => Assert.Equal(nameof(Person.Name), x.Name),
                x => Assert.Equal(nameof(Person.SurName), x.Name),
                x => Assert.Equal(nameof(Person.BirthDate), x.Name)
            );
        }

    }
}
  