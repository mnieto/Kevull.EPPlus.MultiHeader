using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection.Metadata;
using System.Text;
using System.Threading.Tasks;

namespace Kevull.EPPLus.MultiHeader.Test
{
    internal enum Gender
    {
        NotSpecified,
        Male,
        Female
    }

    internal record Person(string Name, string Surname, DateTime BirthDate, Gender Gender, int? NumOfComputers, Uri? Profile, TimeOnly AlarmTime)
    {
        internal static List<Person> BuildPeopleList()
        {
            return new List<Person>
            {
                new Person("Médiamass","Large", DateTime.Parse("2017/05/28"), Gender.Male,  null, null, new TimeOnly(8,0)),
                new Person("Aimée","Bateson", DateTime.Parse("1958/06/07"), Gender.Female, 2, new Uri("https://github.com/"), new TimeOnly(7, 0))
            };
        }
    }


}
