using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection.Metadata;
using System.Text;
using System.Threading.Tasks;

namespace EPPLus.MultiHeader.Test
{
    internal record Person(string Name, string SurName, DateTime BirthDate)
    {
        public int Age
        {
            get
            {
                var today = DateTime.Today;
                var age = (today.Year - BirthDate.Date.Year);
                // Go back to the year in which the person was born in case of a leap year
                if (BirthDate.Date > today.AddYears(-age)) age--;
                return age;
            }
        }
    }
}
