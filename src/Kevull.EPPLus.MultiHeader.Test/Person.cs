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

    internal record Person(string Name, string Surname, DateTime BirthDate, Gender Gender, int? NumOfComputers, Uri? Profile)
    {

    }
}
