using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection.Metadata;
using System.Text;
using System.Threading.Tasks;

namespace EPPLus.MultiHeader.Test
{
    internal record Person(string Name, string Surname, DateTime BirthDate, int? NumOfComputers, Uri? Profile)
    {

    }
}
