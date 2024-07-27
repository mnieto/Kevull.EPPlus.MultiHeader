using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPLus.MultiHeader.Columns
{
    internal class PropertyNames
    {
        public string Name { get; set; } = "";
        public string FullName { get; set; } = "";
        public string? ParentName { get; set; }
        public Type? ParentType { get; set; }
    }
}
