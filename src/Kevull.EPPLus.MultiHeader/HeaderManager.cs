using Kevull.EPPLus.MultiHeader.Columns;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;
using System.Security.Cryptography.Pkcs;

namespace Kevull.EPPLus.MultiHeader
{
    /// <summary>
    /// Stores information about the columns to be shown and build the needed header structure
    /// </summary>
    /// <remarks>
    /// <para>As the source Type can have nested properties, this structure is recursive</para>
    /// <para>Intended for internal use. Use <see cref="HeaderManager{T}"/> instead.</para>
    /// </remarks>
    public class HeaderManager
    {

        /// <summary>
        /// First row used for headers
        /// </summary>
        public int FirstRow = 1;

        /// <summary>
        /// Leftmost column used in the report
        /// </summary>
        public int FirstColumn = 1;

        /// <summary>
        /// List of columns. Each column stores the given customization during the configuration phase
        /// </summary>
        /// <remarks>
        /// <para>This property will reference ALL configured columns, independently of the nested level</para>
        /// <para>To get the properties of this level see the protected property <see cref="DirectColumns"/></para>
        /// </remarks>
        public List<ColumnInfo> Columns { get; set; }

        /// <summary>
        /// If <c>true</c> the configuration will find the end of a previous report and 
        /// </summary>
        public bool AppendToExistingReport { get; internal set; }

        /// <summary>
        /// Porperties, by property name, of the source Type
        /// </summary>
        public Dictionary<string, PropertyInfo>? Properties { get; set; }

        /// <summary>
        /// Shows or not autofilter on last header row
        /// </summary>
        public bool AutoFilter { get; set; } = true;

        /// <summary>
        /// Gets the number of rows of the Header. That is the deep of nested properties in the source Type
        /// </summary>
        public int Height { get; internal set; }
        
        /// <summary>
        /// Gets the number of columns needed to represent the object of the source Type
        /// </summary>
        public int Width => Columns.Sum(x => x.Width);

        /// <summary>
        /// <see cref="Type"/> of the source Type
        /// </summary>
        protected Type ObjectType { get; private set; }

        /// <summary>
        /// Configured columns that belong to this level
        /// </summary>
        protected List<ColumnInfo> DirectColumns => Columns.Where(x => x.Name == x.FullName).ToList();

        /// <summary>
        /// Returns a list of configured columns that belong to the next level. WARN: This operation is NOT idempotent
        /// </summary>
        /// <param name="deep">current deep. That is, the method will return the inmediate children columns</param>
        protected List<ColumnInfo> ChildColumns(int deep)
        {
            var result = new List<ColumnInfo>();
            foreach(var col in Columns.Where(x => x.Deep > deep))
            {
                col.FullName = string.Join('.', col.FullName.Split('.').Skip(deep));
                result.Add(col);
            }
            return result;
        }

        /// <summary>
        /// Ctor. Invoked when using default configuration
        /// </summary>
        /// <remarks>Intended for internal use. Use <see cref="HeaderManager{T}.HeaderManager()"/> instead</remarks>
        public HeaderManager(Type type) : this(type, 1, 1) { }

        /// <summary>
        /// Ctor. Invoked when using configured columns
        /// </summary>
        /// <remarks>Intended for internal use. Use <see cref="HeaderManager{T}.HeaderManager(List{ColumnInfo})"/>i nstead</remarks>
        public HeaderManager(Type type, List<ColumnInfo> columns)
        {
            ObjectType = type;
            Columns = columns;
        }

        /// <summary>
        /// Ctor. Invoked mainly from nested property levels
        /// </summary>
        protected HeaderManager(Type type, int index, int deep, List<ColumnInfo>? columns = null)
        {
            ObjectType = type;
            Columns = columns ?? new List<ColumnInfo>();
            BuildHeaders(index, deep);
        }

        /// <summary>
        /// Build the <see cref="Columns"/> and header structure
        /// </summary>
        internal void BuildHeaders()
        {
            BuildHeaders(FirstColumn, 1);
        }

        private void BuildHeaders(int column, int deep)
        {
            Height = deep;
            var properties = GetProperties(ObjectType.GetTypeInfo());
            var result = BuildFromDefinedColumns();
            result.AddRange(BuildFromRemainingProperties(properties.Where(x => !x.Used)));
            result.RemoveAll(x => x.Ignore);
            result.AddRange(Columns.Where(x => x.Deep > deep));

            Columns = result;
            Properties = properties.Where(x => !x.Ignored)
                .Select(x => x.Info)
                .ToDictionary(x => x.Name, x => x);


            List<ColumnInfo> BuildFromDefinedColumns() {
                var result = new List<ColumnInfo>(DirectColumns.Where(x => x.Order.HasValue));
                foreach (var colInfo in result)
                {
                    if (colInfo.Ignore)
                    {
                        var property = properties.First(x => x.Info.Name == colInfo.Name);
                        property.Ignored = colInfo.Ignore;
                        property.Used = true;
                        continue;
                    }

                    if (!colInfo.IsDynamic)
                    {
                        var property = properties.First(x => x.Info.Name == colInfo.Name);
                        property.Ignored = colInfo.Ignore;
                        if (IsNestedObject(property.Info.PropertyType))
                        {
                            colInfo.Header = new HeaderManager(property.Info.PropertyType, column, deep + 1, ChildColumns(deep));
                            Height = Math.Max(Height, colInfo.Header.Height);
                        }
                        if (colInfo.IsMultiValue)
                        {
                            Height = Math.Max(Height, deep + 1);
                        }
                        if (colInfo.StyleName == null)
                        {
                            colInfo.StyleName = GetDefaultStyleNameForDataType(property.Info);
                        }
                        property.Used = true;
                    }
                    colInfo.Index = column;
                    column += colInfo.Width;
                }
                return result;
            }

            List<ColumnInfo> BuildFromRemainingProperties(IEnumerable<ObjectProperty> properties)
            {
                var result = new List<ColumnInfo>();
                foreach (var property in properties)
                {
                    var colInfo = ColumnInfoFactory(property.Info);
                    if (colInfo.Ignore)
                        continue;
                    result.Add(colInfo);
                    if (IsNestedObject(property.Info.PropertyType))
                    {
                        colInfo.Header = new HeaderManager(property.Info.PropertyType, column, deep + 1, ChildColumns(deep));
                        Height = Math.Max(Height, colInfo.Header.Height);
                    }
                    if (colInfo.IsMultiValue)
                    {
                        Height = Math.Max(Height, deep + 1);
                    }
                    colInfo.Index = column;
                    column += colInfo.Width;
                }
                return result;
            }

        }

        private List<ObjectProperty> GetProperties(TypeInfo typeInfo)
        {
            var result = new List<ObjectProperty>();
            foreach (var item in typeInfo.GetProperties())
            {
                result.Add(new ObjectProperty(item));
            }
            return result;
        }

        private bool IsNestedObject(Type type)
        {
            if (typeof(IDictionary).IsAssignableFrom(type))
            {
                type = type.GenericTypeArguments[1];
            }
            return type != typeof(string) &&
                   (type.IsClass || type.IsInterface) &&
                   type != typeof(Uri);
        }

        private ColumnInfo ColumnInfoFactory(PropertyInfo property)
        {
            var column = DirectColumns.FirstOrDefault(x => x.Name == property.Name && !x.IsDynamic);
            if (column != null)
            {
                return column;
            }
            if (property.PropertyType != typeof(string) && typeof(IEnumerable).IsAssignableFrom(property.PropertyType))   //IList<T>, IDictionary<K,T>, IDictionary and string implement IEnumerable
                throw new InvalidOperationException($"Column {property.Name} is IEnumerable and must be preconfigured. See ConfigurationBuilder.Configure.");

            string? styleName = GetDefaultStyleNameForDataType(property);
            return new ColumnInfo(property.Name, styleName: styleName);
        }

        private string? GetDefaultStyleNameForDataType(PropertyInfo property)
        {
            string? styleName = null;
            if (property.PropertyType == typeof(DateOnly) || property.PropertyType == typeof(DateTime))
            {
                styleName = StyleNames.DateStyleName;
            }
            else if (property.PropertyType == typeof(TimeOnly))
            {
                styleName = StyleNames.TimeStyleName;
            }
            return styleName;
        }

        [DebuggerDisplay("{Info.Name}")]
        private class ObjectProperty
        {
            public ObjectProperty(PropertyInfo property)
            {
                Info = property;
            }
            public PropertyInfo Info { get; set; }
            public bool Used { get; set; }
            public bool Ignored { get; set; }
        }
    }

    /// <summary>
    /// Stores information about the columns to be shown and build the needed header structure
    /// </summary>
    /// <remarks>As the source Type can have nested properties, this structure is recursive</remarks>
    public class HeaderManager<T> : HeaderManager
    {
        /// <summary>
        /// Ctor. Invoked when using default configuration
        /// </summary>
        public HeaderManager() : base(typeof(T)) { }

        /// <summary>
        /// Ctor. Invoked when using configured columns
        /// </summary>
        public HeaderManager(List<ColumnInfo> columns): base(typeof(T), columns) { }
    }
}