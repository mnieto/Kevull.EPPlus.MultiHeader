using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;

namespace Kevull.EPPLus.MultiHeader.Columns
{
    internal class PropertyNameBuilder<T>
    {
        public PropertyNames Build(Expression<Func<T, object?>> expression)
        {
            var result = new PropertyNames();
            var names = new List<string>();
            var memberExpr = expression.Body as MemberExpression;
            var unaryExpr = expression.Body as UnaryExpression;
            if (memberExpr == null && unaryExpr == null)
                throw new InvalidCastException(expression.Body.ToString());

            if (unaryExpr != null)
                memberExpr = unaryExpr!.Operand as MemberExpression;

            result.Name = memberExpr!.Member.Name;
            while (memberExpr != null && memberExpr.Expression != null && memberExpr.Expression.NodeType == ExpressionType.MemberAccess)
            {
                names.Add(memberExpr.Member.Name);
                memberExpr = memberExpr.Expression as MemberExpression;
                if (names.Count == 1 && memberExpr != null)
                {
                    result.ParentName = memberExpr!.Member.Name;
                    result.ParentType = memberExpr!.Type;
                }
            }
            names.Add(memberExpr!.Member.Name);
            names.Reverse();
            
            result.FullName = string.Join('.', names);
            return result;
        }
    }
}
