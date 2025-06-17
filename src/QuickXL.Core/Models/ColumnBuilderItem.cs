using QuickXL.Core.Settings.Columns;
using System.Linq.Expressions;

namespace QuickXL.Core.Models
{
    internal class ColumnBuilderItem<TDto>(
        Expression<Func<TDto, object>> propertySelector, 
        ColumnSettings columnSettings) where TDto : class, new()
    {
        public string HeaderName { get; set; } = columnSettings.HeaderName ?? GetPropertyName(propertySelector);

        public Expression<Func<TDto, object>> PropertySelector { get; set; } = propertySelector;

        public ColumnSettings ColumnSettings { get; set; } = columnSettings;

        public readonly Func<TDto, object?> ValueGetter = propertySelector.Compile();

        private static string GetPropertyName(Expression<Func<TDto, object>> propertySelector)
        {
            if (propertySelector.Body is MemberExpression member)
            {
                return member.Member.Name;
            }

            if (propertySelector.Body is UnaryExpression unary && unary.Operand is MemberExpression unaryMember)
            {
                return unaryMember.Member.Name;
            }

            throw new ArgumentException("Invalid property selector expression");
        }

        public string GetPropertyName() => 
            GetPropertyName(PropertySelector);
    }
}
