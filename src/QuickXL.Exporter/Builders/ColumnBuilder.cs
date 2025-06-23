using Ardalis.GuardClauses;
using QuickXL.Exporter.Helpers;
using QuickXL.Exporter.Models;
using QuickXL.Exporter.Styles;
using System.Linq.Expressions;

namespace QuickXL.Exporter
{
    public sealed class ColumnBuilder<TDto>
        where TDto : class, new()
    {
        internal IList<ColumnBuilderItem<TDto>> ColumnBuilderItems { get; set; }

        internal ExportBuilder<TDto> ExportBuilder;

        internal XLGeneralStyle? XLGeneralStyle { get; set; }

        internal ColumnBuilder(ExportBuilder<TDto> exportBuilder)
        {
            ExportBuilder = exportBuilder;
            ColumnBuilderItems = [];
        }


        public ColumnBuilder<TDto> AddColumn(Expression<Func<TDto, object>> propertySelector, Action<ColumnSettings>? configuration = null)
        {
            Guard.Against.Null(ExportBuilder);
            Guard.Against.Null(ColumnBuilderItems);

            ColumnSettings columnSettings = new();

            configuration?.Invoke(columnSettings);

            ColumnBuilderItems.Add(new(propertySelector, columnSettings));

            return this;
        }

        public ColumnBuilder<TDto> AddGeneralStyle(XLGeneralStyle xlGeneralStyle)
        {
            XLGeneralStyle = xlGeneralStyle;

            return this;
        }

        public byte[] Export()
        {
            return new OpenXmlWorkbookHelper<TDto>()
                       .CreateWorkbook(ExportBuilder);
        }
    }
}