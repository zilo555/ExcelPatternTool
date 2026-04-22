using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using ExcelPatternTool.Contracts;
using ExcelPatternTool.Contracts.NPOI.AdvancedTypes;
using ExcelPatternTool.Contracts.Attributes;
using ExcelPatternTool.Contracts.Models;

namespace ExcelPatternTool.Core.NPOI
{
    public class BaseReader : BaseHandler
    {
        private readonly DataFormatter _dataFormatter = new DataFormatter(CultureInfo.InvariantCulture);

        internal T GetDataToObject<T>(IRow row, List<ColumnMetadata> columns) where T : IExcelEntity
        {
            Type objType = typeof(T);
            return (T)GetDataToObject(objType, row, columns);
        }

        internal object GetDataToObject(Type objType, IRow row, List<ColumnMetadata> columns)
        {
            object result = Activator.CreateInstance(objType);

            for (int j = 0; j < columns.Count; j++)
            {
                if (columns[j].ColumnOrder < 0)
                {
                    continue;
                }
                ICell cell = row.GetCell(columns[j].ColumnOrder);
                if (cell == null)
                {
                    Console.WriteLine($"第 {row.RowNum} 行， 第 {columns[j].ColumnOrder} 列  不符合MetaData定义规范，跳过");
                    continue;
                }
                string colTypeDesc = columns[j].PropType.Name.ToLowerInvariant();

                switch (colTypeDesc)
                {
                    case "string":
                        string tmpStr = ExtractStringFromCell(cell);
                        AssignValue(objType, columns[j].PropName, result, tmpStr);
                        break;
                    case "datetime":
                        DateTime tmpDt = ExtractDateFromCell(cell);
                        AssignValue(objType, columns[j].PropName, result, tmpDt);
                        break;
                    case "int":
                    case "int32":
                        double tmpInt = ExtractNumericFromCell(cell);
                        AssignValue(objType, columns[j].PropName, result, Convert.ToInt32(tmpInt));
                        break;

                    case "decimal":
                        double tmpDecimal = ExtractNumericFromCell(cell);
                        AssignValue(objType, columns[j].PropName, result, Convert.ToDecimal(tmpDecimal));
                        break;
                    case "int64":
                    case "long":
                        double tmpLong = ExtractNumericFromCell(cell);
                        AssignValue(objType, columns[j].PropName, result, Convert.ToInt64(tmpLong));
                        break;

                    case "double":
                        double tmpDb = ExtractNumericFromCell(cell);
                        AssignValue(objType, columns[j].PropName, result, Convert.ToDouble(tmpDb));
                        break;
                    case "single":
                        double tmpSingle = ExtractNumericFromCell(cell);
                        AssignValue(objType, columns[j].PropName, result, Convert.ToSingle(tmpSingle));
                        break;
                    case "boolean":
                    case "bool":
                        bool tmpBool = ExtractBooleanFromCell(cell);
                        AssignValue(objType, columns[j].PropName, result, Convert.ToBoolean(tmpBool));
                        break;
                    case "formulatedtype`1":
                        var gType = columns[j].PropType.GenericTypeArguments.FirstOrDefault();
                        var tmpFormularted = ExtractAdvancedFromCell(cell, gType, typeof(FormulatedType<>));
                        if (cell.CellType != CellType.Formula)
                        {
                            switch (colTypeDesc)
                            {
                                case "string":
                                    tmpFormularted.SetValue(ExtractStringFromCell(cell));
                                    break;
                                case "datetime":
                                    tmpFormularted.SetValue(ExtractStringFromCell(cell));
                                    break;
                                case "int":
                                case "int32":
                                    tmpFormularted.SetValue(ExtractNumericFromCell(cell));
                                    break;

                                case "decimal":
                                    tmpFormularted.SetValue(ExtractNumericFromCell(cell));
                                    break;
                                case "int64":
                                case "long":
                                    tmpFormularted.SetValue(ExtractNumericFromCell(cell));
                                    break;

                                case "double":
                                    tmpFormularted.SetValue(ExtractNumericFromCell(cell));
                                    break;
                                case "single":
                                    tmpFormularted.SetValue(ExtractNumericFromCell(cell));
                                    break;
                                case "boolean":
                                case "bool":
                                    tmpFormularted.SetValue(ExtractBooleanFromCell(cell));
                                    break;
                            }
                        }
                        else
                        {
                            (tmpFormularted as IFormulatedType).Formula = cell.CellFormula;
                        }

                        AssignValue(objType, columns[j].PropName, result, tmpFormularted);
                        break;

                    case "commentedtype`1":
                        var commentedType = columns[j].PropType.GenericTypeArguments.FirstOrDefault();
                        var tmpCommented = ExtractAdvancedFromCell(cell, commentedType, typeof(CommentedType<>));
                        if (cell.CellComment != null)
                        {
                            (tmpCommented as ICommentedType).Comment = cell.CellComment.String.String;
                        }

                        AssignValue(objType, columns[j].PropName, result, tmpCommented);
                        break;

                    case "styledtype`1":
                        var styledType = columns[j].PropType.GenericTypeArguments.FirstOrDefault();
                        var tmpStyled = ExtractAdvancedFromCell(cell, styledType, typeof(StyledType<>));

                        (tmpStyled as IStyledType).StyleMetadata = CellStyleToMeta(cell.CellStyle);
                        AssignValue(objType, columns[j].PropName, result, tmpStyled);
                        break;

                    case "fulladvancedtype`1":
                        var fullarmedType = columns[j].PropType.GenericTypeArguments.FirstOrDefault();
                        var tmpFullarmed = ExtractAdvancedFromCell(cell, fullarmedType, typeof(FullAdvancedType<>));
                        if (cell.CellType != CellType.Formula)
                        {
                            switch (colTypeDesc)
                            {
                                case "string":
                                    tmpFullarmed.SetValue(ExtractStringFromCell(cell));
                                    break;
                                case "datetime":
                                    tmpFullarmed.SetValue(ExtractStringFromCell(cell));
                                    break;
                                case "int":
                                case "int32":
                                    tmpFullarmed.SetValue(ExtractNumericFromCell(cell));
                                    break;

                                case "decimal":
                                    tmpFullarmed.SetValue(ExtractNumericFromCell(cell));
                                    break;
                                case "int64":
                                case "long":
                                    tmpFullarmed.SetValue(ExtractNumericFromCell(cell));
                                    break;

                                case "double":
                                    tmpFullarmed.SetValue(ExtractNumericFromCell(cell));
                                    break;
                                case "single":
                                    tmpFullarmed.SetValue(ExtractNumericFromCell(cell));
                                    break;
                                case "boolean":
                                case "bool":
                                    tmpFullarmed.SetValue(ExtractBooleanFromCell(cell));
                                    break;
                            }

                        }
                        else
                        {
                            if (cell.CellFormula != null)
                            {
                                (tmpFullarmed as IFormulatedType).Formula = cell.CellFormula;
                            }
                        }
                        if (cell.CellComment != null)
                        {
                            (tmpFullarmed as ICommentedType).Comment = cell.CellComment.String.String;
                        }
                        (tmpFullarmed as IStyledType).StyleMetadata = CellStyleToMeta(cell.CellStyle);

                        AssignValue(objType, columns[j].PropName, result, tmpFullarmed);
                        break;
                    default:
                        double tmpDef = ExtractNumericFromCell(cell);
                        AssignValue(objType, columns[j].PropName, result, tmpDef);
                        break;

                }
            }

            AssignValue(objType, "RowNumber", result, row.RowNum);

            return result;
        }

        private bool ExtractBooleanFromCell(ICell cell)
        {
            if (cell == null)
            {
                return false;
            }

            if (cell.CellType == CellType.Boolean)
            {
                return cell.BooleanCellValue;
            }

            if (cell.CellType == CellType.Error || cell.CellType == CellType.Blank || cell.CellType == CellType._None)
            {
                return false;
            }

            string formattedValue = GetFormattedCellValue(cell);
            if (string.IsNullOrWhiteSpace(formattedValue))
            {
                return false;
            }

            if (bool.TryParse(formattedValue, out bool boolValue))
            {
                return boolValue;
            }

            if (double.TryParse(formattedValue, NumberStyles.Any, CultureInfo.InvariantCulture, out double numericValue)
                || double.TryParse(formattedValue, NumberStyles.Any, CultureInfo.CurrentCulture, out numericValue))
            {
                return numericValue > 0;
            }

            return false;
        }

        private double ExtractNumericFromCell(ICell cell)
        {
            if (cell == null)
            {
                return 0;
            }

            if (cell.CellType == CellType.Boolean)
            {
                return cell.BooleanCellValue ? 1 : 0;
            }

            if (cell.CellType == CellType.Error)
            {
                return cell.ErrorCellValue;
            }

            string formattedValue = GetFormattedCellValue(cell);
            if (string.IsNullOrWhiteSpace(formattedValue))
            {
                return 0;
            }

            if (double.TryParse(formattedValue, NumberStyles.Any, CultureInfo.InvariantCulture, out double value)
                || double.TryParse(formattedValue, NumberStyles.Any, CultureInfo.CurrentCulture, out value))
            {
                return value;
            }

            return 0;
        }

        private DateTime ExtractDateFromCell(ICell cell)
        {
            if (cell == null)
            {
                return default;
            }

            if ((cell.CellType == CellType.Numeric || cell.CellType == CellType.Formula) && DateUtil.IsCellDateFormatted(cell))
            {
                return cell.DateCellValue ?? default;
            }

            string formattedValue = GetFormattedCellValue(cell);
            if (string.IsNullOrWhiteSpace(formattedValue))
            {
                return default;
            }

            if (DateTime.TryParse(formattedValue, CultureInfo.CurrentCulture, DateTimeStyles.None, out DateTime currentCultureDate)
                || DateTime.TryParse(formattedValue, CultureInfo.InvariantCulture, DateTimeStyles.None, out currentCultureDate))
            {
                return currentCultureDate;
            }

            return default;
        }

        private string ExtractStringFromCell(ICell cell)
        {
            if (cell == null)
            {
                return string.Empty;
            }

            if (cell.CellType == CellType.Error)
            {
                return "Error Code:" + cell.ErrorCellValue.ToString();
            }

            return GetFormattedCellValue(cell);
        }

        private string GetFormattedCellValue(ICell cell)
        {
            try
            {
                return _dataFormatter.FormatCellValue(cell) ?? string.Empty;
            }
            catch (Exception ex)
            {
                throw new Exception($"读取单元格格式化文本失败。Row={cell.RowIndex}, Col={cell.ColumnIndex}, CellType={cell.CellType}", ex);
            }
        }

        private IAdvancedType ExtractDateFromFomular<T>(ICell cell, Type iType) where T : struct
        {
            var value = IAdvancedTypeFactory(iType, typeof(T));
            if (cell.CellType == CellType.Formula)
            {
                var TType = typeof(T);
                string colTypeDesc = TType.Name.ToLowerInvariant();
                T realValue;
                switch (colTypeDesc)
                {
                    case "string":
                        string tmpStr = ExtractStringFromCell(cell);
                        realValue = (T)Convert.ChangeType(tmpStr, TType);
                        break;
                    case "datetime":
                        DateTime tmpDt = ExtractDateFromCell(cell);
                        realValue = (T)Convert.ChangeType(tmpDt, TType);
                        break;
                    case "int":
                    case "int32":
                        double tmpInt = ExtractNumericFromCell(cell);
                        realValue = (T)Convert.ChangeType(tmpInt, TType);
                        break;

                    case "decimal":
                        double tmpDecimal = ExtractNumericFromCell(cell);
                        realValue = (T)Convert.ChangeType(tmpDecimal, TType);
                        break;
                    case "int64":
                    case "long":
                        double tmpLong = ExtractNumericFromCell(cell);
                        realValue = (T)Convert.ChangeType(tmpLong, TType);
                        break;

                    case "double":
                        double tmpDb = ExtractNumericFromCell(cell);
                        realValue = (T)Convert.ChangeType(tmpDb, TType);
                        break;
                    case "single":
                        double tmpSingle = ExtractNumericFromCell(cell);
                        realValue = (T)Convert.ChangeType(tmpSingle, TType);
                        break;
                    case "boolean":
                    case "bool":
                        bool tmpBool = ExtractBooleanFromCell(cell);
                        realValue = (T)Convert.ChangeType(tmpBool, TType);
                        break;
                    default:
                        realValue = new T();
                        break;

                }

                value.SetValue(realValue);
            }

            return value;
        }


        private IAdvancedType IAdvancedTypeFactory(Type iType, Type GenericType)
        {
            var type = iType.MakeGenericType(GenericType);
            IAdvancedType result = Activator.CreateInstance(type) as IAdvancedType;
            return result;
        }

        private IAdvancedType ExtractAdvancedFromCell(ICell cell, Type type, Type iType)
        {
            var value = IAdvancedTypeFactory(iType, type);

            var TType = type;
            string colTypeDesc = TType.Name.ToLowerInvariant();
            switch (colTypeDesc)
            {
                case "string":
                    string tmpStr = ExtractStringFromCell(cell);
                    value.SetValue(tmpStr);
                    break;
                case "datetime":
                    DateTime tmpDt = ExtractDateFromCell(cell);
                    value.SetValue(tmpDt);
                    break;
                case "int":
                case "int32":
                    double tmpInt = ExtractNumericFromCell(cell);
                    value.SetValue(Convert.ToInt32(tmpInt));
                    break;

                case "decimal":
                    double tmpDecimal = ExtractNumericFromCell(cell);
                    value.SetValue(Convert.ToDecimal(tmpDecimal));
                    break;
                case "int64":
                case "long":
                    double tmpLong = ExtractNumericFromCell(cell);
                    value.SetValue(Convert.ToInt64(tmpLong));
                    break;

                case "double":
                    double tmpDb = ExtractNumericFromCell(cell);
                    value.SetValue(tmpDb);
                    break;
                case "single":
                    double tmpSingle = ExtractNumericFromCell(cell);
                    value.SetValue(Convert.ToSingle(tmpSingle));
                    break;
                case "boolean":
                case "bool":
                    bool tmpBool = ExtractBooleanFromCell(cell);
                    value.SetValue(Convert.ToBoolean(tmpBool));
                    break;
                default:
                    value = new FormulatedType<int>();
                    break;
            }
            return value;
        }


        private void AssignValue(Type objType, string propertyName, object instance, object data)
        {
            objType.InvokeMember(propertyName,
                BindingFlags.DeclaredOnly |
                BindingFlags.Public | BindingFlags.NonPublic |
                BindingFlags.Instance | BindingFlags.SetProperty, null, instance, new object[] { data });
        }

        internal List<ColumnMetadata> GetTypeDefinition(Type type)
        {
            List<ColumnMetadata> columns = new List<ColumnMetadata>();
            foreach (var prop in type.GetProperties())
            {
                var tmp = new ColumnMetadata();
                var attrs = Attribute.GetCustomAttributes(prop);
                tmp.PropName = prop.Name;
                tmp.PropType = prop.PropertyType;
                tmp.ColumnName = prop.Name;
                tmp.ColumnOrder = -1;
                foreach (var attr in attrs)
                {
                    if (attr is ImportableAttribute)
                    {
                        ImportableAttribute attribute = (ImportableAttribute)attr;
                        tmp.ColumnName = attribute.Name;
                        tmp.ColumnOrder = attribute.Order;
                        tmp.Ignore = attribute.Ignore;
                    }
                }
                if (!tmp.Ignore)
                {
                    columns.Add(tmp);
                }
            }
            return columns.OrderBy(x => x.ColumnOrder).ToList();
        }
    }
}
