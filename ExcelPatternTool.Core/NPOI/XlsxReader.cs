using System;
using System.Collections.Generic;
using System.IO;
using ExcelPatternTool.Contracts;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ExcelPatternTool.Core.NPOI
{
    public class XlsxReader : BaseReader, IReader
    {
        MemoryStream mem;
        private ISheet sheet;
        public XlsxReader(byte[] data)
        {
            mem = new MemoryStream(data);
            var document = new XSSFWorkbook(mem);
            Document = document;
        }

        public XlsxReader(string filePath)
        {
            mem = null;
            var document = new XSSFWorkbook(filePath);
            Document = document;
        }

        public IEnumerable<T> ReadRows<T>(IImportOption importOption) where T : IExcelEntity
        {

            var columns = GetTypeDefinition(typeof(T));
            sheet = Document.GetSheet(importOption.SheetName);
            if (sheet == null)
            {
                throw new Exception($"没找到名称为{importOption.SheetName}的Sheet");

            }
            int firstRow = sheet.FirstRowNum;
            int lastRow = sheet.LastRowNum;
            int startRow = firstRow + importOption.SkipRows;
            List<T> result = new(Math.Max(lastRow - startRow + 1, 0));
            for (int i = startRow; i <= lastRow; i++)
            {

                T objectInstance;
                IRow row = sheet.GetRow(i);
                if (row != null)
                {
                    try
                    {
                        objectInstance = GetDataToObject<T>(row, columns);

                    }
                    catch (Exception e)
                    {
                        throw new Exception($"处理行失败,位置{row.RowNum}:{e.Message}", e);
                    }
                    result.Add(objectInstance);
                }
            }
            return result;

        }

        public IEnumerable<T> ReadRows<T>(int sheetNumber, int rowsToSkip) where T : IExcelEntity
        {
            var columns = GetTypeDefinition(typeof(T));
            if (sheetNumber < 0 || sheetNumber >= Document.NumberOfSheets)
            {
                throw new Exception($"没找到Index为{sheetNumber}的Sheet");
            }
            sheet = Document.GetSheetAt(sheetNumber);
            int firstRow = sheet.FirstRowNum;
            int lastRow = sheet.LastRowNum;
            int startRow = firstRow + rowsToSkip;
            List<T> result = new(Math.Max(lastRow - startRow + 1, 0));
            for (int i = startRow; i <= lastRow; i++)
            {

                T objectInstance;
                IRow row = sheet.GetRow(i);
                if (row != null)
                {
                    try
                    {
                        objectInstance = GetDataToObject<T>(row, columns);

                    }
                    catch (Exception e)
                    {
                        throw new Exception($"处理行失败,位置{row.RowNum}:{e.Message}", e);
                    }
                    result.Add(objectInstance);
                }

            }
            return result;
        }


        public IEnumerable<IExcelEntity> ReadRows(Type entityType, IImportOption importOption)
        {

            var columns = GetTypeDefinition(entityType);
            sheet = Document.GetSheet(importOption.SheetName);
            if (sheet == null)
            {
                throw new Exception($"没找到名称为{importOption.SheetName}的Sheet");

            }
            int firstRow = sheet.FirstRowNum;
            int lastRow = sheet.LastRowNum;
            int startRow = firstRow + importOption.SkipRows;
            List<IExcelEntity> result = new(Math.Max(lastRow - startRow + 1, 0));
            for (int i = startRow; i <= lastRow; i++)
            {

                IExcelEntity objectInstance;
                IRow row = sheet.GetRow(i);
                if (row != null)
                {
                    try
                    {
                        objectInstance = (IExcelEntity)GetDataToObject(entityType, row, columns);

                    }
                    catch (Exception e)
                    {
                        throw new Exception($"处理行失败,位置{row.RowNum}:{e.Message}", e);
                    }
                    result.Add(objectInstance);
                }
            }
            return result;

        }

        public IEnumerable<IExcelEntity> ReadRows(Type entityType, int sheetNumber, int rowsToSkip)
        {
            var columns = GetTypeDefinition(entityType);
            if (sheetNumber < 0 || sheetNumber >= Document.NumberOfSheets)
            {
                throw new Exception($"没找到Index为{sheetNumber}的Sheet");
            }
            sheet = Document.GetSheetAt(sheetNumber);
            int firstRow = sheet.FirstRowNum;
            int lastRow = sheet.LastRowNum;
            int startRow = firstRow + rowsToSkip;
            List<IExcelEntity> result = new(Math.Max(lastRow - startRow + 1, 0));
            for (int i = startRow; i <= lastRow; i++)
            {

                IExcelEntity objectInstance;
                IRow row = sheet.GetRow(i);
                if (row != null)
                {
                    try
                    {
                        objectInstance = (IExcelEntity)GetDataToObject(entityType, row, columns);
                    }
                    catch (Exception e)
                    {
                        throw new Exception($"处理行失败,位置{row.RowNum}:{e.Message}", e);
                    }
                    result.Add(objectInstance);
                }

            }
            return result;
        }

    }
}
