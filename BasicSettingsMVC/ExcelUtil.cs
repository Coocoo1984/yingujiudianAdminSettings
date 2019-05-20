using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace BasicSettingsMVC
{

    public static class ExcelUtil
    {
        //商品
        public const string GoodsModelName = "Goods";
        public const string GoodsDataTableName = "商品";
        public const int GoodsRowStarIndex = 6;
        public const int GoodsColumnStarIndex = 2;
        public static readonly string[] GoodsSheetHeader = { "商品名称", "计量单位", "商品类目", "采购类别", "规格", "是否启用" };
        public static readonly string[] GoodsModelOnlyMappedPropertyArray = { "Name", "GoodsUnitId", "GoodsClassId", "BizTypeId", "Specification", "Disable" };
        public static readonly string[] GoodsModelPropertyArray = { "Name", "GoodsUnitName", "GoodsClassName", "BizTypeName", "Specification", "DisableForShow" };
        //未隐射属性名, <DB隐射属性名, excel列名>
        public static Dictionary<string, Tuple<string, string>> GoodsDictionary = new Dictionary<string, Tuple<string, string>>(){
            { GoodsModelPropertyArray[0], new Tuple<string,string>(GoodsModelOnlyMappedPropertyArray[0], GoodsSheetHeader[0]) },
            { GoodsModelPropertyArray[1], new Tuple<string,string>(GoodsModelOnlyMappedPropertyArray[1], GoodsSheetHeader[1]) },
            { GoodsModelPropertyArray[2], new Tuple<string,string>(GoodsModelOnlyMappedPropertyArray[2], GoodsSheetHeader[2]) },
            { GoodsModelPropertyArray[3], new Tuple<string,string>(GoodsModelOnlyMappedPropertyArray[3], GoodsSheetHeader[3]) },
            { GoodsModelPropertyArray[4], new Tuple<string,string>(GoodsModelOnlyMappedPropertyArray[4], GoodsSheetHeader[4]) },
            { GoodsModelPropertyArray[5], new Tuple<string,string>(GoodsModelOnlyMappedPropertyArray[5], GoodsSheetHeader[5]) }
        };

        //计量单位
        public const string GoodsUnitModelName = "GoodsUnit";
        public const string GoodsUnitDataTableName = "计量单位";
        public const int GoodsUnitRowStarIndex = 6;
        public const int GoodsUnitColumnStarIndex = 2;
        public static readonly string[] GoodsUnitSheetHeader = { "计量单位名称", "规格" };
        public static readonly string[] GoodsUnitModelOnlyMappedPropertyArray = { "Name", "Desc" };
        public static readonly string[] GoodsUnitModelPropertyArray = { "Name", "Desc" };
        //未隐射属性名, <DB隐射属性名, excel列名>
        public static Dictionary<string, Tuple<string, string>> GoodsUnitDictionary = new Dictionary<string, Tuple<string, string>>(){
            { GoodsUnitModelPropertyArray[0], new Tuple<string,string>(GoodsUnitModelOnlyMappedPropertyArray[0], GoodsUnitSheetHeader[0]) },
            { GoodsUnitModelPropertyArray[1], new Tuple<string,string>(GoodsUnitModelOnlyMappedPropertyArray[1], GoodsUnitSheetHeader[1]) }
        };

        //商品类目
        public const string GoodsClassModelName = "GoodsClass";        
        public const string GoodsClassDataTableName = "商品类目";
        public const int GoodsClassRowStarIndex = 6;
        public const int GoodsClassColumnStarIndex = 2;
        public static readonly string[] GoodsClassSheetHeader = { "商品类目名称", "规格", "采购类别", "是否启用" };
        public static readonly string[] GoodsClassModelOnlyMappedPropertyArray = { "Name", "Specification", "BizTypeId","Disable" };
        public static readonly string[] GoodsClassModelPropertyArray = { "Name", "Specification", "BizTypeName", "DisableForShow" };
        //未隐射属性名, <DB隐射属性名, excel列名>
        public static Dictionary<string, Tuple<string, string>> GoodsClassDictionary = new Dictionary<string, Tuple<string, string>>(){
            { GoodsClassModelPropertyArray[0], new Tuple<string,string>(GoodsClassModelOnlyMappedPropertyArray[0], GoodsClassSheetHeader[0]) },
            { GoodsClassModelPropertyArray[1], new Tuple<string,string>(GoodsClassModelOnlyMappedPropertyArray[1], GoodsClassSheetHeader[1]) },
            { GoodsClassModelPropertyArray[2], new Tuple<string,string>(GoodsClassModelOnlyMappedPropertyArray[2], GoodsClassSheetHeader[2]) },
            { GoodsClassModelPropertyArray[3], new Tuple<string,string>(GoodsClassModelOnlyMappedPropertyArray[3], GoodsClassSheetHeader[3]) }
        };

        //采购类别
        public const string BizTypeModelName = "BizType";
        public const string BizTypeDataTableName = "采购类别";
        public const int BizTypeRowStarIndex = 6;
        public const int BizTypeColumnStarIndex = 2;
        public static readonly string[] BizTypeSheetHeader = { "采购类别名称", "规格", "是否启用" };
        public static readonly string[] BizTypeModelOnlyMappedPropertyArray = { "Name", "Desc", "Disable" };
        public static readonly string[] BizTypeModelPropertyArray = { "Name", "Desc", "DisableForShow" };
        //未隐射属性名, <DB隐射属性名, excel列名>
        public static Dictionary<string, Tuple<string, string>> BizTypeDictionary = new Dictionary<string, Tuple<string, string>>(){
            { BizTypeModelPropertyArray[0], new Tuple<string,string>(BizTypeModelOnlyMappedPropertyArray[0], BizTypeSheetHeader[0]) },
            { BizTypeModelPropertyArray[1], new Tuple<string,string>(BizTypeModelOnlyMappedPropertyArray[1], BizTypeSheetHeader[1]) },
            { BizTypeModelPropertyArray[2], new Tuple<string,string>(BizTypeModelOnlyMappedPropertyArray[2], BizTypeSheetHeader[2]) }
        };

        //权限
        public const string RsPermissionModelName = "Usr";
        public const string RsPermissionDataTableName = "权限";
        public const int RsPermissionRowStarIndex = 6;
        public const int RsPermissionColumnStarIndex = 2;
        public static readonly string[] RsPermissionSheetHeader = {
            "微信ID",
            "报价明细查看",
            "报价初审",
            "报价复审",
            "采购初审",
            "采购复审",
            "采购三审",
            "退货审核",
            "库管",
            "报表导出"
        };
        public static readonly string[] RsPermissionModelOnlyMappedPropertyArray = {
            "WechatID",
            "QuoteDetailRead",
            "QuoteAudit",
            "QuoteAudit2",
            "PurchaceAudit",
            "PurchaceAudit2",
            "PurchaceAudit3",
            "ChargeBackAudit",
            "DepotAdmin",
            "ReportExport"
        };
        public static readonly string[] RsPermissionModelPropertyArray = {
            "WechatID",
            "QuoteDetailRead",
            "QuoteAudit",
            "QuoteAudit2",
            "PurchaceAudit",
            "PurchaceAudit2",
            "PurchaceAudit3",
            "ChargeBackAudit",
            "DepotAdmin",
            "ReportExport"
        };
        //未隐射属性名, <DB隐射属性名, excel列名>
        public static Dictionary<string, Tuple<string, string>> RsPermissionDictionary = new Dictionary<string, Tuple<string, string>>(){
            { RsPermissionModelPropertyArray[0], new Tuple<string,string>(RsPermissionModelOnlyMappedPropertyArray[0], RsPermissionSheetHeader[0]) },
            { RsPermissionModelPropertyArray[1], new Tuple<string,string>(RsPermissionModelOnlyMappedPropertyArray[1], RsPermissionSheetHeader[1]) },
            { RsPermissionModelPropertyArray[2], new Tuple<string,string>(RsPermissionModelOnlyMappedPropertyArray[2], RsPermissionSheetHeader[2]) },
            { RsPermissionModelPropertyArray[3], new Tuple<string,string>(RsPermissionModelOnlyMappedPropertyArray[3], RsPermissionSheetHeader[3]) },
            { RsPermissionModelPropertyArray[4], new Tuple<string,string>(RsPermissionModelOnlyMappedPropertyArray[4], RsPermissionSheetHeader[4]) },
            { RsPermissionModelPropertyArray[5], new Tuple<string,string>(RsPermissionModelOnlyMappedPropertyArray[5], RsPermissionSheetHeader[5]) },
            { RsPermissionModelPropertyArray[6], new Tuple<string,string>(RsPermissionModelOnlyMappedPropertyArray[6], RsPermissionSheetHeader[6]) },
            { RsPermissionModelPropertyArray[7], new Tuple<string,string>(RsPermissionModelOnlyMappedPropertyArray[7], RsPermissionSheetHeader[7]) },
            { RsPermissionModelPropertyArray[8], new Tuple<string,string>(RsPermissionModelOnlyMappedPropertyArray[8], RsPermissionSheetHeader[8]) }
        };


        /// <summary>
        /// 从文件流读取DataTable数据源
        /// </summary>
        /// <param name="stream"></param>
        /// <returns></returns>
        public static DataTable GetDataTable(Stream stream)
        {
            XSSFWorkbook workbook;
            try
            {
                workbook = new XSSFWorkbook(stream);
            }
            catch (Exception exception)
            {
                throw exception;
            }
            ISheet sheetAt = workbook.GetSheetAt(0);
            IEnumerator rowEnumerator = sheetAt.GetRowEnumerator();
            DataTable table = new DataTable();
            IRow row = sheetAt.GetRow(0);
            if (row != null)
            {
                for (int j = 0; j < row.LastCellNum; j++)
                {
                    ICell cell = row.GetCell(j);
                    if (cell == null)
                    {
                        table.Columns.Add("cell" + j.ToString());
                    }
                    else
                    {
                        table.Columns.Add(cell.ToString());
                    }
                }
            }
            int count = table.Columns.Count;
            for (int i = 0; rowEnumerator.MoveNext(); i++)
            {
                if (i > 0)
                {
                    IRow current = (IRow)rowEnumerator.Current;
                    DataRow row3 = table.NewRow();
                    for (int k = 0; k < count; k++)
                    {
                        ICell cell2 = current.GetCell(k);
                        if (cell2 == null)
                        {
                            row3[k] = null;
                        }
                        else
                        {
                            row3[k] = cell2.ToString();
                        }
                    }
                    table.Rows.Add(row3);
                }
            }
            return table;
        }

        /// <summary>
        /// 从Sheet按照位置读取表格返回DataTable
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="rowStartIndex">内容开始行index[0-n]</param>
        /// <param name="columnsStartIndex">内容开始列index[0-n]</param>
        /// <param name="headers">表头(一维)</param>
        /// <returns></returns>
        public static DataTable GetDataTable(ISheet sheet, int rowStartIndex, int columnsStartIndex, IEnumerable<string> headers, string dataTableName)
        {
            DataTable result;
            if (string.IsNullOrWhiteSpace(dataTableName))
            {
                result = new DataTable();
            }
            else
            {
                result = new DataTable(dataTableName);
            }
             

            if (headers != null && headers.Count() > 0)
            {
                if (rowStartIndex < 0)
                {
                    rowStartIndex = 0;
                }
                foreach (var item in headers)
                {
                    result.Columns.Add(item, item.GetType());
                }

                IEnumerator rowEnumerator = sheet.GetRowEnumerator();
                for (int i = 0; rowEnumerator.MoveNext(); i++)
                {
                    if (i >= sheet.FirstRowNum + rowStartIndex - 1)
                    {
                        IRow currentRow = (IRow)rowEnumerator.Current;
                        if (currentRow.Cells.Count >= result.Columns.Count)
                        {
                            DataRow dataRow = result.NewRow();
                            for (int j = 0; j < columnsStartIndex + result.Columns.Count; j++)
                            {
                                if (j >= columnsStartIndex)
                                {
                                    ICell cell = currentRow.GetCell(j);
                                    cell?.SetCellType(CellType.String);
                                    if (cell != null &&  cell.StringCellValue.Length > 0)
                                    {
                                        dataRow[j - columnsStartIndex] = cell.StringCellValue;
                                    }
                                    else
                                    {
                                        dataRow = null;
                                    }
                                }
                            }
                            if (dataRow != null)
                            {
                                result.Rows.Add(dataRow);
                            }
                        }
                    }
                }
            }
            return result;
        }

        /// <summary>
        /// 从文件流按照固定规则读取出DataSet集合
        /// </summary>
        /// <param name="stream"></param>
        /// <returns></returns>
        public static DataSet GetDataSet(Stream stream)
        {
            DataSet result = new DataSet();
            XSSFWorkbook workbook;
            try
            {
                workbook = new XSSFWorkbook(stream);
            }
            catch (Exception exception)
            {
                throw exception;
            }

            IEnumerator sheetEnumerator = workbook.GetEnumerator();
            for (int i = 0; sheetEnumerator.MoveNext(); i++)
            {
                //Sheet循环读取
                ISheet sheet = (ISheet)sheetEnumerator.Current;
                switch (sheet.SheetName)
                {
                    case BizTypeDataTableName:
                        result.Tables.Add(ExcelUtil.GetDataTable(sheet, BizTypeRowStarIndex, BizTypeColumnStarIndex, BizTypeSheetHeader, BizTypeDataTableName));
                        break;
                    case GoodsClassDataTableName:
                        result.Tables.Add(ExcelUtil.GetDataTable(sheet, GoodsClassRowStarIndex, GoodsClassColumnStarIndex, GoodsClassSheetHeader, GoodsClassDataTableName));
                        break;
                    case GoodsDataTableName:
                        result.Tables.Add(ExcelUtil.GetDataTable(sheet, GoodsRowStarIndex, GoodsColumnStarIndex, GoodsSheetHeader, GoodsDataTableName));
                        break;
                    case GoodsUnitDataTableName:
                        result.Tables.Add(ExcelUtil.GetDataTable(sheet, GoodsUnitRowStarIndex, GoodsUnitColumnStarIndex, GoodsUnitSheetHeader, GoodsUnitDataTableName));
                        break;
                    case RsPermissionDataTableName:
                        result.Tables.Add(ExcelUtil.GetDataTable(sheet, RsPermissionRowStarIndex, RsPermissionColumnStarIndex, RsPermissionSheetHeader, RsPermissionDataTableName));
                        break;
                }

            }
            return result;
        }

        /// <summary>
        /// 从DataSet按照固定规则写入文件模板
        /// </summary>
        /// <param name="data">DataSet数据（预定的格式）</param>
        /// <param name="stream">模板excel对象（从预定的格式模板生成）</param>
        /// <returns></returns>
        public static int SetDataSet2Workbook(DataSet data, XSSFWorkbook workbook)
        {
            int result = 0;
            if (data != null && data.Tables.Count >  0) 
            {
                IEnumerator sheetEnumerator = workbook.GetEnumerator();
                for (int i = 0; sheetEnumerator.MoveNext(); i++)
                {
                    //Sheet循环写入
                    ISheet sheet = (ISheet)sheetEnumerator.Current;
                    if (data.Tables.Contains(sheet.SheetName))
                    {
                        switch (sheet.SheetName)
                        {
                            case BizTypeDataTableName:
                                result += ExcelUtil.SetDataTable2Sheet(sheet, BizTypeRowStarIndex, BizTypeColumnStarIndex, BizTypeSheetHeader, data.Tables[BizTypeDataTableName]);
                                break;
                            case GoodsClassDataTableName:
                                result += ExcelUtil.SetDataTable2Sheet(sheet, GoodsClassRowStarIndex, GoodsClassColumnStarIndex, GoodsClassSheetHeader, data.Tables[GoodsClassDataTableName]);
                                break;
                            case GoodsDataTableName:
                                result += ExcelUtil.SetDataTable2Sheet(sheet, GoodsRowStarIndex, GoodsColumnStarIndex, GoodsSheetHeader, data.Tables[GoodsDataTableName]);
                                break;
                            case GoodsUnitDataTableName:
                                result += ExcelUtil.SetDataTable2Sheet(sheet, GoodsUnitRowStarIndex, GoodsUnitColumnStarIndex, GoodsUnitSheetHeader, data.Tables[GoodsUnitDataTableName]);
                                break;
                            case RsPermissionDataTableName:
                                result += ExcelUtil.SetDataTable2Sheet(sheet, RsPermissionRowStarIndex, RsPermissionColumnStarIndex, RsPermissionSheetHeader, data.Tables[RsPermissionDataTableName]);
                                break;
                        }
                    }

                }
            }

            return result;
        }

        /// <summary>
        /// 从DataTable写入到Sheet固定区域
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="rowStartIndex">内容开始行index[0-n]</param>
        /// <param name="columnsStartIndex">内容开始列index[0-n]</param>
        /// <param name="headers">表头(一维)</param>
        /// <returns>导出的行数量</returns>
        public static int SetDataTable2Sheet(ISheet sheet, int rowStartIndex, int columnsStartIndex, IEnumerable<string> headers, DataTable dataTable)
        {
            int result = 0;
            int xStart = sheet.FirstRowNum + rowStartIndex - 1;
            
            if (dataTable != null && dataTable.Rows.Count > 0)
            {
                int x = 0;
                foreach(DataRow dr in dataTable.Rows)
                {
                    int y = 0;
                    foreach (DataColumn dc in dataTable.Columns)
                    {
                        ICell cell = sheet.GetRow(xStart+x).GetCell(columnsStartIndex + y);
                        cell.SetCellValue(dr[dc.ColumnName].ToString());
                        y++;
                    }
                    x++;
                }
                result++;
            }

            return result;
        }

    }
}
