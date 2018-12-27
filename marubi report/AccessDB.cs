using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Data;

namespace Marubi
{
    public delegate void StatusDelegate(string status);

    public enum ExcelDataType 
    {
        Product = 1,
        PackingMaterial =2,
        Bom = 3,
        Report = 4,
        Order = 5,
    }

    public class AccessDB
    {
        const string productTableName = "ProductInventory";
        const string partTableName = "PartInventory";
        const string bomTableName = "BOM";
        const string reportTableName = "Product4Report";

        public static StatusDelegate statusDelegate;

        public static DataSet GetDataSource(string xlsConnectionString, ExcelDataType xlsDataType) 
        {
            string workSheetName = "";
            string sql = "";
            string tableName = "";

            if (string.IsNullOrEmpty(xlsConnectionString))
                throw new ArgumentNullException("xlsConnectionString");

            if (xlsDataType == ExcelDataType.Product) 
            {
                workSheetName = "成品库存";
                tableName = productTableName;
                sql = "select [NO#] as RowId, 商品编号 as ProductId, 商品名称 as ProductName, 可用库存数量 as AvailableQty from [{0}]";
            } 
            else if (xlsDataType == ExcelDataType.PackingMaterial) 
            {
                workSheetName = "包材库存";
                tableName = partTableName;
                sql = "select [NO#] as RowId, 商品编号 as PartId, 商品名称 as PartName, 可用库存数量 as AvailableQty from [{0}]";
            }
            else if (xlsDataType == ExcelDataType.Bom) 
            {
                workSheetName = "BOM";
                tableName = bomTableName;
                sql = "select 产品编码 as ProductId, 产品名称 as ProductName, 规格 as ProductSpec, 物料编码 as PartId, 物料名称 as PartName from [{0}]";
            }
            else if (xlsDataType == ExcelDataType.Report) 
            {
                workSheetName = "计划产品";
                tableName = reportTableName;
                sql = "select 商品编号 as ProductId from [{0}]";
            }
            else if (xlsDataType == ExcelDataType.Order) 
            {
                workSheetName = "订单明细";
                tableName = "OrderDetails";
                sql = "select * from [{0}]";
            }

            workSheetName = workSheetName + "$";
            sql = string.Format(sql, workSheetName);

            return DataSource(xlsConnectionString, sql, tableName, workSheetName);
        }         

        private static DataSet DataSource(string xlsConnectionString, string selectCommandText, string dataSetName, string workSheetName)
        {
            using (OleDbConnection xlsConnection = new OleDbConnection(xlsConnectionString))
            {
                try
                {
                    xlsConnection.Open();

                    if (!WorkSheetValidation(xlsConnection, workSheetName))
                    {
                        string message = string.Format("导入的 Excel 数据文件中不存在表单 [{0}], 请确认！", workSheetName.TrimEnd(new char[] { '$' }));

                        throw new Exception(message);
                    }

                    OleDbDataAdapter xlsDataAdapter = new OleDbDataAdapter(selectCommandText, xlsConnection);
                    DataSet productDataSet = new DataSet();
                    xlsDataAdapter.Fill(productDataSet, dataSetName);

                    xlsConnection.Close();

                    return productDataSet;
                }
                catch
                {
                    xlsConnection.Close();
                    throw;
                }
            }
        }

        private static bool WorkSheetValidation(OleDbConnection oleDbConnection, object key)
        {
            DataTable workSheetTable = oleDbConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            workSheetTable.PrimaryKey = new DataColumn[] { workSheetTable.Columns["table_name"] };
            return workSheetTable.Rows.Contains(key);
        }

        public static int SaveDataSource(string accessConnectionString, DataSet sourceDataSet, ExcelDataType xlsDataType, MainForm mainForm)
        {
            string tableName = "";

            if (string.IsNullOrEmpty(accessConnectionString))
                throw new ArgumentNullException("accessConnectionString");

            if (xlsDataType == ExcelDataType.Product)
            {
                tableName = productTableName;
            }
            else if (xlsDataType == ExcelDataType.PackingMaterial)
            {
                tableName = partTableName;
            }
            else if (xlsDataType == ExcelDataType.Bom)
            {
                tableName = bomTableName;
            }
            else if (xlsDataType == ExcelDataType.Report)
            {
                tableName = reportTableName;
            }

            using (OleDbConnection accessConnection = new OleDbConnection(accessConnectionString))
            {
                try
                {
                    accessConnection.Open();

                    Delete(accessConnection, tableName);

                    string selectCommadText = string.Format("select * from {0}", tableName);

                    OleDbDataAdapter mdbDataAdapter = new OleDbDataAdapter(selectCommadText, accessConnection);
                    OleDbCommandBuilder cb = new OleDbCommandBuilder(mdbDataAdapter);
                    DataSet destinationDataSet = new DataSet(tableName);
                    mdbDataAdapter.Fill(destinationDataSet, tableName);

                    if (xlsDataType == ExcelDataType.Bom)
                    {
                        //BOM 表单中,有合并的单元格,需要拆分调整数据
                        SplitColumnData(sourceDataSet);
                    }
                    else if (xlsDataType == ExcelDataType.Product || xlsDataType == ExcelDataType.PackingMaterial)
                    {
                        //删除多行表头
                        DeleteColumnRow(sourceDataSet, mainForm);
                    }

                    foreach (DataRow row in sourceDataSet.Tables[tableName].Rows)
                    {
                        DataRow newRow = destinationDataSet.Tables[tableName].NewRow();

                        foreach (DataColumn col in sourceDataSet.Tables[0].Columns)
                        {
                            string colName = col.ColumnName;

                            //XLS 文件中编号（NO.)信息不需要保存到数据库,跳过
                            if (colName.ToUpper() == "ROWID")
                                continue;
                            newRow[colName] = row[colName];
                        }
                        newRow["UpdtDate"] = DateTime.Now;
                        destinationDataSet.Tables[tableName].Rows.Add(newRow);
                    }
                    accessConnection.Close();
                    return mdbDataAdapter.Update(destinationDataSet, tableName);

                }
                catch
                {
                    accessConnection.Close();
                    throw;
                }

            }
        }

        private static bool IsNumberic(string s)
        {
            try
            {
                int var = Convert.ToInt32(s);
                return true;
            }
            catch
            {
                return false;
            }
        }

        //原始 EXCEL 表单数据中有多行表头, 需要删除
        private static void DeleteColumnRow(DataSet dataSet, MainForm mainForm)
        {
            DataTable dt = dataSet.Tables[0];
            for (int i = 0; i < dt.Rows.Count - 1; i++)
            {
                DataRow row = dataSet.Tables[0].Rows[i];
                if (!IsNumberic(row["RowId"].ToString()))
                {
                    string rowData = "";
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        rowData += " " + dt.Rows[i][j].ToString();
                    }
                    dt.Rows[i].Delete();
                    mainForm.SetProcessStatus(string.Format("    行 {0} 被忽略 ！ 数据: {1}", i + 2, rowData));
                }
            }

            dataSet.AcceptChanges();
        }

        //BOM 表单中,有合并的单元格,需要拆分调整数据
        private static void SplitColumnData(DataSet bomDataSet)
        {
            DataTable bomTable = bomDataSet.Tables[0];

            string productId = "";
            string productName = "";
            string productSpec = "";

            for (int i = 0; i < bomTable.Rows.Count; i++)
            {
                DataRow row = bomTable.Rows[i];
                string pid = row["ProductId"].ToString();

                if (!string.IsNullOrEmpty(pid))
                {
                    productId = pid;
                    productName = row["ProductName"].ToString();
                    productSpec = row["ProductSpec"].ToString();
                }
                else
                {
                    row["ProductId"] = productId;
                    row["ProductName"] = productName;
                    row["ProductSpec"] = productSpec;
                }
            }

            bomDataSet.AcceptChanges();
        }

        private static void Delete(OleDbConnection oleDbConnection, string tableName)
        {
            using (OleDbCommand command = oleDbConnection.CreateCommand())
            {
                command.CommandText = string.Format("delete from {0}", tableName);
                command.CommandType = CommandType.Text;
                command.ExecuteNonQuery();
            }
        }

        public static DataTable GetReportTable(string accessConnectionString)
        {
            using (OleDbConnection accessConnection = new OleDbConnection(accessConnectionString))
            {
                try
                {
                    accessConnection.Open();

                    string selectCommadText = "SELECT Product4Report.ProductId as 产品编号, BOM.ProductName as 产品名称, BOM.ProductSpec as 规格, "
                                            + "ProductInventory.AvailableQty as 成品合计, BOM.PartId as 物料编码, BOM.PartName as 物料名称, "
                                            + "PartInventory.AvailableQty as 包材库存 "
                                            + "FROM ((Product4Report LEFT JOIN BOM ON Product4Report.ProductId = BOM.ProductId) "
                                            + "LEFT JOIN ProductInventory ON BOM.ProductId = ProductInventory.ProductId) "
                                            + "LEFT JOIN PartInventory ON BOM.PartId = PartInventory.PartId ORDER BY BOM.ProductId";
                     
                    OleDbDataAdapter adapter = new OleDbDataAdapter(selectCommadText, accessConnection);
                    OleDbCommandBuilder cb = new OleDbCommandBuilder(adapter);
                    DataSet ds = new DataSet();
                    adapter.Fill(ds);
                    accessConnection.Close();

                    return ds.Tables[0];
                }
                catch
                {
                    accessConnection.Close();
                    throw;
                }
            }
        }
    }
}
