using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelManageLib.Common;

namespace ExcelManageLib.ExcelDataAccess
{
    public class ExcelTableDataAccess
    {
        /// <summary>
        /// using the OLEDB.Excel to get DS from Excel,
        /// </summary>
        /// <param name="Path"></param>
        /// <param name="SheetName"></param>
        /// <param name="ColumnName"></param>
        /// <param name="HDR"></param>
        /// <param name="WhereSql"></param>
        /// <returns></returns>
        public DataSet ExcelToDS(string Path, string SheetName, string ColumnName, string HDR, string WhereSql = null)
        {
            string strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Path + ";Extended Properties=Excel 12.0;";
            if (HDR.Equals("true")) strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Path + ";Extended Properties='Excel 12.0;HDR=NO;IMEX=1';";
            DataSet ds = null;
            OleDbConnection conn = new OleDbConnection(strConn);
            var excutedSqlString = String.Empty;
            try
            {
                conn.Open();
                string strExcel = "";
                OleDbDataAdapter myCommand = null;

                strExcel = "select " + ColumnName + " from [" + SheetName + "$]";
                if (!string.IsNullOrEmpty(WhereSql)) strExcel = strExcel + " " + WhereSql;
                excutedSqlString = strExcel;
                myCommand = new OleDbDataAdapter(strExcel, strConn);
                ds = new DataSet();
                myCommand.Fill(ds, "table1");
                conn.Close();
            }
            catch (Exception ex)
            {
               conn.Close();
               var ErrorMessage = string.Format("方法：{0},关键参数：{1}，SQl Script {2},错误提示：{3},堆栈信息：{4}", "ExcelToDS", Path, excutedSqlString, ex.Message, ex.StackTrace);
               Log.Write(ErrorMessage);
               throw ex;
            }
            return ds;
        }


        /// <summary>
        ///  using the OLEDB.Excel to get DS from Excel, the head clumn name is not  first row.
        /// </summary>
        /// <param name="Path"></param>
        /// <param name="SheetName"></param>
        /// <param name="columnNameStr"></param>
        /// <param name="deepHeader"></param>
        /// <returns></returns>
        public DataSet ExcelToDS(string Path, string SheetName, string columnNameStr, int deepHeader = 8)
        {

            try
            {
                DataSet ds = ExcelToDS(Path, SheetName, "*", "true", "");

                string[] columnNameList = columnNameStr.Split(',');
                int[] columnNumberList = new int[columnNameList.Length];

                for (int i = 0; i < columnNumberList.Length; i++)
                {
                    columnNumberList[i] = -1;  //如果匹配不了标题，默认值为-1
                }

                DataTable dt = ds.Tables[0];
                int colCount = dt.Rows[0].ItemArray.Count();

                for (int i = 0; i < deepHeader; i++)  //遍历每行
                {
                    var currentDataRow = dt.Rows[i];
                    for (int j = 0; j < colCount; j++)  //遍历每列
                    {
                        string headerName = Convert.ToString(dt.Rows[i].ItemArray[j]);
                        for (int k = 0; k < columnNameList.Length; k++)
                        {
                            string columnName = columnNameList[k];
                            if (columnName == headerName) //如果父或子标题等于用户需要显示的标题相同，记录下列index
                            {
                                columnNumberList[k] = j;
                            }
                        }
                    }
                }


                ///Create A DataSet to return 
                DataTable retDataTable = new DataTable();
                retDataTable.TableName = SheetName;
                for (int i = 0; i < columnNameList.Length; i++)
                {
                    retDataTable.Columns.Add(columnNameList[i], typeof(String));
                }

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    retDataTable.Rows.Add();
                    for (int j = 0; j < columnNumberList.Length; j++)
                    {

                        if (columnNumberList[j] != -1)
                        {
                            retDataTable.Rows[i][j] = Convert.ToString(dt.Rows[i][columnNumberList[j]]);
                        }
                        else
                        {
                            retDataTable.Rows[i][j] = "N/A";
                        }

                    }
                }

                DataSet retDataSet = new DataSet();
                retDataSet.Tables.Add(retDataTable);

                return retDataSet;
            }
            catch (Exception ex)
            {
                var ErrorMessage = string.Format("方法：{0},关键参数：{1} ,错误提示：{2},堆栈信息：{3}", "ExcelToDS", Path, ex.Message, ex.StackTrace);
                Log.Write(ErrorMessage);
                throw ex;
            }

        }


        /// <summary>
        ///  using the OLEDB.Excel to get clumn name index,
        /// </summary>
        /// <param name="Path"></param>
        /// <param name="SheetName"></param>
        /// <param name="columnNameStr"></param>
        /// <param name="deepHeader"></param>
        /// <returns></returns>
        public int[] GetClumnNameIndexArray(string Path, string SheetName, string columnNameStr, int deepHeader = 2)
        {

            try
            {
                DataSet ds = ExcelToDS(Path, SheetName, "top "+ deepHeader + " * " , "true", "");

                string[] columnNameList = columnNameStr.Split(',');
                int[] columnNumberList = new int[columnNameList.Length];

                for (int i = 0; i < columnNumberList.Length; i++)
                {
                    columnNumberList[i] = -1;  //如果匹配不了标题，默认值为-1
                }

                DataTable dt = ds.Tables[0];
                int colCount = dt.Rows[0].ItemArray.Count();

                for (int i = 0; i < deepHeader; i++)  //遍历每行
                {
                    var currentDataRow = dt.Rows[i];
                    for (int j = 0; j < colCount; j++)  //遍历每列
                    {
                        string headerName = Convert.ToString(dt.Rows[i].ItemArray[j]);
                        for (int k = 0; k < columnNameList.Length; k++)
                        {
                            string columnName = columnNameList[k];
                            if (columnName.Trim() == headerName.Trim()) //如果父或子标题等于用户需要显示的标题相同，记录下列index
                            {
                                columnNumberList[k] = j;
                                break;
                            }
                        }
                    }
                }
                return columnNumberList;
            }
            catch (Exception ex)
            {
                var ErrorMessage = string.Format("方法：{0},关键参数：{1} ,错误提示：{2},堆栈信息：{3}", "ExcelToDS", Path, ex.Message, ex.StackTrace);
                Log.Write(ErrorMessage);
                throw ex;
            }

        }

        /// <summary>
        ///   using the OLEDB.Excel to excute mutil sql in the Excel. 
        /// </summary>
        /// <param name="sql"></param>
        /// <param name="Path"></param>
        /// <returns></returns>
        public string ExecuteTransaction(List<string> sql, string Path)
        {
            string result = "true";
            string ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Path + ";Extended Properties=Excel 12.0;";
            //string ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Path + ";Extended Properties=Excel 12.0;HDR=Yes;IMEX=1;";
            using (OleDbConnection connection =
                       new OleDbConnection(ConnectionString))
            {
                OleDbCommand command = new OleDbCommand();
                OleDbTransaction transaction = null;
                // Set the Connection to the new OleDbConnection.
                command.Connection = connection;
                // Open the connection and execute the transaction.
                var excutedSqlString = String.Empty;
                //  var count = 0;
                try
                {
                    connection.Open();
                    // Start a local transaction
                    transaction = connection.BeginTransaction();
                    // Assign transaction object for a pending local transaction.
                    //  command.Connection = connection;
                    command.Transaction = transaction;

                    foreach (var item in sql)
                    {
                        excutedSqlString = item;
                        //   LogMess(temp);
                        command.CommandText = item;
                        command.ExecuteNonQuery();
                    }
                    // Execute the commands.
                    transaction.Commit();

                    // Commit the transaction.
                }
                catch (Exception ex)
                {
                    transaction.Rollback();
                    connection.Close();
                    var ErrorMessage = string.Format("方法:{0},关键参数:{1}，SQl Script:{2}, 错误提示:{3},堆栈信息:{4}", "ExecuteTransaction", Path, excutedSqlString, ex.Message, ex.StackTrace);
                    Log.Write(ErrorMessage);
                    throw ex;

                }
                connection.Close();
                return result;
            }
            // The connection is automatically closed when the
            // code exits the using block.
        }

        /// <summary>
        ///   using the OLEDB.Excel to excute sql in the Excel. 
        /// </summary>
        /// <param name="sql"></param>
        /// <param name="Path"></param>
        /// <returns></returns>
        public string ExecuteSql(string sql, string Path)
        {
            string result = "true";
            string ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Path + ";Extended Properties=Excel 12.0;";
            //string ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Path + ";Extended Properties=Excel 12.0;HDR=Yes;IMEX=1;";
            using (OleDbConnection connection =
                       new OleDbConnection(ConnectionString))
            {
                OleDbCommand command = new OleDbCommand();
                // Set the Connection to the new OleDbConnection.
                command.Connection = connection;
                // Open the connection and execute the transaction.
                var excutedSqlString = String.Empty;
                //  var count = 0;
                try
                {
                    connection.Open();
                    excutedSqlString = sql;
                    command.CommandText = sql;
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {

                    connection.Close();
                    var ErrorMessage = string.Format("方法:{0},关键参数:{1}，SQl Script:{2}, 错误提示:{3},堆栈信息:{4}", "ExecuteSql", Path, excutedSqlString, ex.Message, ex.StackTrace);
                    Log.Write(ErrorMessage);
                    throw ex;

                }
                connection.Close();
                return result;
            }
        }

        private object missing = System.Reflection.Missing.Value;

        /// <summary>
        /// using the Interop.Excel to insert data into the Excel. 
        /// </summary>
        /// <param name="password"></param>
        /// <param name="excelFilePath"></param>
        /// <param name="insertDataTable"></param>
        /// <param name="cloumNameIndexs"></param>
        public void InsertDataIntoExcel(string password, string excelFilePath, DataTable insertDataTable,int[] cloumNameIndexs)
        {
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook wb = null;
            Microsoft.Office.Interop.Excel.Worksheet ws = null;
            Microsoft.Office.Interop.Excel.Range range = null;

            try
            {
                if (insertDataTable.Rows.Count > 0)
                {
                    excel.Visible = false;//设置调用引用的 Excel文件是否可见
                    excel.DisplayAlerts = false;

                    if (String.IsNullOrEmpty(password))
                    {
                        wb = excel.Workbooks.Open(excelFilePath);
                    }
                    else
                    {
                        wb = excel.Workbooks.Open(excelFilePath, missing, false, missing, password, password, missing, missing, missing, missing, missing, missing, missing, Type.Missing, Type.Missing);
                    }

                     ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets[1];

                    int topRowCount = ws.UsedRange.Rows.Count;
                    int colCount = ws.UsedRange.Columns.Count;

                    if (insertDataTable.Rows.Count > 0)
                    {
                        var rowCount = insertDataTable.Rows.Count;
                        var columnCount = insertDataTable.Rows[0].ItemArray.Count();
                        var startIndex = topRowCount + 1;
                        #region old updating method              
                        //Log.Write("------------------------insert single data into Excel ---------------------");
                        //for (int i = 0; i < rowCount; i++)
                        //{
                        //    var rowItem = insertDataTable.Rows[i];

                        //    for (int j = 0; j < columnCount; j++)
                        //    {
                        //        ws.Cells[startIndex + i, cloumNameIndexs[j] + 1].Value = rowItem[j];
                        //    }                  
                        //}
                        //Log.Write("------------------------insert single data into Excel ---------------------");
                        #endregion

                        #region  insert multiple data into Excel 
                        Log.Write("------------------------insert multiple data into Excel ---------------------");
                        int maxClumnNameIndex = 0;
                        for (int i = 1; i < cloumNameIndexs.Length; i++)
                        {
                            if (maxClumnNameIndex < cloumNameIndexs[i])
                            {
                                maxClumnNameIndex = cloumNameIndexs[i];
                            }
                        }

                        object[,] objValue = new object[rowCount, maxClumnNameIndex + 2];

                        //for (int i = 0; i < rowCount; i++)
                        //{
                        //    var rowItem = insertDataTable.Rows[i];
                        //    for (int j = 0; j < maxClumnNameIndex + 2; j++)
                        //    {
                        //        objValue[i, j] = String.Empty;
                        //    }
                        //}
                        for (int i = 0; i < rowCount; i++)
                        {
                            var rowItem = insertDataTable.Rows[i];
                            for (int j = 0; j < columnCount; j++)
                            {
                                if (rowItem[j] != null)
                                {
                                    var val = Convert.ToString(rowItem[j]);
                                    objValue[i, cloumNameIndexs[j]] = String.IsNullOrEmpty(val) ? "" : val;
                                }

                            }
                        }

                        range = (Microsoft.Office.Interop.Excel.Range)ws.Cells[startIndex, 1];
                        range = range.get_Resize(rowCount, maxClumnNameIndex + 2);
                        range.FormulaArray = objValue;
                        Log.Write("------------------------insert multiple data into Excel ---------------------");
                        #endregion
                    }
                    excel.ActiveWorkbook.Save();
                }
            }
            catch (Exception ex)
            {
                var ErrorMessage = string.Format("方法：{0}, 错误提示：{1},堆栈信息：{2}", "InsertDataIntoExcel", ex.Message, ex.StackTrace);
                Log.Write(ErrorMessage);
                throw ex;
            }
            finally
            {
                if (wb != null)
                    wb.Close(true, Type.Missing, Type.Missing);

                if (excel != null)
                {
                    excel.Workbooks.Close();
                    excel.Quit();
                }

                if (range != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
                    range = null;
                }

                if (ws != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ws);
                    ws = null;
                }

                if (wb != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wb);
                    wb = null;
                }
                if (excel != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
                    excel = null;
                }
                CommonMethod.KillAllExcelProcess();
                GC.Collect();
            }

        }

        /// <summary>
        /// using the Interop.Excel to update the Excel. 
        /// </summary>
        /// <param name="password"></param>
        /// <param name="excelFilePath"></param>
        /// <param name="updateDataTable">The Index colmun is addtional colmun in the updateDataTable,the colmun specified the location of data</param>
        /// <param name="cloumNameIndexs"></param>
        /// <param name="isUpdataCol"></param>
        /// <param name="SheetIndex"></param>
        public void UpdateDataInExcel(string password, string excelFilePath, DataTable updateDataTable, int[] cloumNameIndexs,bool[] isUpdataCol,int SheetIndex=1)
        {
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook wb = null;
            Microsoft.Office.Interop.Excel.Worksheet ws = null;
            Microsoft.Office.Interop.Excel.Range range = null;
            try
            {

                if (updateDataTable.Rows.Count > 0)
                {
                    Log.Write("----------------------------------UpdateOrderList----------------------------------");

                    excel.Visible = false;//设置调用引用的 Excel文件是否可见
                    excel.DisplayAlerts = false;

                    if (String.IsNullOrEmpty(password))
                    {
                        wb = excel.Workbooks.Open(excelFilePath);
                    }
                    else
                    {
                        wb = excel.Workbooks.Open(excelFilePath, missing, false, missing, password, password, missing, missing, missing, missing, missing, missing, missing, Type.Missing, Type.Missing);
                    }

                     ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets[1];

                    #region  update data

                    var rowCount = updateDataTable.Rows.Count;

                    var columnCount = updateDataTable.Rows[0].ItemArray.Count();
                 
                    for (int i = 0; i < rowCount; i++)
                    {
                        var rowItem = updateDataTable.Rows[i];
                        var index = Convert.ToInt32(rowItem["Index"]);
        
                        for(int j = 0; j < columnCount;j++)
                        {
                            if (isUpdataCol[j])
                            {
                                if (rowItem[j] != null)
                                {
                                    var val = Convert.ToString(rowItem[j]);
                                    ws.Cells[index + 2, cloumNameIndexs[j] + 1].Value = String.IsNullOrEmpty(val)?"":val;
                                }
                            }

                        }                   
                    }
                    #endregion

                    excel.ActiveWorkbook.Save();
                }
            }
            catch (Exception ex)
            {
                var ErrorMessage = string.Format("方法：{0}, 错误提示：{1},堆栈信息：{2}", "InsertDataIntoExcel", ex.Message, ex.StackTrace);
                Log.Write(ErrorMessage);
                throw ex;
            }
            finally
            {
                if (wb != null)
                    wb.Close(true, Type.Missing, Type.Missing);

                if (excel != null)
                {
                    excel.Workbooks.Close();
                    excel.Quit();
                }

                if (range != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
                    range = null;
                }

                if (ws != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ws);
                    ws = null;
                }

                if (wb != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wb);
                    wb = null;
                }
                if (excel != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
                    excel = null;
                }
                CommonMethod.KillAllExcelProcess();
                GC.Collect();
            }



        }

        /// <summary>
        /// Using the Interop.Excel to get DS Data
        /// </summary>
        /// <param name="password"></param>
        /// <param name="excelFilePath"></param>
        /// <param name="orderingClumnNames"></param>
        /// <param name="OrderDTClumnNameIndex"></param>
        /// <param name="sheetInde"></param>
        /// <returns></returns>
        public DataSet ExcelToDS(string password, string excelFilePath, string orderingClumnNames, int[] OrderDTClumnNameIndex, int sheetInde = 1)
        {

            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook wb = null;
            Microsoft.Office.Interop.Excel.Worksheet ws = null;
            Microsoft.Office.Interop.Excel.Range range = null;

            try
            {
                excel.Visible = false;//设置调用引用的 Excel文件是否可见
                excel.DisplayAlerts = false;
  
                if (String.IsNullOrEmpty(password))
                {
                    wb = excel.Workbooks.Open(excelFilePath);
                }
                else
                {
                    wb = excel.Workbooks.Open(excelFilePath, missing, false, missing, password, password, missing, missing, missing, missing, missing, missing, missing, Type.Missing, Type.Missing);
                }
  
                ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets[sheetInde];

                int rowCount = ws.UsedRange.Rows.Count;
                int colCount = ws.UsedRange.Columns.Count;

                int maxClumnNameIndex = 0;
                for (int i = 0; i < OrderDTClumnNameIndex.Length; i++)
                {
                    if (maxClumnNameIndex < OrderDTClumnNameIndex[i])
                    {
                        maxClumnNameIndex = OrderDTClumnNameIndex[i];
                    }
                }

                range = (Microsoft.Office.Interop.Excel.Range)ws.Cells[2, 1];
                range = range.get_Resize(rowCount, maxClumnNameIndex + 2);
                object[,] objArray = range.Value2;

                DataTable orderDataTable = new DataTable();
                orderDataTable.TableName = "table1";
                var columnNameList = orderingClumnNames.Split(',');
                for (int i = 0; i < columnNameList.Length; i++)
                {
                    orderDataTable.Columns.Add(columnNameList[i], typeof(String));
                }

                for (int i = 0; i < rowCount - 1; i++)
                {
                    orderDataTable.Rows.Add();
                    for (int j = 0; j < columnNameList.Length; j++)
                    {

                        if (OrderDTClumnNameIndex[j] != -1)
                        {
                            var objVal = objArray[i + 1, OrderDTClumnNameIndex[j] + 1];
                            if (objVal != null)
                            {
                                var text = objVal.ToString();
                                orderDataTable.Rows[i][j] = String.IsNullOrEmpty(text) ? "" : text;
                            }
                        }
                        else
                        {
                            orderDataTable.Rows[i][j] = "N/A";
                        }
                    }
                }

                DataSet retDataSet = new DataSet();
                retDataSet.Tables.Add(orderDataTable);
                return retDataSet;
            }
            catch (Exception ex)
            {
                var ErrorMessage = string.Format("方法：{0}, 错误提示：{1},堆栈信息：{2}", "GetOrderDT", ex.Message, ex.StackTrace);
                Log.Write(ErrorMessage);
                throw ex;
            }
            finally
            {
                if (wb != null)
                    wb.Close(true, Type.Missing, Type.Missing);

                if (excel != null)
                {
                    excel.Workbooks.Close();
                    excel.Quit();
                }

                if (range != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
                    range = null;
                }

                if (ws != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ws);
                    ws = null;
                }

                if (wb != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wb);
                    wb = null;
                }
                if (excel != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
                    excel = null;
                }
                CommonMethod.KillAllExcelProcess();
                GC.Collect();
            }
        }

        /// <summary>
        /// using the OleDb to create excel Table
        /// </summary>
        /// <param name="excelFilePath"></param>
        /// <param name="dt"></param>
        public void CreateExcelTable(string excelFilePath, DataTable dt)
        {
            try
            {
                List<string> sqlList = new List<string>();

                //创建表格字段
                StringBuilder sqlScript = new StringBuilder();
                sqlScript.Append("CREATE TABLE ").Append("[" + dt.TableName + "]");
                sqlScript.Append("(");
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    if (dt.Columns[i].DataType == typeof(double))
                        sqlScript.Append("[" + dt.Columns[i].ColumnName + "] numeric(18,2),");
                    else
                        sqlScript.Append("[" + dt.Columns[i].ColumnName + "] text,");
                }
                sqlScript = sqlScript.Remove(sqlScript.Length - 1, 1);
                sqlScript.Append(")");
                sqlList.Add(sqlScript.ToString());
                ExecuteTransaction(sqlList, excelFilePath);
            }
            catch (Exception ex)
            {
                var ErrorMessage = string.Format("方法：{0},关键参数：{1}，错误提示：{2},堆栈信息：{3}", "createExcelTabl", excelFilePath, ex.Message, ex.StackTrace);
                Log.Write(ErrorMessage);
                throw ex;
            }
        }

        /// <summary>
        /// using the OleDb to insert data into Excel
        /// </summary>
        /// <param name="excelFilePath"></param>
        /// <param name="dt"></param>
        public void InsertDTintoTable(string excelFilePath, DataTable dt)
        {
            try
            {
                List<string> insertSqlList = new List<string>();
                StringBuilder sqlStr = new StringBuilder();
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    sqlStr.Clear();
                    StringBuilder strvalue = new StringBuilder();
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        strvalue.Append("'" + dt.Rows[i][j].ToString() + "'");
                        if (j != dt.Columns.Count - 1)
                        {
                            strvalue.Append(",");
                        }
                    }
                    var tempStr = String.Format("insert into [{0}] values({1})", dt.TableName, strvalue.ToString());
                    insertSqlList.Add(tempStr);
                }
                ExecuteTransaction(insertSqlList, excelFilePath);
            }
            catch (Exception ex)
            {
                var ErrorMessage = string.Format("方法：{0},关键参数：{1}，错误提示：{2},堆栈信息：{3}", "intoDTintoTable", excelFilePath, ex.Message, ex.StackTrace);
                Log.Write(ErrorMessage);
                throw ex;
            }

        }

        /// <summary>
        ///  using the OleDb to Delete DataTable
        /// </summary>
        /// <param name="Path">路径</param>
        /// <param name="dt">DataTable</param>
        /// <param name="HDR">HDR true:读取column false:column行作为内容读取</param>
        public void DropTable(string sheetName, string Path, string HDR = "true")
        {
            string HDRStr = "YES";
            if (!HDR.Equals("true")) HDRStr = "NO";
            string strCon = string.Empty;
            FileInfo file = new FileInfo(Path);
            string extension = file.Extension;
            switch (extension)
            {
                case ".xls":
                    strCon = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Path + ";Extended Properties=Excel 8.0;HDR=" + HDRStr + ";";
                    break;
                case ".xlsx":
                    strCon = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Path + ";Extended Properties='Excel 12.0;HDR=" + HDRStr + ";IMEX=0;'";
                    break;
                default:
                    strCon = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Path + ";Extended Properties='Excel 8.0;HDR=" + HDRStr + ";IMEX=0;'";
                    break;
            }
            using (System.Data.OleDb.OleDbConnection con = new System.Data.OleDb.OleDbConnection(strCon))
            {

                con.Open();

                System.Data.OleDb.OleDbCommand cmd;
                try
                {
                    cmd = new System.Data.OleDb.OleDbCommand(string.Format("drop table {0}", sheetName), con);    //覆盖文件时可能会出现Table 'Sheet1' already exists.所以这里先删除了一下
                    cmd.ExecuteNonQuery();
                }
                catch (Exception ex)
                {

                }
                finally
                {
                    con.Close();
                }
            }

        }
    }
}
