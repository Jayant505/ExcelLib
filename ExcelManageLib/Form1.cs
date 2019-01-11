using ExcelManageLib.DataModel;
using ExcelManageLib.ExcelDataAccess;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelManageLib
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            var dt = new DataTable();
            ExcelTableDataAccess dataAccess = new ExcelTableDataAccess();

            if (dt == null)
            {
                return;
            }

            if (dt.Columns.Count > 0)
            {
                dataAccess.CreateExcelTable(ToolConfiguration.ordingExcelFilePath, dt);

            }

            if (dt.Columns.Count > 0 && dt.Rows.Count > 0)
            {
                dataAccess.InsertDTintoTable(ToolConfiguration.ordingExcelFilePath, dt);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ExcelTableDataAccess excelDataAccess = new ExcelTableDataAccess();
            var orderDS = excelDataAccess.ExcelToDS(ToolConfiguration.ordingExcelFilePath, ToolConfiguration.ordingSheetName, "仓源,Dept,Cat,[Stock Cat],[SKU state],Remark,[supplier code],[supplier name],LT,SKU,Des", "", "");
            var cloumNameIndexs = excelDataAccess.GetClumnNameIndexArray(ToolConfiguration.ordingExcelFilePath, "Sheet1", "仓源,Dept,Cat,Stock Cat,SKU state,Remark,supplier code,supplier name,LT,SKU,Des", 2);
            excelDataAccess.InsertDataIntoExcel(ToolConfiguration.orderingExcelPassword, ToolConfiguration.ordingExcelFilePath, orderDS.Tables[0], cloumNameIndexs);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            ExcelTableDataAccess excelDataAccess = new ExcelTableDataAccess();
            string cloumNameDT = "仓源,Dept,Cat,[Stock Cat],SKU state,Remark,[supplier code],[supplier name],LT,SKU,Des";
            string cloumNameUpdateDT = "Index,仓源,Dept,Cat,Stock Cat,SKU state,Remark,supplier code,supplier name,LT,SKU,Des";

            var cloumNameIndexs = excelDataAccess.GetClumnNameIndexArray(ToolConfiguration.ordingExcelFilePath, "Sheet1", cloumNameDT, 2);
            var isUpdate = new bool[cloumNameIndexs.Count()];

            var updateDataTable = new DataTable();

            string[] columnNameList = cloumNameUpdateDT.Split(',');
            for (int i = 0; i < columnNameList.Length; i++)
            {
                updateDataTable.Columns.Add(columnNameList[i], typeof(String));          
            }
            var createRowIndex = 1;
            updateDataTable.Rows.Add();
            updateDataTable.Rows[createRowIndex]["Index"] = "1";  //The Index colmun is addtional colmun in the updateDataTable,the colmun specified the location of data
            updateDataTable.Rows[createRowIndex]["仓源"] = "";
            updateDataTable.Rows[createRowIndex]["Dept"] = "";
            updateDataTable.Rows[createRowIndex]["Cat"] = "";
            updateDataTable.Rows[createRowIndex]["Stock_cat"] = "";
            updateDataTable.Rows[createRowIndex]["SKU_state"] = "";
            updateDataTable.Rows[createRowIndex]["Remark"] = "";
            updateDataTable.Rows[createRowIndex]["supplier code"] = "";
            updateDataTable.Rows[createRowIndex]["supplier name"] = "";
            updateDataTable.Rows[createRowIndex]["LT"] = "";
            updateDataTable.Rows[createRowIndex]["SKU"] = "";
            updateDataTable.Rows[createRowIndex]["Des"] = "";


            excelDataAccess.UpdateDataInExcel(ToolConfiguration.orderingExcelPassword, ToolConfiguration.ordingExcelFilePath, updateDataTable, cloumNameIndexs, isUpdate);
        }
    }
}
