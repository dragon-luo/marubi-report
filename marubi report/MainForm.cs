using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;
using System.Threading;

namespace Marubi
{
    public partial class MainForm : Form
    {
        private const string DRIVER_4_EXCEL_2010 = "Microsoft.Ace.OLEDB.12.0";
        private const string DRIVER_4_EXCEL_2003 = "Microsoft.Jet.OLEDB.4.0";

        private const string REPORT_FILE_NAME = "marubi inventory report.xls";
        private const string DATA_FILE_NAME = "marubi inventory.xls";
        private const string ACCESS_DB_FILE = "marubi.mdb";
        
        private const string EXCEL_CONNECTION_STRING = "Provider={0};Data Source={1};Extended Properties=Excel 8.0";        
        private const string ACCESS_CONNNECTION_STRING = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0}";

        private string startupPath = Path.Combine(Application.StartupPath, "access db");

        private string xlsConnectionString = "";

        delegate void ProcessStatusDelegate(string message);
        delegate void SetButtonEnableDelegate(bool enable);

        #region Construct(s)
        public MainForm()
        {
            InitializeComponent();
        }
        #endregion

        #region Properties
        private string DataFile { get; set; }
        
        private string ReportFile { get; set; }

        private string AccessDB { get; set; }
        #endregion

        private void InitializeUI()
        {            
            DataFile = Path.Combine(startupPath, DATA_FILE_NAME);
            ReportFile = Path.Combine(startupPath, REPORT_FILE_NAME);

            InitializeFile();
            SetConnectionString(DataFile, AccessDB);
        }

        private void InitializeFile()
        {
            TxtRptFile.Text = ReportFile;
            TxtDataFile.Text = DataFile;
        }

        private void SetConnectionString(string dataFile, string accessDB)
        {
            string driver4Excel = DRIVER_4_EXCEL_2010;
            FileInfo file = new FileInfo(dataFile);
            if (file.Extension == ".xls")
                driver4Excel = DRIVER_4_EXCEL_2003;
            else if (file.Extension == ".xlsx")                                                
                driver4Excel = DRIVER_4_EXCEL_2010;
            
            xlsConnectionString = string.Format(EXCEL_CONNECTION_STRING, driver4Excel, dataFile);
            accessConnectionString = string.Format(ACCESS_CONNNECTION_STRING, accessDB);
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            InitializeUI();                     
        }

        private void BtnBrow_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();                     
            openFileDialog.Filter = "Excel(2000-2007) (*.xls)|*.xls|Excel(2010) (*.xlsx)|*.xlsx";
            openFileDialog.FilterIndex = 0;
            openFileDialog.RestoreDirectory = true;
            openFileDialog.Title = "选择 Excel 数据文件";
            openFileDialog.InitialDirectory = startupPath;

            if (DialogResult.OK == openFileDialog.ShowDialog())
            {
                DataFile = openFileDialog.FileName;
                this.TxtDataFile.Text = DataFile;
                SetConnectionString();
            }
        }

        

        private void BtnImport_Click(object sender, EventArgs e)
        {
            FileInfo file = new FileInfo(xlsFullFilePath);
            if (!file.Exists)
            {
                MessageBox.Show("所选择的 Excel 数据文件不存在, 请确认！", "数据导入", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
           
            lstStatus.Items.Clear();
            lstStatus.Items.Add("");
            Thread thread = new Thread(new ThreadStart(ImportData));
            thread.Start();           
        }

        private void ImportData()
        {
            try
            {
                SetButtonEnable(false);
                var st = new System.Diagnostics.Stopwatch();

                SetProcessStatus("开始导入源数据，请稍后 ...");

                SetProcessStatus("  正在读取成品数据 ...");
                st.Start();
                DataSet productDataSet = AccessDB.GetDataSource(xlsConnectionString, ExcelDataType.Product);
                st.Stop();
                SetProcessStatus(string.Format("    已成功读取成品数据 {0} 条, 用时 {1} 秒.", productDataSet.Tables[0].Rows.Count, st.Elapsed.TotalSeconds));

                SetProcessStatus("  正在读取包材数据 ...");
                st.Start();
                DataSet partDataSet = AccessDB.GetDataSource(xlsConnectionString, ExcelDataType.PackingMaterial);
                st.Stop();
                SetProcessStatus(string.Format("    已成功读取包材数据 {0} 条, 用时 {1} 秒.", partDataSet.Tables[0].Rows.Count, st.Elapsed.TotalSeconds));

                SetProcessStatus("  正在读取 BOM 数据 ...");
                st.Start();
                DataSet bomDataSet = AccessDB.GetDataSource(xlsConnectionString, ExcelDataType.Bom);
                st.Stop();
                SetProcessStatus(string.Format("    已成功读取 BOM 数据 {0} 条, 用时 {1} 秒.", bomDataSet.Tables[0].Rows.Count, st.Elapsed.TotalSeconds));

                SetProcessStatus("  正在保存成品数据 ...");
                st.Start();
                int productsAffected = AccessDB.SaveDataSource(ACCESS_CONNNECTION_STRING, productDataSet, ExcelDataType.Product, this);
                st.Stop();
                SetProcessStatus(string.Format("    已成功保存成品数据 {0} 条, 用时 {1} 秒！", productsAffected, st.Elapsed.TotalSeconds));

                SetProcessStatus("  正在保存包材数据 ...");
                st.Start();
                int partsAffected = AccessDB.SaveDataSource(ACCESS_CONNNECTION_STRING, partDataSet, ExcelDataType.PackingMaterial, this);
                st.Stop();
                SetProcessStatus(string.Format("    已成功保存包材数据 {0} 条, 用时 {1} 秒！", partsAffected, st.Elapsed.TotalSeconds));

                SetProcessStatus("  正在保存 BOM 数据 ...");
                st.Start();
                int bomAffected = AccessDB.SaveDataSource(ACCESS_CONNNECTION_STRING, bomDataSet, ExcelDataType.Bom, this);
                st.Stop();
                SetProcessStatus(string.Format("    已成功保存 BOM 数据 {0} 条, 用时 {1} 秒！", bomAffected, st.Elapsed.TotalSeconds));

                SetProcessStatus("导入源数据完成.");
                SetButtonEnable(true);
            }
            catch (Exception ex)
            {
                SetProcessStatus(string.Format("导入源数据出现异常, {0}", ex.Message));                
                MessageBox.Show(ex.Message, "数据导入", MessageBoxButtons.OK, MessageBoxIcon.Information);
                SetButtonEnable(true);
            }
        }

        #region SetProcessStatus
        public void SetProcessStatus(string message)
        {
            if (this.InvokeRequired)
            {
                ProcessStatusDelegate statusDelegate = new ProcessStatusDelegate(SetProcessStatus);
                this.Invoke(statusDelegate, new object[] { message });
            }
            else
            {
                this.lstStatus.Items.Insert(lstStatus.Items.Count - 1, message);
                this.lstStatus.SelectedIndex = this.lstStatus.Items.Count - 1; 
                this.lstStatus.SelectedIndex = -1;
            }
        }
        #endregion

        #region SetButtonEnable
        public void SetButtonEnable(bool enable)
        {
            if (this.InvokeRequired)
            {
                SetButtonEnableDelegate enableDelegate = new SetButtonEnableDelegate(SetButtonEnable);
                this.Invoke(enableDelegate, new object[] { enable });
            }
            else
            {
                BtnImport.Enabled = enable;
                BtnGenReport.Enabled = enable;
                BtnBrow.Enabled = enable;
            }
        }
        #endregion

        private void BtnGenReport_Click(object sender, EventArgs e)
        {
            FileInfo file = new FileInfo(xlsFullFilePath);
            if (!file.Exists)
            {
                MessageBox.Show("所选择的 Excel 数据文件不存在, 请确认！", "生成报表", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            lstStatus.Items.Clear();
            lstStatus.Items.Add("");
            Thread thread = new Thread(new ThreadStart(ExportData));
            thread.Start();
        }        

        private void ExportData()
        {
            try
            {
                SetButtonEnable(false);
                
                var st = new System.Diagnostics.Stopwatch();

                SetProcessStatus("开始生成报表，请稍后 ...");

                SetProcessStatus("  正在读取需要产生报表的商品编号 ...");
                st.Start();
                DataSet reportDataSet = AccessDB.GetDataSource(xlsConnectionString, ExcelDataType.Report);
                st.Stop();
                SetProcessStatus(string.Format("    已成功读取商品编号数据 {0} 条, 用时 {1} 秒.", reportDataSet.Tables[0].Rows.Count, st.Elapsed.TotalSeconds));

                SetProcessStatus("  正在保存需要产生报表的商品编号 ...");
                st.Start();
                int reportsAffected = AccessDB.SaveDataSource(ACCESS_CONNNECTION_STRING, reportDataSet, ExcelDataType.Report, this);
                st.Stop();
                SetProcessStatus(string.Format("    已成功保存商品编号数据 {0} 条, 用时 {1} 秒.", reportsAffected, st.Elapsed.TotalSeconds));

                SetProcessStatus("  正在生成报表 ...");
                st.Start();
                DataTable dt = AccessDB.GetReportTable(ACCESS_CONNNECTION_STRING);
                if (dt != null && dt.Rows.Count > 0)
                    ExcelUtils.Export(dt, rptFullFilePath, "常规产品");
                st.Stop();
                SetProcessStatus(string.Format("    生成报表用时 {1} 秒.", reportsAffected, st.Elapsed.TotalSeconds));

                SetProcessStatus("    生成报表成功！");
                SetProcessStatus("生成报表完成.");

                SetButtonEnable(true);
            }
            catch (Exception ex)
            {
                SetProcessStatus(string.Format("生成报表出现异常, {0}", ex.Message));                
                MessageBox.Show(ex.Message, "生成报表", MessageBoxButtons.OK, MessageBoxIcon.Information);
                SetButtonEnable(true);
            }
           
        }        
    }
}
