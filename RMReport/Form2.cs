using System;
using System.Data;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Reflection;
using System.Collections;
using System.Data.SQLite;
using ADODB;
using System.IO;

namespace RMReport
{
    public partial class Form2 : Form
    {
        //创建控件; hwnd为控件句柄，大于0为嵌入报表
        [DllImport("MyReportMachine.dll")]
        private static extern void MCreate(int hwnd);

        //关闭控件并释放
        [DllImport("MyReportMachine.dll")]
        private static extern void MClose();

        //清除所有变量与数据对象
        [DllImport("MyReportMachine.dll")]
        private static extern void MClearAll();

        //将变量传入报表
        //   objName:变量名称;
        //   objValue:变量值(可以是字符串,数值,日期,是否等类型)
        [DllImport("MyReportMachine.dll")]
        private static extern void MAddV(object objName, object objValue);

        //动态将数据集传入报表
        //   objName:数据集名称;
        //   objValue:Recordset数据集
        [DllImport("MyReportMachine.dll")]
        private static extern void MAddData(object objName, ADODB.Recordset rsData);

        //控件中有数据源指定数据源对应关系(最多10个数据源）(不超过10个数据源时，请使用这个方法)
        //   intNumber:数据源编号(0-9)共12个数据源
        //   objValue:Recordset数据集
        //   objName:数据集别名
        [DllImport("MyReportMachine.dll")]
        private static extern void MSetData(int intNumber, ADODB.Recordset rsData, object objName);

        //报表打印
        //   intKind: =0:报表预览；=1:打印报表；=2:报表设计；=3:报表准备；=4:显示准备的报表；
        //   intShowDialog: =0:隐藏打印对话框, <>0:显示打印对话框
        //   intProgress: =0:隐藏报表加载进度条, <>0:显示报表加载进度条
        //   objFileName:报表文件路径
        //   objPrinter:打印名称,="":默认打印
        [DllImport("MyReportMachine.dll")]
        private static extern void MPrintReport(int intKind, int intShowDialog, int intProgress, object objFileName, object objPrinter);

        //--------------------------------------------------------以下是报表附属属性，可以不用调用
        //返回报表页数
        [DllImport("MyReportMachine.dll")]
        private static extern int MReportCount();

        //设报表预览按钮显示状态
        //   intZoom：缩放按钮
        //   intLoad：导出按钮
        //   intSave：保存按钮
        //   intPrint：打印按钮
        //   intPageSetup：报表页面设置按钮
        //   intExit：退出报表预览按钮
        //   intSaveToXls：导出到Execl按钮
        //   intExport：导出按钮
        //   intNavigator：导航按钮
        [DllImport("MyReportMachine.dll")]
        private static extern void MPreviewButtons(int intZoom, int intLoad, int intSave, int intPrint, int intPageSetup, int intExit, int intSaveToXls, int intExport, int intNavigator);
        //设置报表语言(默认中文)
        [DllImport("MyReportMachine.dll")]
        private static extern void MLanguage(object strFileName);

        //设置预览模式；0:模式预览  1:嵌入预览
        [DllImport("MyReportMachine.dll")]
        private static extern void MSetPreview(int intKind);

        //数据库文件
        string dbFile;
        //查询数据
        Hashtable dataDict;
        //报表模板
        string rmf;
        //打印动作
        int action;

        public Form2(string rmf, string dbFile, Hashtable dataDict, int action)
        {
            InitializeComponent();
            this.WindowState = FormWindowState.Minimized;
            this.ShowInTaskbar = false;
            base.SetVisibleCore(false);
            this.dbFile = dbFile;
            this.rmf = rmf;
            this.dataDict = dataDict;
            this.action = action;
        }

        //DataTable转RecordSet
        public Recordset DsToRs(DataTable table)
        {

            Recordset rs = new RecordsetClass();
            System.Array ArrA = System.Array.CreateInstance(typeof(string), table.Columns.Count);

            foreach (DataColumn dc in table.Columns)
            {
                ArrA.SetValue(dc.ColumnName, dc.Ordinal);
                rs.Fields._Append(dc.ColumnName, GetDataType(dc.DataType), -1, FieldAttributeEnum.adFldIsNullable);

            }
            rs.Open(Missing.Value, Missing.Value, CursorTypeEnum.adOpenUnspecified, LockTypeEnum.adLockUnspecified, -1);
            //rs.Open(null,null, CursorTypeEnum.adOpenDynamic, LockTypeEnum.adLockOptimistic, 0);
            // int i = 0;
            foreach (DataRow dr in table.Rows)
            {

                rs.AddNew(Missing.Value, Missing.Value); //object o;
                                                         // rs.AddNew(rs.Fields.,dr.ItemArray);

                //rs.AddNew(ArrA, dr.ItemArray);
                for (int i = 0; i < table.Columns.Count; i++)
                {
                    rs.Fields[i].Value = dr[i];
                }
            }
            return rs;
        }

        //转类型
        public static DataTypeEnum GetDataType(Type dataType)
        {
            switch (dataType.ToString())
            {
                case "System.Boolean": return DataTypeEnum.adBoolean;

                case "System.Char": return DataTypeEnum.adChar;
                case "System.DateTime": return DataTypeEnum.adDBTimeStamp;// .adDate;
                case "System.Decimal": return DataTypeEnum.adDecimal;// .adNumeric;
                case "System.Double": return DataTypeEnum.adDouble;
                case "System.Int16": return DataTypeEnum.adSmallInt;
                case "System.Int32": return DataTypeEnum.adInteger;
                case "System.Int64": return DataTypeEnum.adBigInt;
                case "System.Single": return DataTypeEnum.adSingle;
                case "System.String": return DataTypeEnum.adVarWChar;
                case "System.Byte[]": return DataTypeEnum.adVarBinary;
                case "System.Object": return DataTypeEnum.adVariant;
                case "System.Guid": return DataTypeEnum.adGUID;

                case "System.Byte": return DataTypeEnum.adUnsignedTinyInt;
                case "System.SByte": return DataTypeEnum.adTinyInt;
                case "System.UInt16": return DataTypeEnum.adUnsignedSmallInt;
                case "System.UInt32": return DataTypeEnum.adUnsignedInt;
                case "System.UInt64": return DataTypeEnum.adUnsignedBigInt;


                default: throw new Exception("没有对应的数据类型");

            }
        }

        ///<summary>
        /// 执行SQL返回数据; StrSql:查询语句; OleDbCon:数据库连接; T_Data:数据表
        ///</summary>
        public void My_OpenSql(string StrSql, SQLiteConnection conn, DataTable T_Data)
        {
            if (conn.State != ConnectionState.Open)
            {
                conn.Open();
            }
            SQLiteDataAdapter adapter = new SQLiteDataAdapter(StrSql, conn);
            T_Data.Clear();
            T_Data.Columns.Clear();
            try
            {
                adapter.Fill(T_Data);
            }
            catch (Exception e)
            {
                MessageBox.Show("错误信息为：\n"+e.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Environment.Exit(1);
            }
            conn.Close();
            adapter.Dispose();
        }

        //加载
        private void Form1_Load(object sender, EventArgs e)
        {
            string dbPath = Application.StartupPath + @"\db\" + this.dbFile;
            if (!File.Exists(dbPath))
            {
                MessageBox.Show("数据库不存在！\n" + dbPath, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Environment.Exit(1);
            }
            string connString = "Data Source=" + dbPath + ";Version=3;";
            this.conn = new SQLiteConnection(connString);
            PrintReport(this.action);
            Environment.Exit(0);
        }

        //报表打印
        private void PrintReport(int IntKind)
        {
            string StrReport, StrSql;
            DataTable DtTemp;
            StrReport = Application.StartupPath + @"\rmf\" + this.rmf;
            if (!File.Exists(StrReport))
            {
                MessageBox.Show("报表文件不存在！\n" + StrReport, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Environment.Exit(1);
            }

            MCreate(0); //报表创建 ； Hwnd可以为0 ； ***必须有
            MClearAll(); //清除所有数据源及变量

            //2、数据源的设法方法
            //动态加载数据源
            //使用动态加载数据源时,请先清除数据源

            //控件中有数据源指定数据源对应关系(最多10个数据源）(不超过10个数据源时，请使用这个方法)
            //Rm.MSetData(MSetData(IntNumber As Long, RsData As Object, StrName as String)
            //参数说明：
            //1、参数1(IntNumber):数据源编号(0-9)共10个数据源
            //2、参数2(RsData)是Recordset数据源;
            //3、参数3(StrName)数据源别名
            int i = 0;
            foreach (DictionaryEntry d in this.dataDict)
            {
                StrSql = (string)d.Value;
                DtTemp = new DataTable();
                this.My_OpenSql(StrSql, conn, DtTemp);
                MSetData(i++, this.DsToRs(DtTemp), (string)d.Key);
            }

            //报表打印
            //MPrintReport(IntKind As Long, IntShowDialog As Long, IntProgress As Long, StrFileName as String, StrPrinter  as String)
            //参数说明:
            //1、参数1(IntKind): =0:报表预览；=1:打印报表；=2:报表设计；=3:报表准备；=4:显示准备的报表；
            //2、参数2(IntShowDialog): =0:隐藏打印对话框, <>0:显示打印对话框
            //3、参数3(IntProgress): =0:隐藏报表加载进度条, <>0:显示报表加载进度条
            //4、参数4(StrFileName):报表名称
            //5、参数5(StrPrinter):打印名称,="":默认打印

            switch (IntKind)
            {
                case 1:
                    {
                        MPrintReport(1, 1, 0, StrReport, "");
                    }
                    break;
                case 2:
                    {
                        MPrintReport(2, 1, 0, StrReport, "Adobe PDF");//"Adobe PDF") '指定打印名称:"Adobe PDF"
                    }
                    break;
                default:
                    {
                        MPrintReport(0, 1, 0, StrReport, "");
                    }
                    break;
            };

            MClose(); //卸载 ；***必须有
        }
    }
}
