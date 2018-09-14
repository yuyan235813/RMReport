using System;
using System.Collections;
using System.Windows.Forms;

namespace RMReport
{
    static class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main(string[] args) 
        {
            var arguments = CommandLineArgumentParser.Parse(args);
            string dbFile = "balance.db";
            string dbStr = "data:select * from t_balance;";
            string rmf = "过称单(标准式).rmf";
            int action = 2;
            int argv = 4;
            if (arguments.Has("-whosyourdaddy"))
            {
                argv = 4;
            }
            else
            {
                if (arguments.Has("-d"))
                {
                    dbFile = arguments.Get("-d").Next;
                    argv++;
                }
                if (arguments.Has("-s"))
                {
                    dbStr = arguments.Get("-s").Next;
                    argv++;
                }
                if (arguments.Has("-r"))
                {
                    rmf = arguments.Get("-r").Next;
                    argv++;
                }
                if (arguments.Has("-a"))
                {
                    action = Convert.ToInt32(arguments.Get("-a").Next);
                    argv++;
                }
            }
            if(argv != 4)
            {
                MessageBox.Show("参数错误！", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Environment.Exit(1);
            }

            Hashtable dataDict = new Hashtable();
            // 分号分隔
            string[] dbStrPair = dbStr.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (string item in dbStrPair)
            {
                // 冒号分隔
                string[] kv = item.Split(new char[] { ':' }, StringSplitOptions.RemoveEmptyEntries);
                dataDict.Add(kv[0], kv[1]);
            }
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Form f = new Form2(rmf, dbFile, dataDict, action);
            Application.Run(f);
        }
    }
}
