using System;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace version2
{
    class Program
    {
        [STAThread]
        public static void Main()
        {
            Helper helper = new Helper();
            Common common = new Common();
            WriteWord writeWord = new WriteWord();

            //*********************************选取word模板*************************************
            helper.Guide(1);
            string WordTempleteFile = "Word 2007 Files(*.docx)|*.docx|Word 2003 Files(*.doc)|*.doc";
            string WordTempleteTitle = "Word模板";
            string WordTempletePath = common.OpenFile(WordTempleteFile, WordTempleteTitle);
            //    //*********************************选取Excel数据源模板******************************
            helper.Guide(2);
            string OpenExcelFilter = "xls|*.xls|xlsx|*.xlsx|csv|*.csv";
            string Title = "数据源";
            string ExcelSourcePath = common.OpenFile(OpenExcelFilter, Title);
            //*********************************选取保存生成word的路径*************************************
            Console.Clear();
            Console.WriteLine("");
            Console.WriteLine("****- 您选择的数据源路径为{0}-****", ExcelSourcePath);
            Console.WriteLine("");
            Console.WriteLine("****- 您选择的模板路径为{0} -*****", WordTempletePath);

            if (!common.UserChooseOption("请检查控制台打印出的信息是否正确？"))
            {
                Application.Restart();
                Environment.Exit(0);
                Console.WriteLine("请重新选择！");
            }
            Console.Clear();
            helper.Guide(3);
            string Description = "生成存放模板的文件夹";
            string SavedPath = common.OpenFolderBrowreDialog(Description) + @"\" ;
            Console.WriteLine("");
            Console.WriteLine("**************** 生成的文档路径为{0} ***************", SavedPath);

            if (!common.UserChooseOption("请检查控制台打印出的信息是否正确？"))
            {

                Application.Restart();
                Environment.Exit(0);
                Console.WriteLine("请重新选择！");

            }
            //处理保存的文件名称，防止重复
            //若用户选择文件夹目录不存在则创建，并提示
            try
            {
                if (!Directory.Exists(SavedPath))
                {
                    Directory.CreateDirectory(SavedPath);                   
                }
                //else if (!File.Exists(SavedPath + @"\" + "log.txt"))
                //{
                //    FileStream fs = new FileStream(SavedPath + @"\" + "log.txt", FileMode.CreateNew);
                //}
            }
            catch { }
            //*********************************生成word文件*************************************
            DateTime StartTime = DateTime.Now.ToLocalTime().ToLocalTime();

            StringBuilder log = new StringBuilder(writeWord.Write(ExcelSourcePath, WordTempletePath, SavedPath)); //处理日志文件
            
            //5.保存给出提示信息
            //*********************************转换完成*************************************
            DateTime EndTime = DateTime.Now.ToLocalTime().ToLocalTime();

            double UseTime = (EndTime - StartTime).TotalSeconds;
            log.Append("\r\n");
            log.Append("\r\n");
            log.Append("用时：");
            log.Append(UseTime.ToString("F2"));
            log.Append("秒");

            string writelog =log.ToString();

            var output = new StreamWriter(SavedPath + @"\" + "log.txt", false, Encoding.UTF8);

            output.Write(writelog);
            output.Close();
            output.Dispose();

            helper.Guide(4, UseTime.ToString("F2"));

            helper.Guide();
        }
    }
}
