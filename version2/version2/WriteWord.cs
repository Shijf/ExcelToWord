using System;
using System.Data;
using System.Text;
using Xceed.Words.NET;

namespace version2
{
    class WriteWord
    {
        string Value = "";//在模板中书签填入的对应值
        string Bookmark = "";//对应模板中的书签
        string WriteLog = "";//写入日志文件
        string CustomFileName = "";//初始化自定义文件名称

        /// <summary>
        /// 将数据源标题对应的值写入相应的书签
        /// </summary>
        /// <param name="ExcelSourcePath">数据源所在路径</param>
        /// <param name="WordTempletePath">模板所在路径</param>
        /// <param name="SavedPath">生成文件保存的路径</param>
        /// <returns>日志文件</returns>
        public string Write(string ExcelSourcePath, string WordTempletePath, string SavedPath)
        {
            ExcelHelper excel = new ExcelHelper(); // 实例化读取表格助手类
            DocX document = DocX.Load(WordTempletePath);//实例化Word处理函数库

            DataSet result = excel.ReadExcelToTable(ExcelSourcePath);

            Common common = new Common();//公用函数库
            StringBuilder log = new StringBuilder("欢迎使用ExcelToWord:" +"\r\n"); //处理日志文件
            
            int Rows = result.Tables[0].Rows.Count;
            int Columns = result.Tables[0].Columns.Count;
            int CustomRows = result.Tables[1].Rows.Count;
            int CustomColumns = result.Tables[1].Columns.Count;
            Console.Clear();
            ConsoleColor colorBack = Console.BackgroundColor;
            ConsoleColor colorFore = Console.ForegroundColor;

            if (result != null && result.Tables.Count > 0)
            {
                Console.WriteLine("");
                Console.WriteLine("**************** 您选择的数据源路径为{0} **************", ExcelSourcePath);
                Console.WriteLine("");
                Console.WriteLine("**************** 您选择的模板路径为{0} ***************", WordTempletePath);          
                Console.WriteLine("");
                Console.WriteLine("**************** 所有数据已加载完毕 ***************");
                Console.WriteLine("");
                int count = Rows;
                int index = 1;
                double prePercent = 0;
                //绘制界面
                Console.WriteLine("********************* 开始生成 ********************");
                log.Append("您选择的数据源路径为:"+ ExcelSourcePath + "\r\n");
                log.Append("您选择的模板路径为:" + WordTempletePath + "\r\n");
                log.Append("生成的文档路径为:" + SavedPath + "\r\n");
                log.Append("**************************************************");
                log.Append("\r\n");
                log.Append("生成的文档如下:" + "\r\n");
                Console.BackgroundColor = ConsoleColor.DarkCyan;
                for (int i = 0; ++i <= 50;)
                {
                    Console.Write(" ");
                }
                Console.WriteLine(" ");
                Console.BackgroundColor = colorBack;
                Console.WriteLine("已完成0%");
                Console.WriteLine("**************************************************");
                for (int i = 1; i < Rows; i++)//循环行数
                {

                    //绘制界面
                    index++;
                    double percent;
                    if (index <= count)
                    {
                        percent = (double)index / count;
                        percent = Math.Ceiling(percent * 100);
                    }
                    else
                    {
                        percent = 1;
                        percent = Math.Ceiling(percent * 100);
                    }
                    

                    for (int j = 0; j < Columns; j++)//循环列数
                    {
                        Bookmark = result.Tables[0].Rows[0][j].ToString();//获取书签 
                        if ((Bookmark.ToString() != ""))//判断当前单元格是否有值
                        {
                            Value = result.Tables[0].Rows[i][j].ToString();//获取书签对应的值  

                            var TempleteBookmark = document.Bookmarks[Bookmark];

                            if (TempleteBookmark != null)
                            {
                                TempleteBookmark.SetText(Value);
                            }
                        }
                    }
                    for (int o = 0; o < CustomColumns; o++)
                    {
                        if (result.Tables[1].Rows[0][o].ToString() != "")
                        {
                            if (i > (CustomRows - 1))//在自定义文件名称中 防止有行数没有对应数据源
                            {
                                Random ran = new Random();
                                CustomFileName = "第" + (i + 1) + "未填写自定义文件名称" + ran.Next(1, Rows * 10);
                            }
                            else
                            {
                                if (result.Tables[1].Rows[i][o].ToString() != "")//取对应行的自定义文件名称
                                {
                                    CustomFileName = CustomFileName + result.Tables[1].Rows[i][o].ToString();
                                }
                                else//在自定义文件名称中 防止有行数没有对应数据源
                                {
                                    Random ran = new Random();
                                    CustomFileName = "第" + i + "未填写自定义文件名称" + ran.Next(1, Rows * 10);
                                }
                            }                    

                        }
                    }
                    // 开始控制进度条和进度变化
                    for (int k = Convert.ToInt32(prePercent); k <= percent; k++)
                    {
                        //绘制进度条进度                 
                        Console.BackgroundColor = ConsoleColor.Yellow;//设置进度条颜色                
                        Console.SetCursorPosition(k / 2, 8);//设置光标位置,参数为第几列和第几行                
                        Console.Write(" ");//移动进度条                
                        Console.BackgroundColor = colorBack;//恢复输出颜色                
                        //更新进度百分比,原理同上.                
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.SetCursorPosition(0, 9);
                        if (k == 100)
                        {
                            Console.Write("已完成{0}%，所有文档已生成", k);
                        }
                        else {
                            Console.Write("已完成{0}%，请勿关闭程序", k);
                        }
                        
                        Console.ForegroundColor = colorFore;
                    }
                    document.SaveAs(SavedPath + CustomFileName + @".docx");//生成文档最终的名称

                    log.Append("\r\n" );
                    log.Append(CustomFileName);
                    log.Append(".docx");
    
                    CustomFileName = "";

                    prePercent = percent;

                    Console.SetCursorPosition(0, 11);
                }
            }

            Console.WriteLine("");
            Console.WriteLine("**************** 生成的文档路径为{0} ***************", SavedPath);
            WriteLog = log.ToString();
            return WriteLog;
        }
        /// <summary>
        /// 格式化 生成的word文件名称
        /// </summary>
        /// <param name="SavedPath"></param>
        /// <param name="Code"></param>
        /// <returns></returns>
        public string WordSavedPath(string SavedPath, string Code = "没有自定义文件")
        {
            return SavedPath + @"/" + Code;
        }
    }
}
