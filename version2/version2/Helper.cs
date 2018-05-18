using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace version2
{
    class Helper
    {
     
        #region 步骤引导
        public void Guide(int step = 1000, string UserTime = "0")
        {

            switch (step)
            {
                case 1:
                    Console.WriteLine("");
                    Console.WriteLine("\n 请选择您要生成数据的模板");
                    break;
                case 2:
                    Console.WriteLine("");
                    Console.WriteLine("\n 请选择您要读取的数据来源");
                    break;
                case 3:
                    Console.WriteLine("");
                    Console.WriteLine("\n 请选择您要保存生成的文档路径");
                    break;               
                case 4:
                    Console.WriteLine("");
                    Console.WriteLine("\n 用时：" + UserTime + "秒");
                    Console.WriteLine("\n 感谢您的使用，联系邮箱：shijf_job@163.com");
                    Console.WriteLine("\n 按任意键退出...");
                    Console.ReadKey();
                    break;
                default:
                    break;
            }

        }
        #endregion

        public Helper()
        {
            String[] colorNames = ConsoleColor.GetNames(typeof(ConsoleColor));
        }

        public void Instruction()
        {



        }
    }
}
