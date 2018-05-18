using System;
using System.Windows.Forms;

namespace version2
{
    class Common
    {
        #region 选取文件对话框

        public string OpenFile(string Filter, string Title)
        {

            string ExcelSourcePath = "";

            OpenFileDialog fd = new OpenFileDialog
            {
                Filter = Filter,
                FilterIndex = 1,
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                Title = "\n 请选择" + Title,
                Multiselect = false,
                ReadOnlyChecked = false
            };

            if (fd.ShowDialog() != DialogResult.OK)
            {
                string info = "\n 您没有选择" + Title;
                Console.WriteLine(info);
                string check = "\n 请选择您要的操作";
                Console.WriteLine(check);
                if (UserChooseOption(info + "是否继续操作（程序将重启）"))
                {
                    Application.Restart();
                    Environment.Exit(0);
                }
                else
                {
                    Environment.Exit(0);
                }
            }
            if (fd.FileName != null)
            {
                ExcelSourcePath = fd.FileName;//文件的全路径       
            }
            return ExcelSourcePath;

        }
        #endregion

        #region  选取文件夹位置对话框

        public string OpenFolderBrowreDialog(string Description, string SelectedPath = "")
        {
            FolderBrowserDialog openFolder = new FolderBrowserDialog
            {
                Description = Description,
                ShowNewFolderButton = true
            };

            if (openFolder.ShowDialog() != DialogResult.OK)
            {
                string info = "\n 您没有选择要保存的文件夹！";
                Console.WriteLine(info);
                if (UserChooseOption(info + "! 是否继续操作（程序将重启）"))
                {
                    Application.Restart();
                    Environment.Exit(0);
                }
                else
                {
                    Environment.Exit(0);
                }
            }
            else
            {
                //提示信息
                //记录选中的目录  
                SelectedPath = openFolder.SelectedPath;
            }
            return SelectedPath;
        }
        #endregion

        #region  判断用户选择
        public bool UserChooseOption(string info)
        {
            Console.Beep();
            bool ChooseOption = false;

            if (MessageBox.Show(info, "温馨提示", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                ChooseOption = true;
            }

            return ChooseOption;

        }
        #endregion

    }
}
