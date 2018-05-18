using System.IO;
using ExcelDataReader;
using System.Data;
namespace version2
{
    class ExcelHelper
    {
        #region 
        /// <summary>
        /// 根据excle的路径把第一个sheel中的内容放入datatable
        /// </summary>
        /// <param name="SourcePath">数据源路径</param>
        /// <returns></returns>
        public DataSet ReadExcelToTable(string SourcePath)//excel存放的路径
        {
            string ExcelExtension = System.IO.Path.GetExtension(SourcePath);//文件的扩展名 //读取Excel数据

            using (FileStream stream = new FileStream(SourcePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                IExcelDataReader reader = null;
                
                if (ExcelExtension == ".xls")//判断当前的数据源文件格式
                {
                    reader = ExcelReaderFactory.CreateBinaryReader(stream);
                }
                else if (ExcelExtension == ".xlsx")
                {
                    reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                }
                else if (ExcelExtension == ".csv")
                {
                    reader = ExcelReaderFactory.CreateCsvReader(stream);
                }            
                DataSet result = reader.AsDataSet();

                return result;
            }
        }
        #endregion
    }
}
