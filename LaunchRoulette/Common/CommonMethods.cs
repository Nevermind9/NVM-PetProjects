using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Excel;
using LaunchRoulette.Entities;

namespace LaunchRoulette.Common
{
    public class CommonMethods
    {
        public static bool IsPathRelatedToExcelFile(string fullPath)
        {
            return fullPath != string.Empty
                   && (fullPath.EndsWith(".xls")
                       || fullPath.EndsWith(".xlsx"));
        }

        /// <summary>
        /// Method reads the first column of excel file and returns a list of rows read
        /// </summary>
        /// <param name="excelFilePath"></param>
        /// <returns></returns>
        public static List<PersonObject> GetContenderList(string excelFilePath)
        {
            bool uplod = true;
            string fleUpload = Path.GetExtension(excelFilePath);
            var contenderList = new List<PersonObject>();
            if (fleUpload.Trim().ToLower() == ".xls" || fleUpload.Trim().ToLower() == ".xlsx")
            {
                
                try
                {
                    var dt = xlsInsert(excelFilePath);
                    
                    var colCount = 0;

                    if (dt.Columns.Count == 0)
                    {
                        MessageBox.Show("The excel file is empty.");
                        return new List<PersonObject>();
                    }
                    else if (dt.Columns.Count == 1)
                        colCount = 1;
                    else
                        colCount = 2;


                    foreach (DataRow row in dt.Rows)
                    {
                        var nameOfUser = row[0] as string;
                        if (nameOfUser == null) continue;

                        string mailOfUser = string.Empty;
                        if(colCount == 2)
                        mailOfUser = row[1] as string ?? string.Empty;

                        contenderList.Add(new PersonObject()
                        {
                            FullName = nameOfUser.Trim(),
                            EMail = mailOfUser.Trim()
                        });
                    }


                    
                }
                catch (Exception exception)
                {
                    MessageBox.Show(string.Format("An issue occured in: ReadExcelFile. Exception stack: {0}.", exception.StackTrace));
                }
            }

            return contenderList;
        }

        /// <summary>
        /// Method transforms the excel into DataTable and returns it.
        /// </summary>
        /// <param name="pth"></param>
        /// <returns></returns>
        public static DataTable xlsInsert(string pth)
        {
            FileStream stream = File.Open(pth, FileMode.Open, FileAccess.Read);
            IExcelDataReader excelReader;

            if (Path.GetExtension(pth).ToLower().Equals(".xls"))
            {
                excelReader = ExcelReaderFactory.CreateBinaryReader(stream);
            }
            else if (Path.GetExtension(pth).ToLower().Equals(".xlsx"))
            {
                excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            }
            else
            {
                throw new Exception("The file must be an .xls or .xlsx file.");
            }

            excelReader.IsFirstRowAsColumnNames = true;
            DataSet result = excelReader.AsDataSet();

            excelReader.Close();
            return result.Tables[0];
        }

        public static void SaveSettingsToConfig(string key,string value)
        {
            var configFile = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            var settings = configFile.AppSettings.Settings;
            settings[key].Value = value;
            
            configFile.Save(ConfigurationSaveMode.Modified);
            ConfigurationManager.RefreshSection("appSettings");
        }
    }
}
