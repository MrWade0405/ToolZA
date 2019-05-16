using LinqToExcel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;

namespace toolZA
{
    class Program
    {
        static void Main(string[] args)
        {
            const string rootDirName = @"D:\DictionaryForFullStack\DictionaryForFullStack";
            string connectionString = ConfigurationManager.ConnectionStrings["LearningLanguages"].ConnectionString;
            
            if (Directory.Exists(rootDirName))
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();

                    string[] rootDirFiles = Directory.GetFiles(rootDirName);

                    List<string> columnNames = GetColumnNames(rootDirFiles[0]);

                    var categoriesRootTable = ExcelSelect(rootDirFiles[0]).ToList();
                    var languagesTable = ExcelSelect(rootDirFiles[1]).ToList();
                    var testNamesTable = ExcelSelect(rootDirFiles[2]).ToList();

                    string[] picturesCategories = Directory.GetFiles(rootDirName + "\\pictures");

                    string[] rootDirDirs = Directory.GetDirectories(rootDirName);

                    int iteratorCatRows = 0;

                    FillLanguageTable(languagesTable, connection, columnNames);

                    foreach (string categoryDir in rootDirDirs)
                    {
                        DirectoryInfo catDirInfo = new DirectoryInfo(categoryDir);
                        if ((catDirInfo.Name == "Test_Icons") || (catDirInfo.Name == "pictures")) continue;

                        var pictureCategory = Array.Find(picturesCategories, s => s.ToLower().Contains(catDirInfo.Name.ToLower()));
                        SqlCommand commandCat = new SqlCommand();
                        commandCat.Connection = connection;
                        commandCat.Parameters.Add(new SqlParameter("@name", catDirInfo.Name));
                        commandCat.Parameters.Add(new SqlParameter("@picture", pictureCategory));
                        commandCat.CommandText = "INSERT INTO Categories (name, picture) VALUES (@name, @picture); SELECT SCOPE_IDENTITY()";
                        var idCategory = commandCat.ExecuteScalar();

                        FillTranslationTable(languagesTable, connection, columnNames, idCategory, categoriesRootTable, iteratorCatRows);

                        iteratorCatRows++;

                        string[] picturesSubCategories = Directory.GetFiles(categoryDir + "\\pictures");

                        //ConsoleWriteExcel(categoryDir);

                        string[] subCategoriesDirs = Directory.GetDirectories(categoryDir);

                        int iteratorSubCatRows = 0;
                        foreach (string subCategoryDir in subCategoriesDirs)
                        {
                            DirectoryInfo subCategoryDirInfo = new DirectoryInfo(subCategoryDir);
                            if (subCategoryDirInfo.Name == "pictures") continue;

                            var categoryTable = ExcelSelect(Directory.GetFiles(categoryDir)[0]).ToList();

                            //ConsoleWriteExcel(subCategoryDirInfo.FullName);
                            var pictureSubCategory = Array.Find(picturesSubCategories, s => s.ToLower().Contains(subCategoryDirInfo.Name.ToLower()));
                            SqlCommand commandSubCat = new SqlCommand();
                            commandSubCat.Connection = connection;
                            commandSubCat.Parameters.Add(new SqlParameter("@name", subCategoryDirInfo.Name));
                            commandSubCat.Parameters.Add(new SqlParameter("@parent_id", idCategory));
                            commandSubCat.Parameters.Add(new SqlParameter("@picture", pictureSubCategory));
                            commandSubCat.CommandText = $"INSERT INTO Categories (name, parent_id, picture) VALUES (@name, @parent_id, @picture); SELECT SCOPE_IDENTITY()";

                            var idSubCategory = commandSubCat.ExecuteScalar();

                            FillTranslationTable(languagesTable, connection, columnNames, idSubCategory, categoryTable, iteratorSubCatRows);

                            iteratorSubCatRows++;
                            //string[] subCategoryDirDirs = Directory.GetDirectories(subCategoryDirInfo.FullName);
                            //string picturesDir = new DirectoryInfo(subCategoryDirDirs[0]).FullName;

                            ////ConsoleWriteExcel(picturesDir);

                            //string pronounceDir = new DirectoryInfo(subCategoryDirDirs[1]).FullName;
                            //string[] pronounceSubDirs = Directory.GetDirectories(pronounceDir);

                            //foreach (string pronounceSubDir in pronounceSubDirs)
                            //{
                            //    DirectoryInfo pronounceSubDirInfo = new DirectoryInfo(pronounceSubDir);
                            //    string[] pronounceSubDirFiles = Directory.GetFiles(pronounceSubDirInfo.FullName);

                            //    foreach (string pronounceSubDirFile in pronounceSubDirFiles)
                            //    {
                            //        FileInfo pronounceSubDirFileInfo = new FileInfo(pronounceSubDirFile);
                            //    }
                            //}
                        }
                    }





                }
            }
            else
            {
                Console.WriteLine("Not Found Folder");
            }
            Console.WriteLine("Finished!!!");
            Console.ReadKey();
        }
        public static string GetTranslation(ExcelTable excelTable, string column)
        {
            switch (column)
            {
                case "UA":
                    return excelTable.UA;
                    break;
                case "RU":
                    return excelTable.RU;
                    break;
                case "ENU":
                    return excelTable.ENU;
                    break;
                case "GER":
                    return excelTable.GER;
                    break;
                case "CHI":
                    return excelTable.CHI;
                    break;
                case "POR":
                    return excelTable.POR;
                    break;
                case "SPA":
                    return excelTable.SPA;
                    break;
                case "POL":
                    return excelTable.POL;
                    break;
                default:
                    return "";
                    break;
            }
            
        }
        /*public static void ConsoleWriteExcel(string nameDir)
        {
            DirectoryInfo dirInfo = new DirectoryInfo(nameDir);
            if (dirInfo.Name == "pictures")
            {
                string[] filesPictures = Directory.GetFiles(dirInfo.FullName);
                foreach (string filePicture in filesPictures)
                {
                    FileInfo filePictureInfo = new FileInfo(filePicture);
                }
                return;
            }
            if (dirInfo.Name == "Test_Icons")
            {
                string dirTestIcons = Directory.GetDirectories(dirInfo.FullName)[0];
                string dirWhiteIcons = Directory.GetDirectories(dirTestIcons)[1];
                string[] filesWhiteIcons = Directory.GetFiles(dirWhiteIcons);
                foreach (string fileWhiteIcon in filesWhiteIcons)
                {
                    FileInfo fileWhiteIconInfo = new FileInfo(fileWhiteIcon);
                    if (fileWhiteIconInfo.Name.StartsWith("Test"))
                    {

                    }
                }
                return;
            }
            string exсelTableFile = Directory.GetFiles(nameDir)[0];

            //ExcelSelect(exсelTableFile);
        }*/
        public static IQueryable<ExcelTable> ExcelSelect(string nameFile)
        {
            FileInfo fileInfO = new FileInfo(nameFile);

            var excel = new ExcelQueryFactory(fileInfO.FullName);
            var worksheetList = excel.GetWorksheetNames().ToList();
            var excelTable = from c in excel.Worksheet<ExcelTable>(worksheetList[0])
                                 where c.Name != ""
                                 select c;
            return excelTable;
        }
        public static List<string> GetColumnNames(string nameFile)
        {
            FileInfo fileInfO = new FileInfo(nameFile);
            var excel = new ExcelQueryFactory(fileInfO.FullName);
            var worksheetList = excel.GetWorksheetNames().ToList();
            return excel.GetColumnNames(worksheetList[0]).ToList();
        }
        public static void FillLanguageTable (List<ExcelTable> languagesTable, SqlConnection connection, List<string> columnNames)
        {
            int iteratorLangTableColumns = 1;
            foreach (var lang in languagesTable)
            {
                if (!String.IsNullOrEmpty(lang.Name))
                {
                    SqlCommand commandLang = new SqlCommand();
                    commandLang.Connection = connection;
                    commandLang.Parameters.Add(new SqlParameter("@name", columnNames[iteratorLangTableColumns]));
                    commandLang.CommandText = "INSERT INTO Languages (name) VALUES (@name)";
                    commandLang.ExecuteNonQuery();
                    iteratorLangTableColumns++;
                }
            }
        }
        public static void FillTranslationTable (List<ExcelTable> languagesTable, SqlConnection connection, List<string> columnNames, object id, List<ExcelTable> transTable, int iteratorTransTable)
        {
            int iteratorLangTableColumns = 1;
            foreach (var lang in languagesTable)
            {
                if (!String.IsNullOrEmpty(lang.Name))
                {
                    SqlCommand commandLang = new SqlCommand();
                    commandLang.Connection = connection;
                    commandLang.CommandText = $"SELECT Id FROM Languages WHERE name='{columnNames[iteratorLangTableColumns]}'";
                    SqlDataReader idLanguageQuery = commandLang.ExecuteReader();
                    idLanguageQuery.Read();
                    var idLanguage = idLanguageQuery.GetValue(0);
                    idLanguageQuery.Close();

                    SqlCommand commandCatTrans = new SqlCommand();
                    commandCatTrans.Connection = connection;
                    commandCatTrans.Parameters.Add(new SqlParameter("@category_id", id));
                    commandCatTrans.Parameters.Add(new SqlParameter("@lang_id", idLanguage));
                    commandCatTrans.Parameters.Add(new SqlParameter("@translation", GetTranslation(transTable[iteratorTransTable], columnNames[iteratorLangTableColumns])));
                    commandCatTrans.CommandText = "INSERT INTO CategoriesTranslations (category_id, lang_id, translation) VALUES (@category_id, @lang_id, @translation); SELECT SCOPE_IDENTITY()";
                    var idCategoryTranslation = commandCatTrans.ExecuteScalar();
                    iteratorLangTableColumns++;
                }
            }
        }
    }
}