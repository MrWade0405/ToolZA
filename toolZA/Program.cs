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

                    FillTranslationLanguagesTable(languagesTable, connection, columnNames);

                    FillTestTable(testNamesTable, connection, columnNames, languagesTable, rootDirName);

                    foreach (string categoryDir in rootDirDirs)
                    {
                        DirectoryInfo catDirInfo = new DirectoryInfo(categoryDir);

                        if ((catDirInfo.Name == "Test_Icons") || (catDirInfo.Name == "pictures")) continue;

                        var pictureCategory = Array.Find(picturesCategories, s => s.ToLower().Contains(catDirInfo.Name.ToLower()));

                        SqlCommand commandCat = new SqlCommand();
                        commandCat.Connection = connection;
                        commandCat.Parameters.Add(new SqlParameter("@name", catDirInfo.Name));
                        commandCat.Parameters.Add(new SqlParameter("@picture", pictureCategory));
                        commandCat.CommandText = "INSERT INTO Categories (name, picture) VALUES (@name, @picture); " +
                                                 "SELECT SCOPE_IDENTITY()";
                        var idCategory = commandCat.ExecuteScalar();

                        FillTranslationTable(languagesTable, connection, columnNames, idCategory, categoriesRootTable, iteratorCatRows,
                                             "CategoriesTranslations", "category_id", null);

                        iteratorCatRows++;

                        string[] picturesSubCategories = Directory.GetFiles(categoryDir + "\\pictures");

                        string[] subCategoriesDirs = Directory.GetDirectories(categoryDir);

                        int iteratorSubCatRows = 0;

                        foreach (string subCategoryDir in subCategoriesDirs)
                        {
                            DirectoryInfo subCategoryDirInfo = new DirectoryInfo(subCategoryDir);

                            if (subCategoryDirInfo.Name == "pictures") continue;

                            var categoryTable = ExcelSelect(Directory.GetFiles(categoryDir)[0]).ToList();
                            var pictureSubCategory = Array.Find(picturesSubCategories, s => s.ToLower().Contains
                                                                                       (subCategoryDirInfo.Name.ToLower()));

                            SqlCommand commandSubCat = new SqlCommand();
                            commandSubCat.Connection = connection;
                            commandSubCat.Parameters.Add(new SqlParameter("@name", subCategoryDirInfo.Name));
                            commandSubCat.Parameters.Add(new SqlParameter("@parent_id", idCategory));
                            commandSubCat.Parameters.Add(new SqlParameter("@picture", pictureSubCategory));
                            commandSubCat.CommandText = "INSERT INTO Categories (name, parent_id, picture) " +
                                                        "VALUES (@name, @parent_id, @picture); SELECT SCOPE_IDENTITY()";
                            var idSubCategory = commandSubCat.ExecuteScalar();

                            FillTranslationTable(languagesTable, connection, columnNames, idSubCategory, categoryTable,
                                                 iteratorSubCatRows, "CategoriesTranslations", "category_id", null);

                            iteratorSubCatRows++;

                            var subCategoryTable = ExcelSelect(Directory.GetFiles(subCategoryDir)[0]).ToList();

                            string[] subCategoryDirDirs = Directory.GetDirectories(subCategoryDirInfo.FullName);

                            FillWordsTable(subCategoryTable, connection, columnNames, languagesTable, subCategoryDirDirs, idSubCategory);
                        }
                    }
                }
            }
            else
            {
                Console.WriteLine("Not Found Folder");
            }
            Console.WriteLine("Data loaded to database");
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
        public static IQueryable<ExcelTable> ExcelSelect(string nameFile)
        {
            FileInfo fileInfO = new FileInfo(nameFile);

            var excel = new ExcelQueryFactory(fileInfO.FullName);
            var worksheetList = excel.GetWorksheetNames().ToList();

            if (nameFile.Contains("Languages"))
            {
                return from c in excel.Worksheet<ExcelTable>(worksheetList[0])
                       where c.Name != ""
                       select c;
            }
            else
            {
                return from c in excel.Worksheet<ExcelTable>(worksheetList[0])
                       where c.Name != ""
                       orderby c.Name
                       select c;
            }
        }
        public static List<string> GetColumnNames(string nameFile)
        {
            FileInfo fileInfO = new FileInfo(nameFile);

            var excel = new ExcelQueryFactory(fileInfO.FullName);
            var worksheetList = excel.GetWorksheetNames().ToList();

            return excel.GetColumnNames(worksheetList[0]).ToList();
        }
        public static void FillWordsTable(List<ExcelTable> subCategoryTable, SqlConnection connection, List<string> columnNames,
                                          List<ExcelTable> languagesTable, string[] subCategoryDirDirs, object idSubCategory)
        {
            string picturesDir = new DirectoryInfo(subCategoryDirDirs[0]).FullName;
            string[] picturesFiles = Directory.GetFiles(picturesDir);

            string pronounceDir = new DirectoryInfo(subCategoryDirDirs[1]).FullName;
            string[] pronounceSubDirs = Directory.GetDirectories(pronounceDir);

            DirectoryInfo soundsDir = null;
            string[] soundsFiles = null;

            if (subCategoryDirDirs.Length == 3)
            {
                soundsDir = new DirectoryInfo(subCategoryDirDirs[2]);
                soundsFiles = Directory.GetFiles(soundsDir.FullName);
            }

            int iteratorWords = 0;

            foreach (var row in subCategoryTable)
            {
                SqlCommand commandWord = new SqlCommand();
                commandWord.Connection = connection;
                commandWord.Parameters.Add(new SqlParameter("@name", row.Name));
                commandWord.Parameters.Add(new SqlParameter("@category_id", idSubCategory));
                commandWord.Parameters.Add(new SqlParameter("@picture", picturesFiles[iteratorWords]));

                string soundFile = soundsDir is null ? null : Array.Find(soundsFiles, s => s.ToLower().Contains(row.Name.ToLower()));

                if (soundFile is null)
                {
                    commandWord.CommandText = "INSERT INTO Words (name, category_id, picture) VALUES (@name, @category_id, @picture); " +
                                              "SELECT SCOPE_IDENTITY()";
                }
                else
                {
                    commandWord.Parameters.Add(new SqlParameter("@sound", soundFile));
                    commandWord.CommandText = "INSERT INTO Words (name, category_id, picture, sound) " +
                                              "VALUES (@name, @category_id, @picture, @sound); SELECT SCOPE_IDENTITY()";
                }

                var idWord = commandWord.ExecuteScalar();

                FillTranslationTable(languagesTable, connection, columnNames, idWord, subCategoryTable, iteratorWords,
                                     "WordTranslations", "word_id", pronounceSubDirs);

                iteratorWords++;
            }
        }
        public static void FillLanguageTable(List<ExcelTable> languagesTable, SqlConnection connection, List<string> columnNames)
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
        public static void FillTestTable(List<ExcelTable> testNamesTable, SqlConnection connection, List<string> columnNames,
                                         List<ExcelTable> languagesTable, string rootDirName)
        {
            string dirTestIcons = Directory.GetDirectories(rootDirName + "\\Test_Icons")[0];
            string dirWhiteIcons = Directory.GetDirectories(dirTestIcons)[1];
            string[] filesWhiteIcons = Directory.GetFiles(dirWhiteIcons);

            int iteratorTestIcons = 2;

            foreach (var row in testNamesTable)
            {
                SqlCommand commandTest = new SqlCommand();
                commandTest.Connection = connection;
                commandTest.Parameters.Add(new SqlParameter("@name", row.Name));
                commandTest.Parameters.Add(new SqlParameter("@icon", filesWhiteIcons[iteratorTestIcons]));
                commandTest.CommandText = "INSERT INTO Tests (name, icon) VALUES (@name, @icon); SELECT SCOPE_IDENTITY()";
                var idTest = commandTest.ExecuteScalar();

                FillTranslationTable(languagesTable, connection, columnNames, idTest, testNamesTable, iteratorTestIcons - 2,
                                     "TestTranslations", "test_id", null);

                iteratorTestIcons++;
            }
        }
        public static void FillTranslationTable(List<ExcelTable> languagesTable, SqlConnection connection, List<string> columnNames,
                                                 object id, List<ExcelTable> transTable, int iteratorTransTable, string table,
                                                 string column, string[] pronounceSubDirs)
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

                    SqlCommand commandTrans = new SqlCommand();
                    commandTrans.Connection = connection;
                    commandTrans.Parameters.Add(new SqlParameter("@id", id));
                    commandTrans.Parameters.Add(new SqlParameter("@lang_id", idLanguage));
                    commandTrans.Parameters.Add(new SqlParameter("@translation", GetTranslation(transTable[iteratorTransTable],
                                                                                                columnNames[iteratorLangTableColumns])));

                    if (table == "WordTranslations")
                    {
                        DirectoryInfo pronounceSubDirInfo = new DirectoryInfo
                            (Array.Find(pronounceSubDirs, s => s.ToLower().Contains(columnNames[iteratorLangTableColumns].ToLower())));
                        string[] pronounceSubDirFiles = Directory.GetFiles(pronounceSubDirInfo.FullName);
                        string pronounceFile = Array.Find(pronounceSubDirFiles, s => s.ToLower().Trim().Contains
                                                                                (transTable[iteratorTransTable].Name.ToLower().Trim()));

                        commandTrans.Parameters.Add(new SqlParameter("@pronounce", pronounceFile));
                        commandTrans.CommandText = $"INSERT INTO {table} ({column}, lang_id, translation, pronounce) " +
                                                   $"VALUES (@id, @lang_id, @translation, @pronounce)";
                    }
                    else
                    {
                        commandTrans.CommandText = $"INSERT INTO {table} ({column}, lang_id, translation) " +
                                                   $"VALUES (@id, @lang_id, @translation)";
                    }

                    commandTrans.ExecuteNonQuery();

                    iteratorLangTableColumns++;
                }
            }
        }
        public static void FillTranslationLanguagesTable(List<ExcelTable> languagesTable, SqlConnection connection,
                                                          List<string> columnNames)
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

                    FillTranslationTable(languagesTable, connection, columnNames, idLanguage, languagesTable,
                                         iteratorLangTableColumns - 1, "LanguageTranslations", "native_lang_id", null);

                    iteratorLangTableColumns++;
                }
            }
        }
    }
}