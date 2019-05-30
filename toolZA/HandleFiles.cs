using LinqToExcel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.IO;
using System.Linq;

namespace toolZA
{
    public enum Columns
    {
        UA = 1,
        RU,
        ENU,
        GER,
        CHI,
        POR,
        SPA,
        POL
    }

    class HandleFiles
    {
        public HandleFiles()
        {
            _rootDirName = @"D:\DictionaryForFullStack\DictionaryForFullStack";
            _connectionString = ConfigurationManager.ConnectionStrings["LearningLanguages"]?.ConnectionString;
            _connection = new SqlConnection(_connectionString);
        }

        private string _rootDirName;

        private string _connectionString;

        private SqlConnection _connection;

        private const string PicturesString = "pictures";
        private const string TestIconsString = "Test_Icons";
        private const string CategoriesTranslationsString = "CategoriesTranslations";
        private const string CategoryIdString = "category_id";
        private const string LanguagesString = "Languages";
        private const string TestsNamesString = "TestsNames";
        private const string WordTranslationsString = "WordTranslations";
        private const string WordIdString = "word_id";
        private const string TestTranslationsString = "TestTranslations";
        private const string TestIdString = "test_id";
        private const string LanguageTranslationsString = "LanguageTranslations";
        private const string NativeLangIdString = "native_lang_id";

        public void FillDB()
        {
            if (Directory.Exists(_rootDirName))
            {
                try
                {
                    using (_connection)
                    {
                        _connection.Open();

                        string[] rootDirFiles = Directory.GetFiles(_rootDirName);

                        List<string> columnNames = GetColumnNames(rootDirFiles[0]);

                        List<ExcelTable> categoriesRootTable = ExcelSelect(rootDirFiles[0]).ToList();
                        List<ExcelTable> languagesTable = ExcelSelect(rootDirFiles[1]).ToList();
                        List<ExcelTable> testNamesTable = ExcelSelect(rootDirFiles[2]).ToList();

                        string[] picturesCategories = Directory.GetFiles(_rootDirName + "\\pictures");

                        string[] rootDirDirs = Directory.GetDirectories(_rootDirName);

                        int iteratorCatRows = 0;

                        FillLanguageTable(languagesTable, columnNames);

                        FillTranslationLanguagesTable(languagesTable, columnNames);

                        FillTestTable(testNamesTable, columnNames, languagesTable, _rootDirName);

                        foreach (string categoryDir in rootDirDirs)
                        {
                            DirectoryInfo catDirInfo = new DirectoryInfo(categoryDir);

                            if ((catDirInfo.Name == TestIconsString) || (catDirInfo.Name == PicturesString)) continue;

                            string pictureCategory = Array.Find(picturesCategories, s => s.ToLower().Contains(catDirInfo.Name.ToLower()));

                            SqlCommand commandCat = new SqlCommand();
                            commandCat.Connection = _connection;
                            commandCat.Parameters.Add(new SqlParameter("@name", catDirInfo.Name));
                            commandCat.Parameters.Add(new SqlParameter("@picture", pictureCategory.Substring(3)));
                            commandCat.CommandText = "INSERT INTO Categories (name, picture) VALUES (@name, @picture); " +
                                                     "SELECT SCOPE_IDENTITY()";
                            object idCategory = commandCat.ExecuteScalar();

                            FillTranslationTable(languagesTable, columnNames, idCategory, categoriesRootTable, iteratorCatRows,
                                                 CategoriesTranslationsString, CategoryIdString, null);

                            iteratorCatRows++;

                            string[] picturesSubCategories = Directory.GetFiles(categoryDir + "\\pictures");

                            string[] subCategoriesDirs = Directory.GetDirectories(categoryDir);

                            int iteratorSubCatRows = 0;

                            foreach (string subCategoryDir in subCategoriesDirs)
                            {
                                DirectoryInfo subCategoryDirInfo = new DirectoryInfo(subCategoryDir);

                                if (subCategoryDirInfo.Name == PicturesString) continue;

                                List<ExcelTable> categoryTable = ExcelSelect(Directory.GetFiles(categoryDir)[0]).ToList();
                                string pictureSubCategory = Array.Find(picturesSubCategories, s => s.ToLower().Contains
                                                                                           (subCategoryDirInfo.Name.ToLower()));

                                SqlCommand commandSubCat = new SqlCommand();
                                commandSubCat.Connection = _connection;
                                commandSubCat.Parameters.Add(new SqlParameter("@name", subCategoryDirInfo.Name));
                                commandSubCat.Parameters.Add(new SqlParameter("@parent_id", idCategory));
                                commandSubCat.Parameters.Add(new SqlParameter("@picture", pictureSubCategory.Substring(3)));
                                commandSubCat.CommandText = "INSERT INTO Categories (name, parent_id, picture) " +
                                                            "VALUES (@name, @parent_id, @picture); SELECT SCOPE_IDENTITY()";
                                object idSubCategory = commandSubCat.ExecuteScalar();

                                FillTranslationTable(languagesTable, columnNames, idSubCategory, categoryTable, iteratorSubCatRows,
                                                     CategoriesTranslationsString, CategoryIdString, null);

                                iteratorSubCatRows++;

                                List<ExcelTable> subCategoryTable = ExcelSelect(Directory.GetFiles(subCategoryDir)[0]).ToList();

                                string[] subCategoryDirDirs = Directory.GetDirectories(subCategoryDirInfo.FullName);

                                FillWordsTable(subCategoryTable, columnNames, languagesTable, subCategoryDirDirs, idSubCategory);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
            else
            {
                Console.WriteLine("Not Found Folder");
            }
            Console.WriteLine("Data loaded to database");
            Console.ReadKey();
        }

        private string GetTranslation(ExcelTable excelTable, int columnId)
        {
            switch (columnId)
            {
                case (int)Columns.UA:
                    return excelTable.UA;
                case (int)Columns.RU:
                    return excelTable.RU;
                case (int)Columns.ENU:
                    return excelTable.ENU;
                case (int)Columns.GER:
                    return excelTable.GER;
                case (int)Columns.CHI:
                    return excelTable.CHI;
                case (int)Columns.POR:
                    return excelTable.POR;
                case (int)Columns.SPA:
                    return excelTable.SPA;
                case (int)Columns.POL:
                    return excelTable.POL;
                default:
                    return "";
            }
        }

        private IQueryable<ExcelTable> ExcelSelect(string nameFile)
        {
            FileInfo fileInfO = new FileInfo(nameFile);

            var excel = new ExcelQueryFactory(fileInfO.FullName);
            List<string> worksheetList = excel.GetWorksheetNames().ToList();

            if (nameFile.Contains(LanguagesString) || nameFile.Contains(TestsNamesString))
            {
                return from c in excel.Worksheet<ExcelTable>(worksheetList[0])
                       where !String.IsNullOrEmpty(c.Name)
                       select c;
            }
            else
            {
                return from c in excel.Worksheet<ExcelTable>(worksheetList[0])
                       where !String.IsNullOrEmpty(c.Name)
                       orderby c.Name
                       select c;
            }
        }

        private List<string> GetColumnNames(string nameFile)
        {
            FileInfo fileInfO = new FileInfo(nameFile);

            var excel = new ExcelQueryFactory(fileInfO.FullName);
            List<string> worksheetList = excel.GetWorksheetNames().ToList();

            return excel.GetColumnNames(worksheetList[0]).ToList();
        }

        private void FillWordsTable(List<ExcelTable> subCategoryTable, List<string> columnNames, List<ExcelTable> languagesTable,
                                    string[] subCategoryDirDirs, object idSubCategory)
        {
            string picturesDir = new DirectoryInfo(subCategoryDirDirs[0]).FullName;
            string[] picturesFiles = Directory.GetFiles(picturesDir);

            string pronounceDir = new DirectoryInfo(subCategoryDirDirs[1]).FullName;
            string[] pronounceSubDirs = Directory.GetDirectories(pronounceDir);

            DirectoryInfo soundsDir = null;
            string[] soundsFiles = null;

            //if there is a folder with sound
            if (subCategoryDirDirs.Length == 3)
            {
                soundsDir = new DirectoryInfo(subCategoryDirDirs[2]);
                soundsFiles = Directory.GetFiles(soundsDir.FullName);
            }

            int iteratorWords = 0;

            foreach (var row in subCategoryTable)
            {
                SqlCommand commandWord = new SqlCommand();
                commandWord.Connection = _connection;
                commandWord.Parameters.Add(new SqlParameter("@name", row.Name));
                commandWord.Parameters.Add(new SqlParameter("@category_id", idSubCategory));
                commandWord.Parameters.Add(new SqlParameter("@picture", picturesFiles[iteratorWords].Substring(3)));

                string soundFile = soundsDir is null ? null : Array.Find(soundsFiles, s => s.ToLower().Contains(row.Name.ToLower()));

                if (String.IsNullOrEmpty(soundFile))
                {
                    commandWord.CommandText = "INSERT INTO Words (name, category_id, picture) VALUES (@name, @category_id, @picture); " +
                                              "SELECT SCOPE_IDENTITY()";
                }
                else
                {
                    commandWord.Parameters.Add(new SqlParameter("@sound", soundFile.Substring(3)));
                    commandWord.CommandText = "INSERT INTO Words (name, category_id, picture, sound) " +
                                              "VALUES (@name, @category_id, @picture, @sound); SELECT SCOPE_IDENTITY()";
                }

                object idWord = commandWord.ExecuteScalar();

                FillTranslationTable(languagesTable, columnNames, idWord, subCategoryTable, iteratorWords, WordTranslationsString,
                                     WordIdString, pronounceSubDirs);

                iteratorWords++;
            }
        }

        private void FillLanguageTable(List<ExcelTable> languagesTable, List<string> columnNames)
        {
            //skip the "Name" field
            int iteratorLangTableColumns = 1;

            foreach (var lang in languagesTable)
            {
                if (!String.IsNullOrEmpty(lang.Name))
                {
                    SqlCommand commandLang = new SqlCommand();
                    commandLang.Connection = _connection;
                    commandLang.Parameters.Add(new SqlParameter("@name", columnNames[iteratorLangTableColumns]));
                    commandLang.CommandText = "INSERT INTO Languages (name) VALUES (@name)";
                    commandLang.ExecuteNonQuery();

                    iteratorLangTableColumns++;
                }
            }
        }

        private void FillTestTable(List<ExcelTable> testNamesTable, List<string> columnNames, List<ExcelTable> languagesTable,
                                   string rootDirName)
        {
            string dirTestIcons = Directory.GetDirectories(rootDirName + "\\Test_Icons")[0];
            string dirWhiteIcons = Directory.GetDirectories(dirTestIcons)[0];
            string[] filesWhiteIcons = Directory.GetFiles(dirWhiteIcons);

            //skip Checkbox_No and Checkbox_Yes icons
            int iteratorTestIcons = 2;

            foreach (var row in testNamesTable)
            {
                SqlCommand commandTest = new SqlCommand();
                commandTest.Connection = _connection;
                commandTest.Parameters.Add(new SqlParameter("@name", row.Name));
                commandTest.Parameters.Add(new SqlParameter("@icon", filesWhiteIcons[iteratorTestIcons].Substring(3)));
                commandTest.CommandText = "INSERT INTO Tests (name, icon) VALUES (@name, @icon); SELECT SCOPE_IDENTITY()";
                object idTest = commandTest.ExecuteScalar();

                FillTranslationTable(languagesTable, columnNames, idTest, testNamesTable, iteratorTestIcons - 2,
                                     TestTranslationsString, TestIdString, null);

                iteratorTestIcons++;
            }
        }

        private void FillTranslationTable(List<ExcelTable> languagesTable, List<string> columnNames, object id,
                                          List<ExcelTable> transTable, int iteratorTransTable, string table, string column,
                                          string[] pronounceSubDirs)
        {
            //skip the "Name" field
            int iteratorLangTableColumns = 1;
            List<int> idLanguages = GetListIdLanguages();

            foreach (var lang in languagesTable)
            {
                if (!String.IsNullOrEmpty(lang.Name))
                {
                    SqlCommand commandTrans = new SqlCommand();
                    commandTrans.Connection = _connection;
                    commandTrans.Parameters.Add(new SqlParameter("@id", id));
                    commandTrans.Parameters.Add(new SqlParameter("@lang_id", idLanguages[iteratorLangTableColumns - 1]));
                    commandTrans.Parameters.Add(new SqlParameter("@translation", GetTranslation(transTable[iteratorTransTable],
                                                                                                iteratorLangTableColumns)));

                    if (table == WordTranslationsString)
                    {
                        DirectoryInfo pronounceSubDirInfo = new DirectoryInfo
                            (Array.Find(pronounceSubDirs, s => s.ToLower().Contains(columnNames[iteratorLangTableColumns].ToLower())));
                        string[] pronounceSubDirFiles = Directory.GetFiles(pronounceSubDirInfo.FullName);
                        string pronounceFile = Array.Find(pronounceSubDirFiles, s => s.ToLower().Trim().Contains
                                                                                (transTable[iteratorTransTable].Name.ToLower().Trim()));

                        commandTrans.Parameters.Add(new SqlParameter("@pronounce", pronounceFile.Substring(3)));
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

        private void FillTranslationLanguagesTable(List<ExcelTable> languagesTable, List<string> columnNames)
        {
            //skip the "Name" field
            int iteratorLangTableColumns = 1;
            List<int> idLanguages = GetListIdLanguages();

            foreach (var lang in languagesTable)
            {
                if (!String.IsNullOrEmpty(lang.Name))
                {
                    FillTranslationTable(languagesTable, columnNames, idLanguages[iteratorLangTableColumns - 1], languagesTable,
                                         iteratorLangTableColumns - 1, LanguageTranslationsString, NativeLangIdString, null);

                    iteratorLangTableColumns++;
                }
            }
        }

        private List<int> GetListIdLanguages()
        {
            List<int> idLanguages = new List<int>();

            SqlCommand commandLang = new SqlCommand();
            commandLang.Connection = _connection;
            commandLang.CommandText = $"SELECT Id FROM Languages";
            SqlDataReader idLanguageQuery = commandLang.ExecuteReader();

            if (idLanguageQuery.HasRows)
            {
                while (idLanguageQuery.Read())
                {
                    idLanguages.Add((int)idLanguageQuery.GetValue(0));
                }
            }

            idLanguageQuery.Close();

            return idLanguages;
        }
    }
}
