using LinqToExcel;
using System;
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
            //Console.OutputEncoding = Encoding.UTF8;
            const string rootDirName = @"D:\DictionaryForFullStack\DictionaryForFullStack";
            string connectionString = ConfigurationManager.ConnectionStrings["LearningLanguages"].ConnectionString;
            
            if (Directory.Exists(rootDirName))
            {
                //Console.WriteLine("Categories:");
                string[] rootDirDirs = Directory.GetDirectories(rootDirName);

                foreach (string categoryDir in rootDirDirs)
                {
                    ConsoleWriteExcel(categoryDir);

                    DirectoryInfo catDirInfo = new DirectoryInfo(categoryDir);
                    if (catDirInfo.Name == "Test_Icons") continue;

                    string[] subCategoriesDirs = Directory.GetDirectories(categoryDir);
                    //Console.WriteLine("SubCategories:");

                    foreach (string subCategoryDir in subCategoriesDirs)
                    {
                        DirectoryInfo subCategoryDirInfo = new DirectoryInfo(subCategoryDir);

                        ConsoleWriteExcel(subCategoryDirInfo.FullName);

                        if (subCategoryDirInfo.Name != "pictures")
                        {
                            string[] subCategoryDirDirs = Directory.GetDirectories(subCategoryDirInfo.FullName);
                            string picturesDir = new DirectoryInfo(subCategoryDirDirs[0]).FullName;

                            ConsoleWriteExcel(picturesDir);

                            //Console.WriteLine("Pronounce: ");
                            string pronounceDir = new DirectoryInfo(subCategoryDirDirs[1]).FullName;
                            string[] pronounceSubDirs = Directory.GetDirectories(pronounceDir);

                            foreach (string pronounceSubDir in pronounceSubDirs)
                            {
                                DirectoryInfo pronounceSubDirInfo = new DirectoryInfo(pronounceSubDir);
                                //Console.WriteLine(pronounceSubDirInfo.Name);
                                string[] pronounceSubDirFiles = Directory.GetFiles(pronounceSubDirInfo.FullName);

                                foreach (string pronounceSubDirFile in pronounceSubDirFiles)
                                {
                                    FileInfo pronounceSubDirFileInfo = new FileInfo(pronounceSubDirFile);
                                    //Console.WriteLine(pronounceSubDirFile.Name);
                                }
                            }
                        }
                    }

                    //Console.WriteLine();

                }
                
                //Console.WriteLine();
                //Console.WriteLine("RootExcel:");
                string[] rootDirFiles = Directory.GetFiles(rootDirName);





                //////
                //string file = Directory.GetFiles(rootDirName)[0];
                //FileInfo fileInfo = new FileInfo(file);
                //var excel = new ExcelQueryFactory(fileInfo.FullName);
                //var worksheetList = excel.GetWorksheetNames().ToList();
                //var categoriesRoot = from c in excel.Worksheet<CategoriesRoot>(worksheetList[0])
                //                     select c;
                //using (SqlConnection connection = new SqlConnection(connectionString))
                //{
                //    connection.Open();
                //    SqlCommand command = new SqlCommand();
                //    command.Connection = connection;
                //    foreach (var i in categoriesRoot)
                //    {
                //        if (!String.IsNullOrEmpty(i.Name))
                //        {
                //            command.CommandText = $"INSERT INTO Categories (name, parentId, picture) VALUES ('{i.Name}', '{null}', '{fileInfo.FullName}')";
                //            //command.CommandText = $"DELETE FROM Categories WHERE name='{i.Name}'";

                //            command.ExecuteNonQuery();
                //        }
                //    }
                //}
                ////////








                foreach (string rootDirFile in rootDirFiles)
                {
                    ExcelSelect(rootDirFile);
                }
            }
            else
            {
                Console.WriteLine("Not Found Folder");
            }
            Console.WriteLine("Finished!!!");
            Console.ReadKey();
        }
        public static void ConsoleWriteExcel(string nameDir)
        {
            DirectoryInfo dirInfo = new DirectoryInfo(nameDir);
            if (dirInfo.Name == "pictures")
            {
                //Console.WriteLine("Pictures: ");
                string[] filesPictures = Directory.GetFiles(dirInfo.FullName);
                foreach (string filePicture in filesPictures)
                {
                    FileInfo filePictureInfo = new FileInfo(filePicture);
                    //Console.WriteLine(filePictureInfo.Name);
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
                        //Console.WriteLine(fileWhiteIconInfo.Name);
                    }
                }
                return;
            }
            //Console.WriteLine(dirInfo.Name);
            string exсelTableFile = Directory.GetFiles(nameDir)[0];

            ExcelSelect(exсelTableFile);
        }
        public static void ExcelSelect(string nameFile)
        {
            FileInfo fileInfO = new FileInfo(nameFile);
            //Console.WriteLine(fileInfO.Name);
            var excel = new ExcelQueryFactory(fileInfO.FullName);
            var worksheetList = excel.GetWorksheetNames().ToList();
            var excelTable = from c in excel.Worksheet<ExcelTable>(worksheetList[0])
                                 select c;
            foreach (var row in excelTable)
            {
                if (!String.IsNullOrEmpty(row.Name))
                {
                    //Console.WriteLine($"{i.Name}    {i.UA}    {i.RU}    " +
                    //    $"{i.ENU}    {i.GER}    {i.CHI}    {i.POR}    {i.SPA}    {i.POL}");
                }
            }
        }


    }
}