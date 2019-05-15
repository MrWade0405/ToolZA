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
            Console.OutputEncoding = Encoding.UTF8;
            const string rootDirName = @"D:\DictionaryForFullStack\DictionaryForFullStack";
            string connectionString = ConfigurationManager.ConnectionStrings["LearningLanguages"].ConnectionString;
            

            if (Directory.Exists(rootDirName))
            {
                Console.WriteLine("Categories:");
                string[] RootDirs = Directory.GetDirectories(rootDirName);

                foreach (string s in RootDirs)
                {
                    ConsoleWriteExcel(s);

                    DirectoryInfo rootDirInfo = new DirectoryInfo(s);
                    if (rootDirInfo.Name == "Test_Icons") continue;

                    string[] dirs = Directory.GetDirectories(s);
                    Console.WriteLine("SubCategories:");

                    foreach (string d in dirs)
                    {
                        DirectoryInfo dirInfo = new DirectoryInfo(d);
                        ConsoleWriteExcel(dirInfo.FullName);

                        if (dirInfo.Name != "pictures")
                        {
                            string[] subDirs = Directory.GetDirectories(dirInfo.FullName);
                            string picturesWordsDir = new DirectoryInfo(subDirs[0]).FullName;
                            ConsoleWriteExcel(picturesWordsDir);

                            Console.WriteLine("Pronounce: ");
                            string pronounceWordsDir = new DirectoryInfo(subDirs[1]).FullName;
                            string[] pronounceWordsSubDirs = Directory.GetDirectories(pronounceWordsDir);
                            foreach (string p in pronounceWordsSubDirs)
                            {
                                DirectoryInfo subDirsPronounceInfo = new DirectoryInfo(p);
                                Console.WriteLine(subDirsPronounceInfo.Name);
                                string[] filesSubDirsPronounce = Directory.GetFiles(subDirsPronounceInfo.FullName);

                                foreach (string k in filesSubDirsPronounce)
                                {
                                    FileInfo fileInf = new FileInfo(k);
                                    Console.WriteLine(fileInf.Name);
                                }
                            }
                        }
                    }

                    Console.WriteLine();

                }
                
                Console.WriteLine();
                Console.WriteLine("RootExcel:");
                string[] files = Directory.GetFiles(rootDirName);





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








                foreach (string s in files)
                {
                    ExcelSelect(s);
                }
            }
            else
            {
                Console.WriteLine("Not Found Folder");
            }
            Console.ReadKey();
        }
        public static void ConsoleWriteExcel(string nameDir)
        {
            DirectoryInfo RootDirInfo = new DirectoryInfo(nameDir);
            if (RootDirInfo.Name == "pictures")
            {
                Console.WriteLine("Pictures: ");
                string[] filesPictures = Directory.GetFiles(RootDirInfo.FullName);
                foreach (string s in filesPictures)
                {
                    FileInfo fileInf = new FileInfo(s);
                    Console.WriteLine(fileInf.Name);
                }
                return;
            }
            if (RootDirInfo.Name == "Test_Icons")
            {
                string dirTestIcons = Directory.GetDirectories(RootDirInfo.FullName)[0];
                string dirWhiteIcons = Directory.GetDirectories(dirTestIcons)[1];
                string[] filesWhiteIcons = Directory.GetFiles(dirWhiteIcons);
                foreach (string s in filesWhiteIcons)
                {
                    FileInfo fileInf = new FileInfo(s);
                    if (fileInf.Name.StartsWith("Test"))
                        Console.WriteLine(fileInf.Name);
                }
                return;
            }
            Console.WriteLine(RootDirInfo.Name);
            string fileCategories = Directory.GetFiles(nameDir)[0];
            ExcelSelect(fileCategories);

        }
        public static void ExcelSelect(string nameFile)
        {
            FileInfo fileInf = new FileInfo(nameFile);
            Console.WriteLine(fileInf.Name);
            var excel = new ExcelQueryFactory(fileInf.FullName);
            var worksheetList = excel.GetWorksheetNames().ToList();
            var categoriesRoot = from c in excel.Worksheet<CategoriesRoot>(worksheetList[0])
                                 select c;
            foreach (var i in categoriesRoot)
            {
                if (!String.IsNullOrEmpty(i.Name))
                    Console.WriteLine($"{i.Name}    {i.UA}    {i.RU}    " +
                        $"{i.ENU}    {i.GER}    {i.CHI}    {i.POR}    {i.SPA}    {i.POL}");
            }
        }


    }
}