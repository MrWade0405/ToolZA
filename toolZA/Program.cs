using LinqToExcel;
using System;
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

            if (Directory.Exists(rootDirName))
            {
                Console.WriteLine("Categories:");
                string[] RootDirs = Directory.GetDirectories(rootDirName);

                foreach (string s in RootDirs)
                {
                    ConsoleWriteExcel(s);

                    string[] dirs = Directory.GetDirectories(s);
                    Console.WriteLine("SubCategories:");

                    foreach (string d in dirs)
                    {
                        DirectoryInfo dirInfo = new DirectoryInfo(d);
                        if (dirInfo.Name == "pictures") continue;
                        Console.WriteLine(dirInfo.Name);

                        ConsoleWriteExcel(d);


                    }

                    Console.WriteLine();

                }
                
                Console.WriteLine();
                Console.WriteLine("RootExcel:");
                string[] files = Directory.GetFiles(rootDirName);
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
            if (RootDirInfo.Name == "pictures") return;
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