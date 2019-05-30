using System;

namespace toolZA
{
    class Program
    {
        static void Main(string[] args)
        {
            HandleFiles handleFiles = new HandleFiles();

            handleFiles.FillDB();

            Console.ReadKey();
        }   
    }
}