using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;
using System.IO;
using System.Text.RegularExpressions;

namespace FunWithWord
{
    class Program
    {
        const int BORDEROFUNVISIBLECHARS = 32;
        
        static void Main(string[] args)
        {
            string inputdirectory = "";
            if (args.Length == 0)
            {
                Console.WriteLine("Please, run this app with parameters./nWrite directory with your data:");
                inputdirectory = Console.ReadLine();
            }
            if (inputdirectory == "") inputdirectory = args[0];
            Console.WriteLine("Input directory: {0}", inputdirectory);
            string outputdirectory = (args.Length < 2) ? inputdirectory : args[1];
            Console.WriteLine("Output directory: {0}", outputdirectory);
            DirectoryInfo di = new DirectoryInfo(inputdirectory);       //found all good files in directory (.doc and .rtf)
            List<FileInfo> fil = di.GetFiles("*.rtf").ToList<FileInfo>();
            fil =  fil.Concat(di.GetFiles("*.doc").ToList<FileInfo>()).ToList<FileInfo>();
#if (DEBUG)
            #region Estimate files List
            Console.WriteLine("#### List of Estimate files ####");
            foreach (FileInfo fi in fil)
            {
                Console.WriteLine("Estimate: {0}", fi.Name);
            }
            Console.WriteLine("#### End of Estimate list ####");
            #endregion
#endif
            WordTableParser wtp = new WordTableParser();
            ExcelOutput exo = new ExcelOutput();
            foreach (FileInfo fi in fil)        //for each file - parsing
            {
                Console.WriteLine();
                Console.WriteLine("Parsing " + fi.Name);
                Estimate currentEstimate = wtp.Parsing(inputdirectory + "\\" + fi.Name);
                Console.WriteLine("Parsing " + fi.Name + " complite");
                exo.FillWith(currentEstimate);
            }
            exo.Close(outputdirectory);
            Console.WriteLine("DONE!");
            Console.ReadLine();
        }

    }
}
