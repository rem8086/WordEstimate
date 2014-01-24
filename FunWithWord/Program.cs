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
        const int TABLENUMBER = 3;

        static void Main(string[] args)
        {

            if (Regex.IsMatch("итого  по  смете", "итого[\\s]{1,}по[\\s]{1,}смете")) Console.WriteLine("YES");
            Console.ReadLine();
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
            DirectoryInfo di = new DirectoryInfo(inputdirectory);
            List<FileInfo> fil = di.GetFiles("*.doc").ToList<FileInfo>();
            fil.Concat(di.GetFiles("*.rtf").ToList<FileInfo>());
            List<Estimate> estimateList = new List<Estimate>();
            ExcelOutput exo = new ExcelOutput();
            foreach (FileInfo fi in fil)
            {
                Console.WriteLine("Parsing "+inputdirectory+"\\"+fi.Name);
                WordTableParser wtp = new WordTableParser();
                Estimate currentEstimate = wtp.Parsing(inputdirectory + "\\" + fi.Name, TABLENUMBER);
                estimateList.Add(currentEstimate);
                Console.WriteLine("Parsing complite");
                exo.FillWith(currentEstimate);
            }
            exo.Close(outputdirectory);
            Console.ReadLine();
        }

        static string RemoveUnvisibleCharacters(string inputstring)
        {
            string str = "";
            for (int i = 0; i < inputstring.Length; i++)
            {
                if (Convert.ToInt32(inputstring[i]) >= BORDEROFUNVISIBLECHARS)
                    str += inputstring[i];
            }
            return str;
        }
    }
}
