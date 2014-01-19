using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
//using Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.IO;
using System.Text.RegularExpressions;

namespace FunWithWord
{
    class Program
    {
        static void Main(string[] args)
        {
            string inputdirectory = "";
            if (args.Length == 0)
            {
                Console.WriteLine("Please, run this app with parameters./nWrite directory with your data:");
                inputdirectory = Console.ReadLine();
            }
            if (inputdirectory == "") inputdirectory = args[0]; //Console.ReadLine();
            string outputdirectory = (args.Length < 2) ? inputdirectory : args[1];
            DirectoryInfo di = new DirectoryInfo(inputdirectory);
            List<FileInfo> fil = di.GetFiles("*.doc").ToList<FileInfo>();
            List<Estimate> estimateList = new List<Estimate>();
            Application ap = new Application();
            ap.Visible = true;
            Workbook wb = ap.Workbooks.Add();
            ExcelOutput exo = new ExcelOutput(wb);
            foreach (FileInfo fi in fil)
            {
                Console.WriteLine("Parsing "+inputdirectory+"\\"+fi.Name);
                WordTableParser wtp = new WordTableParser();
                Estimate currentEstimate = wtp.Parsing(inputdirectory + "\\" + fi.Name, 3);
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
                if (Convert.ToInt32(inputstring[i]) >= 32)
                    str += inputstring[i];
            }
            return str;
        }
    }
}
