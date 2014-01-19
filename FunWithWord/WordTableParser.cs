using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace FunWithWord
{
    class WordTableParser
    {
        WordTable tableForParse;

        public WordTableParser()
        {
            tableForParse = new WordTable();
        }

        public Estimate Parsing(string filepath, int tableNumber)
        {
            tableForParse.ConnectToDocment(filepath);
            tableForParse.ChooseTable(tableNumber);
            Estimate es = new Estimate(filepath.Substring(filepath.Length - 10));
            Console.WriteLine("No way" + tableForParse.GetElementCount);
            string isnumberpattern = "^[0-9]{1,}\\.$";
            for (int i = 1; i < tableForParse.GetElementCount; i++)
            {
                if ((tableForParse[i].ColumnIndex == 1) &&
                    (IsInScheme(RemoveUnvisibleCharacters(tableForParse[i].Range.Text), isnumberpattern)))
                {
                    Console.WriteLine("Try to parse element #{0}", tableForParse[i].Range.Text.Replace(Convert.ToChar(7),'.'));
                    es.Add(ParsingString(i));
                    //Console.WriteLine(es[es.StringCount - 1].ToString());
                }
            }
            return es;
        }
        // Хуита нахуй передлать потом нормально
        EstimateString ParsingString(int firstCell)
        {
            string numberpattern = "^[0-9]{1,}";
            int number = Convert.ToInt32(ShemePart(RemoveUnvisibleCharacters(tableForParse[firstCell].Range.Text), numberpattern));
            EstimateString resultString = new EstimateString(number);
            string namecaption = tableForParse[firstCell + 1].Range.Text;
            int divider = namecaption.IndexOf(Convert.ToChar(13));
            resultString.Name = namecaption.Substring(1, divider - 1);
            resultString.Caption = namecaption.Substring(divider + 1, namecaption.Length - divider-1);
            if (NormalizeNumber(tableForParse[firstCell+2].Range.Text) != "")
                resultString.Volume = Convert.ToDouble(NormalizeNumber(tableForParse[firstCell + 2].Range.Text));
            if (NormalizeNumber(tableForParse[firstCell + 12].Range.Text) != "")
                resultString.CurrentWorkers = Convert.ToDouble(NormalizeNumber(tableForParse[firstCell + 12].Range.Text));
            if (NormalizeNumber(tableForParse[firstCell + 13].Range.Text) != "")
                resultString.CurrentMachine = Convert.ToDouble(NormalizeNumber(tableForParse[firstCell + 13].Range.Text));
            if (NormalizeNumber(tableForParse[firstCell + 20].Range.Text) != "")
                resultString.CurrentMaterials = Convert.ToDouble(NormalizeNumber(tableForParse[firstCell + 20].Range.Text));
            if (NormalizeNumber(tableForParse[firstCell + 21].Range.Text) != "")
                resultString.CurrentMachineWorkers = Convert.ToDouble(NormalizeNumber(tableForParse[firstCell + 21].Range.Text));
            return resultString;
        }

        string NormalizeNumber(string inputstring)
        {
            string outputstring = RemoveUnvisibleCharacters(inputstring);
            outputstring = outputstring.Replace(".", ",");
            outputstring = outputstring.Replace(" ", "");
            return outputstring;
        }

        string RemoveUnvisibleCharacters(string inputstring)
        {
            string str = "";
            for (int i = 0; i < inputstring.Length; i++)
            {
                if (Convert.ToInt32(inputstring[i]) >= 32)
                    str += inputstring[i];
            }
            return str;
        }

        bool IsInScheme(string inputString, string pattern)
        {
            return Regex.IsMatch(inputString, pattern);
        }

        string ShemePart(string inputstring, string pattern)
        {
            return Regex.Match(inputstring, pattern).Value;
        }
    }
}
