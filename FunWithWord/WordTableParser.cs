using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace FunWithWord
{
    class WordTableParser
    {
        const string ISNUMBERPATTERN = "^[0-9]{1,}\\.$";
        const int NUMBERCOLUMN = 1;
        const string ISRESUMEPATTERN = "ИТОГО[\\s]{1,}ПО[\\s]{1,}СМЕТЕ";
        const string ISEQUIPMENTPATTERN = "СТОИМОСТЬ[\\s]{1,}ОБОРУДОВАНИЯ";
        const string ISTRANSPORTPATTERN = "ТРАНСПОРТНЫЕ РАСХОДЫ";
        const string ISDEPOTPATTERN = "ЗАГОТОВИТЕЛЬНО-СКЛАДСКИЕ РАСХОДЫ";
        const string ISTOTALPATTERN = "ВСЕГО[\\s]{1,}ПО[\\s]{1,}СМЕТЕ";
        const string ISOVERHEADSPATTERN = "ВСЕГО НАКЛАДНЫЕ РАСХОДЫ";
        const string ISESTIMATEPROFITPATTERN = "ВСЕГО СМЕТНАЯ ПРИБЫЛЬ";
        
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
            Console.WriteLine("Strings count: /while  don't know/ " + tableForParse.GetElementCount);
            double equip = 0, depot = 0, transport = 0;
            bool isResumeFinding = false;
            for (int i = 1; i < tableForParse.GetElementCount; i++)
            {
                if ((tableForParse[i].ColumnIndex == NUMBERCOLUMN) &&
                    (IsInScheme(RemoveUnvisibleCharacters(tableForParse[i].Range.Text), ISNUMBERPATTERN)))
                {
                    Console.Write("Parse element #{0}... ", RemoveUnvisibleCharacters(tableForParse[i].Range.Text));
                    EstimateString parsingEsS = ParsingString(i);
                    if (parsingEsS != null)
                    {
                        es.Add(parsingEsS);
                        Console.WriteLine("Parsing complete");
                    }
                    else { Console.WriteLine("Parse ERROR!"); }
                }
                if (IsInScheme(tableForParse[i].Range.Text, ISRESUMEPATTERN))
                {
                    Console.WriteLine("I find resume");
                    EstimateString resumeString = ParsingResume(i);
                    if (resumeString != null)
                    {
                        es.AddResumeString(resumeString);
                        Console.WriteLine("Add resume");
                    }
                    isResumeFinding = true;
                }
                if (IsInScheme(tableForParse[i].Range.Text, ISTOTALPATTERN)) {Console.WriteLine("I find total"); es.TotalEstimateCost = ParsingCost(i);}
                if ((isResumeFinding)&&(IsInScheme(tableForParse[i].Range.Text, ISOVERHEADSPATTERN))) { Console.WriteLine("I find overhread"); es.Overheads = ParsingCost(i); }
                if ((isResumeFinding)&&(IsInScheme(tableForParse[i].Range.Text, ISESTIMATEPROFITPATTERN))) { Console.WriteLine("I find profit"); es.EstimateProfit = ParsingCost(i); }
                if (IsInScheme(tableForParse[i].Range.Text, ISEQUIPMENTPATTERN)) equip = ParsingCost(i);
                if (IsInScheme(tableForParse[i].Range.Text, ISDEPOTPATTERN)) depot = ParsingCost(i);
                if (IsInScheme(tableForParse[i].Range.Text, ISTRANSPORTPATTERN)) transport = ParsingCost(i);
            }
            es.AddEquipment(equip, transport, depot);
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
            try
            {
                if (NormalizeNumber(tableForParse[firstCell + 2].Range.Text) != "")
                    resultString.Volume = Convert.ToDouble(NormalizeNumber(tableForParse[firstCell + 2].Range.Text));
                if (NormalizeNumber(tableForParse[firstCell + 12].Range.Text) != "")
                    resultString.CurrentWorkers = Convert.ToDouble(NormalizeNumber(tableForParse[firstCell + 12].Range.Text));
                if (NormalizeNumber(tableForParse[firstCell + 13].Range.Text) != "")
                    resultString.CurrentMachine = Convert.ToDouble(NormalizeNumber(tableForParse[firstCell + 13].Range.Text));
                if (NormalizeNumber(tableForParse[firstCell + 20].Range.Text) != "")
                    resultString.CurrentMaterials = Convert.ToDouble(NormalizeNumber(tableForParse[firstCell + 20].Range.Text));
                if (NormalizeNumber(tableForParse[firstCell + 21].Range.Text) != "")
                    resultString.CurrentMachineWorkers = Convert.ToDouble(NormalizeNumber(tableForParse[firstCell + 21].Range.Text));
            }
            catch { return null; }
            return resultString;
        }

        EstimateString ParsingResume(int firstCell)
        {
            EstimateString resultString = new EstimateString(0);
            resultString.Name = tableForParse[firstCell].Range.Text;
            resultString.Volume = 0;
            try
            {
                if (NormalizeNumber(tableForParse[firstCell + 7].Range.Text) != "")
                    resultString.CurrentWorkers = Convert.ToDouble(NormalizeNumber(tableForParse[firstCell + 12].Range.Text));
                if (NormalizeNumber(tableForParse[firstCell + 8].Range.Text) != "")
                    resultString.CurrentMachine = Convert.ToDouble(NormalizeNumber(tableForParse[firstCell + 13].Range.Text));
                if (NormalizeNumber(tableForParse[firstCell + 14].Range.Text) != "")
                    resultString.CurrentMaterials = Convert.ToDouble(NormalizeNumber(tableForParse[firstCell + 20].Range.Text));
                if (NormalizeNumber(tableForParse[firstCell + 15].Range.Text) != "")
                    resultString.CurrentMachineWorkers = Convert.ToDouble(NormalizeNumber(tableForParse[firstCell + 21].Range.Text));
            }
            catch { return null; }
            return resultString;
        }

        double ParsingCost(int firstCell)
        {
            double resultcost = 0;
            try
            {
                if (NormalizeNumber(tableForParse[firstCell + 6].Range.Text) != "")
                    resultcost = Convert.ToDouble(NormalizeNumber(tableForParse[firstCell + 12].Range.Text));
            }
            catch { return 0; }
            return resultcost;
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
            string str = inputstring;
            for (int i = 0; i < inputstring.Length; i++)
            {
                if (Convert.ToInt32(inputstring[i]) < 32)
                    str = str.Replace(Convert.ToString(inputstring[i]), "");
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
