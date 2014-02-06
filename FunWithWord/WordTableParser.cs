//#define DEBUG
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;


namespace FunWithWord
{
    class WordTableParser
    {
        const string CONFIGPATH = "config.ini";
        const string TABLENUMBER = "TableNumber";
        const string ISNUMBERPATTERN = "StringNumberPattern";
        const string ISRESUMEPATTERN = "ResumeStringPattern";
        const string ISEQUIPMENTPATTERN = "EquipmentStringPattern";
        const string ISTRANSPORTPATTERN = "TransportCostPattern";
        const string ISDEPOTPATTERN = "DepotCostPattern";
        const string ISTOTALPATTERN = "TotalStringPattern";
        const string ISOVERHEADSPATTERN = "OverheadStringPattern";
        const string ISESTIMATEPROFITPATTERN = "EstimateProfitStringPattern";

        const string ELEMENTSSCHEME = "ElementsScheme";
        const string ESTIMATESTRINGSCHEME = "EstimateStringScheme";
        const string RESUMETRINGSCHEME = "ResumeStringScheme";
        const string COMMONSTRINCOSTSCHEME = "CommonStringCostScheme";

        const string NUMBER = "Number";
        const string NAME = "Name";
        const string VOLUME = "Volume";
        const string COST = "Cost";    
        const string PAY = "Pay";	
        const string MACHINE = "Machine";
	    const string MATERIAL = "Material";
        const string MACHINEPAY = "MachinePay";

        
        WordTable tableForParse;
        Dictionary<string, string> configDictionary;
        Dictionary<string, int> stringShiftDictionary;
        int stringShiftFirst, stringShiftLength;
        Dictionary<string, int> resumeShiftDictionary;
        int resumeShiftFirst, resumeShiftLength;
        Dictionary<string, int> costShiftDictionary;
        int costShiftFirst, costShiftLength;

        public WordTableParser()
        {
            tableForParse = new WordTable();
            ConfigParse config = new ConfigParse(CONFIGPATH);
            configDictionary = config.Parsing();
            ElementTemplate templ = new ElementTemplate(configDictionary[ELEMENTSSCHEME]);
            stringShiftFirst = 0; stringShiftLength = 0;
            stringShiftDictionary = templ.ValuesShift(configDictionary[ESTIMATESTRINGSCHEME], NUMBER, 
                new string[] { NAME, VOLUME, COST, PAY, MACHINE, MATERIAL, MACHINEPAY }, out stringShiftFirst, out stringShiftLength);
            resumeShiftFirst = 0; resumeShiftLength = 0;
            resumeShiftDictionary = templ.ValuesShift(configDictionary[RESUMETRINGSCHEME], NAME,
                new string[] { COST, PAY, MACHINE, MATERIAL, MACHINEPAY }, out resumeShiftFirst, out resumeShiftLength);
            costShiftFirst = 0; costShiftLength = 0;
            costShiftDictionary = templ.ValuesShift(configDictionary[COMMONSTRINCOSTSCHEME], NAME,
                new string[] { COST }, out costShiftFirst, out costShiftLength);
            templ.Dispose();
#if (DEBUG)
            #region Configuration files structure
            Console.WriteLine("##### Configuration files #####");
            foreach (KeyValuePair<string, string> pair in configDictionary)
            {
                Console.WriteLine("{0} = {1}", pair.Key, pair.Value);
            }
            Console.WriteLine("###### Template for {0} #####", configDictionary[ESTIMATESTRINGSCHEME]);
            foreach (KeyValuePair<string, int> pair in stringShiftDictionary)
            {
                Console.WriteLine("{0} shift is {1}", pair.Key, pair.Value);
            }
            Console.WriteLine("Column of main element is {0}", stringShiftFirst);
            Console.WriteLine("Length of block is {0}", stringShiftLength);
            Console.WriteLine("###### Template for {0} #####", configDictionary[RESUMETRINGSCHEME]);
            foreach (KeyValuePair<string, int> pair in resumeShiftDictionary)
            {
                Console.WriteLine("{0} shift is {1}", pair.Key, pair.Value);
            }
            Console.WriteLine("Column of main element is {0}", resumeShiftFirst);
            Console.WriteLine("Length of block is {0}", resumeShiftLength);
            Console.WriteLine("###### Template for {0} #####", configDictionary[COMMONSTRINCOSTSCHEME]);
            foreach (KeyValuePair<string, int> pair in costShiftDictionary)
            {
                Console.WriteLine("{0} shift is {1}", pair.Key, pair.Value);
            }
            Console.WriteLine("Column of main element is {0}", costShiftFirst);
            Console.WriteLine("Length of block is {0}", costShiftLength);
            Console.WriteLine("###### End configuration ########");
            #endregion
#endif
        }

        public Estimate Parsing(string filepath)
        {
            tableForParse.ConnectToDocment(filepath);
            tableForParse.ChooseTable(Convert.ToInt32(configDictionary[TABLENUMBER]));
            Estimate es = new Estimate(filepath.Substring(filepath.LastIndexOf("\\")+1));
            double equip = 0, depot = 0, transport = 0;
            bool isResumeFind = false;
            for (int i = 1; i < tableForParse.GetElementCount; i++)
            {
                if ((tableForParse[i].ColumnIndex != Convert.ToInt32(stringShiftFirst)) && (tableForParse[i].ColumnIndex != Convert.ToInt32(resumeShiftFirst)) &&
                    (tableForParse[i].ColumnIndex != Convert.ToInt32(costShiftFirst))) continue;
                if ((tableForParse[i].ColumnIndex == Convert.ToInt32(stringShiftFirst)) &&
                    (IsInScheme(RemoveUnvisibleCharacters(tableForParse[i].Range.Text), configDictionary[ISNUMBERPATTERN])) && !(isResumeFind))
                {
                    Console.Write("Parse element #{0}...\t", RemoveUnvisibleCharacters(tableForParse[i].Range.Text));
                    EstimateString parsingEsS = ParsingString(i);
                    if (parsingEsS != null)
                    {
                        es.Add(parsingEsS);
                        Console.WriteLine("Parsing complete");
                    }
                    else { Console.WriteLine("Parse ERROR!"); }
                    i += Convert.ToInt32(stringShiftLength);
                    continue;
                }
                else if (IsInScheme(tableForParse[i].Range.Text, configDictionary[ISRESUMEPATTERN]))
                {
                    Console.Write("Parse resume...\t");
                    EstimateString resumeString = ParsingResume(i);
                    if (resumeString != null)
                    {
                        es.AddResumeString(resumeString);
                        Console.WriteLine("Parsing complete");
                    }
                    else { Console.WriteLine("Parse ERROR!"); }
                    isResumeFind = true;
                    i += Convert.ToInt32(resumeShiftLength);
                    continue;
                }
                else if ((isResumeFind) && (IsInScheme(tableForParse[i].Range.Text, configDictionary[ISTOTALPATTERN]))) es.TotalEstimateCost = ParsingCost(i);
                else if ((isResumeFind)&&(IsInScheme(tableForParse[i].Range.Text, configDictionary[ISOVERHEADSPATTERN]))) es.Overheads = ParsingCost(i);
                else if ((isResumeFind)&&(IsInScheme(tableForParse[i].Range.Text, configDictionary[ISESTIMATEPROFITPATTERN]))) es.EstimateProfit = ParsingCost(i);
                else if (IsInScheme(tableForParse[i].Range.Text, configDictionary[ISEQUIPMENTPATTERN])) equip = ParsingCost(i);
                else if (IsInScheme(tableForParse[i].Range.Text, configDictionary[ISDEPOTPATTERN])) depot = ParsingCost(i);
                else if (IsInScheme(tableForParse[i].Range.Text, configDictionary[ISTRANSPORTPATTERN])) transport = ParsingCost(i);
            }
            es.AddEquipment(equip, transport, depot);
            tableForParse.DisconnectFromDocument();
            return es;
        }

        EstimateString ParsingString(int firstCell)
        {
            string numberpattern = "^[0-9]{1,}";
            int number = Convert.ToInt32(ShemePart(RemoveUnvisibleCharacters(tableForParse[firstCell].Range.Text), numberpattern));
            EstimateString resultString = new EstimateString(number);
            string namecaption = tableForParse[firstCell + stringShiftDictionary[NAME]].Range.Text;
            int divider = namecaption.IndexOf(Convert.ToChar(13));
            resultString.Name = namecaption.Substring(0, divider - 1);
            resultString.Caption = namecaption.Substring(divider + 1, namecaption.Length - divider-1);
            try
            {
                if (NormalizeNumber(tableForParse[firstCell + stringShiftDictionary[VOLUME]].Range.Text) != "")
                    resultString.Volume = Convert.ToDouble(NormalizeNumber(tableForParse[firstCell + stringShiftDictionary[VOLUME]].Range.Text));
                if (NormalizeNumber(tableForParse[firstCell + stringShiftDictionary[COST]].Range.Text) != "")
                    resultString.CurrentCost = Convert.ToDouble(NormalizeNumber(tableForParse[firstCell + stringShiftDictionary[COST]].Range.Text));
                if (NormalizeNumber(tableForParse[firstCell + stringShiftDictionary[PAY]].Range.Text) != "")
                    resultString.CurrentWorkers = Convert.ToDouble(NormalizeNumber(tableForParse[firstCell + stringShiftDictionary[PAY]].Range.Text));
                if (NormalizeNumber(tableForParse[firstCell + stringShiftDictionary[MACHINE]].Range.Text) != "")
                    resultString.CurrentMachine = Convert.ToDouble(NormalizeNumber(tableForParse[firstCell + stringShiftDictionary[MACHINE]].Range.Text));
                if (NormalizeNumber(tableForParse[firstCell + stringShiftDictionary[MATERIAL]].Range.Text) != "")
                    resultString.CurrentMaterials = Convert.ToDouble(NormalizeNumber(tableForParse[firstCell + stringShiftDictionary[MATERIAL]].Range.Text));
                if (NormalizeNumber(tableForParse[firstCell + stringShiftDictionary[MACHINEPAY]].Range.Text) != "")
                    resultString.CurrentMachineWorkers = Convert.ToDouble(NormalizeNumber(tableForParse[firstCell + stringShiftDictionary[MACHINEPAY]].Range.Text));
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
                if (NormalizeNumber(tableForParse[firstCell + resumeShiftDictionary[COST]].Range.Text) != "")
                    resultString.CurrentCost = Convert.ToDouble(NormalizeNumber(tableForParse[firstCell + resumeShiftDictionary[COST]].Range.Text));
                if (NormalizeNumber(tableForParse[firstCell + resumeShiftDictionary[PAY]].Range.Text) != "")
                    resultString.CurrentWorkers = Convert.ToDouble(NormalizeNumber(tableForParse[firstCell + resumeShiftDictionary[PAY]].Range.Text));
                if (NormalizeNumber(tableForParse[firstCell + resumeShiftDictionary[MACHINE]].Range.Text) != "")
                    resultString.CurrentMachine = Convert.ToDouble(NormalizeNumber(tableForParse[firstCell + resumeShiftDictionary[MACHINE]].Range.Text));
                if (NormalizeNumber(tableForParse[firstCell + resumeShiftDictionary[MATERIAL]].Range.Text) != "")
                    resultString.CurrentMaterials = Convert.ToDouble(NormalizeNumber(tableForParse[firstCell + resumeShiftDictionary[MATERIAL]].Range.Text));
                if (NormalizeNumber(tableForParse[firstCell + resumeShiftDictionary[MACHINEPAY]].Range.Text) != "")
                    resultString.CurrentMachineWorkers = Convert.ToDouble(NormalizeNumber(tableForParse[firstCell + resumeShiftDictionary[MACHINEPAY]].Range.Text));
            }
            catch { return null; }
#if (DEBUG)
                #region resume parsing
                if  (resultString.CurrentCost != resultString.CurrentWorkers + resultString.CurrentMachine + resultString.CurrentMaterials)
                    Console.WriteLine("Cost not equals sum of elements");
                else
                Console.WriteLine("Resume string parse good");
                #endregion
#endif
            return resultString;
        }

        double ParsingCost(int firstCell)
        {
            double resultcost = 0;
            try
            {
                if (NormalizeNumber(tableForParse[firstCell + costShiftDictionary[COST]].Range.Text) != "")
                    resultcost = Convert.ToDouble(NormalizeNumber(tableForParse[firstCell + costShiftDictionary[COST]].Range.Text));
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
