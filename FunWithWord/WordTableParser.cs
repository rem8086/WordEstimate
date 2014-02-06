//#define DEBUG
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Word;
using System.Threading;


namespace FunWithWord
{
    class WordTableParser
    {
        const string CONFIGPATH = "config.ini";         // program constants - name of config file, and names of config file elements
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
        Dictionary<string, string> configDictionary;    // dictionary with config elements with values from config file
                                                            //three dictionaries with relative positions of needed elements (like volume, cost, machine cost, material cost)
        Dictionary<string, int> stringShiftDictionary;      //for every estimate string (each elementary work in estimate)
        int stringShiftFirst, stringShiftLength;
        Dictionary<string, int> resumeShiftDictionary;      //for total estimate resume 
        int resumeShiftFirst, resumeShiftLength;
        Dictionary<string, int> costShiftDictionary;        //for the other elements like equipment, estimate profit etc. (only cost pulling)
        int costShiftFirst, costShiftLength;

        public WordTableParser()            //fillind dictionaries with values from config file, end Excel templates file
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

        public Estimate Parsing(string filepath)    //main function - pull estimate data from Word file
        {
            tableForParse.ConnectToDocment(filepath);
            tableForParse.ChooseTable(Convert.ToInt32(configDictionary[TABLENUMBER]));
            Estimate es = new Estimate(filepath.Substring(filepath.LastIndexOf("\\")+1));
            double equip = 0, depot = 0, transport = 0;
            bool isResumeFind = false;
            foreach (Cell currentCell in tableForParse.SelectedTable.Cells)         //check every cell in table and by regexp pattern try to find interesting element
            {
                if ((currentCell.ColumnIndex != Convert.ToInt32(stringShiftFirst)) && (currentCell.ColumnIndex != Convert.ToInt32(resumeShiftFirst)) &&
                    (currentCell.ColumnIndex != Convert.ToInt32(costShiftFirst))) continue;
                string currentCellText = RemoveUnvisibleCharacters(currentCell.Range.Text);
                if ((currentCell.ColumnIndex == Convert.ToInt32(stringShiftFirst)) &&
                    (Regex.IsMatch(currentCellText, configDictionary[ISNUMBERPATTERN])) && !(isResumeFind)) //like number of estimate string
                {
                    Console.Write("Parse element #{0}...\t", currentCellText);
                    EstimateString parsingEsS = ParsingString(currentCell);                                 //then parse this and next cells and add result into Estimate
                    if (parsingEsS != null)
                    {
                        es.Add(parsingEsS);
                        Console.WriteLine("Parsing complete");
                    }
                    else { Console.WriteLine("Parse ERROR!"); }
                    //i += Convert.ToInt32(stringShiftLength);
                    continue;
                }
                else if (Regex.IsMatch(currentCellText, configDictionary[ISRESUMEPATTERN]))                 //like total result string
                {
                    Console.Write("Parse resume...\t");
                    EstimateString resumeString = ParsingResume(currentCell);                               //parse this too
                    if (resumeString != null)
                    {
                        es.AddResumeString(resumeString);
                        Console.WriteLine("Parsing complete");
                    }
                    else { Console.WriteLine("Parse ERROR!"); }
                    isResumeFind = true;
                    //i += Convert.ToInt32(resumeShiftLength);
                    continue;
                }                                                                                           //like the other strings, necessary to us
                else if ((isResumeFind) && (Regex.IsMatch(currentCellText, configDictionary[ISTOTALPATTERN]))) es.TotalEstimateCost = ParsingCost(currentCell);
                else if ((isResumeFind) && (Regex.IsMatch(currentCellText, configDictionary[ISOVERHEADSPATTERN]))) es.Overheads = ParsingCost(currentCell);
                else if ((isResumeFind) && (Regex.IsMatch(currentCellText, configDictionary[ISESTIMATEPROFITPATTERN]))) es.EstimateProfit = ParsingCost(currentCell);
                else if (Regex.IsMatch(currentCellText, configDictionary[ISEQUIPMENTPATTERN])) equip = ParsingCost(currentCell);
                else if (Regex.IsMatch(currentCellText, configDictionary[ISDEPOTPATTERN])) depot = ParsingCost(currentCell);
                else if (Regex.IsMatch(currentCellText, configDictionary[ISTRANSPORTPATTERN])) transport = ParsingCost(currentCell);
                 
            }
            es.AddEquipment(equip, transport, depot);
            tableForParse.DisconnectFromDocument();
            return es;
        }

        EstimateString ParsingString(Cell firstCell)    //function for parsing part of table and return string of estimate
        {
            string numberpattern = "^[0-9]{1,}";
            int number = Convert.ToInt32(Regex.Match(RemoveUnvisibleCharacters(firstCell.Range.Text), numberpattern).Value);
            EstimateString resultString = new EstimateString(number);
            string namecaption = CellShift(firstCell, stringShiftDictionary[NAME]).Range.Text;  //found name and caption by positions of cells versus cell with number
            int divider = namecaption.IndexOf(Convert.ToChar(13));
            resultString.Name = namecaption.Substring(0, divider - 1);
            resultString.Caption = namecaption.Substring(divider + 1, namecaption.Length - divider-1);
            try
            {                                                                                   //as well found another data
                if (NormalizeNumber(CellShift(firstCell, stringShiftDictionary[VOLUME]).Range.Text) != "")
                    resultString.Volume = Convert.ToDouble(NormalizeNumber(CellShift(firstCell, stringShiftDictionary[VOLUME]).Range.Text));
                if (NormalizeNumber(CellShift(firstCell, stringShiftDictionary[COST]).Range.Text) != "")
                    resultString.CurrentCost = Convert.ToDouble(NormalizeNumber(CellShift(firstCell, stringShiftDictionary[COST]).Range.Text));
                if (NormalizeNumber(CellShift(firstCell, stringShiftDictionary[PAY]).Range.Text) != "")
                    resultString.CurrentWorkers = Convert.ToDouble(NormalizeNumber(CellShift(firstCell, stringShiftDictionary[PAY]).Range.Text));
                if (NormalizeNumber(CellShift(firstCell, stringShiftDictionary[MACHINE]).Range.Text) != "")
                    resultString.CurrentMachine = Convert.ToDouble(NormalizeNumber(CellShift(firstCell, stringShiftDictionary[MACHINE]).Range.Text));
                if (NormalizeNumber(CellShift(firstCell, stringShiftDictionary[MATERIAL]).Range.Text) != "")
                    resultString.CurrentMaterials = Convert.ToDouble(NormalizeNumber(CellShift(firstCell, stringShiftDictionary[MATERIAL]).Range.Text));
                if (NormalizeNumber(CellShift(firstCell, stringShiftDictionary[MACHINEPAY]).Range.Text) != "")
                    resultString.CurrentMachineWorkers = Convert.ToDouble(NormalizeNumber(CellShift(firstCell, stringShiftDictionary[MACHINEPAY]).Range.Text));
            }
            catch { return null; }
            return resultString;
        }

        EstimateString ParsingResume(Cell firstCell)    //like ParsingsString function return resume estimate string 
        {
            EstimateString resultString = new EstimateString(0);
            resultString.Name = firstCell.Range.Text;
            resultString.Volume = 0;
            try
            {
                if (NormalizeNumber(CellShift(firstCell, resumeShiftDictionary[COST]).Range.Text) != "")
                    resultString.CurrentCost = Convert.ToDouble(NormalizeNumber(CellShift(firstCell, resumeShiftDictionary[COST]).Range.Text));
                if (NormalizeNumber(CellShift(firstCell, resumeShiftDictionary[PAY]).Range.Text) != "")
                    resultString.CurrentWorkers = Convert.ToDouble(NormalizeNumber(CellShift(firstCell, resumeShiftDictionary[PAY]).Range.Text));
                if (NormalizeNumber(CellShift(firstCell, resumeShiftDictionary[MACHINE]).Range.Text) != "")
                    resultString.CurrentMachine = Convert.ToDouble(NormalizeNumber(CellShift(firstCell, resumeShiftDictionary[MACHINE]).Range.Text));
                if (NormalizeNumber(CellShift(firstCell, resumeShiftDictionary[MATERIAL]).Range.Text) != "")
                    resultString.CurrentMaterials = Convert.ToDouble(NormalizeNumber(CellShift(firstCell, resumeShiftDictionary[MATERIAL]).Range.Text));
                if (NormalizeNumber(CellShift(firstCell, resumeShiftDictionary[MACHINEPAY]).Range.Text) != "")
                    resultString.CurrentMachineWorkers = Convert.ToDouble(NormalizeNumber(CellShift(firstCell, resumeShiftDictionary[MACHINEPAY]).Range.Text));
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

        double ParsingCost(Cell firstCell)      //and another parsing function for the other elemtns (only cost getting)
        {
            double resultcost = 0;
            try
            {
                if (NormalizeNumber(CellShift(firstCell, costShiftDictionary[COST]).Range.Text) != "")
                    resultcost = Convert.ToDouble(NormalizeNumber(CellShift(firstCell, costShiftDictionary[COST]).Range.Text));
            }
            catch { return 0; }
            return resultcost;
        }

        Cell CellShift(Cell inputCell, int shift) //return cell of table, shifted versus inputCell on needed count of cells
        {
            if (shift < 1) return null;
            Cell shiftCell = inputCell;
            int i = 0;
            do
            {
                shiftCell = shiftCell.Next;
                i++;
            } while (i < shift);
            return shiftCell;
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

    }
}
