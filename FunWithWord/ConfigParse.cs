using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace FunWithWord
{
    class ConfigParse
    {
        StreamReader configFile;
        Dictionary<string, string> configDictionary;

        public ConfigParse(string path)
        {
            try
            {
                configFile = new StreamReader(path);
                configDictionary = new Dictionary<string, string>();
                DefaultDictionaryCreate();
            }
            catch (FileNotFoundException e)
            { 
               Console.WriteLine(e.Message);
            }
        }

        void DefaultDictionaryCreate()
        {
            configDictionary.Add("TableNumber", "3");
            configDictionary.Add("ElementsScheme", "PartsOfEstimate_Scheme.xls");
            configDictionary.Add("EstimateStringScheme", "String");
            configDictionary.Add("ResumeStringScheme", "Resume");
            configDictionary.Add("CommonStringCostScheme", "Cost");
            configDictionary.Add("StringNumberPattern", "^[0-9]{1,}\\.$");
            //configDictionary.Add("NumberColumn", "1");
            //configDictionary.Add("StringLength", "21");
            configDictionary.Add("ResumeStringPattern", "ИТОГО[\\s]{1,}ПО[\\s]{1,}СМЕТЕ");
            configDictionary.Add("EquipmentStringPattern", "СТОИМОСТЬ[\\s]{1,}ОБОРУДОВАНИЯ");
            configDictionary.Add("TransportCostPattern", "ТРАНСПОРТНЫЕ[\\s]{1,}РАСХОДЫ");
            configDictionary.Add("DepotCostPattern", "ЗАГОТОВИТЕЛЬНО-СКЛАДСКИЕ[\\s]{1,}РАСХОДЫ");
            configDictionary.Add("TotalStringPattern", "ВСЕГО[\\s]{1,}ПО[\\s]{1,}СМЕТЕ");
            configDictionary.Add("OverheadStringPattern", "ВСЕГО НАКЛАДНЫЕ РАСХОДЫ");
            configDictionary.Add("EstimateProfitStringPattern", "ВСЕГО СМЕТНАЯ ПРИБЫЛЬ");
            //configDictionary.Add("NameColumn", "2");
        }

        public Dictionary<string, string> Parsing()
        {
            while (!configFile.EndOfStream)
            {
                string currentstring = configFile.ReadLine();
                if ((currentstring.Length > 0) && (currentstring[0] != '#'))
                {
                    int spaceindex = currentstring.IndexOf(' ');
                    if (spaceindex > 0)
                    {
                        string currentKey = currentstring.Substring(0, spaceindex);
                        string currentValue = currentstring.Substring(spaceindex+1);
                        if (configDictionary.ContainsKey(currentKey))
                        {
                            configDictionary[currentKey] = currentValue;
                        }
                    }
                }
            }
            configFile.Close();
            return configDictionary;
        }
    }
}
