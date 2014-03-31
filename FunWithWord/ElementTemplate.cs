using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;

namespace FunWithWord
{
    class ElementTemplate : IDisposable //class for work with file with estimate elements templates
    {                                   //templates are used for finding needed cells into estimate table

        Workbook templatesWorkbook;

        public ElementTemplate(string filename)
        {
            Application ap = new Application();
            templatesWorkbook = ap.Workbooks.Open(Environment.CurrentDirectory + "\\" + filename, ReadOnly: true);
        }
                                
                                    //function for searching elements positions about root element in template document
        public Dictionary<string, int> ValuesShift(string sheetName,        //name of sheet in document for searching
                                                    string rootElName,      //name of root element about which search is making
                                                    string[] shiftElNames,  //array of elements, which relative positions we need to find
                                                    out int rootElColumn)   //returned column of root element for relief of searching this in whole estimate
        { 
            Worksheet stringSheet = templatesWorkbook.Worksheets[sheetName];            
            Range mainRange = stringSheet.get_Range(sheetName);                 //name of range with elements equals name of worksheet
            Dictionary<string, int> shiftDictionary = new Dictionary<string,int>();
            int stringElCount = 0;
            rootElColumn = 0;
            int rootElPosition = 0;
            for (int i = 1; i <= mainRange.Rows.Count; i++)
			{
			    for (int j = 1; j <= mainRange.Columns.Count; j++)          //go throw the range and calc cells count
			    {
                    int r = mainRange.Row + i - 1;
                    int c = mainRange.Column + j - 1;
                    if (IsFirstCellInMerge(mainRange, i, j))
                    {
                        stringElCount++;
                    }
                    if (mainRange.Cells[i, j].Value == rootElName)
                    {
                        rootElPosition = stringElCount;
                        rootElColumn = j;
                    }
                    foreach (string str in shiftElNames)        //all founded elements add into dictionary with number of this element cell
	                {
                         if (mainRange.Cells[i, j].Value == str)
                             shiftDictionary.Add(str, stringElCount);
	                }
			    }
			}
            stringElCount--;
            if (rootElPosition == 0) throw new Exception(String.Format("Root element {0} not found", rootElName)); 
            for (int i = 0; i < shiftElNames.Length; i++)
            {
                if (shiftDictionary.ContainsKey(shiftElNames[i]))
                    shiftDictionary[shiftElNames[i]] -= rootElPosition;         //transform numbers versus number of root element
            }
            return shiftDictionary;
        }

        bool IsFirstCellInMerge(Range r, int row, int column)       //true - if first cell in merged area
        {
            if (!r.Cells[row, column].MergeCells) return true;
            if ((r.Cells[row,column].MergeArea.Row - r.Row + 1 < row)||
                (r.Cells[row,column].MergeArea.Column - r.Column + 1 < column)) return false;
            return true;
        }

        public void Dispose()
        {
            if (templatesWorkbook != null) templatesWorkbook.Close();
        }
    }
}
