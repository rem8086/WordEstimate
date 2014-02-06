using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;

namespace FunWithWord
{
    class ElementTemplate : IDisposable
    {

        Workbook templatesWorkbook;

        public ElementTemplate(string filename)
        {
            Application ap = new Application();
            templatesWorkbook = ap.Workbooks.Open(Environment.CurrentDirectory + "\\" + filename, ReadOnly: true);
        }

        public Dictionary<string, int> ValuesShift(string sheetName, string rootElName, string[] shiftElNames, out int rootElColumn, out int stringElCount)
        {
            Worksheet stringSheet = templatesWorkbook.Worksheets[sheetName];            
            Range mainRange = stringSheet.get_Range(sheetName);
            //Range rootElementRange = mainRange.Find(rootElName, LookIn: XlFindLookIn.xlValues, LookAt: XlLookAt.xlWhole,
            //                                        SearchOrder: XlSearchOrder.xlByColumns);
            //if (rootElementRange == null) throw new Exception(String.Format("Root Element \"{0}\" not found", rootElName));
            Dictionary<string, int> shiftDictionary = new Dictionary<string,int>();
            stringElCount = 0;
            rootElColumn = 0;
            int rootElPosition = 0;
            for (int i = 1; i <= mainRange.Rows.Count; i++)
			{
			    for (int j = 1; j <= mainRange.Columns.Count; j++)
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
                    foreach (string str in shiftElNames)
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
                    shiftDictionary[shiftElNames[i]] -= rootElPosition;
            }
            return shiftDictionary;
        }

        bool IsFirstCellInMerge(Range r, int row, int column)
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
