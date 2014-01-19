using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;

namespace FunWithWord
{
    class ExcelOutput
    {
        Estimate inputEstimate;
        Workbook currentWorkbook;
        Worksheet wsEstimateStrings;
        int fillingRowsCount;

        public ExcelOutput(Workbook wb)
        {
            currentWorkbook = wb;
            fillingRowsCount = 1;
            wsEstimateStrings = (Worksheet)currentWorkbook.Worksheets.Add();
            wsEstimateStrings.Name = "FullEstimateData";
        }

        public void Close(string path)
        {
            currentWorkbook.SaveAs(path+"\\output"+DateTime.Now.ToString().Replace(':','_')+".xls");
            currentWorkbook.Close(SaveChanges: true);
        }

        public void FillWith(Estimate es)
        {
            inputEstimate = es;
            for (int i = 0; i < inputEstimate.StringCount; i++)
            {
                string[] data = inputEstimate[i].ToStringArray();
                Range c = (Range)wsEstimateStrings.Cells[fillingRowsCount, 1];
                c.Value = inputEstimate.Name;
                for (int j = 0; j < data.Length; j++)
                {
                    Range ran = (Range) wsEstimateStrings.Cells[fillingRowsCount, j+2]; 
                    ran.Value = data[j];
                }
                fillingRowsCount++;
            }
        }
    }
}
