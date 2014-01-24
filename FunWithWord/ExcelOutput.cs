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
        int wsEstimateStringsRowCount;
        Worksheet wsMaterialStrings;
        int wsMaterialStringsRowCount;
        Worksheet wsResumes;
        int wsResumesRowCount;

        public ExcelOutput()
        {
            Application ap = new Application();
            ap.Visible = true;
            currentWorkbook = ap.Workbooks.Add();
            wsEstimateStringsRowCount = 1;
            wsEstimateStrings = (Worksheet)currentWorkbook.Worksheets.Add();
            wsEstimateStrings.Name = "FullEstimateData";
            wsMaterialStringsRowCount = 1;
            wsMaterialStrings = (Worksheet)currentWorkbook.Worksheets.Add();
            wsMaterialStrings.Name = "MaterialsData";
            wsResumesRowCount = 1;
            wsResumes = (Worksheet)currentWorkbook.Worksheets.Add();
            wsResumes.Name = "EstimateResumes";
        }

        public void Close(string path)
        {
            currentWorkbook.SaveAs(path+"\\output"+DateTime.Now.ToString().Replace(':','_')+".xls");
            currentWorkbook.Close(SaveChanges: true);
        }

        public void FillWith(Estimate es)
        {
            inputEstimate = es;
            MainDataFill();
            MaterialStringsFill();
            ResumesFill();
        }

        void MainDataFill()
        {
            for (int i = 0; i < inputEstimate.StringCount; i++)
            {
                string[] data = inputEstimate[i].ToStringArray();
                Range c = (Range)wsEstimateStrings.Cells[wsEstimateStringsRowCount, 1];
                c.Value = inputEstimate.Name;
                for (int j = 0; j < data.Length; j++)
                {
                    Range ran = (Range)wsEstimateStrings.Cells[wsEstimateStringsRowCount, j + 2];
                    ran.Value = data[j];
                }
                wsEstimateStringsRowCount++;
            }
        }

        void MaterialStringsFill()
        {
            foreach (EstimateString ess in inputEstimate.EstimateMaterials())
            {
                string[] data = ess.ToStringArray();
                Range c = (Range)wsMaterialStrings.Cells[wsMaterialStringsRowCount, 1];
                c.Value = ess.Name;
                for (int j = 0; j < data.Length; j++)
                {
                    Range ran = (Range)wsMaterialStrings.Cells[wsMaterialStringsRowCount, j + 2];
                    ran.Value = data[j];
                }
                wsMaterialStringsRowCount++;
            }
        }

        void ResumesFill()
        {
            Range c1 = (Range)wsResumes.Cells[wsResumesRowCount, 1];
                c1.Value = inputEstimate.ResumeString.CurrentWorkers;
            Range c2 = (Range)wsResumes.Cells[wsResumesRowCount, 2];
                c2.Value = inputEstimate.ResumeString.CurrentMachine;
            Range c3 = (Range)wsResumes.Cells[wsResumesRowCount, 3];
                c3.Value = inputEstimate.ResumeString.CurrentMaterials;
            Range c4 = (Range)wsResumes.Cells[wsResumesRowCount, 4];
                c4.Value = inputEstimate.ResumeString.CurrentCost;
            Range c5 = (Range)wsResumes.Cells[wsResumesRowCount, 5];
                c5.Value = inputEstimate.Equipment.EquipmentCost;
            Range c6 = (Range)wsResumes.Cells[wsResumesRowCount, 6];
                c6.Value = inputEstimate.Equipment.DepotCost;
            Range c7 = (Range)wsResumes.Cells[wsResumesRowCount, 7];
                c7.Value = inputEstimate.Equipment.TransportCost;
            Range c8 = (Range)wsResumes.Cells[wsResumesRowCount, 8];
                c8.Value = inputEstimate.Overheads;
            Range c9 = (Range)wsResumes.Cells[wsResumesRowCount, 9];
                c9.Value = inputEstimate.EstimateProfit;
            Range c10 = (Range)wsResumes.Cells[wsResumesRowCount, 10];
                c10.Value = inputEstimate.TotalEstimateCost;
            wsResumesRowCount++;
        }
    }
}
