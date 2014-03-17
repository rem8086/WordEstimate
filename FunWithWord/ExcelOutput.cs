using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;

namespace FunWithWord
{
    class ExcelOutput       //class for work with output excel file
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
            TitlesFill();
        }

        public void Close(string path)
        {
            currentWorkbook.SaveAs(path+"\\output"+DateTime.Now.ToString().Replace(':','_')+".xls");
            currentWorkbook.Close(SaveChanges: true);
        }

        public void FillWith(Estimate es) //main procedure of estimate data filling 
        {
            inputEstimate = es;
            MainDataFill();
            MaterialStringsFill();
            ResumesFill();
        }

        void TitlesFill()       //titles and formating
        {
            wsEstimateStrings.Cells[wsEstimateStringsRowCount, 1].Value = "Estimate Name";
            wsEstimateStrings.Cells[wsEstimateStringsRowCount, 2].Value = "Number";
            wsEstimateStrings.Cells[wsEstimateStringsRowCount, 3].Value = "Name";
            wsEstimateStrings.Cells[wsEstimateStringsRowCount, 4].Value = "Caption";
            wsEstimateStrings.Cells[wsEstimateStringsRowCount, 5].Value = "Volume";
            wsEstimateStrings.Cells[wsEstimateStringsRowCount, 6].Value = "Pay";
            wsEstimateStrings.Cells[wsEstimateStringsRowCount, 7].Value = "Machine";
            wsEstimateStrings.Cells[wsEstimateStringsRowCount, 8].Value = "PayMachine";
            wsEstimateStrings.Cells[wsEstimateStringsRowCount, 9].Value = "Materials";
            wsEstimateStrings.Cells[wsEstimateStringsRowCount, 10].Value = "Cost";
            wsEstimateStrings.Rows[wsEstimateStringsRowCount].Font.Bold = true;
            wsEstimateStringsRowCount++;
            wsMaterialStrings.Cells[wsMaterialStringsRowCount, 1].Value = "Estimate Name";
            wsMaterialStrings.Cells[wsMaterialStringsRowCount, 2].Value = "Number";
            wsMaterialStrings.Cells[wsMaterialStringsRowCount, 3].Value = "Name";
            wsMaterialStrings.Cells[wsMaterialStringsRowCount, 4].Value = "Caption";
            wsMaterialStrings.Cells[wsMaterialStringsRowCount, 5].Value = "Volume";
            wsMaterialStrings.Cells[wsMaterialStringsRowCount, 9].Value = "Materials";
            wsMaterialStrings.Cells[wsMaterialStringsRowCount, 10].Value = "Cost";
            wsMaterialStrings.Rows[wsMaterialStringsRowCount].Font.Bold = true;
            wsMaterialStringsRowCount++;
            wsResumes.Cells[wsResumesRowCount, 1].Value = "Estimate Name";
            wsResumes.Cells[wsResumesRowCount, 2].Value = "Pay";
            wsResumes.Cells[wsResumesRowCount, 3].Value = "Machine";
            wsResumes.Cells[wsResumesRowCount, 4].Value = "PayMachine";
            wsResumes.Cells[wsResumesRowCount, 5].Value = "Materials";
            wsResumes.Cells[wsResumesRowCount, 6].Value = "Cost";
            wsResumes.Cells[wsResumesRowCount, 7].Value = "Equipment";
            wsResumes.Cells[wsResumesRowCount, 8].Value = "Depot";
            wsResumes.Cells[wsResumesRowCount, 9].Value = "Transport";
            wsResumes.Cells[wsResumesRowCount, 10].Value = "Overhead";
            wsResumes.Cells[wsResumesRowCount, 11].Value = "Profit";
            wsResumes.Cells[wsResumesRowCount, 12].Value = "Total";
            wsResumes.Rows[wsResumesRowCount].Font.Bold = true;
            wsResumesRowCount++;
        }

        void MainDataFill()         //page with all strings
        {
            for (int i = 0; i < inputEstimate.StringCount; i++)
            {
                string[] data = inputEstimate[i].ToStringArray();
                wsEstimateStrings.Cells[wsEstimateStringsRowCount, 1].Value = inputEstimate.Name;
                for (int j = 0; j < data.Length; j++)
                {
                    wsEstimateStrings.Cells[wsEstimateStringsRowCount, j + 2].NumberFormat = "@";
                    wsEstimateStrings.Cells[wsEstimateStringsRowCount, j + 2].Value = data[j];
                }
                wsEstimateStringsRowCount++;
            }
        }

        void MaterialStringsFill()  //page with only material strings
        {
            foreach (EstimateString ess in inputEstimate.EstimateMaterials())
            {
                string[] data = ess.ToStringArray();
                wsMaterialStrings.Cells[wsMaterialStringsRowCount, 1].Value = inputEstimate.Name;
                for (int j = 0; j < data.Length; j++)
                {
                    wsMaterialStrings.Cells[wsMaterialStringsRowCount, j + 2].NumberFormat = "@";
                    wsMaterialStrings.Cells[wsMaterialStringsRowCount, j + 2].Value = data[j];
                }
                wsMaterialStringsRowCount++;
            }
        }

        void ResumesFill()          //page with resumes
        {
            for (int i = 1; i < 13; i++)
            {
                wsResumes.Cells[wsResumesRowCount, i].NumberFormat = "@";
            }
            wsResumes.Cells[wsResumesRowCount, 1].Value = inputEstimate.Name;
            wsResumes.Cells[wsResumesRowCount, 2].Value = inputEstimate.ResumeString.CurrentWorkers;
            wsResumes.Cells[wsResumesRowCount, 3].Value = inputEstimate.ResumeString.CurrentMachine;
            wsResumes.Cells[wsResumesRowCount, 4].Value = inputEstimate.ResumeString.CurrentMachineWorkers;
            wsResumes.Cells[wsResumesRowCount, 5].Value = inputEstimate.ResumeString.CurrentMaterials;
            wsResumes.Cells[wsResumesRowCount, 6].Value = inputEstimate.ResumeString.CurrentCost;
            wsResumes.Cells[wsResumesRowCount, 7].Value = inputEstimate.Equipment.EquipmentCost;
            wsResumes.Cells[wsResumesRowCount, 8].Value = inputEstimate.Equipment.DepotCost;
            wsResumes.Cells[wsResumesRowCount, 9].Value = inputEstimate.Equipment.TransportCost;
            wsResumes.Cells[wsResumesRowCount, 10].Value = inputEstimate.Overheads;
            wsResumes.Cells[wsResumesRowCount, 11].Value = inputEstimate.EstimateProfit;
            wsResumes.Cells[wsResumesRowCount, 12].Value = inputEstimate.TotalEstimateCost;
            wsResumesRowCount++;
        }          
    }
}
