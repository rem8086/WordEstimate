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
            wsEstimateStrings.Cells[wsEstimateStringsRowCount, 1].Value = "File Name";
            wsEstimateStrings.Cells[wsEstimateStringsRowCount, 2].Value = "Estimate Code";
            wsEstimateStrings.Cells[wsEstimateStringsRowCount, 3].Value = "Estimate Name";
            wsEstimateStrings.Cells[wsEstimateStringsRowCount, 4].Value = "Number";
            wsEstimateStrings.Cells[wsEstimateStringsRowCount, 5].Value = "Name";
            wsEstimateStrings.Cells[wsEstimateStringsRowCount, 6].Value = "Caption";
            wsEstimateStrings.Cells[wsEstimateStringsRowCount, 7].Value = "Volume";
            wsEstimateStrings.Cells[wsEstimateStringsRowCount, 8].Value = "Pay";
            wsEstimateStrings.Cells[wsEstimateStringsRowCount, 9].Value = "Machine";
            wsEstimateStrings.Cells[wsEstimateStringsRowCount, 10].Value = "PayMachine";
            wsEstimateStrings.Cells[wsEstimateStringsRowCount, 11].Value = "Materials";
            wsEstimateStrings.Cells[wsEstimateStringsRowCount, 12].Value = "Cost";
            wsEstimateStrings.Cells[wsEstimateStringsRowCount, 13].Value = "Overheads";
            wsEstimateStrings.Cells[wsEstimateStringsRowCount, 14].Value = "Profit";
            wsEstimateStrings.Cells[wsEstimateStringsRowCount, 15].Value = "Total";
            wsEstimateStrings.Rows[wsEstimateStringsRowCount].Font.Bold = true;
            wsEstimateStringsRowCount++;
            wsMaterialStrings.Cells[wsMaterialStringsRowCount, 1].Value = "File Name";
            wsMaterialStrings.Cells[wsMaterialStringsRowCount, 2].Value = "Estimate Code";
            wsMaterialStrings.Cells[wsMaterialStringsRowCount, 3].Value = "Estimate Name";
            wsMaterialStrings.Cells[wsMaterialStringsRowCount, 4].Value = "Number";
            wsMaterialStrings.Cells[wsMaterialStringsRowCount, 5].Value = "Name";
            wsMaterialStrings.Cells[wsMaterialStringsRowCount, 6].Value = "Caption";
            wsMaterialStrings.Cells[wsMaterialStringsRowCount, 7].Value = "Volume";
            wsMaterialStrings.Cells[wsMaterialStringsRowCount, 11].Value = "Materials";
            wsMaterialStrings.Cells[wsMaterialStringsRowCount, 12].Value = "Cost";
            wsMaterialStrings.Rows[wsMaterialStringsRowCount].Font.Bold = true;
            wsMaterialStringsRowCount++;
            wsResumes.Cells[wsResumesRowCount, 1].Value = "File Name";
            wsResumes.Cells[wsResumesRowCount, 2].Value = "Estimate Code";
            wsResumes.Cells[wsResumesRowCount, 3].Value = "Estimate Name";
            wsResumes.Cells[wsResumesRowCount, 4].Value = "Pay";
            wsResumes.Cells[wsResumesRowCount, 5].Value = "Machine";
            wsResumes.Cells[wsResumesRowCount, 6].Value = "PayMachine";
            wsResumes.Cells[wsResumesRowCount, 7].Value = "Materials";
            wsResumes.Cells[wsResumesRowCount, 8].Value = "Cost";
            wsResumes.Cells[wsResumesRowCount, 9].Value = "Equipment";
            wsResumes.Cells[wsResumesRowCount, 10].Value = "Depot";
            wsResumes.Cells[wsResumesRowCount, 11].Value = "Transport";
            wsResumes.Cells[wsResumesRowCount, 12].Value = "Overhead";
            wsResumes.Cells[wsResumesRowCount, 13].Value = "Profit";
            wsResumes.Cells[wsResumesRowCount, 14].Value = "Total";
            wsResumes.Rows[wsResumesRowCount].Font.Bold = true;
            wsResumesRowCount++;
        }

        void MainDataFill()         //page with all strings
        {
            for (int i = 0; i < inputEstimate.StringCount; i++)
            {
                string[] data = inputEstimate[i].ToStringArray();
                wsEstimateStrings.Cells[wsEstimateStringsRowCount, 1].Value = inputEstimate.FileName;
                wsEstimateStrings.Cells[wsEstimateStringsRowCount, 2].Format = "@";
                wsEstimateStrings.Cells[wsEstimateStringsRowCount, 2].Value = inputEstimate.Code;
                wsEstimateStrings.Cells[wsEstimateStringsRowCount, 3].Value = inputEstimate.Name;
                for (int j = 0; j < data.Length; j++)
                {
                    wsEstimateStrings.Cells[wsEstimateStringsRowCount, j + 4].NumberFormat = "@";
                    wsEstimateStrings.Cells[wsEstimateStringsRowCount, j + 4].Value = data[j];
                }
                wsEstimateStringsRowCount++;
            }
        }

        void MaterialStringsFill()  //page with only material strings
        {
            foreach (EstimateString ess in inputEstimate.EstimateMaterials())
            {
                string[] data = ess.ToStringArray();
                wsMaterialStrings.Cells[wsMaterialStringsRowCount, 1].Value = inputEstimate.FileName;
                wsMaterialStrings.Cells[wsMaterialStringsRowCount, 2].Format = "@";
                wsMaterialStrings.Cells[wsMaterialStringsRowCount, 2].Value = inputEstimate.Code;
                wsMaterialStrings.Cells[wsMaterialStringsRowCount, 3].Value = inputEstimate.Name;
                for (int j = 0; j < data.Length; j++)
                {
                    wsMaterialStrings.Cells[wsMaterialStringsRowCount, j + 4].NumberFormat = "@";
                    wsMaterialStrings.Cells[wsMaterialStringsRowCount, j + 4].Value = data[j];
                }
                wsMaterialStringsRowCount++;
            }
        }

        void ResumesFill()          //page with resumes
        {
            string[] data = inputEstimate.ToStringArray();
            for (int i = 0; i < data.Length; i++)
            {
                wsResumes.Cells[wsResumesRowCount, i + 1].NumberFormat = "@";
                wsResumes.Cells[wsResumesRowCount, i + 1].Value = data[i];
            }
            wsResumesRowCount++;
        }          
    }
}
