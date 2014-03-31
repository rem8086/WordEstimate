using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;
using Microsoft.Office.Interop.Word;

namespace FunWithWord
{
    class WordTable     //serving class for work with Word document
    {
        int wordApplicationProcessId;
        Document wordDocument;
        Selection selectedTable;

        public WordTable()
        {
            wordApplicationProcessId = 0;
            Document wordDocument = new Document();
        }

        public void ConnectToDocment(string pathToDoc)
        {
            List<int> processIDList = new List<int>();
            foreach (Process p in Process.GetProcessesByName("WINWORD"))
            {
                processIDList.Add(p.Id);
            }
            Application ap = new Application();
            wordDocument = ap.Documents.Open(pathToDoc, ReadOnly: true, Visible: true);
            foreach (Process p in Process.GetProcessesByName("WINWORD"))
            {
                if (!processIDList.Contains(p.Id)) wordApplicationProcessId = p.Id;
            }
            
        }

        public Cell this[int index]
        {
            get { return SelectedTable.Cells[index]; }
        }

        public void ChooseTable(int tableNumber)
        {
            wordDocument.Tables[tableNumber].Select();
            selectedTable = wordDocument.ActiveWindow.Panes[1].Selection;
        }

        public Selection SelectedTable
        {
            get { return selectedTable; }
        }

        public Cell GetElement(int index)
        {
            return selectedTable.Cells[index];
        }

        public int GetElementCount
        {
            get { return SelectedTable.Cells.Count; }
        }

        public void DisconnectFromDocument()
        {
            ((_Document)wordDocument).Close();
            if (wordApplicationProcessId != 0) Process.GetProcessById(wordApplicationProcessId).Kill();
        }
    }
}
