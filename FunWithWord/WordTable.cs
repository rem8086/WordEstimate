using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;
using Microsoft.Office.Interop.Word;

namespace FunWithWord
{
    class WordTable
    {
        int wordApplicationProcessId;
        Document wordDocument;
        Table wordTable;

        public Table WTable
        {
            get { return wordTable; }
        }

        public WordTable()
        {
            wordApplicationProcessId = 0;
            Document wordDocument = new Document();
            //wordTable = new Table();
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
            get { return wordTable.Range.Cells[index]; }
        }

        public void ChooseTable(int tableNumber)
        {
            wordTable = wordDocument.Tables[tableNumber];
        }

        public Cell GetElement(int index)
        {
            return wordTable.Range.Cells[index];
        }

        public int GetElementCount
        {
            get { return wordTable.Range.Cells.Count; }
        }

        public void DisconnectFromDocument()
        {
            ((_Document)wordDocument).Close();
            if (wordApplicationProcessId != 0) Process.GetProcessById(wordApplicationProcessId).Kill();
        }
    }
}
