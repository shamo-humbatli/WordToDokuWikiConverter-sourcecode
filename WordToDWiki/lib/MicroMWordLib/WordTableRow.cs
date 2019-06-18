using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SHWordTable
{
    public class WordTableRow
    {
        private List<WordTableCell> prp_TableCells = null;
        private int prp_Count = -1;


        public List<WordTableCell> TableCells { get => prp_TableCells; set => prp_TableCells = value; }
        public int Count { get => prp_Count; }

        public WordTableRow(WordTableCell[] WTCells)
        {
            if (TableCells == null)
            {
                TableCells = new List<WordTableCell>();
            }

            TableCells.AddRange(WTCells);
            prp_Count += WTCells.Length;
        }

        public WordTableRow()
        {
            if (TableCells == null)
            {
                TableCells = new List<WordTableCell>();
            }
        }

        public void AddCell(WordTableCell WTCell)
        {
            TableCells.Add(WTCell);
            prp_Count++;
        }

        public void AddCell(WordTableCell[] WTCell)
        {
            TableCells.AddRange(WTCell);
            prp_Count += WTCell.Length;
        }

    }
}
