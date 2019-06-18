using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SHWordTable
{
    public class WordTable
    {

        private static string prp_SeparatorFor_Paragraph = "/pgraph/";

        private int prp_RowCount = 0;
        private int prp_ColumnCount = 0;
        private string prp_TableName = string.Empty;

        private List<WordTableRow> prp_TableRows = null;

        public int RowCount { get => prp_RowCount; }
        public int ColumnCount { get => prp_ColumnCount; }
        public List<WordTableRow> TableRows { get => prp_TableRows; }
        public string TableName { get => prp_TableName; set => prp_TableName = value; }
        public static string SeparatorFor_Paragraph { get => prp_SeparatorFor_Paragraph; }

        public WordTable()
        {
            prp_TableRows = new List<WordTableRow>();
        }
        public WordTable(WordTableRow[] WTRow)
        {
            prp_TableRows = new List<WordTableRow>();

            this.prp_TableRows.AddRange(WTRow);
            this.prp_RowCount += WTRow.Length;
        }

        public void AddRow(WordTableRow WTRow)
        {
            this.prp_TableRows.Add(WTRow);
            this.prp_RowCount++;
        }

        public WordTableCell GetCell(int WTRowIndex, int WTColumnIndex)
        {
            if(WTRowIndex < 0 | WTColumnIndex < 0)
            {
                return null;
            }

            if(WTRowIndex >= this.RowCount)
            {
                return null;
            }

            if(WTColumnIndex >= this.TableRows[WTRowIndex].TableCells.Count)
            {
                return null;
            }

            return this.TableRows[WTRowIndex].TableCells[WTColumnIndex];
        }
    }
}
