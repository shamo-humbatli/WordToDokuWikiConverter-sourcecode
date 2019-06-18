using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SHWordTable
{
    public enum TableCellValueType
    {
        Null = 0,
        Empty = 1,
        Text = 2
    }

    public enum TableCellType
    {
        Normal = 0,
        H_Merged = 1,
        V_Merged = 2,
        HV_Merged = 3
    }

    public class WordTableCell
    {
        private TableCellType prp_WordTableCellType = TableCellType.Normal;
        private TableCellValueType prp_WordTableCellValueType = TableCellValueType.Text;
        private string prp_WordTableCellValue = string.Empty;

        public TableCellType WordTableCellType { get => prp_WordTableCellType; set => prp_WordTableCellType = value; }
        public TableCellValueType WordTableCellValueType { get => prp_WordTableCellValueType; set => prp_WordTableCellValueType = value; }
        public string WordTableCellValue { get => prp_WordTableCellValue; set => prp_WordTableCellValue = value; }
        

        public WordTableCell()
        {
            this.WordTableCellType = TableCellType.Normal;
            this.WordTableCellValue = null;
            this.WordTableCellValueType = TableCellValueType.Null;
        }

        public WordTableCell(string _CellValue, TableCellType _CellType = TableCellType.Normal, TableCellValueType _CellValueType = TableCellValueType.Text)
        {
            this.WordTableCellType = _CellType;
            this.WordTableCellValue = _CellValue;
            this.WordTableCellValueType = _CellValueType;
        }

        public void CreateCell(string _CellValue, TableCellType _CellType = TableCellType.Normal, TableCellValueType _CellValueType = TableCellValueType.Text)
        {
            this.WordTableCellType = _CellType;
            this.WordTableCellValue = _CellValue;
            this.WordTableCellValueType = _CellValueType;
        }
    }
}
