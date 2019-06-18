using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace SHWordTable
{
    public class WordTableReader
    {
        public static WordTable GetTableFromTableXMLData(string XMLContent)
        {
            XmlDocument XMLDoc = new XmlDocument();
            XMLDoc.LoadXml(XMLContent);

            WordTable WTable = new WordTable();
            WTable.TableName = "Table_" + Guid.NewGuid().ToString();

            //XmlReader MyXMLRdr = XmlReader.Create(XMLContent);

            // Getting table rows
            XmlNodeList AllTableRows = XMLDoc.GetElementsByTagName(WordXMLTags.WordTagName_TableRow);

            List<bool> VMerge = new List<bool>();

            int RowCounter = 0;
            foreach (XmlNode trow in AllTableRows)
            {
                // Getting table cells
                XmlNodeList AllTableCells = trow.ChildNodes;
                WordTableRow WTRow = new WordTableRow();

                int CellCounter = 0;

                foreach (XmlNode tcell in AllTableCells)
                {
                    //XmlNodeList CellChilds = cxmln.ChildNodes; 
                    if (tcell.Name == WordXMLTags.WordTagName_TableCell)
                    {
                        int HMerge_CellCount = 0;
                        WordTableCell WTCell = new WordTableCell();

                        if (RowCounter == 0)
                        {
                            VMerge.Add(false);
                        }

                        XmlDocument Prgs = new XmlDocument();
                        Prgs.LoadXml(tcell.OuterXml);

                        XmlNodeList Paragraphs = Prgs.GetElementsByTagName(WordXMLTags.WordTagName_Paragraph);

                        //if (HMerge_CellCount > 1)
                        //{
                        //    WTCell.WordTableCellType = TableCellType.H_Merged;
                        //    HMerge_CellCount--;
                        //}
                        //else
                        //{
                        //    WTCell.WordTableCellType = TableCellType.Normal;
                        //}

                        if (VMerge[CellCounter] == true)
                        {
                            if (WTCell.WordTableCellType == TableCellType.H_Merged)
                            {
                                WTCell.WordTableCellType = TableCellType.HV_Merged;
                            }
                            else
                            {
                                WTCell.WordTableCellType = TableCellType.V_Merged;
                            }
                        }
                        else
                        {
                            WTCell.WordTableCellType = TableCellType.Normal;
                        }

                        foreach (XmlNode prg in Paragraphs)
                        {
                            XmlDocument PrgXML = new XmlDocument();
                            PrgXML.LoadXml(prg.OuterXml);

                            XmlNodeList wtextlines = PrgXML.GetElementsByTagName(WordXMLTags.WordTagName_Text);
                            if (wtextlines != null)
                            {
                                foreach (XmlNode _textn in wtextlines)
                                {
                                    if(_textn != null)
                                    {
                                        if (_textn.InnerText == string.Empty)
                                        {
                                            WTCell.WordTableCellValue += string.Empty;
                                        }
                                        else if(_textn.InnerText == null)
                                        {
                                            WTCell.WordTableCellValue = null;
                                        }
                                        else
                                        {
                                            string _txt = _textn.InnerText;

                                            _txt = _txt.Replace("\r", string.Empty);
                                            _txt = _txt.Replace("\n", string.Empty);
                                            _txt = _txt.Replace("\a", string.Empty);
                                            _txt = _txt.Replace("\v", string.Empty);
                                            _txt = _txt.Replace("\t", string.Empty);

                                            WTCell.WordTableCellValue += _txt;
                                            WTCell.WordTableCellValueType = TableCellValueType.Text;
                                        }
                                    }
                                }
                            }
                            if (WTCell.WordTableCellValue != null && WTCell.WordTableCellValue != string.Empty)
                            {
                                WTCell.WordTableCellValue += WordTable.SeparatorFor_Paragraph;
                            }

                            PrgXML.RemoveAll();
                        }

                        if (WTCell.WordTableCellValue != null && WTCell.WordTableCellValue != string.Empty)
                        {
                            WTCell.WordTableCellValue = WTCell.WordTableCellValue.Substring(0, WTCell.WordTableCellValue.LastIndexOf(WordTable.SeparatorFor_Paragraph));
                        }

                        XmlNode _GSpanNode = Prgs.GetElementsByTagName(WordXMLTags.WordTagName_GridSpan)[0];
                        XmlNode _VMergeNode = Prgs.GetElementsByTagName(WordXMLTags.WordTagName_VerticalMerge)[0];

                        if (_GSpanNode != null)
                        {
                            if (_GSpanNode.Attributes[WordXMLTags.WordTagAttrib_Value] != null)
                            {
                                HMerge_CellCount = Convert.ToInt32(_GSpanNode.Attributes[WordXMLTags.WordTagAttrib_Value].Value);

                                if (RowCounter == 0)
                                {
                                    VMerge.AddRange(Enumerable.Repeat<bool>(false, HMerge_CellCount - 1).ToArray());
                                }
                                else
                                {
                                    for (int bi = CellCounter + 1; bi < CellCounter + (HMerge_CellCount - 1); bi++)
                                    {
                                        VMerge[CellCounter] = false;
                                    }
                                }
                            }
                        }
                        else
                        {
                            HMerge_CellCount = 0;
                        }

                        if (_VMergeNode != null)
                        {
                            if (_VMergeNode.Attributes[WordXMLTags.WordTagAttrib_Value] != null)
                            {
                                if (_VMergeNode.Attributes[WordXMLTags.WordTagAttrib_Value].Value.ToLower() == WordXMLTags.WordTagAttribValue_Restart | _VMergeNode.Attributes[WordXMLTags.WordTagAttrib_Value].Value.ToLower() == WordXMLTags.WordTagAttribValue_Continue)
                                {
                                    VMerge[CellCounter] = true;
                                }
                            }
                        }
                        else
                        {
                            VMerge[CellCounter] = false;
                        }


                        WTRow.AddCell(WTCell);

                        CellCounter++;
                        CellCounter += HMerge_CellCount - 1;

                        Prgs.RemoveAll();

                        while (HMerge_CellCount > 1)
                        {
                            WTRow.AddCell(new WordTableCell(null, TableCellType.H_Merged, TableCellValueType.Null));
                            HMerge_CellCount--;
                        }
                    }
                }

                WTable.AddRow(WTRow);
                RowCounter++;
            }

            XMLDoc.RemoveAll();
            VMerge.Clear();

            return WTable;
        }
    }

}
