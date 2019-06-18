using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using MicroMWordLib.WordTable;
using MicroMWordLib.WordImage;
using MicroMWordLib.WordList;
using MicroMWordLib.WordParagraph;
using MicroMWordLib.WordText;
using MicroMWordLib.WordAdditionalElement;
using MicroMWordLib;

namespace DokuWikiFormatter
{
    public class DWikiFormatter
    {
        private string MyOutputWDirectory = string.Empty;
        private string MyOutputWDirectoryName = string.Empty;
        private string MyOutputDirectory = string.Empty;
        private string MyMediaFilesSource = string.Empty;
        private string MyOriginalInputFileName = string.Empty;

        private string prp_TableOfContentsString = "Table Of Contents";
        private string prp_TableOffiguresString = "Table Of Figures";

        private bool prp_ExportTOC = false;
        private bool prp_ExportTOF = false;
        private bool prp_AddFooterInfo = true;
        private DWikiSyntax.ListOrder prp_ListOrder = DWikiSyntax.ListOrder.Unordered;
        private bool prp_ShiftOrdering = true;
        private bool prp_TableFirstRowIsHeader = true;
        private bool prp_ConsiderImageSize = true;

        public string OriginalInputFileName { get => MyOriginalInputFileName; set => MyOriginalInputFileName = value; }
        public bool ExportTOC { get => prp_ExportTOC; set => prp_ExportTOC = value; }
        public bool ExportTOF { get => prp_ExportTOF; set => prp_ExportTOF = value; }
        public string TableOfContentsString { get => prp_TableOfContentsString; set => prp_TableOfContentsString = value; }
        public string TableOffiguresString { get => prp_TableOffiguresString; set => prp_TableOffiguresString = value; }
        public bool AddFooterInfo { get => prp_AddFooterInfo; set => prp_AddFooterInfo = value; }
        public DWikiSyntax.ListOrder ListOrder { get => prp_ListOrder; set => prp_ListOrder = value; }
        public bool ShiftOrdering { get => prp_ShiftOrdering; set => prp_ShiftOrdering = value; }
        public bool TableFirstRowIsHeader { get => prp_TableFirstRowIsHeader; set => prp_TableFirstRowIsHeader = value; }
        public bool ConsiderImageSize { get => prp_ConsiderImageSize; set => prp_ConsiderImageSize = value; }

        public DWikiFormatter(string OutputFolder, string MediaFilesSource)
        {
            MyOutputDirectory = OutputFolder;
            MyMediaFilesSource = MediaFilesSource;
        }

        public string ExportToDWikiContent(IWBaseElement[] in_WBElements, string OutputWFolderName)
        {
            try
            {

                if (in_WBElements == null || OutputWFolderName == null)
                {
                    return null;
                }

                MyOutputWDirectoryName = OutputWFolderName;
                //MyOutputWDirectory = MyOutputDirectory + "\\" + MyOutputWDirectoryName;
                //DirectoryInfo DInfo = new DirectoryInfo(MyOutputWDirectory);
                // Log
                //if (DInfo.Exists == false)
                //{
                //    DInfo.Create();
                //    // Log
                //    Logger.AddLog(new LogContent() { LogSubject = "[DokuWiki Exporting]", LogMessage = "Folder created: " + MyOutputWDirectory, LogSeverity = LoggerParameters.LogSeverity.DEBUG });
                //}

                StringBuilder out_DWikiOutput = new StringBuilder();

                foreach (IWBaseElement BElem in in_WBElements)
                {

                    if (BElem.GetType() == typeof(WParagraph))
                    {
                        out_DWikiOutput.Append(FormatWParagraph(BElem as WParagraph));
                        if ((BElem as WParagraph).ParagraphStyle == ParagraphStyle.SimpleContainer)
                        {
                            out_DWikiOutput.AppendLine(DWikiSyntax.TableNewLine);
                        }
                        else
                        {
                            out_DWikiOutput.AppendLine();
                        }
                    }
                    else if (BElem.GetType() == typeof(WList))
                    {
                        out_DWikiOutput.AppendLine(FormatWList(BElem as WList, ListOrder, ShiftOrdering));
                    }
                    else if (BElem.GetType() == typeof(WTable))
                    {
                        out_DWikiOutput.AppendLine(FormatWTable(BElem as WTable, TableFirstRowIsHeader));
                    }
                    else if (BElem.GetType() == typeof(WAElement))
                    {
                        WAElement AddElem = BElem as WAElement;

                        if (AddElem.ElementType == WAElement.WAElementType.TableOfContents && ExportTOC == true)
                        {
                            out_DWikiOutput.AppendLine(FormatWAElement(AddElem, TableOfContentsString));
                        }
                        else if(AddElem.ElementType == WAElement.WAElementType.TableOfFigures && ExportTOF == true)
                        {
                            out_DWikiOutput.AppendLine(FormatWAElement(AddElem, TableOffiguresString));
                        }
                    }
                }

                if (AddFooterInfo == true)
                {
                    out_DWikiOutput.AppendLine(Environment.NewLine + DWikiSyntax.Line);
                    out_DWikiOutput.Append("This file has been exported by **SHWordToDWiki**.");
                    out_DWikiOutput.Append(DWikiSyntax.TableNewLine);
                    out_DWikiOutput.Append("<nowiki>Export date              : " + DateTime.Now.ToString() + "</nowiki>");
                    out_DWikiOutput.Append(DWikiSyntax.TableNewLine);
                    out_DWikiOutput.Append("<nowiki>Original input file name : " + (MyOriginalInputFileName == string.Empty ? "Undefined" : MyOriginalInputFileName) + "</nowiki>");
                    out_DWikiOutput.Append(DWikiSyntax.TableNewLine);
                    out_DWikiOutput.Append("<nowiki>Output file name         : " + MyOutputWDirectoryName + "</nowiki>");
                    out_DWikiOutput.Append(DWikiSyntax.TableNewLine);
                    out_DWikiOutput.AppendLine("<nowiki>Computer name: " + Environment.MachineName + "</nowiki>");
                }
                string svdFPath = SaveAs(MyOutputDirectory, MyOutputWDirectoryName, out_DWikiOutput.ToString(), MyMediaFilesSource);

                return out_DWikiOutput.ToString();
            }
            catch (Exception Exp)
            {
                return null;
            }
        }

        public string GetDWikiOutput(IWBaseElement[] in_WBElements)
        {
            if (in_WBElements == null)
            {
                return null;
            }

            StringBuilder out_DWikiOutput = new StringBuilder();

            foreach (IWBaseElement BElem in in_WBElements)
            {
                string rslt_FormattedText = string.Empty;

                if (BElem.GetType() == typeof(WParagraph))
                {
                    rslt_FormattedText = FormatWParagraph(BElem as WParagraph);
                }
                else if (BElem.GetType() == typeof(WList))
                {
                    rslt_FormattedText = FormatWList(BElem as WList);
                }
                else if (BElem.GetType() == typeof(WTable))
                {
                    rslt_FormattedText = FormatWTable(BElem as WTable);
                }
            }
            return out_DWikiOutput.ToString();
        }

        public string FormatWList(WList in_Wlist, DWikiSyntax.ListOrder in_ListOrder = DWikiSyntax.ListOrder.Unordered, bool ShiftOrdering = true)
        {
            if (in_Wlist == null)
            {
                return null;
            }

            StringBuilder out_FList = new StringBuilder();
            string ListOrderSymbol = DWikiSyntax.UnorderedList;
            switch (in_ListOrder)
            {
                case DWikiSyntax.ListOrder.Ordered:
                    ListOrderSymbol = DWikiSyntax.OrderedList;
                    break;
                case DWikiSyntax.ListOrder.Unordered:
                    ListOrderSymbol = DWikiSyntax.UnorderedList;
                    break;
            }

            for (int LIIndx = 0; LIIndx < in_Wlist.ListItems.Count; LIIndx++)
            {
                out_FList.Append(DWikiSyntax.TwoWhitespace + ListOrderSymbol);
                bool FAdded = false;

                foreach (WListItemElement IElem in in_Wlist.ListItems[LIIndx].ItemElements)
                {
                    if (FAdded == true)
                    {
                        out_FList.Append(DWikiSyntax.TwoWhitespace);
                    }

                    if (IElem.ListItemElement.GetType() == typeof(WParagraph))
                    {
                        out_FList.Append(FormatWParagraph(IElem.ListItemElement as WParagraph));
                    }
                    else if (IElem.ListItemElement.GetType() == typeof(WTable))
                    {
                        string _DWTblCnt = FormatWTable(IElem.ListItemElement as WTable, TableFirstRowIsHeader);

                        _DWTblCnt = _DWTblCnt.Replace(DWikiSyntax.NoFormatting_b, string.Empty);
                        _DWTblCnt = _DWTblCnt.Replace(DWikiSyntax.NoFormatting_e, string.Empty);

                        _DWTblCnt = _DWTblCnt.Replace(Environment.NewLine, DWikiSyntax.NoFormatting_e + DWikiSyntax.TableNewLine + DWikiSyntax.NoFormatting_b);

                        out_FList.Append(DWikiSyntax.NoFormatting_b + _DWTblCnt + DWikiSyntax.NoFormatting_e);
                    }
                    out_FList.Append(DWikiSyntax.TableNewLine);
                    FAdded = (FAdded == false) ? true : FAdded;
                }

                if (out_FList.Length > DWikiSyntax.TableNewLine.Length)
                {
                    out_FList.Remove(out_FList.Length - DWikiSyntax.TableNewLine.Length, DWikiSyntax.TableNewLine.Length);
                }

                if (in_Wlist.ListItems[LIIndx].ChildList != null)
                {
                    DWikiSyntax.ListOrder LOrder = in_ListOrder;
                    if (ShiftOrdering == true)
                    {
                        if (in_ListOrder == DWikiSyntax.ListOrder.Ordered)
                        {
                            LOrder = DWikiSyntax.ListOrder.Unordered;
                        }
                        else
                        {
                            LOrder = DWikiSyntax.ListOrder.Ordered;
                        }
                    }
                    string _LChldItms = FormatWList(in_Wlist.ListItems[LIIndx].ChildList, LOrder, ShiftOrdering);
                    out_FList.Append(Environment.NewLine + DWikiSyntax.TwoWhitespace);
                    _LChldItms = _LChldItms.Replace(Environment.NewLine, Environment.NewLine + DWikiSyntax.TwoWhitespace);
                    out_FList.Append(_LChldItms);

                }

                out_FList.Append(Environment.NewLine);
            }

            if (out_FList.Length > Environment.NewLine.Length)
            {
                out_FList.Remove(out_FList.Length - Environment.NewLine.Length, Environment.NewLine.Length);
            }

            return out_FList.ToString();
        }

        public string FormatWTable(WTable in_WTable, bool FirstRowIsHeader = false)
        {
            StringBuilder out_FTable = new StringBuilder(string.Empty);

            try
            {
                for (int row = 0; row < in_WTable.TableRows.Count; row++)
                {
                    for (int col = 0; col < in_WTable.TableRows[row].TableCells.Count; col++)
                    {
                        WTableCell wtc = in_WTable.GetCell(row, col);

                        if (wtc.WordTableCellType == TableCellType.H_Merged)
                        {
                            if (row == 0 && FirstRowIsHeader == true)
                            {
                                out_FTable.Append(DWikiSyntax.TableHeader);
                            }
                            else
                            {
                                out_FTable.Append(DWikiSyntax.TableCell);
                            }
                        }
                        else if (wtc.WordTableCellType == TableCellType.V_Merged)
                        {
                            out_FTable.Append(DWikiSyntax.TableCell + DWikiSyntax.TableCell_VMerge);
                        }
                        else if (wtc.WordTableCellType == TableCellType.Normal)
                        {
                            if (row == 0 && FirstRowIsHeader == true)
                            {
                                out_FTable.Append(DWikiSyntax.TableHeader);
                            }
                            else
                            {
                                out_FTable.Append(DWikiSyntax.TableCell);
                            }

                            bool talign = false;
                            Alignment walign = Alignment.Left;
                            IWBaseElement[] AllWElems = wtc.CellElements.ToArray();

                            if (AllWElems != null)
                            {
                                if (talign == false)
                                {
                                    int _celii = 0;
                                    Alignment _calgn = Alignment.Center;
                                    while (_celii < wtc.CellElements.Count)
                                    {
                                        if (wtc.CellElements[_celii].GetType() == typeof(WParagraph))
                                        {
                                            _calgn = (wtc.CellElements[_celii] as WParagraph).Alingment;
                                            break;
                                        }
                                        _celii++;
                                    }

                                    switch (_calgn)
                                    {
                                        case Alignment.Left:
                                            out_FTable.Append(string.Empty);
                                            walign = Alignment.Left;
                                            break;
                                        case Alignment.Right:
                                            out_FTable.Append(DWikiSyntax.TwoWhitespace);
                                            walign = Alignment.Right;
                                            break;
                                        case Alignment.Center:
                                            out_FTable.Append(DWikiSyntax.TwoWhitespace);
                                            walign = Alignment.Center;
                                            break;
                                        default:
                                            walign = Alignment.Left;
                                            break;
                                    }

                                    talign = true;
                                }

                                foreach (IWBaseElement WEl in AllWElems)
                                {
                                    if (WEl.GetType() == typeof(WParagraph))
                                    {
                                        (WEl as WParagraph).ParagraphStyle = ParagraphStyle.SimpleContainer;
                                        out_FTable.Append(FormatWParagraph(WEl as WParagraph));
                                        out_FTable.Append(DWikiSyntax.TableNewLine);
                                    }
                                    else if (WEl.GetType() == typeof(WTable))
                                    {
                                        string TblDWikiCnt = FormatWTable(WEl as WTable, TableFirstRowIsHeader);

                                        if (walign != Alignment.Left)
                                        {
                                            out_FTable.Remove(out_FTable.Length - 1, 1);
                                        }
                                        //TblDWikiCnt = TblDWikiCnt.Replace(Environment.NewLine, DWkiSyntax.TableNewLine);

                                        TblDWikiCnt = TblDWikiCnt.Replace(DWikiSyntax.NoFormatting_b, string.Empty);
                                        TblDWikiCnt = TblDWikiCnt.Replace(DWikiSyntax.NoFormatting_e, string.Empty);

                                        TblDWikiCnt = TblDWikiCnt.Replace(Environment.NewLine, DWikiSyntax.NoFormatting_e + DWikiSyntax.TableNewLine + DWikiSyntax.NoFormatting_b);
                                        out_FTable.Append(DWikiSyntax.NoFormatting_b + TblDWikiCnt + DWikiSyntax.NoFormatting_e);
                                        out_FTable.Append(DWikiSyntax.TableNewLine);

                                    }
                                    else
                                    {
                                        string LstDWikiCnt = FormatWList(WEl as WList);

                                        if (walign != Alignment.Left)
                                        {
                                            out_FTable.Remove(out_FTable.Length - 1, 1);
                                        }

                                        LstDWikiCnt = LstDWikiCnt.Replace(DWikiSyntax.NoFormatting_b, string.Empty);
                                        LstDWikiCnt = LstDWikiCnt.Replace(DWikiSyntax.NoFormatting_e, string.Empty);

                                        LstDWikiCnt = LstDWikiCnt.Replace(Environment.NewLine, DWikiSyntax.NoFormatting_e + DWikiSyntax.TableNewLine + DWikiSyntax.NoFormatting_b);
                                        out_FTable.Append(DWikiSyntax.NoFormatting_b + LstDWikiCnt + DWikiSyntax.NoFormatting_e);

                                        out_FTable.Append(DWikiSyntax.TableNewLine);
                                    }
                                }

                                if (out_FTable.Length > DWikiSyntax.TableNewLine.Length)
                                {
                                    out_FTable.Remove(out_FTable.Length - DWikiSyntax.TableNewLine.Length, DWikiSyntax.TableNewLine.Length);
                                }

                                if (talign == true)
                                {
                                    switch (walign)
                                    {
                                        case Alignment.Left:
                                            out_FTable.Append(DWikiSyntax.TwoWhitespace);
                                            break;
                                        case Alignment.Right:
                                            out_FTable.Append(string.Empty);
                                            break;
                                        case Alignment.Center:
                                            out_FTable.Append(DWikiSyntax.TwoWhitespace);
                                            break;
                                        default: break;
                                    }

                                }
                            }

                        }

                        if (col == in_WTable.TableRows[0].TableCells.Count - 1)
                        {
                            if (row == 0 && FirstRowIsHeader == true)
                            {
                                out_FTable.Append(DWikiSyntax.TableHeader);
                            }
                            else
                            {
                                out_FTable.Append(DWikiSyntax.TableCell);
                            }
                        }
                    }
                    out_FTable.Append(Environment.NewLine);
                }
            }
            catch (Exception Exp)
            {
                return null;
            }
            return out_FTable.ToString();
        }

        public string FormatWImage(WImage in_WImage, string PIdentifier)
        {
            string out_FImage = string.Empty;
            if (in_WImage.Width >= 0 && in_WImage.Height >= 0 && ConsiderImageSize == true)
            {
                out_FImage = DWikiSyntax.IEImage_b + PIdentifier + "_mfiles" + ":" + Path.GetFileName(in_WImage.ImagePath) + DWikiSyntax.ImageSizeAdd + in_WImage.Width + DWikiSyntax.ImageSizeSeparator + in_WImage.Height + DWikiSyntax.IEImage_e;
            }
            else
            {
                out_FImage = DWikiSyntax.IEImage_b + PIdentifier + "_mfiles" + ":" + Path.GetFileName(in_WImage.ImagePath) + DWikiSyntax.IEImage_e;
            }
            return out_FImage;
        }

        public string FormatWParagraph(WParagraph in_WParag)
        {
            StringBuilder out_WParag = new StringBuilder();
            if (in_WParag == null)
            {
                return null;
            }

            bool isHeader = false;
            string headerText = string.Empty;
            if (in_WParag.ParagraphStyle.ToString().ToLower().Contains("heading"))
            {
                isHeader = true;
                switch (in_WParag.ParagraphStyle)
                {
                    case ParagraphStyle.WordHeading1:
                        out_WParag.Append(DWikiSyntax.Heading1);
                        headerText = DWikiSyntax.Heading1;
                        break;
                    case ParagraphStyle.WordHeading2:
                        out_WParag.Append(DWikiSyntax.Heading2);
                        headerText = DWikiSyntax.Heading2;
                        break;
                    case ParagraphStyle.WordHeading3:
                        out_WParag.Append(DWikiSyntax.Heading3);
                        headerText = DWikiSyntax.Heading3;
                        break;
                    case ParagraphStyle.WordHeading4:
                        out_WParag.Append(DWikiSyntax.Heading4);
                        headerText = DWikiSyntax.Heading4;
                        break;
                    case ParagraphStyle.WordHeading5:
                        out_WParag.Append(DWikiSyntax.Heading5);
                        headerText = DWikiSyntax.Heading5;
                        break;
                    default:
                        out_WParag.Append(DWikiSyntax.Heading5);
                        headerText = DWikiSyntax.Heading5;
                        break;
                }
            }

            foreach (IWParagraph WPEl in in_WParag.Elements)
            {
                if (WPEl.GetType() == typeof(WImage))
                {
                    if (string.IsNullOrEmpty((WPEl as WImage).ImagePath) == false)
                    {
                        out_WParag.Append(DWikiSyntax.Whitespace);
                        out_WParag.Append(FormatWImage(WPEl as WImage, MyOutputWDirectoryName));
                        out_WParag.Append(DWikiSyntax.Whitespace);
                    }

                    //CopyImage((WPEl as WImage).ImagePath);
                }
                else
                {
                    WTextPart WTPart = WPEl as WTextPart;
                    if (isHeader == false)
                    {
                        if (WTPart.Bold == true)
                        {
                            out_WParag.Append(DWikiSyntax.BoldText_b);
                        }

                        if (WTPart.Italic == true)
                        {
                            out_WParag.Append(DWikiSyntax.ItalicText_b);
                        }

                        if (WTPart.Underline == true)
                        {
                            out_WParag.Append(DWikiSyntax.UnderLinedText_b);
                        }
                    }

                    if (isHeader == true)
                    {
                        out_WParag.Append(WTPart.Text);
                    }
                    else
                    {
                        out_WParag.Append(DWikiSyntax.NoFormatting_b + WTPart.Text + DWikiSyntax.NoFormatting_e);
                    }

                    if (isHeader == false)
                    {

                        if (WTPart.Underline == true)
                        {
                            out_WParag.Append(DWikiSyntax.UnderLinedText_e);
                        }

                        if (WTPart.Italic == true)
                        {
                            out_WParag.Append(DWikiSyntax.ItalicText_e);
                        }

                        if (WTPart.Bold == true)
                        {
                            out_WParag.Append(DWikiSyntax.BoldText_e);
                        }
                    }
                }
            }


            if (isHeader == true)
            {
                out_WParag.Append(headerText);
            }

            return out_WParag.ToString();
        }

        public string FormatWAElement(WAElement in_WAElement, string in_HeaderText)
        {
            StringBuilder out_WAElemWikiCnt = new StringBuilder();

            if (in_WAElement == null)
            {
                return null;
            }

            int FCount = in_WAElement.Lines[0].Fields.Count;

            in_HeaderText = (in_HeaderText == null) ? string.Empty : in_HeaderText;
            out_WAElemWikiCnt.Append(DWikiSyntax.TableHeader + DWikiSyntax.TwoWhitespace + in_HeaderText + DWikiSyntax.TwoWhitespace);

            for (int hdr = 0; hdr < FCount; hdr++)
            {
                out_WAElemWikiCnt.Append(DWikiSyntax.TableHeader);
            }

            out_WAElemWikiCnt.Append(Environment.NewLine);

            foreach (WAElementLine ElLine in in_WAElement.Lines)
            {
                foreach (WAElementLineField LnField in ElLine.Fields)
                {
                    out_WAElemWikiCnt.Append(DWikiSyntax.TableCell);

                    foreach (WTextPart WTextPrt in LnField.Elements)
                    {
                        WTextPart WTPart = WTextPrt as WTextPart;
                        if (WTPart.Bold == true)
                        {
                            out_WAElemWikiCnt.Append(DWikiSyntax.BoldText_b);
                        }
                        else if (WTPart.Italic == true)
                        {
                            out_WAElemWikiCnt.Append(DWikiSyntax.ItalicText_b);
                        }
                        else if (WTPart.Underline == true)
                        {
                            out_WAElemWikiCnt.Append(DWikiSyntax.UnderLinedText_b);
                        }

                        out_WAElemWikiCnt.Append(DWikiSyntax.NoFormatting_b + WTPart.Text + DWikiSyntax.NoFormatting_e);


                        if (WTPart.Bold == true)
                        {
                            out_WAElemWikiCnt.Append(DWikiSyntax.BoldText_e);
                        }
                        else if (WTPart.Italic == true)
                        {
                            out_WAElemWikiCnt.Append(DWikiSyntax.ItalicText_e);
                        }
                        else if (WTPart.Underline == true)
                        {
                            out_WAElemWikiCnt.Append(DWikiSyntax.UnderLinedText_e);
                        }

                    }

                    out_WAElemWikiCnt.Append(DWikiSyntax.TableCell);
                }

                out_WAElemWikiCnt.Append(Environment.NewLine);
            }

            return out_WAElemWikiCnt.ToString();
        }

        public string SaveAs(string OutputFolderName, string FileName, string Data, string MediaFFToCopyFrom)
        {
            string cpth = string.Empty;

            try
            {
                //FileName = FileName.Replace("_", string.Empty);
                StringBuilder DWTFile = new StringBuilder(string.Empty);

                Directory.CreateDirectory(OutputFolderName + "\\" + FileName);
                Directory.CreateDirectory(OutputFolderName + "\\" + FileName + "\\" + FileName + "_mfiles");

                cpth = OutputFolderName + "\\" + FileName + "\\" + FileName + ".txt";

                File.WriteAllText(cpth, Data);

                DirectoryInfo di = new DirectoryInfo(MediaFFToCopyFrom);
                foreach (FileInfo finfo in di.GetFiles())
                {
                    finfo.CopyTo(OutputFolderName + "\\" + FileName + "\\" + FileName + "_mfiles" + "\\" + finfo.Name, true);
                }
            }
            catch
            {
                return string.Empty;
            }
            return cpth;
        }

        private bool CopyImage(string ImgPath)
        {
            try
            {
                string ImgOutput = MyOutputWDirectory + "\\" + MyOutputWDirectoryName;
                DirectoryInfo DInfo = new DirectoryInfo(ImgOutput);

                if (DInfo.Exists == false)
                {
                    DInfo.Create();
                }

                FileInfo FInfo = new FileInfo(ImgPath);
                FInfo.CopyTo(ImgOutput);
            }
            catch
            {
                return false;
            }

            return true;
        }
    }
}
