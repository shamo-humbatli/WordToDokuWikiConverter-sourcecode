using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using DokuWikiFormatter;
namespace WordToDWiki
{
    public class ConfigBuilder
    {
        private string prp_Directories_OutputDirectory = "Output";
        private string prp_Directories_WorkingDirectory = "WDirectory";
        //private string prp_Directories_LoggingDirectory = "Logs";

        private int prp_SHLogger_LogLineCount = 500;
        private int prp_SHLogger_LogMessageLength = 500;
        private string prp_SHLogger_LogLineStartString = ">>";
        private int prp_SHLogger_LogSubjetLength = 50;

        private bool prp_DWikiExporter_AddFooterInfo = true;
        private bool prp_DWikiExporter_ExportTOC = false;
        private bool prp_DWikiExporter_ExportTOF = false;
        private DWikiSyntax.ListOrder prp_DWikiExporter_ListOrder = DWikiSyntax.ListOrder.Unordered;
        private bool prp_DWikiExporter_ShiftOrdering = true;
        private bool prp_SHMicroMWordLib_JoinImageSelections = false;
        private bool prp_DokuWikiExporter_TableFirstRowIsHeader = true;
        private bool prp_DokuWikiExporter_ConsiderImageSize = true;

        public string Directories_OutputDirectory { get => prp_Directories_OutputDirectory; }
        public string Directories_WorkingDirectory { get => prp_Directories_WorkingDirectory; }
        //public string Directories_LoggingDirectory { get => prp_Directories_LoggingDirectory; }
        public int SHLogger_LogLineCount { get => prp_SHLogger_LogLineCount; }
        public int SHLogger_LogMessageLength { get => prp_SHLogger_LogMessageLength; }
        public string SHLogger_LogLineStartString { get => prp_SHLogger_LogLineStartString; }
        public bool DWikiExporter_AddFooterInfo { get => prp_DWikiExporter_AddFooterInfo; }
        public bool DWikiExporter_ExportTOC { get => prp_DWikiExporter_ExportTOC; }
        public bool DWikiExporter_ExportTOF { get => prp_DWikiExporter_ExportTOF; }
        public int  SHLogger_LogSubjetLength { get => prp_SHLogger_LogSubjetLength; }
        public DWikiSyntax.ListOrder DWikiExporter_ListOrder { get => prp_DWikiExporter_ListOrder; }
        public bool DWikiExporter_ShiftOrdering { get => prp_DWikiExporter_ShiftOrdering; }
        public bool SHMicroMWordLib_JoinImageSelections { get => prp_SHMicroMWordLib_JoinImageSelections; }
        public bool DokuWikiExporter_TableFirstRowIsHeader { get => prp_DokuWikiExporter_TableFirstRowIsHeader; }
        public bool DokuWikiExporter_ConsiderImageSize { get => prp_DokuWikiExporter_ConsiderImageSize; }

        public ConfigBuilder()
        {

        }

        public Hashtable BuildConfig(string ConfFilePath)
        {
            try
            {
                Hashtable out_HTable = new Hashtable();

                XmlDocument ConfFile = new XmlDocument();
                ConfFile.Load(ConfFilePath);

                XmlNode nd_SubTags = ConfFile.GetElementsByTagName("Directories")[0];

                if (nd_SubTags != null)
                {
                    foreach (XmlNode chld in nd_SubTags.ChildNodes)
                    {
                        if (chld.Name == "OutputDirectory")
                        {
                            prp_Directories_OutputDirectory = chld.InnerText;
                            out_HTable.Add("Output", Directories_OutputDirectory);
                        }
                        else if (chld.Name == "WorkingDirectory")
                        {
                            prp_Directories_WorkingDirectory = chld.InnerText;
                            out_HTable.Add("WorkingDirectory", Directories_WorkingDirectory);
                        }
                        //else if (chld.Name == "LoggingDirectory")
                        //{
                        //    //prp_Directories_LoggingDirectory = chld.InnerText;
                        //    //out_HTable.Add("LoggingDirectory", Directories_LoggingDirectory);
                        //}
                    }
                }

                nd_SubTags = ConfFile.GetElementsByTagName("Logger")[0];

                if (nd_SubTags != null)
                {
                    foreach (XmlNode chld in nd_SubTags.ChildNodes)
                    {
                        if (chld.Name == "LogLineCount")
                        {
                            prp_SHLogger_LogLineCount = Convert.ToInt32(chld.InnerText);
                            prp_SHLogger_LogLineCount = prp_SHLogger_LogLineCount < 0 ? 0 : prp_SHLogger_LogLineCount;
                            out_HTable.Add("LogLineCount", SHLogger_LogLineCount);
                        }
                        else if (chld.Name == "LogMessageLength")
                        {
                            prp_SHLogger_LogMessageLength = Convert.ToInt32(chld.InnerText);
                            out_HTable.Add("LogMessageLength", SHLogger_LogMessageLength);
                        }
                        else if (chld.Name == "LogSubjetLength")
                        {
                            prp_SHLogger_LogSubjetLength = Convert.ToInt32(chld.InnerText);
                            out_HTable.Add("LogSubjetLength", SHLogger_LogSubjetLength);
                        }
                        else if (chld.Name == "LogLineStartString")
                        {
                            prp_SHLogger_LogLineStartString = chld.InnerText;
                            out_HTable.Add("LogLineStartString", SHLogger_LogLineStartString);
                        }
                    }

                    nd_SubTags = ConfFile.GetElementsByTagName("DokuWikiExporter")[0];

                    if (nd_SubTags != null)
                    {
                        foreach (XmlNode chld in nd_SubTags.ChildNodes)
                        {
                            if (chld.Name == "AddFooterInfo")
                            {
                                prp_DWikiExporter_AddFooterInfo = chld.InnerText == "true" ? true : false;
                                out_HTable.Add("AddFooterInfo", DWikiExporter_AddFooterInfo);
                            }
                            else if (chld.Name == "ExportTOC")
                            {
                                prp_DWikiExporter_ExportTOC = chld.InnerText == "true" ? true : false;
                                out_HTable.Add("ExportTOC", DWikiExporter_ExportTOC);
                            }
                            else if (chld.Name == "ExportTOF")
                            {
                                prp_DWikiExporter_ExportTOF = chld.InnerText == "true" ? true : false;
                                out_HTable.Add("ExportTOF", DWikiExporter_ExportTOF);
                            }
                            else if (chld.Name == "ListOrder")
                            {
                                prp_DWikiExporter_ListOrder = chld.InnerText == "Ordered" ? DWikiSyntax.ListOrder.Ordered : DWikiSyntax.ListOrder.Unordered;
                                out_HTable.Add("ListOrder", DWikiExporter_ListOrder);
                            }
                            else if (chld.Name == "ShiftOrdering")
                            {
                                prp_DWikiExporter_ShiftOrdering = chld.InnerText == "true" ? true : false;
                                out_HTable.Add("ShiftOrdering", DWikiExporter_ShiftOrdering);
                            }
                            else if (chld.Name == "TableFirstRowIsHeader")
                            {
                                prp_DokuWikiExporter_TableFirstRowIsHeader = chld.InnerText == "true" ? true : false;
                                out_HTable.Add("TableFirstRowIsHeader", DWikiExporter_ShiftOrdering);
                            }
                            else if (chld.Name == "ConsiderImageSize")
                            {
                                prp_DokuWikiExporter_ConsiderImageSize = chld.InnerText == "true" ? true : false;
                                out_HTable.Add("ConsiderImageSize", DWikiExporter_ShiftOrdering);
                            }

                            


                        }
                    }

                    nd_SubTags = ConfFile.GetElementsByTagName("MicroMWordLib")[0];

                    if (nd_SubTags != null)
                    {
                        foreach (XmlNode chld in nd_SubTags.ChildNodes)
                        {
                            if (chld.Name == "JoinImageSelections")
                            {
                                prp_SHMicroMWordLib_JoinImageSelections = chld.InnerText == "true" ? true : false;
                                out_HTable.Add("JoinImageSelections", DWikiExporter_AddFooterInfo);
                            }
                        }
                    }
                }

                return out_HTable;
            }
            catch (Exception Exp)
            {
                return null;
            }
        }
    }
}
