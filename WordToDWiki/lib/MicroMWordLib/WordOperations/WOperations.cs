using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using WordApp = Microsoft.Office.Interop.Word;
using MicroMWordLib.WordContentSelection;
using MicroMWordLib.WordImage;
using MicroMWordLib.WordList;
using MicroMWordLib.WordParagraph;
using MicroMWordLib.WordTable;
using MicroMWordLib.WordText;
using System.IO;
using LittleLyreLogger;
using Microsoft.Office.Interop.Word;

namespace MicroMWordLib.WordOperations
{
    public class WOperations
    {
        private WordApp.Application MyWordApplication = null;
        private WordApp.Document MyWordDocument = null;
        private string prp_ImageFileBaseName = "Images";
        private string MyWordFilePath = string.Empty;
        private string MyWorkingFolder = string.Empty;
        private Object ref_Missing = WParameters.Missing;
        private ILittleLyreLogger prp_Logger = null;
        private bool prp_JoinImages = false;

        public string WorkingFolder { get => MyWorkingFolder; set => MyWorkingFolder = value; }
        public ILittleLyreLogger Logger
        {
            get => prp_Logger;
            set => prp_Logger = value;
        }
        public string ImageFileBaseName { get => prp_ImageFileBaseName; set => prp_ImageFileBaseName = value; }
        public bool JoinImages { get => prp_JoinImages; set => prp_JoinImages = value; }

        public WOperations()
        {
            prp_Logger = new LineLogger();
        }

        public WOperations(string WordFilePath, string WFolder)
        {
            MyWordFilePath = WordFilePath;


            ImageFileBaseName = Path.GetFileNameWithoutExtension(WordFilePath);

            ImageFileBaseName = ImageFileBaseName.Trim();
            ImageFileBaseName = ImageFileBaseName.Replace(" ", string.Empty);

            ImageFileBaseName = (ImageFileBaseName.Length > 200) ? ImageFileBaseName.Substring(0, ImageFileBaseName.Length - 50) + "..." : ImageFileBaseName;

            MyWorkingFolder = WFolder;

            prp_Logger = new ListLogger();
        }



        public IWBaseElement[] GetWordElements()
        {

            // Log
            AddLog(Logger, "Document Operation", "Opening word document (Visible: false, ReadOnly: true) ...");
            MyWordApplication = new WordApp.Application();
            MyWordDocument = MyWordApplication.Documents.Open(MyWordFilePath, ReadOnly: false, Visible: false, OpenAndRepair: false, AddToRecentFiles: false, Revert: false, NoEncodingDialog: true, ConfirmConversions: false);

            try
            {
                //MyWordDocument = MyWordApplication.Documents.Open(MyWordFilePath, ref_Missing, false, false, ref_Missing, ref_Missing, ref_Missing, ref_Missing, ref_Missing, ref_Missing, ref_Missing, false, ref_Missing, ref_Missing, ref_Missing, ref_Missing);
  
                // Log
                AddLog(Logger, "Document Operation", "Word document opened.");

                List<IWBaseElement> WElementsList = new List<IWBaseElement>();

                // Logging
                AddLog(Logger, "Getting WElements", "Trying to get all fist level word paragraphs.");

                WParagraph[] WParags = WParagraphReader.GetAllParagraphs(MyWordApplication, MyWordDocument, Logger);

                // Logging
                AddLog(Logger, "Getting Wlements", "Trying to export all shapes as image");
                WImage[] WImgs = null;
                if (JoinImages == true)
                {
                    WImgs = WImageExporter.ExportImages(MyWordApplication, MyWordDocument, WCSelectionOperations.JoinSelections(WImage.GetAllContentSelections(MyWordApplication, MyWordDocument, Logger), 10, "image_"), MyWorkingFolder, ImageFileBaseName, Logger);
                }
                else
                {
                    WImgs = WImageExporter.ExportImages(MyWordApplication, MyWordDocument, MyWorkingFolder, ImageFileBaseName, Logger);
                }

                // Logging
                AddLog(Logger, "Getting WElements", "Trying to get all fist level word lists.");

                WList[] WLists = WListReader.GetAllLists(MyWordApplication, MyWordDocument, Logger);

                // Logging
                AddLog(Logger, "Getting WElements", "Trying to get all fist level word tables.");
                WTable[] WTables = WTableReader.GetAllTables(MyWordApplication, MyWordDocument, Logger);

                #region Logging
                AddLog(Logger, "Content Selection Recovery", "Trying to  recover content selection of inner elements of all lists...");
                if (WLists.Length < 1)
                {

                    AddLog(Logger, "Content Selection Recovery", "There is no list for inner content selection recovery.");
                }
                else
                {
                    AddLog(Logger, "Content Selection Recovery", "List count for inner content selection recovery : " + WLists.Length);
                }
                #endregion
                for (int li = 0; li < WLists.Length; li++)
                {
                    #region Logging
                    AddLog(Logger, "Selection Recovery", "List [" + (li + 1) + "/" + WLists.Length + "]" + " Recovering inner selection...", LoggerParameters.LogSeverity.DEBUG);
                    #endregion

                    int RCnt = WLists[li].RecoverInnerContentSelection(WParagraph.GetAllContentSelectionsForRange(MyWordApplication, MyWordDocument, WLists[li].ContentSelection.ContentSelectionStart, WLists[li].ContentSelection.ContentSelectionEnd));

                    // Logging
                    AddLog(Logger, "Selection Recovery", "List [" + (li + 1) + "/" + WLists.Length + "]" + "Recovered paragraph content selection count: " + RCnt, LoggerParameters.LogSeverity.DEBUG);
                }

                #region Logging
                AddLog(Logger, "Selection Recovery", "Trying to  recover content selection of inner elements of all tables...");
                if (WLists.Length < 1)
                {

                    AddLog(Logger, "Selection Recovery", "There is no table for inner content selection recovery.");
                }
                else
                {
                    AddLog(Logger, "Selection Recovery", "Table count for inner content selection recovery : " + WTables.Length);
                }
                #endregion
                for (int tbl = 0; tbl < WTables.Length; tbl++)
                {
                    #region Logging
                    AddLog(Logger, "Selection Recovery", "Table [" + (tbl + 1) + "/" + WTables.Length + "]" + " Recovering inner selection...", LoggerParameters.LogSeverity.DEBUG);
                    #endregion

                    int RCnt = WTables[tbl].RecoverInnerContentSelection(WParagraph.GetAllContentSelectionsForRange(MyWordApplication, MyWordDocument, WTables[tbl].ContentSelection.ContentSelectionStart, WTables[tbl].ContentSelection.ContentSelectionEnd));

                    // Logging
                    AddLog(Logger, "Selection Recovery", "Table [" + (tbl + 1) + "/" + WTables.Length + "]" + "Recovery count: " + RCnt, LoggerParameters.LogSeverity.DEBUG);
                }

                #region Logging
                AddLog(Logger, "Image Recovery", "Trying to  recover images of the all paragraphs...");
                if (WParags.Length < 1)
                {

                    AddLog(Logger, "Image Recovery", "There is no paragraph for image recovery.");
                }
                else
                {
                    AddLog(Logger, "Image Recovery", "Paragraph count for image recovery : " + WParags.Length);
                }
                #endregion
                WParags = WParagraph.RecoverImages(WParags, WImgs);
                // Logging
                AddLog(Logger, "Image Recovery", "Image recovery ended for the paragraphs.", LoggerParameters.LogSeverity.DEBUG);

                #region Logging
                AddLog(Logger, "Image Recovery", "Trying to  recover images of the all lists...");
                if (WLists.Length < 1)
                {

                    AddLog(Logger, "Image Recovery", "There is no list for image recovery.");
                }
                else
                {
                    AddLog(Logger, "Image Recovery", "List count for image recovery : " + WLists.Length);
                }
                #endregion
                WLists = WList.RecoverImages(WLists, WImgs);
                // Logging
                AddLog(Logger, "Image Recovery", "Image recovery ended for the lists.", LoggerParameters.LogSeverity.DEBUG);

                #region Logging
                AddLog(Logger, "Image Recovery", "Trying to  recover images of the all tables...");
                if (WLists.Length < 1)
                {

                    AddLog(Logger, "Image Recovery", "There is no table for image recovery.");
                }
                else
                {
                    AddLog(Logger, "Image Recovery", "Table count for image recovery : " + WTables.Length);
                }
                #endregion

                WTables = WTable.RecoverImages(WTables, WImgs);

                // Logging
                AddLog(Logger, "Image Recovery", "Image recovery ended for the tables.", LoggerParameters.LogSeverity.DEBUG);

                // Log
                AddLog(Logger, "Document Operation", "Closing word document (SaveChanges: false) ...");
             
                MyWordDocument.Close(WdSaveOptions.wdDoNotSaveChanges, ref_Missing, ref_Missing);
                MyWordApplication.Quit(WdSaveOptions.wdDoNotSaveChanges, ref_Missing, ref_Missing);

                int RRslt = System.Runtime.InteropServices.Marshal.ReleaseComObject(MyWordApplication);

                // Log
                AddLog(Logger, "Document Operation", "Opened documents closed and application released. Output: " + RRslt);
                AddLog(Logger, "Getting WElements", "Process ended, returning results.");

                WElementsList.AddRange(WParags);
                WElementsList.AddRange(WLists);
                WElementsList.AddRange(WTables);

                return ArrangeInAscendingOrder(WElementsList.ToArray());
            }
            catch(Exception Exp)
            {
                // Log
                AddLog(Logger, "Main Operation", "Error occured. Message -> " + Exp.Message, LoggerParameters.LogSeverity.ERROR);

        
                // Log
                AddLog(Logger, "Document Operation", "Closing word document (SaveChanges: false) ...");
      
                MyWordDocument.Close(WdSaveOptions.wdDoNotSaveChanges, ref_Missing, ref_Missing);
                MyWordApplication.Quit(WdSaveOptions.wdDoNotSaveChanges, ref_Missing, ref_Missing);

                int RRslt = System.Runtime.InteropServices.Marshal.ReleaseComObject(MyWordApplication);

                // Log
                AddLog(Logger, "Document Operation", "Opened documents closed and application released. Output: " + RRslt);
                AddLog(Logger, "Getting WElements", "Process ended, returning [null]");
                return null;
            }
            finally
            {
                // Log
                AddLog(Logger, "Getting WElements", "-------------");
            }
        }

        public static IWBaseElement[] ArrangeInAscendingOrder(IWBaseElement[] in_WElements)
        {
            List<IWBaseElement> _WElemList = new List<IWBaseElement>(in_WElements);
            _WElemList.Sort((a, b) => a.ContentSelection.ContentSelectionStart.CompareTo(b.ContentSelection.ContentSelectionStart));

            return _WElemList.ToArray();
        }

        private void AddLog(ILittleLyreLogger Logger, string LSubj, string LMsg, LoggerParameters.LogSeverity LSvrt = LoggerParameters.LogSeverity.INFO)
        {
            Logger.AddLog(new LogContent() { LogSubject = LSubj, LogMessage = LMsg, LogSeverity = LSvrt });
        }
    }
}
