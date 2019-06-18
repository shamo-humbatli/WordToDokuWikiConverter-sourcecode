using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Word;
using MicroMWordLib.WordContentSelection;
using LittleLyreLogger;
using MicroMWordLib.WordOperations;

namespace MicroMWordLib.WordImage
{
    public class WImage : IWParagraph
    {
        private WCSelection prp_ContentSelection = null;
        private string prp_WImagePath = null;
        private int prp_Width = -1;
        private int prp_Height = -1;
        public string ImagePath { get => prp_WImagePath; set => prp_WImagePath = value; }
        public WCSelection ContentSelection { get => prp_ContentSelection; set => prp_ContentSelection = value; }
        public int Height { get => prp_Height; set => prp_Height = value; }
        public int Width { get => prp_Width; set => prp_Width = value; }

        public static WCSelection[] GetAllContentSelections(Application MWordApp, Document MWordDocument, ILittleLyreLogger Logger)
        {
            Logger.AddLog(new LogContent() { LogSubject = "Getting Selection", LogMessage = "Trying to get all image content selections...", LogSeverity = LoggerParameters.LogSeverity.INFO });

            List<WCSelection> DraftIShapes = new List<WCSelection>();
            List<WCSelection> DraftShapes = new List<WCSelection>();

            Document DraftDoc = MWordApp.Documents.Add(Visible: true);
            //Document DraftDoc = WordApp.Documents.Add(WParameters.Missing, WParameters.Missing, WParameters.Missing, false);
            MWordDocument.Select();
            MWordApp.Selection.Copy();
            DraftDoc.Range().Paste();
            DraftDoc.Activate();
            // Logging
            Logger.AddLog(new LogContent() { LogSubject = "Getting Selection", LogMessage = "Getting inline shapes...", LogSeverity = LoggerParameters.LogSeverity.INFO });

            //List<WCSelection> List_IWCSelection = new List<WCSelection>();

            if (DraftDoc.InlineShapes.Count == 0)
            {
                // Logging
                Logger.AddLog(new LogContent() { LogSubject = "Getting Selection", LogMessage = "No inline shapes found.", LogSeverity = LoggerParameters.LogSeverity.INFO });
            }
            else
            {
                // Logging
                Logger.AddLog(new LogContent() { LogSubject = "Getting Selection", LogMessage = "Inline shape  count: " + DraftDoc.InlineShapes.Count, LogSeverity = LoggerParameters.LogSeverity.INFO });
            }

            for (int IShpI = 1; IShpI <= DraftDoc.InlineShapes.Count; IShpI++)
            {
                InlineShape IShape = DraftDoc.InlineShapes[IShpI];
                IShape.Select();

                //try
                //{
                //    string cxml = MWordDocument.Range(MWordApp.Selection.Start, MWordApp.Selection.End).XML;

                //    if(string.IsNullOrEmpty(cxml) == false)
                //    {
                //        continue;
                //    }
                //}
                //catch
                //{
                //    // Do nothing
                //}

                WCSelection wcs = new WCSelection();
                wcs.ContentSelectionStart = MWordApp.Selection.Start;
                wcs.ContentSelectionEnd = MWordApp.Selection.End;
                DraftIShapes.Add(wcs);

                // Logging
                Logger.AddLog(new LogContent() { LogSubject = "Getting Selection", LogMessage = "Inline shape [" + IShpI + "/" + DraftDoc.InlineShapes.Count + "]. Content selection -> [" + wcs.ContentSelectionStart + " : " + wcs.ContentSelectionEnd + "]", LogSeverity = LoggerParameters.LogSeverity.DEBUG });

            }

            if (DraftDoc.Shapes.Count == 0)
            {
                // Logging
                Logger.AddLog(new LogContent() { LogSubject = "Getting Selection", LogMessage = "No shapes found.", LogSeverity = LoggerParameters.LogSeverity.INFO });
            }
            else
            {
                // Logging
                Logger.AddLog(new LogContent() { LogSubject = "Getting Selection", LogMessage = "Shape  count: " + DraftDoc.Shapes.Count, LogSeverity = LoggerParameters.LogSeverity.INFO });
            }
            for (int ShpI = 1; ShpI <= DraftDoc.Shapes.Count; ShpI++)
            {
                Shape TShp = DraftDoc.Shapes[ShpI];

                TShp.ConvertToInlineShape().Select();

                //try
                //{
                //    string cxml = MWordDocument.Range(MWordApp.Selection.Start, MWordApp.Selection.End).XML;

                //    if (string.IsNullOrEmpty(cxml) == false)
                //    {
                //        continue;
                //    }
                //}
                //catch
                //{
                //    // Do nothing
                //}

                WCSelection wcs = new WCSelection();
                //ds.ID = "ishape_" + Guid.NewGuid();
                //ds.Index = ShpI;
                wcs.ContentSelectionStart = MWordApp.Selection.Start;
                wcs.ContentSelectionEnd = MWordApp.Selection.End;

                DraftShapes.Add(wcs);
                // Logging
                Logger.AddLog(new LogContent() { LogSubject = "Getting Selection", LogMessage = "Inline shape [" + ShpI + "/" + DraftDoc.Shapes.Count + "]. Content selection [start : end] = " + wcs.ContentSelectionStart + " : " + wcs.ContentSelectionEnd + "]", LogSeverity = LoggerParameters.LogSeverity.DEBUG });
                //// Log info
                //AddToLog("Shape was added as selection. Selection[Start, End] = [" + ds.RStart + ", " + ds.REnd + "]");
            }

            DraftIShapes.Sort((a, b) => a.ContentSelectionStart.CompareTo(b.ContentSelectionStart));
            DraftShapes.Sort((a, b) => a.ContentSelectionStart.CompareTo(b.ContentSelectionStart));

            WCSelection[] NewAArray = WCSelectionOperations.CreateNewArrangedSelectionArray(DraftIShapes.ToArray(), DraftShapes.ToArray());

            //// Logging
            //Logger.AddLog(new LogContent() { LogSubject = "Getting Selection", LogMessage = "Rearranging shape content selections and image content selection...", LogSeverity = LoggerParameters.LogSeverity.INFO });

            //WCSelection[] NewAArray = WCSelectionOperations.CreateNewArrangedSelectionArray(List_IWCSelection.ToArray(), DraftShapes.ToArray());

            //List_IWCSelection.Sort((a, b) => a.ContentSelectionStart.CompareTo(b.ContentSelectionStart));
            // Logging
            Logger.AddLog(new LogContent() { LogSubject = "Getting Selection", LogMessage = "Image content selection sorted according to its start.", LogSeverity = LoggerParameters.LogSeverity.DEBUG });

            for(int ICount = 0; ICount < NewAArray.Length; ICount++)
            {
                NewAArray[ICount].ContentID = "image_" + (ICount + 1);
            }

            MWordDocument.Activate();
            DraftDoc.Close(WdSaveOptions.wdDoNotSaveChanges, WParameters.Missing, WParameters.Missing);
            return NewAArray;
        }
    }
}
