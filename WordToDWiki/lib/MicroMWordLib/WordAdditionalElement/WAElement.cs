﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using LittleLyreLogger;
using Microsoft.Office.Interop.Word;
using MicroMWordLib.WordContentSelection;
using MicroMWordLib.WordOperations;

namespace MicroMWordLib.WordAdditionalElement
{
    public class WAElement : IWBaseElement
    {

        public enum WAElementType
        {
            TableOfContents = 0,
            TableOfFigures = 1
        }

        private List<WAElementLine> prp_Lines;
        private WCSelection prp_ContentSelection;
        private WAElementType prp_WAElementType;

        public List<WAElementLine> Lines { get => prp_Lines; set => prp_Lines = value; }
        public WCSelection ContentSelection { get => prp_ContentSelection; set => prp_ContentSelection = value; }
        public WAElementType ElementType { get => prp_WAElementType; set => prp_WAElementType = value; }

        public WAElement()
        {
            prp_Lines = new List<WAElementLine>();
        }

        public static WCSelection[] GetAllContentSelections(Application MWordApp, Document MWordDocument, ILittleLyreLogger Logger)
        {
            // Log
            Logger.AddLog(new LogContent() { LogSubject = "Getting Selection", LogMessage = "Trying to get all table of content selections..." });
            Document DraftDoc = MWordApp.Documents.Add();
            MWordDocument.Select();
            MWordApp.Selection.Copy();
            DraftDoc.Range().Paste();
            DraftDoc.Activate();

            List<WCSelection> WAElementSelection = new List<WCSelection>();

            for (int tcnt = 1; tcnt <= DraftDoc.TablesOfContents.Count; tcnt++)
            {
                TableOfContents toc = DraftDoc.TablesOfContents[tcnt];
                toc.Range.Select();
                WCSelection wcs = new WCSelection();
                wcs.ContentSelectionStart = MWordApp.Selection.Start;
                wcs.ContentSelectionEnd = MWordApp.Selection.End;
                WAElementSelection.Add(wcs);

                // Log
                Logger.AddLog(new LogContent() { LogSubject = "Getting Selection", LogMessage = "Table Of Contents [" + tcnt + "/" + DraftDoc.TablesOfContents.Count + "] added. Content selection -> [" + wcs.ContentSelectionStart + " : " + wcs.ContentSelectionEnd + "]", LogSeverity = LoggerParameters.LogSeverity.DEBUG });
            }

            for (int tfgr = 1; tfgr <= DraftDoc.TablesOfFigures.Count; tfgr++)
            {
                TableOfFigures tof = DraftDoc.TablesOfFigures[tfgr];
                tof.Range.Select();
                WCSelection wcs = new WCSelection();
                wcs.ContentSelectionStart = MWordApp.Selection.Start;
                wcs.ContentSelectionEnd = MWordApp.Selection.End;
                WAElementSelection.Add(wcs);

                // Log
                Logger.AddLog(new LogContent() { LogSubject = "Getting Selection", LogMessage = "Table Of Figures [" + tfgr + "/" + DraftDoc.TablesOfFigures.Count + "] added. Content selection -> [" + wcs.ContentSelectionStart + " : " + wcs.ContentSelectionEnd + "]", LogSeverity = LoggerParameters.LogSeverity.DEBUG });
            }

            {
                for (int ls = 0; ls < WAElementSelection.Count; ls++)
                {
                    WAElementSelection[ls].ContentID = "WAElement_" + (ls + 1);
                }
            }

            MWordDocument.Activate();
            DraftDoc.Close(WdSaveOptions.wdDoNotSaveChanges, WParameters.Missing, WParameters.Missing);

            return WAElementSelection.ToArray();
        }
    }
}
