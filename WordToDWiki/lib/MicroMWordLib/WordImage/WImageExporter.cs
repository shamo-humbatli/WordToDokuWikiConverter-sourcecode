using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Microsoft.Office.Interop.Word;
using MicroMWordLib.WordContentSelection;
using MicroMWordLib.WordOperations;
using LittleLyreLogger;
using System.IO;
using System.Drawing;
using System.Drawing.Imaging;
using LittleImage;
namespace MicroMWordLib.WordImage
{
    public class WImageExporter
    {

        public static WImage[] ExportImages(Application MWordApp, Document MWordDocument, string OutputFolder, string ImageFileName, ILittleLyreLogger Logger)
        {
            //// Log info
            //AddToLog("Trying to get inline shapes...)");

            //if (MyDraftDoc.InlineShapes.Count < 1)
            //{
            //    AddToLog("There is no inline shapes in the word document.");
            //}
            //else
            //{
            //    AddToLog("Inline shapes count: " + MyDraftDoc.InlineShapes.Count);
            //}
            try
            {
                // Logging
                Logger.AddLog(new LogContent() { LogSubject = "Image Exporting", LogMessage = "Trying get all image content selections...", LogSeverity = LoggerParameters.LogSeverity.INFO });
                WCSelection[] ImageWCSList;
                ImageWCSList = WImage.GetAllContentSelections(MWordApp, MWordDocument, Logger);
                List<WImage> ImageList = new List<WImage>();

                for (int wcsl = 0; wcsl < ImageWCSList.Length; wcsl++)
                {
                    WImage wimg = new WImage();
                    wimg.ImagePath = OutputFolder + "\\" + ImageFileName + "_" + ImageWCSList[wcsl].ContentID + ".png";
                    wimg.ContentSelection = ImageWCSList[wcsl];
                    ImageList.Add(wimg);
                }

                Document DraftDoc = MWordApp.Documents.Add(Visible: false);
                //Document DraftDoc = WordApp.Documents.Add(WParameters.Missing, WParameters.Missing, WParameters.Missing, false);
                MWordDocument.Select();
                MWordApp.Selection.Copy();
                DraftDoc.Range().Paste();
                DraftDoc.Activate();

                {
                    // Logging
                    Logger.AddLog(new LogContent() { LogSubject = "Image Exporting", LogMessage = "Trying to save each image content selection data as image file.", LogSeverity = LoggerParameters.LogSeverity.INFO });
                    DirectoryInfo OutpFol = new DirectoryInfo(OutputFolder);
                    if(OutpFol.Exists == false)
                    {
                        OutpFol.Create();
                    }

                    for (int isel = 0; isel < ImageList.Count; isel++)
                    {

                         //MWordDocument.Range(ImageList[isel].ContentSelection.ContentSelectionStart, ImageList[isel].ContentSelection.ContentSelectionEnd).Select();

                        MWordApp.Selection.Start = ImageList[isel].ContentSelection.ContentSelectionStart;
                        MWordApp.Selection.End = ImageList[isel].ContentSelection.ContentSelectionEnd;

                        byte[] ImgData = MWordApp.Selection.Range.EnhMetaFileBits;
                        MemoryStream TMStream = new MemoryStream(ImgData);
                        Image ImgFS = Image.FromStream(TMStream);
                        ImgFS.Save(ImageList[isel].ImagePath, ImageFormat.Png);
                        ImgFS.Dispose();

                        // Logging
                        Logger.AddLog(new LogContent() { LogSubject = "Image Exporting", LogMessage = "Image content selection -> [" + ImageList[isel].ContentSelection.ContentSelectionStart + " : " + ImageList[isel].ContentSelection.ContentSelectionEnd + "] was saved as image. Path -> [" + ImageList[isel].ImagePath + "]", LogSeverity = LoggerParameters.LogSeverity.DEBUG });
                    }
                }

                // Logging
                Logger.AddLog(new LogContent() { LogSubject = "Image Exporting", LogMessage = "Preparing images to be cropped -> [Getting all paths]", LogSeverity = LoggerParameters.LogSeverity.INFO });

                string[] ImgPaths = ImageList.Select(a => a.ImagePath).ToArray();
                ImageOperations.RemoveTransparentPartsFromImages(ImgPaths, Logger);
                // Logging
                Logger.AddLog(new LogContent() { LogSubject = "Image Exporting", LogMessage = "All images was cropped.", LogSeverity = LoggerParameters.LogSeverity.INFO });

                DraftDoc.Close(WdSaveOptions.wdDoNotSaveChanges, WParameters.Missing, WParameters.Missing);
                return ImageList.ToArray();
            }
            catch (Exception Exp)
            {
                // Logging
                Logger.AddLog(new LogContent() { LogSubject = "Image Exporting", LogMessage = "Error occured during operation. Error message -> [" + Exp.Message + "]", LogSeverity = LoggerParameters.LogSeverity.ERROR });
                return null;
            }
        }

        public static WImage[] ExportImages(Application MWordApp, Document MWordDocument, WCSelection[] WImageSelections, string OutputFolder, string ImageFileName, ILittleLyreLogger Logger)
        {
            //// Log info
            //AddToLog("Trying to get inline shapes...)");

            //if (MyDraftDoc.InlineShapes.Count < 1)
            //{
            //    AddToLog("There is no inline shapes in the word document.");
            //}
            //else
            //{
            //    AddToLog("Inline shapes count: " + MyDraftDoc.InlineShapes.Count);
            //}
            try
            {
                // Logging
                Logger.AddLog(new LogContent() { LogSubject = "Image Exporting", LogMessage = "Trying get all image content selections...", LogSeverity = LoggerParameters.LogSeverity.INFO });
                List<WImage> ImageList = new List<WImage>();

                for (int wcsl = 0; wcsl < WImageSelections.Length; wcsl++)
                {
                    WImage wimg = new WImage();
                    wimg.ImagePath = OutputFolder + "\\" + ImageFileName + "_" + WImageSelections[wcsl].ContentID + ".png";
                    wimg.ContentSelection = WImageSelections[wcsl];
                    ImageList.Add(wimg);
                }

                Document DraftDoc = MWordApp.Documents.Add(Visible: false);
                //Document DraftDoc = WordApp.Documents.Add(WParameters.Missing, WParameters.Missing, WParameters.Missing, false);
                MWordDocument.Select();
                MWordApp.Selection.Copy();
                DraftDoc.Range().Paste();
                DraftDoc.Activate();

                {
                    // Logging
                    Logger.AddLog(new LogContent() { LogSubject = "Image Exporting", LogMessage = "Trying to save each image content selection data as image file.", LogSeverity = LoggerParameters.LogSeverity.INFO });
                    DirectoryInfo OutpFol = new DirectoryInfo(OutputFolder);
                    if (OutpFol.Exists == false)
                    {
                        OutpFol.Create();
                    }

                    for (int isel = 0; isel < ImageList.Count; isel++)
                    {

                        //MWordDocument.Range(ImageList[isel].ContentSelection.ContentSelectionStart, ImageList[isel].ContentSelection.ContentSelectionEnd).Select();

                        MWordApp.Selection.Start = ImageList[isel].ContentSelection.ContentSelectionStart;
                        MWordApp.Selection.End = ImageList[isel].ContentSelection.ContentSelectionEnd;

                        byte[] ImgData = MWordApp.Selection.Range.EnhMetaFileBits;
                        MemoryStream TMStream = new MemoryStream(ImgData);
                        Image ImgFS = Image.FromStream(TMStream);
                        ImgFS.Save(ImageList[isel].ImagePath, ImageFormat.Png);
                        ImgFS.Dispose();

                        // Logging
                        Logger.AddLog(new LogContent() { LogSubject = "Image Exporting", LogMessage = "Image content selection = [" + ImageList[isel].ContentSelection.ContentSelectionStart + " : " + ImageList[isel].ContentSelection.ContentSelectionEnd + "] save as image [path -> " + ImageList[isel].ImagePath + "]", LogSeverity = LoggerParameters.LogSeverity.DEBUG });
                    }
                }

                // Logging
                Logger.AddLog(new LogContent() { LogSubject = "Image Exporting", LogMessage = "Preparing images to be cropped -> [Getting all paths]", LogSeverity = LoggerParameters.LogSeverity.INFO });

                string[] ImgPaths = ImageList.Select(a => a.ImagePath).ToArray();
                ImageOperations.RemoveTransparentPartsFromImages(ImgPaths, Logger);
                DraftDoc.Close(WdSaveOptions.wdDoNotSaveChanges, WParameters.Missing, WParameters.Missing);
                return ImageList.ToArray();
            }
            catch (Exception Exp)
            {
                // Logging
                Logger.AddLog(new LogContent() { LogSubject = "Image Exporting", LogMessage = "Error occured during operation. Error message -> [" + Exp.Message + "]", LogSeverity = LoggerParameters.LogSeverity.ERROR });
                return null;
            }
        }
    }
}
