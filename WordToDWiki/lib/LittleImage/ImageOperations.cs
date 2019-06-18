using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using LittleLyreLogger;

namespace LittleImage
{
    public class ImageOperations
    {
        public static void RemoveTransparentPartsFromImages(string[] ImagePaths, ILittleLyreLogger Logger)
        {
            try
            {
                // Logging
                Logger.AddLog(new LogContent() { LogSubject = "Image Operation", LogMessage = "Trying to remove transparent parts from images...", LogSeverity = LoggerParameters.LogSeverity.INFO });

                if (ImagePaths == null)
                {
                    // Logging
                    Logger.AddLog(new LogContent() { LogSubject = "Image Operation", LogMessage = "Input is [null]. Process cancelled.", LogSeverity = LoggerParameters.LogSeverity.INFO });

                    return;
                }

                int CropY1 = 0, CropX1 = 0, CropY2 = 0, CropX2 = 0;
                bool RFound = false;
                Bitmap mybmp = null;

                // Logging
                Logger.AddLog(new LogContent() { LogSubject = "Image Operation", LogMessage = "Image path count: " + ImagePaths.Length, LogSeverity = LoggerParameters.LogSeverity.INFO });

                for (int li = 0; li < ImagePaths.Length; li++)
                {
                    mybmp = new Bitmap(ImagePaths[li], true);
                    CropY1 = 0;
                    CropX1 = 0;
                    CropY2 = mybmp.Height;
                    CropX2 = mybmp.Width;

                    // Logging
                    Logger.AddLog(new LogContent() { LogSubject = "Image Operation", LogMessage = "Image [" + (li + 1) + "/" + ImagePaths.Length + "] Size -> [" + CropX2 + " : " + CropY2 + "] Path -> [ " + ImagePaths[li] + "]", LogSeverity = LoggerParameters.LogSeverity.DEBUG });

                    //Log info
                    //AddToLog("Intial image [width, height] = " + mybmp.Width + ", " + mybmp.Height);

                    RFound = false;
                    for (int h = 0; h < mybmp.Height; h++)
                    {
                        for (int w = 0; w < mybmp.Width; w++)
                        {
                            if (mybmp.GetPixel(w, h).A != 0)
                            {
                                RFound = true;
                                break;
                            }
                        }
                        if (RFound == false)
                        {
                            CropY1 = h;
                        }
                        else
                        {
                            break;
                        }
                    }

                    RFound = false;
                    for (int h = (mybmp.Height - 1); h >= 0; h--)
                    {
                        for (int w = 0; w < mybmp.Width; w++)
                        {
                            if (mybmp.GetPixel(w, h).A != 0)
                            {
                                RFound = true;
                                break;
                            }
                        }
                        if (RFound == false)
                        {
                            CropY2 = h;
                        }
                        else
                        {
                            break;
                        }
                    }

                    RFound = false;
                    for (int w = (mybmp.Width - 1); w >= 0; w--)
                    {
                        for (int h = 0; h < mybmp.Height; h++)
                        {
                            if (mybmp.GetPixel(w, h).A != 0)
                            {
                                RFound = true;
                                break;
                            }
                        }
                        if (RFound == false)
                        {
                            CropX2 = w;
                        }
                        else
                        {
                            break;
                        }
                    }

                    RFound = false;
                    for (int w = 0; w < mybmp.Width; w++)
                    {
                        for (int h = 0; h < mybmp.Height; h++)
                        {
                            if (mybmp.GetPixel(w, h).A != 0)
                            {
                                RFound = true;
                                break;
                            }
                        }

                        if (RFound == false)
                        {
                            CropX1 = w;
                        }
                        else
                        {
                            break;
                        }
                    }

                    mybmp.Dispose();
                    mybmp = new Bitmap(CropX2 - CropX1, CropY2 - CropY1);

                    Graphics Grp = Graphics.FromImage(mybmp);

                    Grp.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.Bicubic;
                    Grp.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias;

                    Image TImg = Image.FromFile(ImagePaths[li]);
                    Grp.DrawImage(TImg, new RectangleF(0, 0, mybmp.Width, mybmp.Height), new RectangleF(CropX1, CropY1, CropX2 - CropX1, CropY2 - CropY1), GraphicsUnit.Pixel);
                    Grp.Dispose();
                    TImg.Dispose();

                    mybmp.Save(ImagePaths[li], ImageFormat.Png);

                    // Logging
                    Logger.AddLog(new LogContent() { LogSubject = "Image Operation", LogMessage = "Image [" + (li + 1) + "/" + ImagePaths.Length + "] cropped and saved. New Size -> [" + mybmp.Width + " : " + mybmp.Height + "]", LogSeverity = LoggerParameters.LogSeverity.DEBUG });
                }
            }
            catch(Exception Exp)
            {
                // Logging
                Logger.AddLog(new LogContent() { LogSubject = "Image Operation", LogMessage = "Error occured during operation. Error message -> [" + Exp.Message + "]", LogSeverity = LoggerParameters.LogSeverity.ERROR });
            }
        }
    }
}
