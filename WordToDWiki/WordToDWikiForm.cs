using DokuWikiFormatter;
using MicroMWordLib;
using MicroMWordLib.WordOperations;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using LittleLyreLogger;
using System.Diagnostics;
using LittleLyreLogger;

namespace WordToDWiki
{
    public partial class WordToDWikiForm : Form
    {
        public WordToDWikiForm()
        {
            InitializeComponent();
        }


        private List<string> ListOfFiles;
        private OpenFileDialog opdFileOpener;

        private string MyWorkingFolder = "WorkFiles";
        private string DWikiWFolderName = "Untitled";
        private string MyOutputFolder = "Output";
        private string MyConfFolder = "Conf";
        private string MyConfFileName = "WordToDWiki.conf";
        //private FCLib MyWiki = null;

        private Thread ProcessController = null;
        private ILittleLyreLogger MyLogger = null;
        private bool SHMicroMWordLib_JoinImageSelections = false;
        private int LogLCount = 333;
        
        private void frmDokuWikiFC_Load(object sender, EventArgs e)
        {
            // Log
            rtxtLogControl.Text += "-> [Initializing]" + " Trying to load configurations..." + Environment.NewLine;
            ConfigBuilder CBld = new ConfigBuilder();
            try
            {
                object Conf = CBld.BuildConfig(MyConfFolder + "\\" + MyConfFileName);
                if(Conf == null)
                {
                    // Log
                    rtxtLogControl.Text += "-> [Initializing]" + " Something went wront during loading configuration. Check configuration file." + Environment.NewLine;

                    CBld = new ConfigBuilder();

                    // Log
                    rtxtLogControl.Text += "-> [Initializing]" + " Loading default configuration..." + Environment.NewLine;
                }
                MyWorkingFolder = CBld.Directories_WorkingDirectory;
                MyOutputFolder = CBld.Directories_OutputDirectory;
                // Log
                rtxtLogControl.Text += "-> [Initializing]" + " Configuration loaded." + Environment.NewLine;
            }
            catch(Exception Exp)
            {
                // Log
                rtxtLogControl.Text += "-> [Initializing]" + " Error occured during loading configuration(check confiuration file). Message -> " + Exp.Message + Environment.NewLine;

                // Log
                rtxtLogControl.Text += "-> [Initializing]" + " Loading default configuration..." + Environment.NewLine;
                CBld = new ConfigBuilder();

                MyWorkingFolder = CBld.Directories_WorkingDirectory;
                MyOutputFolder = CBld.Directories_OutputDirectory;

                // Log
                rtxtLogControl.Text += "-> [Initializing]" + "Configuration loaded." + Environment.NewLine;
            }


            MyLogger = new LineLogger();
            MyLogger.LogLineCount = CBld.SHLogger_LogLineCount;
            LogLCount = CBld.SHLogger_LogLineCount;
            MyLogger.LogLineStartString = CBld.SHLogger_LogLineStartString;
            MyLogger.LogSubjectLength = CBld.SHLogger_LogSubjetLength;
            MyLogger.LogMessageLength = CBld.SHLogger_LogMessageLength;
            MyLogger.OnLogAdded += GetLog;

            DWFormatter = new DWikiFormatter(MyOutputFolder, MyWorkingFolder);
            DWFormatter.ExportTOC = CBld.DWikiExporter_ExportTOC;
            DWFormatter.ExportTOF = CBld.DWikiExporter_ExportTOF;
            DWFormatter.AddFooterInfo = CBld.DWikiExporter_AddFooterInfo;
            DWFormatter.ListOrder = CBld.DWikiExporter_ListOrder;
            DWFormatter.ShiftOrdering = CBld.DWikiExporter_ShiftOrdering;
            DWFormatter.TableFirstRowIsHeader = CBld.DokuWikiExporter_TableFirstRowIsHeader;
            DWFormatter.ConsiderImageSize = CBld.DokuWikiExporter_ConsiderImageSize;
            SHMicroMWordLib_JoinImageSelections = CBld.SHMicroMWordLib_JoinImageSelections;

            CheckDirectories();

            CheckForIllegalCrossThreadCalls = false;

            ArrangeForm();

            ListOfFiles = new List<string>();
            btnProcessStart.Enabled = false;
            btnAddFile.Enabled = true;

            rtxtLogProcess.BackColor = Color.Black;
            rtxtLogProcess.ForeColor = Color.Lime;

            rtxtLogProcess.ReadOnly = true;
            
            //MyWiki.OnLogAdded += GetLog;

            // Log
            rtxtLogControl.Text += "-> [Initializing]" + " Be sure You have finished and saved your work in 'Microsoft Word'. With starting of exporting process all 'Microsoft Word' documents will be closed immediately." + Environment.NewLine + Environment.NewLine;
        }

        private void GetLog(object sender, object Log)
        {
            if (InvokeRequired)
            {
                BeginInvoke(MyLogger.OnLogAdded, new object[] { sender, Log });
            }
            else
            {
                rtxtLogProcess.Text += Log.ToString() + Environment.NewLine;
                rtxtLogProcess.SelectionStart = rtxtLogProcess.Text.Length - 1;
                rtxtLogProcess.ScrollToCaret();

                if (rtxtLogProcess.Lines.Length > LogLCount)
                {
                    rtxtLogProcess.Clear();
                }
            }
        }


        private void CheckDirectories()
        {
            DirectoryInfo tdi = new DirectoryInfo(Environment.CurrentDirectory + "\\" + MyWorkingFolder);
            if (tdi.Exists == false)
            {
                tdi.Create();
            }
            else
            {
                tdi.Delete(true);
                tdi.Create();
            }

            tdi = new DirectoryInfo(Environment.CurrentDirectory + "\\" + MyOutputFolder);
            if (tdi.Exists == false)
            {
                tdi.Create();
            }
        }


        private void ArrangeForm()
        {
            int margintop = 24;
            int marginleft = 6;
            int marginbetw1 = 3;
            int marginbetw2 = 6;

            lblFiles.Left = marginleft;
            lblFiles.Top = margintop;

            lstListOfFiles.Left = marginleft;
            lstListOfFiles.Top = lblFiles.Bottom + marginbetw1;
            lstListOfFiles.Width = ClientSize.Width - 2 * marginleft - (btnProcessStart.Width + marginbetw2);
            lstListOfFiles.Height = (int)((float)(ClientSize.Height - lblFiles.Top) * 0.2) - marginbetw2 - marginbetw1 - lblFiles.Height;

            btnAddFile.Top = lstListOfFiles.Top;
            btnAddFile.Left = lstListOfFiles.Right + marginbetw2;

            btnProcessStart.Top = btnAddFile.Bottom + marginbetw2;
            btnProcessStart.Left = btnAddFile.Left;

            lblSLog.Left = marginleft;
            lblSLog.Top = lstListOfFiles.Bottom + marginbetw2;

            rtxtLogControl.Left = marginleft;
            rtxtLogControl.Top = lblSLog.Bottom + marginbetw1;
            rtxtLogControl.Width = ClientSize.Width - 2 * marginleft;
            rtxtLogControl.Height = (int)((float)(ClientSize.Height - lblFiles.Top) * 0.4) - marginbetw2 - marginbetw1 - lblSLog.Height;

            lblPLog.Left = marginleft;
            lblPLog.Top = rtxtLogControl.Bottom + marginbetw2;

            rtxtLogProcess.Left = marginleft;
            rtxtLogProcess.Top = lblPLog.Bottom + marginbetw1;
            rtxtLogProcess.Width = ClientSize.Width - 2 * marginleft;
            rtxtLogProcess.Height = (int)((float)(ClientSize.Height - lblFiles.Top) * 0.4) - marginbetw2 - marginbetw1 - lblPLog.Height;

        }

        private void KillWordProcesses()
        {
            try
            {
                rtxtLogControl.Text += "-> [Process Action]" + " Trying to kill word processes..." + Environment.NewLine;
                Process[] AllPrcs = Process.GetProcesses();
                for (int prc = 0; prc < AllPrcs.Length; prc++)
                {
                    if (AllPrcs[prc].ProcessName.ToLower().Contains("winword") == true)
                    {
                        rtxtLogControl.Text += "-> [Process Action]" + " Killing process: [" + AllPrcs[prc].ProcessName + "]" + Environment.NewLine;
                        AllPrcs[prc].Kill();
                        rtxtLogControl.Text += "-> [Process Action]" + " Process killed successfully." + Environment.NewLine;
                    }
                }
            }
            catch(Exception Exp)
            {
                rtxtLogControl.Text += "-> [Process Action]" + " Could not stop running processes of Word. You need to kill them manually." + Environment.NewLine;
                rtxtLogControl.Text += "-> [Process Action]" + " Error occurred. Message -> (" + Exp.Message + ")" + Environment.NewLine;
            }
        }

        private void btnAddFile_Click(object sender, EventArgs e)
        {
            opdFileOpener = new OpenFileDialog();
            opdFileOpener.Filter = "Word files (*.doc, *.docx)|*.doc; *.docx| PDF files (*.pdf) | *.pdf";
            opdFileOpener.Multiselect = true;
            opdFileOpener.Title = "Please select Microsoft Word document(s).";

            if (opdFileOpener.ShowDialog() == DialogResult.OK)
            {
                foreach (string sfl in opdFileOpener.FileNames)
                {
                    ListOfFiles.Add(sfl);
                    lstListOfFiles.Items.Add(Path.GetFileName(sfl));
                }
            }

            if (lstListOfFiles.Items.Count > 0)
            {
                btnProcessStart.Enabled = true;
            }
            else
            {
                btnProcessStart.Enabled = false;
            }
        }

        private bool porcessStarted = false;
        private void btnProcess_Click(object sender, EventArgs e)
        {
            if (porcessStarted == true)
            {
                btnAddFile.Enabled = true;
                ProcessController.Abort();
                btnProcessStart.Text = "Start";
                // Log
                rtxtLogControl.Text += "-> [Process Action]" + " Process aborted." + Environment.NewLine;

                KillWordProcesses();
            }
            else
            {
                btnAddFile.Enabled = false;
                KillWordProcesses();
                ProcessController = new Thread(new ThreadStart(StartProcessing));
                ProcessController.Start();
                porcessStarted = true;
                btnProcessStart.Text = "Cancel";

                // Log
                rtxtLogControl.Text += "-> [Process Action]" + " Process started." + Environment.NewLine;
            }
        }


        private WOperations MyWOperations;
        private DWikiFormatter DWFormatter;

        private void StartProcessing()
        {
            try
            {
                for (int fl = 0; fl < ListOfFiles.Count; fl++)
                {
                    CheckDirectories();

                    // Log
                    rtxtLogControl.Text += "-> [Process Action]" + " Working on: [" + lstListOfFiles.Items[fl] + "]" + Environment.NewLine;
                    MyWOperations = new WOperations(ListOfFiles[fl], MyWorkingFolder);
                    MyWOperations.Logger = MyLogger;
                    MyWOperations.JoinImages = SHMicroMWordLib_JoinImageSelections;

                    // Log
                    rtxtLogControl.Text += "-> [Process Action]" + " Trying to get all word elements for the file " + "[" + lstListOfFiles.Items[fl] + "]" + Environment.NewLine;

                    IWBaseElement[] AllWElements = MyWOperations.GetWordElements();

                    // Log
                    rtxtLogControl.Text += "-> [Process Action]" + " All elements collected for the file " + "[" + lstListOfFiles.Items[fl] + "]" + Environment.NewLine;

                    // Log
                    rtxtLogControl.Text += "-> [Process Action]" + " Trying to export all elements in DokuWiki format for the file " + "[" + lstListOfFiles.Items[fl] + "] ..." + Environment.NewLine;

                    DWikiWFolderName = MyWOperations.ImageFileBaseName;
                    DWFormatter.OriginalInputFileName = lstListOfFiles.Items[fl].ToString();
                    DWFormatter.ExportToDWikiContent(AllWElements, DWikiWFolderName);

                    // Log
                    rtxtLogControl.Text += "-> [Process Action]" + " Exporting has been ended for the file" + "[" + lstListOfFiles.Items[fl] + "]" + Environment.NewLine;

                    // Log
                    rtxtLogControl.Text += "-> [Process Action]" + " Process ended for the file " + "[" + lstListOfFiles.Items[fl] + "]" + Environment.NewLine;

                    // Log
                    if (fl == ListOfFiles.Count - 1)
                    {
                        rtxtLogControl.Text += "-> [Process Action]" + " Process ended.";
                    }
                    rtxtLogControl.Text += Environment.NewLine;

                    DirectoryInfo tmpDInfo = new DirectoryInfo(MyWorkingFolder);
                    if (tmpDInfo.Exists == true)
                    {
                        tmpDInfo.Delete(true);
                    }

                }
            }
            catch (Exception Exp)
            {
                // Log
                rtxtLogControl.Text += "-> [Process Action]" + " Error occured. Message -> [" + Exp.Message + "]" + Environment.NewLine;

                KillWordProcesses();
            }
            finally
            {
                btnProcessStart.Text = "Start";
                porcessStarted = false;
                ProcessController.Abort();                
            }
        }

        private void frmDokuWikiFC_SizeChanged(object sender, EventArgs e)
        {
            ArrangeForm();
        }

        private void frmDokuWikiFC_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (MessageBox.Show("Are you sure to close the program?", "Warning", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
            {
                if (ProcessController != null && (ProcessController.ThreadState == System.Threading.ThreadState.Running || ProcessController.ThreadState == System.Threading.ThreadState.Suspended))
                {
                    ProcessController.Abort();
                }

                KillWordProcesses();
            }
            else
            {
                e.Cancel = true;
            }
        }

        private void cmsi_RSItem_Click(object sender, EventArgs e)
        {
            ListOfFiles.RemoveAt(lstListOfFiles.SelectedIndex);
            lstListOfFiles.Items.RemoveAt(lstListOfFiles.SelectedIndex);

            if(lstListOfFiles.Items.Count < 1)
            {
                btnProcessStart.Enabled = false;
            }
        }


        private void cms_FileList_Opening(object sender, CancelEventArgs e)
        {
            if(porcessStarted == true)
            {
                cmsi_RSItem.Enabled = false;
                cmsi_RAItems.Enabled = false;
            }
            else
            {
                if (lstListOfFiles.Items.Count < 1)
                {
                    cmsi_RSItem.Enabled = false;
                    cmsi_RAItems.Enabled = false;
                }
                else
                {
                    cmsi_RSItem.Enabled = true;
                    cmsi_RAItems.Enabled = true;
                }

                if (lstListOfFiles.SelectedIndex < 0)
                {
                    cmsi_RSItem.Enabled = false;
                }
                else
                {
                    cmsi_RSItem.Enabled = true;
                }

            }
        }

        private void cmsi_RAItems_Click(object sender, EventArgs e)
        {
            lstListOfFiles.Items.Clear();
            ListOfFiles.Clear();
            btnProcessStart.Enabled = false;
        }

        //private bool ExportWordToDWiki()
        //{
        //    try
        //    {

        //    }
        //    catch
        //    {
        //        return false;
        //    }

        //    return true;
        //}

        
        private void tsmiAbout_Click(object sender, EventArgs e)
        {
            AboutForm Ab = new AboutForm();
            Ab.ShowDialog();
        }
    }
}
