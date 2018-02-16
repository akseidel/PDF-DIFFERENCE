using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using System.Diagnostics;
using System.IO;
using System.Printing;
using System.Runtime.InteropServices;
using System.Drawing.Printing;
using iTextSharp.text.pdf;

namespace PDF_DIFFERENCE
{
    public partial class PDF_DIFFERENCE : Form
    {
        #region Declarations
        //private string gsEXELocation = "C:\\Program Files\\gs\\gs9.16\\bin\\gswin64c.exe";
        private string IMconvertExeLocation = "";
        private string usersPrinterName;
        private string restoreOnExitToThisPrntName;
        private string psize = null;
        DataTable DTablePDFsOlder = new DataTable();
        BindingSource SBindPDFsOlder = new BindingSource();

        DataTable DTablePDFsNewer = new DataTable();
        BindingSource SBindPDFsNewer = new BindingSource();

        DataTable DTablePNGResults = new DataTable();
        BindingSource SBindPNGResults = new BindingSource();

        //   string msgFileMissing = "File Is Missing!";
        string msgInProcess = "In Process";
        string msgDone = "Done";
        string msgInQueued = "Queued";
        //  string msgNotProcessed = "Not Proccessed";
        private System.Drawing.Rectangle dragBoxFromMouseDown;   //keep
        private int rowIndexFromMouseDown;                       //keep
        private int rowIndexOfItemUnderMouseToDrop;
        private bool inDragInto = false;
        private string dftTitle = "PDF Difference";
        private string printIt = "Start";

        private string HeadStrForComposite = "Composite_";
        //private bool troubleShoot = true;
        private bool troubleShoot = false;
        private string beVerbose = "-verbose ";
        private string densityVal = "200x200 ";

        #endregion

        public PDF_DIFFERENCE()
        {
            InitializeComponent();

            //gsLocation.GS_EXE_Location = gsEXELocation;
        }

        private void PDF_DIFFERENCE_Load(object sender, EventArgs e)
        {
            GetDefaultPrinter(true);  // true means save this name to startingPrntName
            BindDataToGrid();
            // add a RowDeleted event handler for the table.
            DTablePDFsOlder.RowDeleted += new DataRowChangeEventHandler(DTRowOlder_Deleted);
            DTablePDFsNewer.RowDeleted += new DataRowChangeEventHandler(DTRowNewer_Deleted);
            // add a ChangedRow event handler for the table.
            DTablePDFsOlder.RowChanged += new DataRowChangeEventHandler(DTRow_ChngRow);
            DTablePDFsNewer.RowChanged += new DataRowChangeEventHandler(DTRow_ChngRow);
            labelPDFStatsOlder.Text = "";
            labelPDFStatsNewer.Text = "";
            progressBarPDFDiff.Visible = false;
            progressBarPrint.Visible = false;
            SetRadio();
            buttonPrintCancel.Visible = false;
            buttonCancelDifference.Visible = false;

            if (!CheckForHelpers())
            {
                string msg = "The PDF Difference process rquires both ImageMagick and GhostScript, which were not both found on this computer.";
                msg = msg + " This program is not going to work.";
                MessageBox.Show(msg, "Critical Helpers Not Found!");
            }
          
        }

        private bool CheckForHelpers()
        {
            Helpers hpr = new Helpers();
            string GS_Loc = hpr.WhatIsGSExecutable();
            string MGK_Loc = hpr.WhatIsMagick();
            IMconvertExeLocation = MGK_Loc;
            string helpStat = "ImageMagick: " + MGK_Loc;
            helpStat = helpStat + "\n\n" + "GhostScript: " + GS_Loc;
            labelMessage.Text = helpStat;
            if (helpStat.Contains("NOT FOUND")) { return false; }
            return true;
        }

        private void SetRadio()
        {
            radioButtonLetter.Checked = Properties.Settings.Default.letterChecked;
            radioButtonLedger.Checked = Properties.Settings.Default.ledgerChecked;
        }

        private void BindDataToGrid()
        {
            // older PDF table
            DTablePDFsOlder.Columns.Add(new DataColumn("PDFName", typeof(string)));
            DTablePDFsOlder.Columns.Add(new DataColumn("Status", typeof(string)));
            DTablePDFsOlder.Columns.Add(new DataColumn("Pages", typeof(int)));
            SBindPDFsOlder.DataSource = DTablePDFsOlder;
            dataGridViewPDFSOlder.AutoGenerateColumns = false;
            dataGridViewPDFSOlder.DataSource = SBindPDFsOlder;
            // newer PDF table
            DTablePDFsNewer.Columns.Add(new DataColumn("PDFName", typeof(string)));
            DTablePDFsNewer.Columns.Add(new DataColumn("Status", typeof(string)));
            DTablePDFsNewer.Columns.Add(new DataColumn("Pages", typeof(int)));
            SBindPDFsNewer.DataSource = DTablePDFsNewer;
            dataGridViewPDFSNewer.AutoGenerateColumns = false;
            dataGridViewPDFSNewer.DataSource = SBindPDFsNewer;
            // PNG Results
            DTablePNGResults.Columns.Add(new DataColumn("PNGName", typeof(string)));
            SBindPNGResults.DataSource = DTablePNGResults;
            dataGridViewResults.AutoGenerateColumns = false;
            dataGridViewResults.DataSource = SBindPNGResults;
        }

        private void DTRowOlder_Deleted(object sender, DataRowChangeEventArgs e)
        {
            PageCount();
            if (DTablePDFsOlder.Rows.Count > 0)
            {
                buttonStartDifference.Enabled = true;
                return;
            }
            else
            {
                buttonStartDifference.Enabled = false;
            }
            ZapThumbnail(pictureBoxThumbOlder);
            labelPDFStatsOlder.Text = "";
            EnableProcessIfCan();
        }

        private void DTRowNewer_Deleted(object sender, DataRowChangeEventArgs e)
        {
            PageCount();
            if (DTablePDFsNewer.Rows.Count > 0)
            {
                buttonStartDifference.Enabled = true;
                return;
            }
            else
            {
                buttonStartDifference.Enabled = false;
            }
            ZapThumbnail(pictureBoxThumbNewer);
            labelPDFStatsNewer.Text = "";
            EnableProcessIfCan();
        }

        private void DTRow_ChngRow(object sender, DataRowChangeEventArgs e)
        {
            EnableProcessIfCan();
        }

        private void ZapThumbnail(PictureBox pb)
        {
            if (pb.Image != null)
            {
                pb.Image.Dispose();
                pb.Image = null;
            }
        }

        private void EnableProcessIfCan()
        {
            if (DTablePDFsOlder.Rows.Count != DTablePDFsNewer.Rows.Count)
            {
                buttonStartDifference.Enabled = false;
                progressBarPDFDiff.Visible = false;
                return;
            }
            buttonStartDifference.Enabled = true;
        }

        private string GetDefaultPrinter(bool SaveThisForRestoring = false)
        {
            string currentWindowsPName = LocalPrintServer.GetDefaultPrintQueue().FullName;
            if (SaveThisForRestoring) { restoreOnExitToThisPrntName = currentWindowsPName; }
            string tempUsersPN = Properties.Settings.Default.usersPrinterChoice;
            if (tempUsersPN != "")
            {
                // There is a user printer name, see if it is ok. 
                bool haveprinter = false;
                foreach (string printer in PrinterSettings.InstalledPrinters)
                {
                    if (tempUsersPN.Equals(printer, StringComparison.CurrentCultureIgnoreCase))
                    {
                        haveprinter = true;
                        usersPrinterName = tempUsersPN;
                        break;
                    }
                }
                if (!haveprinter) { usersPrinterName = currentWindowsPName; } // set to current printer
            }
            else
            {
                usersPrinterName = currentWindowsPName;
            }
            labelDfltPrnt.Text = usersPrinterName;
            return usersPrinterName;
        }

        private void SetDefaultPrinterToThis(string printerN)
        {
            myPrinters.SetDefaultPrinter(printerN);
        }

        private void PDF_DIFFERENCE_FormClosing(object sender, FormClosingEventArgs e)
        {
            SetDefaultPrinterToThis(restoreOnExitToThisPrntName);
            Properties.Settings.Default.usersPrinterChoice = usersPrinterName;
            Properties.Settings.Default.letterChecked = radioButtonLetter.Checked;
            Properties.Settings.Default.ledgerChecked = radioButtonLedger.Checked;
            Properties.Settings.Default.Save();
        }

        private void UpdatePNGResultsTable()
        {
            // will always set the path from the new table
            if (DTablePDFsNewer.Rows.Count > 0)
            {
                DTablePNGResults.Clear();
                try
                {
                    DataRow dr0 = DTablePDFsNewer.Rows[0];
                    if (dr0.RowState != DataRowState.Detached)
                    {
                        string TheFirstNewerFilePName = dr0["PDFName"].ToString();

                        IEnumerable<String> PNGResultsList = Directory.GetFiles(Path.GetDirectoryName(TheFirstNewerFilePName),
                                        "*.pdf", SearchOption.TopDirectoryOnly);

                        foreach (string s in PNGResultsList)
                        {
                            if (Path.GetFileName(s).StartsWith(HeadStrForComposite))
                            {
                                DataRow dgRow = DTablePNGResults.NewRow();
                                dgRow["PNGName"] = Path.GetFileName(s);
                                DTablePNGResults.Rows.Add(dgRow);
                            }
                        }
                    }
                }
                catch (Exception) { }
            }
            ShowPNGResultsImageForSelection(dataGridViewResults, pictureBoxComposites);
        }

        private void buttonGetDefaultPrinter_Click(object sender, EventArgs e)
        {
            usersPrinterName = LocalPrintServer.GetDefaultPrintQueue().FullName;
            labelDfltPrnt.Text = usersPrinterName;
        }

        private void buttonStartDifference_Click(object sender, EventArgs e)
        {
            int pcnt = PageCount();
            if (pcnt > 10)
            {
                string msg = "Please confirm. PDFDifference these " + pcnt.ToString() + " ?";
                DialogResult dr = MessageBox.Show(msg, "Just Checking", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
                if (dr != DialogResult.OK) { return; }
            }
            buttonStartDifference.Refresh();
            Text = dftTitle + "   " + "<< PDFDifference is ongoing. >>";
            labelMessage.Text = "Starting composite PDFDifference session";
            labelMessage.Refresh();
            LogMessage();
            PDFDifferenceTheList();
        }

        private void PDFDifferenceTheList()
        {
            ///  We are passing the lists instead of each file.
            List<string> PDFOlderFileList = new List<string>();
            List<string> PDFNewerFileList = new List<string>();

            foreach (DataRow dr in DTablePDFsOlder.Rows)
            {
                if (dr["PDFName"] == null) { continue; }
                String pdfName = dr["PDFName"].ToString();
                if (pdfName.Length == 0) { continue; }
                if (!File.Exists(pdfName)) { continue; }
                PDFOlderFileList.Add(pdfName);
                dr["Status"] = msgInQueued;
            }

            foreach (DataRow dr in DTablePDFsNewer.Rows)
            {
                if (dr["PDFName"] == null) { continue; }
                String pdfName = dr["PDFName"].ToString();
                if (pdfName.Length == 0) { continue; }
                if (!File.Exists(pdfName)) { continue; }
                PDFNewerFileList.Add(pdfName);
                dr["Status"] = msgInQueued;
            }

            // erase any existing logfile
            if (PDFNewerFileList.Count > 0) {
                string thisLog = LogFile(true);
            }

            psize = ThePSize();
            if (troubleShoot) { beVerbose = "-verbose "; } else { beVerbose = " "; }

            /// THIS DOES WORK RIGHT   DO NOT ERASE!!!!
            /// convert -density 150x150 -fill red -opaque black -fuzz 75% +antialias WAS_NEW.pdf Zback%02d.png
            /// convert -density 150x150 -transparent white +antialias WAS_ORG.pdf Zfront%02d.png
            /// composite Zfront00.png ZBACK00.png Zresult00.png
            /// convert Zresult00.png  -background white  -alpha remove  -alpha off  Zwresult00.png

            progressBarPDFDiff.Maximum = 5 * PDFOlderFileList.Count();
            IMArgs thisGSArg = new IMArgs
            {
                PDFOlderFileNameList = PDFOlderFileList,
                PDFNewerFileNameList = PDFNewerFileList,
                IMconvertExeLocation = IMconvertExeLocation
            };

            if (!backgroundWorkerIM.IsBusy)
            {
                buttonCancelDifference.Visible = true;
                backgroundWorkerIM.RunWorkerAsync(thisGSArg);
            }
        }

        private void dataGridViewPDFS_DragOver(object sender, DragEventArgs e)
        {
            if (!inDragInto)
            {
                e.Effect = DragDropEffects.Move;
            }
        }

        private void dataGridViewPDFS_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                inDragInto = true;
                e.Effect = DragDropEffects.Copy;
            }
        }

        private void dataGridViewPDFSOlder_DragDrop(object sender, DragEventArgs e)
        {
            // If the drag operation is a copy then copy.
            if (inDragInto)
            {
                //if (e.Effect == DragDropEffects.Copy) {
                string[] filesAry = (string[])e.Data.GetData(DataFormats.FileDrop);
                foreach (string file in filesAry)
                {
                    if (!Path.GetExtension(file).Equals(".pdf", StringComparison.CurrentCultureIgnoreCase)) { continue; };
                    bool isAlready = DTablePDFsOlder.AsEnumerable().Any(row => file == row.Field<String>("PDFName"));
                    if (!isAlready)
                    {
                        DataRow dgRow = DTablePDFsOlder.NewRow();
                        dgRow["PDFName"] = file;
                        dgRow["Status"] = "";
                        dgRow["Pages"] = PDFpgCount(file);
                        DTablePDFsOlder.Rows.Add(dgRow);
                    }
                }
                PageCount();
                inDragInto = false;
            }
            else
            {
                // If the drag operation was a move then remove and insert the row.
                if (e.Effect == DragDropEffects.Move)
                {
                    // The mouse locations are relative to the screen, so they must be 
                    // converted to client coordinates.
                    Point clientPoint = dataGridViewPDFSOlder.PointToClient(new Point(e.X, e.Y));
                    // Get the row index of the item the mouse is below. 
                    rowIndexOfItemUnderMouseToDrop = dataGridViewPDFSOlder.HitTest(clientPoint.X, clientPoint.Y).RowIndex;
                    if (rowIndexOfItemUnderMouseToDrop < 0) { rowIndexOfItemUnderMouseToDrop = 0; }
                    DataGridViewRow rowToMove = e.Data.GetData(typeof(DataGridViewRow)) as DataGridViewRow;
                    DataRow dtRowBeingMoved = DTablePDFsOlder.Rows[rowToMove.Index];
                    DataRow dgRow = DTablePDFsOlder.NewRow();
                    dgRow["PDFName"] = dtRowBeingMoved["PDFName"];
                    dgRow["Status"] = dtRowBeingMoved["Status"];
                    dgRow["Pages"] = dtRowBeingMoved["Pages"];
                    DTablePDFsOlder.Rows.InsertAt(dgRow, rowIndexOfItemUnderMouseToDrop);
                    DTablePDFsOlder.Rows.Remove(dtRowBeingMoved);
                }
            }
        }

        private void dataGridViewPDFSNewer_DragDrop(object sender, DragEventArgs e)
        {
            // If the drag operation is a copy then copy.
            if (inDragInto)
            {
                //if (e.Effect == DragDropEffects.Copy) {
                string[] filesAry = (string[])e.Data.GetData(DataFormats.FileDrop);
                foreach (string file in filesAry)
                {
                    if (!Path.GetExtension(file).Equals(".pdf", StringComparison.CurrentCultureIgnoreCase)) { continue; };
                    bool isAlready = DTablePDFsNewer.AsEnumerable().Any(row => file == row.Field<String>("PDFName"));
                    if (!isAlready)
                    {
                        DataRow dgRow = DTablePDFsNewer.NewRow();
                        dgRow["PDFName"] = file;
                        dgRow["Status"] = "";
                        dgRow["Pages"] = PDFpgCount(file);
                        DTablePDFsNewer.Rows.Add(dgRow);
                    }
                }
                PageCount();
                inDragInto = false;
            }
            else
            {
                // If the drag operation was a move then remove and insert the row.
                if (e.Effect == DragDropEffects.Move)
                {
                    // The mouse locations are relative to the screen, so they must be 
                    // converted to client coordinates.
                    Point clientPoint = dataGridViewPDFSNewer.PointToClient(new Point(e.X, e.Y));
                    // Get the row index of the item the mouse is below. 
                    rowIndexOfItemUnderMouseToDrop = dataGridViewPDFSNewer.HitTest(clientPoint.X, clientPoint.Y).RowIndex;
                    if (rowIndexOfItemUnderMouseToDrop < 0) { rowIndexOfItemUnderMouseToDrop = 0; }
                    DataGridViewRow rowToMove = e.Data.GetData(typeof(DataGridViewRow)) as DataGridViewRow;
                    DataRow dtRowBeingMoved = DTablePDFsNewer.Rows[rowToMove.Index];
                    DataRow dgRow = DTablePDFsNewer.NewRow();
                    dgRow["PDFName"] = dtRowBeingMoved["PDFName"];
                    dgRow["Status"] = dtRowBeingMoved["Status"];
                    dgRow["Pages"] = dtRowBeingMoved["Pages"];
                    DTablePDFsNewer.Rows.InsertAt(dgRow, rowIndexOfItemUnderMouseToDrop);
                    DTablePDFsNewer.Rows.Remove(dtRowBeingMoved);
                }
            }
        }

        private string ThePSize()
        {
            if (radioButtonLetter.Checked) { return "letter"; }
            if (radioButtonLedger.Checked) { return "ledger"; }
            return "letter";
        }

        private void backgroundWorkerIM_DoWork(object sender, DoWorkEventArgs e)
        {
            bool cleanUpTheFiles = checkBoxCleanUp.Checked;
            IMArgs thisArgs = e.Argument as IMArgs;
            string convertExe = thisArgs.IMconvertExeLocation.ToString(); //  String.Concat("\"", thisArgs.IMconvertExeLocation, "\"");
            List<string> PDFOlderFileNameListToDo = thisArgs.PDFOlderFileNameList;
            List<string> PDFNewerFileNameListToDo = thisArgs.PDFNewerFileNameList;
            int inDx = 0;
            foreach (string olderfName in PDFOlderFileNameListToDo)
            {
                // Comparing the newer file to the older file
                string olderfilePathName = string.Concat(olderfName);
                string newerfilePathName = String.Concat(PDFNewerFileNameListToDo[inDx]);
                string olderFN = Path.GetFileNameWithoutExtension(olderfilePathName);
                string newerFN = Path.GetFileNameWithoutExtension(newerfilePathName);
                int imagePageCNT = PDFpgCount(olderfilePathName);

                // The older file goes in front of the newer one when combined in the compostite process.
                string frontPNGname = String.Concat(Path.GetDirectoryName(olderfName), "\\front_", olderFN, ".png");
                string frontPNGnamePat = String.Concat("front_", olderFN, "*.png");
                // The newer file will be in back of the newer one when combined in the composite process.
                string backPNGname = String.Concat(Path.GetDirectoryName(PDFNewerFileNameListToDo[inDx]), "\\back_", newerFN, ".png");
                string backPNGnamePat = String.Concat("back_", newerFN, "*.png");


                string blackFuzz = String.Concat("-fuzz ",maskedTextBoxBlackFuzz.Text," ");
                string whiteFuzz = String.Concat("-fuzz ", maskedTextBoxWhiteFuzz.Text, " ");
                // The recent issue.
                var processWas_NewArgs = String.Concat(" convert -density ",
                                                 densityVal,
                                                 "-fill red ",
                                                 "-opaque black ",
                                                 blackFuzz,
                                                 "+antialias ",
                                                 beVerbose,
                                                 @String.Concat("\"",newerfilePathName,"\""),
                                                 " ",
                                                 @String.Concat("\"",backPNGname, "\"")
                                                 );
                // The former issue.
                var processWas_OrgArgs = String.Concat(" convert -density ",
                                                 densityVal,
                                                 "-transparent white ",
                                                 whiteFuzz,
                                                 "+antialias ",
                                                 beVerbose,
                                                 @String.Concat("\"", olderfilePathName,"\""),
                                                 " ",
                                                 @String.Concat("\"", frontPNGname, "\"")
                                                 );

                try
                {
                    var IMBackProcessInfo = new ProcessStartInfo
                    {
                        FileName = Path.GetFileName(convertExe),
                        WorkingDirectory = Path.GetDirectoryName(olderfilePathName),
                        Arguments = processWas_NewArgs,
                    };
                    if (!troubleShoot) { IMBackProcessInfo.WindowStyle = ProcessWindowStyle.Hidden; }

                    var IMFrontProcessInfo = new ProcessStartInfo
                    {
                        FileName = convertExe,
                        Arguments = processWas_OrgArgs,
                    };
                    if (!troubleShoot) { IMFrontProcessInfo.WindowStyle = ProcessWindowStyle.Hidden; }

                    //string msg = thisArgs.IMconvertExeLocation + processArgsOlder;
                    //msg = msg + "\n\n" + thisArgs.IMconvertExeLocation + processArgsNewer;
                    //msg = msg + "\n\n" + thisArgs.IMconvertExeLocation + processComposite;
                    //msg = msg + "\n\n" + thisArgs.IMconvertExeLocation + processFinishing;
                    //MessageBox.Show(msg);

                    string progMsg = String.Empty;

                    #region Process the back image

                    #region Set message type for page quantity
                    /// reporting when item starts older
                    if (imagePageCNT == 1)
                    {
                        progMsg = Path.GetFileName(backPNGname);
                    }
                    else
                    {
                        progMsg = backPNGnamePat;
                    }
                    #endregion

                    backgroundWorkerIM.ReportProgress(inDx * 4 + inDx, progMsg + "\n\n" + processWas_NewArgs.Trim());
                    using (var IMbackProcess = Process.Start(IMBackProcessInfo))
                    {
                        IMbackProcess.WaitForExit();
                    }
                    #endregion

                    #region Process the front image

                    //if cancellation is pending, cancel work.  
                    if (backgroundWorkerIM.CancellationPending) { e.Cancel = true; break; }

                    #region Set prog message for page count
                    /// reporting when item starts newer
                    if (imagePageCNT == 1)
                    {
                        progMsg = Path.GetFileName(frontPNGname);
                    }
                    else
                    {
                        progMsg = frontPNGnamePat;
                    }
                    #endregion

                    backgroundWorkerIM.ReportProgress(inDx * 4 + inDx + 1, progMsg + "\n\n" + processWas_OrgArgs.Trim());
                    using (var IMfrontProcess = Process.Start(IMFrontProcessInfo))
                    {
                        IMfrontProcess.WaitForExit();
                    }
                    #endregion

                    //if cancellation is pending, cancel work.  
                    if (backgroundWorkerIM.CancellationPending) { e.Cancel = true; break; }

                    #region The composite section
                    /// The composite processes needs to loop through the pdf pages
                    string compositePNGname = null;
                    string compositePDFname = null;
                    var processPNGComposite = String.Empty;
                    var processPNGFinishing = String.Empty;
                    var processPDFFinishing = String.Empty;

                    #region LoopThroughEachPage
                    for (int i = 1; i <= imagePageCNT; i++)
                    {
                        if (imagePageCNT == 1 && i == 1)
                        {
                            compositePNGname = String.Concat(Path.GetDirectoryName(PDFNewerFileNameListToDo[inDx]),
                                                            "\\", HeadStrForComposite, newerFN, "_and_", olderFN, ".png");
                            compositePDFname = String.Concat(Path.GetDirectoryName(PDFNewerFileNameListToDo[inDx]),
                                                            "\\", HeadStrForComposite, newerFN, "_and_", olderFN, ".pdf");
                        }
                        else // nnames need to be reconstructed based on the page number
                        {
                            compositePNGname = String.Concat(Path.GetDirectoryName(PDFNewerFileNameListToDo[inDx]),
                                                       "\\", HeadStrForComposite, newerFN, "_and_", olderFN, "-", (i - 1).ToString(), ".png");
                            compositePDFname = String.Concat(Path.GetDirectoryName(PDFNewerFileNameListToDo[inDx]),
                                                       "\\", HeadStrForComposite, newerFN, "_and_", olderFN, "-", (i - 1).ToString(), ".pdf");
                            frontPNGname = String.Concat(Path.GetDirectoryName(olderfName), "\\front_", olderFN, "-", (i - 1).ToString(), ".png");
                            backPNGname = String.Concat(Path.GetDirectoryName(PDFNewerFileNameListToDo[inDx]), "\\back_", newerFN, "-", (i - 1).ToString(), ".png");
                        }

                        processPNGComposite = String.Concat(" composite ",
                                           beVerbose,
                                           @String.Concat("\"", frontPNGname,"\""),
                                           " ",
                                           @String.Concat("\"", backPNGname,"\""),
                                           " ",
                                           @String.Concat("\"", compositePNGname,"\"")
                                            );

                        processPNGFinishing = String.Concat(" convert ",
                                                   @String.Concat("\"", compositePNGname, "\""),
                                                   " ",
                                                   "-background white ",
                                                   "-alpha remove ",
                                                   "-alpha off ",
                                                   beVerbose,
                                                   @String.Concat("\"", compositePNGname, "\"")
                                                    );

                        processPDFFinishing = String.Concat(" convert ",
                                                   beVerbose,
                                                   @String.Concat("\"", compositePNGname, "\""),
                                                   " ",
                                                   @String.Concat("\"",compositePDFname,"\"")
                                                    );

                        var IMPNGCompositeProcessInfo = new ProcessStartInfo
                        {
                            FileName = convertExe,
                            Arguments = processPNGComposite,
                        };
                        if (!troubleShoot) { IMPNGCompositeProcessInfo.WindowStyle = ProcessWindowStyle.Hidden; }

                        var IMPNGFinishingProcessInfo = new ProcessStartInfo
                        {
                            FileName = convertExe,
                            Arguments = processPNGFinishing,
                        };
                        if (!troubleShoot) { IMPNGFinishingProcessInfo.WindowStyle = ProcessWindowStyle.Hidden; }

                        var IMPDFFinishingProcessInfo = new ProcessStartInfo
                        {
                            FileName = convertExe,
                            Arguments = processPDFFinishing,
                        };
                        if (!troubleShoot) { IMPDFFinishingProcessInfo.WindowStyle = ProcessWindowStyle.Hidden; }

                        /// reporting when item starts composite
                        backgroundWorkerIM.ReportProgress(inDx * 4 + inDx + 2, Path.GetFileName(compositePNGname) + "\n\n" + processPNGComposite.Trim());
                        using (var IMPNGCompositeProcess = Process.Start(IMPNGCompositeProcessInfo))
                        {
                            IMPNGCompositeProcess.WaitForExit();
                        }

                        //if cancellation is pending, cancel work.  
                        if (backgroundWorkerIM.CancellationPending) { e.Cancel = true; break; }

                        /// reporting when item starts PNG finishing
                        backgroundWorkerIM.ReportProgress(inDx * 4 + inDx + 2, Path.GetFileName(compositePNGname) + "\n\n" + processPNGFinishing.Trim());
                        using (var IMPNGFinishingProcess = Process.Start(IMPNGFinishingProcessInfo))
                        {
                            IMPNGFinishingProcess.WaitForExit();
                        }

                        //if cancellation is pending, cancel work.  
                        if (backgroundWorkerIM.CancellationPending) { e.Cancel = true; break; }

                        /// reporting when item starts PDF finishing
                        backgroundWorkerIM.ReportProgress(inDx * 4 + inDx + 3, Path.GetFileName(compositePDFname) + "\n\n" + processPDFFinishing.Trim());
                        using (var IMPDFFinishingProcess = Process.Start(IMPDFFinishingProcessInfo))
                        {
                            IMPDFFinishingProcess.WaitForExit();
                        }
                        backgroundWorkerIM.ReportProgress(-1, Path.GetFileName(compositePDFname));

                        //if cancellation is pending, cancel work.  
                        if (backgroundWorkerIM.CancellationPending) { e.Cancel = true; break; }

                    }
                    #endregion
                    #endregion

                    /// reporting when item is done
                    backgroundWorkerIM.ReportProgress((inDx * 4 + inDx + 4), Path.GetFileName(compositePDFname));

                    #region Cleanup: delete the temporary back and front image files
                    if (cleanUpTheFiles)
                    {
                        try
                        {
                            IEnumerable<String> frontfileList = Directory.GetFiles(Path.GetDirectoryName(olderfName),
                                frontPNGnamePat, SearchOption.TopDirectoryOnly);
                            foreach (string s in frontfileList)
                            {
                                File.Delete(s);
                            }
                        }
                        catch (Exception) { }
                        try
                        {
                            IEnumerable<String> backfileList = Directory.GetFiles(Path.GetDirectoryName(newerfilePathName),
                                backPNGnamePat, SearchOption.TopDirectoryOnly);
                            foreach (string s in backfileList)
                            {
                                File.Delete(s);
                            }
                        }
                        catch (Exception) { }
                    }
                    #endregion
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error in DoWork");
                    throw;
                }
                //if cancellation is pending, cancel work.  
                if (backgroundWorkerIM.CancellationPending)
                {
                    e.Cancel = true;
                    break;
                }
                inDx++;
            } // end for each
            e.Result = thisArgs;
        }

        private void backgroundWorkerIM_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            int numToDo = DTablePDFsOlder.Rows.Count;
            int stepMSGs = 5;
            int valX = e.ProgressPercentage;
            // progress value comes in as an integer
            progressBarPDFDiff.Visible = true;
            progressBarPDFDiff.Value = valX + 1;
            // we need to decode the vlaX to get what row it corresponds to
            // mod(vlaX,stepmsgs) = what stage in each step 
            int stage = (valX % stepMSGs);
            // row = (vlax - stage)/ stepmsgs
            int row = (valX - stage) / stepMSGs;

            Text = dftTitle + "   " + "<< PDFDifference is ongoing. >>";

            switch (stage)
            {
                case -1:
                    UpdatePNGResultsTable();
                    break;
                case 0:
                    labelMessage.Text = "Stage: " + (stage + 1).ToString() + "  Image: " + e.UserState.ToString();
                    LogMessage();
                    break;
                case 1:
                    labelMessage.Text = "Stage: " + (stage + 1).ToString() + "  Image: " + e.UserState.ToString();
                    LogMessage();
                    break;
                case 2:
                    labelMessage.Text = "Stage: " + (stage + 1).ToString() + "  Image: " + e.UserState.ToString();
                    LogMessage();
                    break;
                case 3:
                    labelMessage.Text = "Stage: " + (stage + 1).ToString() + "  Finshing: " + e.UserState.ToString();
                    LogMessage();
                    break;
                case 4:
                    labelMessage.Text = "Stage: " + (stage + 1).ToString() + "  Completed: " + e.UserState.ToString();
                    LogMessage();
                    break;
            }

            for (int i = 0; i < numToDo; i++)
            {
                if (valX + 1 == numToDo * stepMSGs)
                {
                    DTablePDFsOlder.Rows[i]["Status"] = msgDone;
                    DTablePDFsNewer.Rows[i]["Status"] = msgDone;
                }
                else
                {
                    if (i == row)
                    {
                        DTablePDFsOlder.Rows[i]["Status"] = msgInProcess;
                        DTablePDFsNewer.Rows[i]["Status"] = msgInProcess;
                    }
                    if (i < row)
                    {
                        DTablePDFsOlder.Rows[i]["Status"] = msgDone;
                        DTablePDFsNewer.Rows[i]["Status"] = msgDone;
                    }
                }
            }
        }

        private void backgroundWorkerIM_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled)
            {
                Text = dftTitle + "   " + "Canceled .... The processing is incomplete.";
                labelMessage.Text = "Canceled .... The processing is incomplete.";
            }
            else if (e.Error != null)
            {
                MessageBox.Show("Error. Details: " + (e.Error as Exception).ToString());
                labelMessage.Text = "Error. Details: " + (e.Error as Exception).ToString();
            }
            else
            {
                Text = dftTitle + "   " + "Done creatings PDF difference image files.";
                labelMessage.Text = "Done creatings PDF difference image files.";
            }
            buttonCancelDifference.Visible = false;
            buttonStartDifference.Text = printIt;
            buttonStartDifference.Enabled = true;
            progressBarPDFDiff.Value = progressBarPDFDiff.Minimum;
            progressBarPDFDiff.Visible = false;
            LogMessage();
        }

        class IMArgs
        {
            public List<string> PDFOlderFileNameList { get; set; }
            public List<string> PDFNewerFileNameList { get; set; }
            // public string filename { get; set; }
            public string IMconvertExeLocation { get; set; }
        }

        private void dataGridViewPDFS_MouseMove(object sender, MouseEventArgs e)
        {
            if ((e.Button & MouseButtons.Left) == MouseButtons.Left)
            {
                DataGridView dgv = sender as DataGridView;
                // If the mouse moves outside the rectangle, start the drag.
                if (dragBoxFromMouseDown != Rectangle.Empty &&
                !dragBoxFromMouseDown.Contains(e.X, e.Y))
                {
                    // Proceed with the drag and drop, passing in the list item.                    
                    DragDropEffects dropEffect = dgv.DoDragDrop(
                          dgv.Rows[rowIndexFromMouseDown],
                          DragDropEffects.Move);
                }
            }
        }

        private void dataGridViewPDFS_MouseDown(object sender, MouseEventArgs e)
        {
            DataGridView dgv = sender as DataGridView;
            // Get the index of the item the mouse is below.
            rowIndexFromMouseDown = dgv.HitTest(e.X, e.Y).RowIndex;
            if (rowIndexFromMouseDown != -1)
            {
                // Remember the point where the mouse down occurred. The DragSize indicates the
                //  size that the mouse can move before a drag event should be started.
                Size dragSize = SystemInformation.DragSize;

                // Create a rectangle using the DragSize, with the mouse position being
                // at the center of the rectangle.
                dragBoxFromMouseDown = new Rectangle(
                          new Point(
                            e.X - (dragSize.Width / 2),
                            e.Y - (dragSize.Height / 2)),
                      dragSize);
            }
            else
            {
                // Reset the rectangle if the mouse is not over an item in the ListBox.
                dragBoxFromMouseDown = Rectangle.Empty;
            }
        }

        private void dataGridViewPDFS_SelectionChanged(object sender, EventArgs e)
        {
            if (sender == dataGridViewPDFSOlder)
            {
                ShowPreviewImageForSelection(dataGridViewPDFSOlder, pictureBoxThumbOlder, labelPDFStatsOlder);
            }
            if (sender == dataGridViewPDFSNewer)
            {
                ShowPreviewImageForSelection(dataGridViewPDFSNewer, pictureBoxThumbNewer, labelPDFStatsNewer);
            }
        }

        private void buttonPrinters_Click(object sender, EventArgs e)
        {
            Process.Start("rundll32.exe", "shell32.dll,SHHelpShortcuts_RunDLL PrintersFolder");
        }

        private void ShowPreviewImageForSelection(DataGridView dgv, PictureBox pb, Label theLabel = null)
        {
            pb.Image = null;
            if (dgv.SelectedCells.Count < 1) { return; }
            string selfName = dgv.SelectedRows[0].Cells[0].Value.ToString();

            if (dgv == dataGridViewResults)  // special case
            {
                if (dataGridViewPDFSNewer.Rows.Count > 0)
                {
                    DataRow dr0 = DTablePDFsNewer.Rows[0];
                    // get path
                    string TheFirstNewerFilePName = dr0["PDFName"].ToString();
                    string thePath = Path.GetDirectoryName(TheFirstNewerFilePName);
                    selfName = Path.Combine(thePath, selfName);
                }
            }

            if (File.Exists(selfName))
            {
                Microsoft.WindowsAPICodePack.Shell.ShellFile shellfile; // = new Microsoft.WindowsAPICodePack.Shell.ShellFile();
                shellfile = Microsoft.WindowsAPICodePack.Shell.ShellFile.FromFilePath(selfName);
                pb.Image = shellfile.Thumbnail.LargeBitmap;
                if (theLabel != null) { ReportPagesData(selfName, theLabel); }

            }
        }

        // The process makes both PNG and PDFs as finished products. The PNGs are used for
        // previewing. The list shows only the PDFs. The PDFs are used for printing.
        // here the final file name is created by changing the extension to png.
        private void ShowPNGResultsImageForSelection(DataGridView dgv, PictureBox pb)
        {
            // clearing an existing image
            pb.Image = null;
            // was an item selected
            if (dgv.SelectedCells.Count < 1) { return; }
            if (DTablePDFsNewer.Rows.Count < 1)
            {
                // we cannot get path
                labelMessage.Text = "";
                string msg = "";
                msg = "Unable to get the path for viewing this difference file.";
                labelMessage.Text = msg;
                dgv.ClearSelection();
                return;
            }
            DataRow dr0 = DTablePDFsNewer.Rows[0];
            // are there actually PDFsNewer?
            if (dr0.RowState != DataRowState.Detached)
            {
                // get path
                string TheFirstNewerFilePName = dr0["PDFName"].ToString();
                string thePath = Path.GetDirectoryName(TheFirstNewerFilePName);
                // get name
                string selfName = dgv.SelectedRows[0].Cells[0].Value.ToString();
                String _fname = Path.Combine(thePath, Path.GetFileNameWithoutExtension(selfName)) + ".png";
                // one last check
                if (File.Exists(_fname))
                {
                    // This is the way Microsoft publishes as the way to avoid file locking. The image is scalable.
                    try
                    {
                        FileStream fs = null;
                        fs = new FileStream(_fname, FileMode.Open, FileAccess.Read);
                        pb.Image = Image.FromStream(fs);
                        fs.Close();
                        PageCount();
                    }
                    catch (IOException) { }
                }
                else
                {
                    if (File.Exists(Path.Combine(thePath, selfName)))
                    {
                        labelMessage.Text = "The corresponding PNG image for that PDF is missing. You can still print the PDF.";
                    }
                    else
                    {
                        labelMessage.Text = "Sorry, the list was out of date. What you selected does not exist. The list is now updated.";
                    }
                    UpdatePNGResultsTable();
                }
            }
        }

        private void ReportPagesData(string selfName, Label theLabel)
        {
            PdfReader rdr = new PdfReader(selfName);
            int nump = rdr.NumberOfPages;
            List<iTextSharp.text.Rectangle> sizesUsed = new List<iTextSharp.text.Rectangle>();
            for (int p = 1; p <= nump; p++)
            {
                iTextSharp.text.Rectangle pgsizeU = rdr.GetPageSizeWithRotation(p);
                if (!sizesUsed.Contains(pgsizeU)) { sizesUsed.Add(pgsizeU); }
            }
            string strSizes = "";
            foreach (iTextSharp.text.Rectangle Pgsize in sizesUsed)
            {
                float pht = Pgsize.Height / 72;
                float pwd = Pgsize.Width / 72;
                strSizes = strSizes + pwd.ToString(pwd % 1 == 0 ? "F0" : "F2") + " x " + pht.ToString(pht % 1 == 0 ? "F0" : "F2") + " , ";
            }
            string msg = "";
            if (sizesUsed.Count == 1)
            {
                msg = "Size: " + strSizes + nump.ToString() + " Pages";
            }
            else
            {
                msg = "Sizes: " + strSizes + nump.ToString() + " Pages";
            }
            theLabel.Text = msg;
        }

        private int TablePageCount(DataTable dt)
        {
            int pgCnt = 0;
            foreach (DataRow dr in dt.Rows)
            {
                if (dr["PDFName"] == null) { continue; }
                String pdfName = dr["PDFName"].ToString();
                if (pdfName.Length == 0) { continue; }
                pgCnt += (int)dr["Pages"];
            }
            return pgCnt;
        }

        private int PageCount()
        {
            int pgCntOldr = TablePageCount(DTablePDFsOlder);
            int pgCntNewer = TablePageCount(DTablePDFsNewer);

            labelQty.Text = "Newer Pages: " + pgCntNewer.ToString() + " | Older Pages: " + pgCntOldr.ToString();

            labelMessage.Text = "";
            string msg = "";
            if (pgCntOldr != pgCntNewer)
            {
                msg = "The PDF page count between the newer and older list is not equal.";
                msg = msg + " The PDF pair that does not have equal pages will not be processed.";
                msg = msg + " PDFs are paired according to their order in the newer/older lists.";
            }
            labelMessage.Text = msg;

            return pgCntOldr;
        }

        private int PDFpgCount(string PDFName)
        {
            if (!File.Exists(PDFName)) { return 0; }
            return new PdfReader(PDFName).NumberOfPages;
        }

        private void buttonClearList_Click(object sender, EventArgs e)
        {
            if (sender == buttonClearListOlder)
            {
                DTablePDFsOlder.Clear();
                ZapThumbnail(pictureBoxThumbOlder);
                labelPDFStatsOlder.Text = "";
            }
            if (sender == buttonClearListNewer)
            {
                DTablePDFsNewer.Clear();
                ZapThumbnail(pictureBoxThumbNewer);
                labelPDFStatsNewer.Text = "";
                DTablePNGResults.Clear();
            }
            buttonStartDifference.Enabled = false;
            progressBarPDFDiff.Value = progressBarPDFDiff.Minimum;
            progressBarPDFDiff.Visible = false;
            PageCount();
        }

        /// Use to see if a file exists in path. ImageMagick "convert.exe" is needed but
        /// beware that "convert.exe" is also a windows executable that does something else.
        /// 
        /// Not used. Also Imagemagick 7 no longer uses convert as an exe name.
        public static bool ExistsOnPathEnvironment(string exeName)
        {
            try
            {
                Process p = new Process();
                p.StartInfo.UseShellExecute = false;
                p.StartInfo.FileName = "where";
                p.StartInfo.Arguments = exeName;
                p.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                p.Start();
                p.WaitForExit();
                return p.ExitCode == 0;
            }
            catch (Win32Exception)
            {
                throw new Exception("'where' command is not on path");
            }
        }

        /// Returns the FIRST full path for whatever "where" finds using the windows path
        /// enviromental variable.
        public static string GetFullPath(string exeName)
        {
            try
            {
                Process p = new Process();
                p.StartInfo.UseShellExecute = false;
                p.StartInfo.FileName = "where";
                p.StartInfo.Arguments = exeName;
                p.StartInfo.RedirectStandardOutput = true;
                p.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                p.Start();
                string output = p.StandardOutput.ReadToEnd();
                p.WaitForExit();

                if (p.ExitCode != 0)
                    return null;

                // just return first match
                return output.Substring(0, output.IndexOf(Environment.NewLine));
            }
            catch (Win32Exception)
            {
                throw new Exception("'where' command is not on path");
            }
        }

        private void dataGridViewPDFSNewer_RowStateChanged(object sender, DataGridViewRowStateChangedEventArgs e)
        {
            UpdatePNGResultsTable();
        }

        private void dataGridViewResults_SelectionChanged(object sender, EventArgs e)
        {
            /// The purpose here it to light up the print button. The datagrid gets filled as
            /// the PNGs are made, not the PDF. We do not want the button enabled until after
            /// the PDF is available.
            DataGridView thisDGV = dataGridViewResults;
            if (thisDGV.SelectedRows.Count > 0)
            {
                DataRow dr0 = DTablePDFsNewer.Rows[0];
                //// are there actually PDFsNewer?
                if (dr0.RowState != DataRowState.Detached)
                {
                    // get newers path
                    string theNewersPath = Path.GetDirectoryName(dr0["PDFName"].ToString());
                    // get dataGridViewResults selection name
                    string selFileName = Path.Combine(theNewersPath, thisDGV.SelectedRows[0].Cells[0].Value.ToString());
                    if (File.Exists(selFileName))
                    {
                        buttonPrintSelected.Enabled = true;
                    }
                }
            }
            else
            {
                buttonPrintSelected.Enabled = false;
            }
        }

        private void dataGridViewResults_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            ShowPNGResultsImageForSelection(dataGridViewResults, pictureBoxComposites);
        }

        private void dataGridViewResults_KeyUp(object sender, KeyEventArgs e)
        {
            ShowPNGResultsImageForSelection(dataGridViewResults, pictureBoxComposites);
        }

        private void buttonPrintSelected_MouseClick(object sender, MouseEventArgs e)
        {
            Int32 selectedRowCount = dataGridViewResults.Rows.GetRowCount(DataGridViewElementStates.Selected);
            if (selectedRowCount > 0)
            {
                // get path
                DataRow dr0 = DTablePDFsNewer.Rows[0];
                string TheFirstNewerFilePName = dr0["PDFName"].ToString();
                string thePath = Path.GetDirectoryName(TheFirstNewerFilePName);

                /// The list should be printed in order. Thereo=fore we are 
                /// passing the list instread of each file.
                List<string> PDFFileList = new List<string>();
                for (int i = 0; i < selectedRowCount; i++)
                {
                    string selfName = dataGridViewResults.SelectedRows[i].Cells[0].Value.ToString();
                    String _fname = Path.Combine(thePath, selfName);
                    // one last check
                    if (File.Exists(_fname))
                    {
                        PDFFileList.Add(_fname);
                    }
                }
                PDFFileList.Reverse();
                progressBarPrint.Maximum = 2*PDFFileList.Count;
                string pSize = ThePSize();
                string printerN = labelDfltPrnt.Text;
                GSArgs thisGSArg = new GSArgs
                {
                    PDFFileNameList = PDFFileList,
                    printerN = printerN,
                    pSize = pSize,
                };
                /// Note: we are passing the entire list of pdf documents to print
                if (!backgroundWorkerGS.IsBusy)
                {
                    progressBarPrint.Visible = true;
                    buttonPrintCancel.Visible = true;
                    backgroundWorkerGS.RunWorkerAsync(thisGSArg);
                }


                //for (int i = 0; i < selectedRowCount; i++)
                //{
                //    pSize = ThePSize();
                //    printerN = labelDfltPrnt.Text;
                //    // get name
                //    string selfName = dataGridViewResults.SelectedRows[i].Cells[0].Value.ToString();
                //    String _fname = Path.Combine(thePath, selfName);
                //    // one last check
                //    if (File.Exists(_fname))
                //    {
                //        GS_PDF gspdf = new GS_PDF();
                //        gspdf.GSPrintDocument(_fname, printerN, pSize);
                //    }
                //}

            }
        }

        public static class myPrinters
        {
            [DllImport("winspool.drv", CharSet = CharSet.Auto, SetLastError = true)]
            public static extern bool SetDefaultPrinter(string Name);
        }

        class GSArgs
        {
            public List<string> PDFFileNameList { get; set; }
            public string printerN { get; set; }
            public string pSize { get; set; }
        }

        private void backgroundWorkerGS_DoWork(object sender, DoWorkEventArgs e)
        {
            GSArgs thisArgs = e.Argument as GSArgs;
            string pSize = thisArgs.pSize;
            string printerN = thisArgs.printerN;
            int progressIndX = 0;
            List<string> PDFFileNameListToDo = thisArgs.PDFFileNameList;
            foreach (string _fname in PDFFileNameListToDo)
            {
                progressIndX++;
                // report starting the indX
                backgroundWorkerGS.ReportProgress(progressIndX, Path.GetFileName(_fname));
                GS_PDF gspdf = new GS_PDF();
                gspdf.GSPrintDocument(_fname, printerN, pSize);
                
                //if cancellation is pending, cancel work.  
                if (backgroundWorkerGS.CancellationPending)
                {
                    e.Cancel = true;
                    break;
                }
            }
            e.Result = thisArgs;
        }

        private void backgroundWorkerGS_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            int progressIndX = e.ProgressPercentage;
            progressBarPrint.Value = 2* progressIndX;
            string f = e.UserState.ToString();
            string msg = "Starting to print " + f;
            labelMessage.Text = msg;
            LogMessage();
        }

        private void backgroundWorkerGS_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled)
            {
                labelMessage.Text = "Canceled .... The printing processing is incomplete.";
                buttonPrintCancel.Visible = false;
            }
            else if (e.Error != null)
            {
                labelMessage.Text = "Error. Details: " + (e.Error as Exception).ToString();
                buttonPrintCancel.Visible = false;
            }
            else
            {
                labelMessage.Text = "Done Printing Difference PDFs";
                buttonPrintCancel.Visible = false;
            }
            progressBarPrint.Visible = false;
        }

        private void buttonPrintCancel_Click(object sender, EventArgs e)
        {
            backgroundWorkerGS.CancelAsync();
            labelMessage.Text = "... Got it. The printing cancel is pending ...";
        }

        private void buttonCancelDifference_Click(object sender, EventArgs e)
        {
            backgroundWorkerIM.CancelAsync();
            labelMessage.Text = "... Got it. The PDF-Difference cancel is pending ...";
        }

        private void ANYdataGridView_DoubleClick(object sender, EventArgs e)
        {
            // only process when list is empty
            DataGridView dgv = sender as DataGridView;
            if (dgv.SelectedRows.Count < 1)
            {
                Process.Start("explorer.exe");
            }
        }

        private void ANYdataGridView_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            DataGridView dgv = sender as DataGridView;
            int rowIndex = e.RowIndex;
            DataGridViewRow row = dgv.Rows[rowIndex];
            //MessageBox.Show(@row.Cells[0].Value.ToString());
            string pathName = Path.GetDirectoryName(row.Cells[0].Value.ToString());
            Process.Start("explorer.exe", @pathName);
        }

        private string LogFile(bool makenew = false)
        {
            string logFileName = null;
            DataRow dr0 = DTablePDFsNewer.Rows[0];
            if (dr0.RowState != DataRowState.Detached)
            {
                string theNewersPath = Path.GetDirectoryName(dr0["PDFName"].ToString());
                logFileName = Path.Combine(theNewersPath, "composite_PDF_SessionLog.txt");
            }
            if (File.Exists(logFileName) && makenew)
            {
                try
                {
                    File.Delete(logFileName);
                }
                catch (Exception)
                {
                    labelMessage.Text = "Unable to delete previous logfile: " + logFileName;
                }
            }
            return logFileName;
        }

        private void LogMessage()
        {
            string thisLog = LogFile();
            // The using statement automatically flushes AND CLOSES the stream and calls 
            // IDisposable.Dispose on the stream object.
            using (StreamWriter file =
            new StreamWriter(@thisLog, true))
            {
                DateTime localDate = DateTime.Now;
                file.WriteLine(localDate.ToString());
                file.WriteLine(labelMessage.Text);
                file.WriteLine("\n");
            }
        }


    }
}



