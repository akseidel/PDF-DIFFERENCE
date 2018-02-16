using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Printing;

namespace PDF_DIFFERENCE
{
    class GS_PDF
    {
        private string GS_ExecutableFile = new Helpers().WhatIsGSExecutable();

        private string GetDefaultPrinter()
        {
            string name = LocalPrintServer.GetDefaultPrintQueue().FullName;
            return name;
        }

        public void GSPrintDocument(string fs, string printerName = null, string pSize = null)
        {
            if (!File.Exists(fs)) return;
            var filename = fs ?? string.Empty;
            printerName = printerName ?? GetDefaultPrinter(); //get your printer here
            pSize = pSize ?? "letter"; // default to letter if not specified

            /// "C:\Program Files\gs\gs9.16\bin\gswin64c.exe"
            /// -dAutoRotatePages=/All
            /// -dNOPAUSE
            /// -dBATCH
            /// -sPAPERSIZE=letter
            /// -dFIXEDMEDIA
            /// -dPDFFitPage
            /// -dEmbedAllFonts=true
            /// -dPDFSETTINGS=/prepress
            /// -dNOPLATFONTS
            /// -dNumCopies=1
            /// -dNoCancel
            /// -sDEVICE="mswinpr2"
            /// -sOutputFile="%printer%\\ipp://iprint.wmtao.com\Toshiba 1057 PCL"
            /// "C:\PDFTest\Newers\Composite_301-302_and_MEF-301-MEF-302-0.pdf"

            var processArgs = string.Format("-dAutoRotatePages=/All "
                + "-dNOPAUSE -dBATCH -sPAPERSIZE={0} -dFIXEDMEDIA -dPDFFitPage "
                + "-dEmbedAllFonts=true -dSubsetFonts=true -dPDFSETTINGS=/prepress "
                + "-dNOPLATFONTS -dNumCopies=1 "
                + "-dNoCancel -sDEVICE=\"mswinpr2\" "
                + "-sOutputFile=\"%printer%{1}\" \"{2}\"", pSize, printerName, filename);
            try
            {
                var gsProcessInfo = new ProcessStartInfo
                {   
                    WindowStyle = ProcessWindowStyle.Hidden,
                    FileName = GS_ExecutableFile,
                    // UseShellExecute = false,
                    // RedirectStandardError = true,
                    // RedirectStandardOutput = true,
                    Arguments = processArgs
                };
                //string msg = gswinEXEInstallationLocation + "\n\n";
                //msg = msg + processArgs;
                //MessageBox.Show(msg);
                ////Console.Write(msg);
                using (var gsProcess = Process.Start(gsProcessInfo))
                {
                    // wait for 10 minutes. If not done then kill the process
                    gsProcess.WaitForExit(6000000);
                    if (gsProcess.HasExited == false) { gsProcess.Kill(); }
                    gsProcess.Close();
                }
            }
            catch (Exception)
            {
                throw;
            }
        } // end GSPrintDocument
    }
}

///You should test your options from the commandline first,
///and then translate the successes into your code.

///A PDF file usually does already include page margins. You "often cut" page content
///may result from a PDF which is meant for A4 page size printed on Letter format.

///PDF also uses some internal boxes which organize the page (and object) content: MediaBox, TrimBox, CropBox, Bleedbox.

///There are various options to control for which "media size" Ghostscript renders a given input:

///-dPDFFitPage
///-dUseTrimBox
///-dUseCropBox

///With PDFFitPage Ghostscript will render to the current page device size (usually the default page size).

///With UseTrimBox it will use the TrimBox (and it will at the same time set the PageSize to that value).

///With UseCropBox it will use the CropBox (and it will at the same time set the PageSize to that value).

///By default (give no parameter), Ghostscript will render using the MediaBox.

///Note, you can additionally control the overall size of your output by using "-sPAPERSIZE" 
///(select amongst all pre-defined values Ghostscript knows) or (for more flexibility) 
///use "-dDEVICEWIDTHPOINTS=NNN -dDEVICEHEIGHTPOINTS=NNN" to set up custom page sizes.
///

///Not sure if it helps anyone, but to add the printing documents to a queue instead of immediately 
///printing make changes to the above section with

///startInfo.Arguments = " -dPrinted -dNoCancel=true -dBATCH -dNOPAUSE
///-dNOSAFER -q -dNumCopies=" + Convert.ToString(numberOfCopies) + 
///" -sDEVICE=mswinpr2 -sOutputFile=%printer%" + printerName + " \"" + pdfFullFileName + "\"";

///Pre-requisites: Configure your printer's job type to "Hold Print": In our case we have a Rico Aficio
///MP 4000 Printer and our usage is to run a nightly job to print a bunch of PDF files generated through SSRS.

///// <summary>
///// Prints the PDF.
///// </summary>
///// <param name="ghostScriptPath">The ghost script path. Eg "C:\Program Files\gs\gs8.71\bin\gswin32c.exe"</param>
///// <param name="numberOfCopies">The number of copies.</param>
///// <param name="printerName">Name of the printer. Eg \\server_name\printer_name</param>
///// <param name="pdfFileName">Name of the PDF file.</param>
///// <returns></returns>
//public bool PrintPDF(string ghostScriptPath,
//                    int numberOfCopies,
//                    string printerName,
//                    string pdfFileName)
//{
//    ProcessStartInfo startInfo = new ProcessStartInfo();
//    startInfo.Arguments = " -dPrinted -dBATCH -dNOPAUSE -dNOSAFER -q -dNumCopies="
//                            + Convert.ToString(numberOfCopies)
//                            + " -sDEVICE=ljet4 -sOutputFile=\"\\\\spool\\"
//                            + printerName + "\" \""
//                            + pdfFileName + "\" ";
//    startInfo.FileName = ghostScriptPath;
//    startInfo.UseShellExecute = false;
//    startInfo.RedirectStandardError = true;
//    startInfo.RedirectStandardOutput = true;
//    Process process = Process.Start(startInfo);
//    Console.WriteLine(process.StandardError.ReadToEnd() + process.StandardOutput.ReadToEnd());
//    process.WaitForExit(30000);
//    if (process.HasExited == false) process.Kill();
//    process.Close();
//    return process.ExitCode == 0;
//}

