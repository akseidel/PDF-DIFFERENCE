using System;
using System.IO;
using System.ComponentModel;
using System.Diagnostics;


namespace PDF_DIFFERENCE
{
    class Helpers
    {
        private string GS_EXE_LOCATION = "C:\\Program Files\\gs\\gs9.16\\bin\\gswin64c.exe";

        public string WhatIsGSExecutable()
        {
            if (File.Exists(GS_EXE_LOCATION))
            {
                return GS_EXE_LOCATION;
            }
            else
            {
                return "NOT FOUND => " + GS_EXE_LOCATION;
            }
        }

        internal string WhatIsMagick()
        {
            string thisMagick = GetFullPathWithMatch("magick", "ImageMagick");
            if (thisMagick != null)
            {
                return thisMagick;
            }
            else
            {
                return "NOT FOUND => " + "ImageMagick";
            }
        }

        /// Returns the full path for whatever "where" finds using the windows path
        /// enviromental variable where that path contains the matchTo string.
        /// Used here to find the convert.exe that belongs to ImageMagick.
        public static string GetFullPathWithMatch(string exeName, string matchTo)
        {
            try
            {
                var pi = new ProcessStartInfo
                {
                    UseShellExecute = false,
                    FileName = "WHERE",
                    Arguments = exeName,
                    WindowStyle = ProcessWindowStyle.Hidden,
                    RedirectStandardOutput = true
                };
                using (var piProcess = Process.Start(pi))
                {
                    piProcess.WaitForExit();
                    string output = piProcess.StandardOutput.ReadToEnd();


                    string[] lines = output.Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);
                    foreach (string s in lines)
                    {
                        bool contains = s.IndexOf(matchTo, StringComparison.OrdinalIgnoreCase) >= 0;
                        if (contains)
                        {
                            return s;
                        }
                    }
                    return null;
                }

                //Process p = new Process();
                //p.StartInfo.UseShellExecute = false;
                //p.StartInfo.FileName = "where";
                //p.StartInfo.Arguments = exeName;
                //p.StartInfo.RedirectStandardOutput = true;
                //p.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                //p.Start();
                //string output = p.StandardOutput.ReadToEnd();
                //p.WaitForExit();
                //if (p.ExitCode != 0) { return null; }
                //string[] lines = output.Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);
                //foreach (string s in lines)
                //{
                //    bool contains = s.IndexOf(matchTo, StringComparison.OrdinalIgnoreCase) >= 0;
                //    if (contains)
                //    {
                //        return s;
                //    }
                //}
                //return null;
            }
            catch (Win32Exception)
            {
                throw new Exception("'where' command is not on path");
            }
        }
    }
}
