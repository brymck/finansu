/* 
 * FinAnSu
 * http://code.google.com/p/finansu/
 * 
 * Copyright 2011, Bryan McKelvey
 * Licensed under the MIT license
 * http://www.opensource.org/licenses/mit-license.php
 */

using System;
using System.ComponentModel;
using System.IO;
using System.Net;
using System.Threading;
using System.Windows.Forms;

namespace FinAnSu.Controls
{
    public partial class DownloadForm : Form
    {
        private string fileName = Environment.GetEnvironmentVariable("UserProfile") + "\\Desktop\\finansu.zip";

        public DownloadForm()
        {
            InitializeComponent();
            ProgressWorker.RunWorkerAsync();
        }

        private void ProgressWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            string suffix = "";

            if (Environment.Is64BitProcess)
            {
                suffix = "x64";
            }
            else
            {
                suffix = "x86";
            }

            Uri url = new Uri(string.Format("https://github.com/downloads/brymck/finansu/FinAnSu-{0}_{1}.zip",
                FinAnSu.Main.LatestVersion(), suffix));

            // first, we need to get the exact size (in bytes) of the file we are downloading
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
            HttpWebResponse response = (HttpWebResponse)request.GetResponse();
            response.Close();

            // gets the size of the file in bytes
            Int64 fileSize = response.ContentLength;
            Int64 runningByteTotal = 0;

            // use the webclient object to download the file
            using (WebClient client = new WebClient())
            {
                // open the file at the remote URL for reading
                using (Stream streamRemote = client.OpenRead(url))
                {
                    // using the FileStream object, we can write the downloaded bytes to the file system
                    using (Stream streamLocal = new FileStream(fileName, FileMode.Create, FileAccess.Write, FileShare.None))
                    {
                        // loop the stream and get the file into the byte buffer
                        int byteSize = 0;
                        byte[] byteBuffer = new byte[fileSize];
                        while ((byteSize = streamRemote.Read(byteBuffer, 0, byteBuffer.Length)) > 0)
                        {
                            // write the bytes to the file system at the file path specified
                            streamLocal.Write(byteBuffer, 0, byteSize);
                            runningByteTotal += byteSize;

                            // calculate the progress out of a base "100"
                            double dIndex = (double)(runningByteTotal);
                            double dTotal = (double)byteBuffer.Length;
                            double dProgressPercentage = (dIndex / dTotal);
                            int iProgressPercentage = (int)(dProgressPercentage * 100);

                            // update the progress bar
                            ProgressWorker.ReportProgress(iProgressPercentage);
                        }
                        // clean up the file stream
                        streamLocal.Close();
                    }

                    // close the connection to the remote server
                    streamRemote.Close();
                }
            }
        }

        private void ProgressWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            StatusLabel.Text = string.Format("Downloading to desktop... {0}% complete", e.ProgressPercentage);
            ProgressBar.Value = e.ProgressPercentage;
        }

        private void ProgressWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            StatusLabel.Text = "Downloading to desktop... 100% complete";
            ProgressBar.Value = 100;
            Thread.Sleep(1000);
            this.Close();
        }
    }
}