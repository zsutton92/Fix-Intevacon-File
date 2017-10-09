using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using System.IO;
using System.Text.RegularExpressions;
using System.Diagnostics;

namespace Fix_Intevacon_File
{
    public partial class frmMain : Form
    {
        public frmMain()
        {
            InitializeComponent();
            this.AllowDrop = true;
            this.DragEnter += new DragEventHandler(frmMain_DragEnter);
            this.DragDrop += new DragEventHandler(frmMain_DragDrop);
        }

        void frmMain_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop)) e.Effect = DragDropEffects.Copy;
        }

        void frmMain_DragDrop(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            foreach (string file in files) ReadCsv(file);
        }

        private void ReadCsv(string file)
        {
            string tempPath = Path.GetDirectoryName(file) + "\\" + Path.GetFileNameWithoutExtension(file) + "_modified.csv";
            var delimiter = ",";
            var firstLineContainsHeaders = true;
            var lineNumber = 0;
            string sourcePath = file;

            var splitExpression = new Regex(@"(" + delimiter + @")(?=(?:[^""]|""[^""]*"")*$)");

            using (var writer = new StreamWriter(tempPath))
            using (var reader = new StreamReader(sourcePath))
            {
                string line = null;
                string[] headers = null;
                if (firstLineContainsHeaders)
                {
                    line = reader.ReadLine();
                    lineNumber++;

                    if (string.IsNullOrEmpty(line)) return; // file is empty;

                    headers = splitExpression.Split(line).Where(s => s != delimiter).ToArray();

                    List<string> newHeaders = new List<string>(headers);
                    newHeaders.RemoveAt(newHeaders.Count - 1);
                    newHeaders.RemoveAt(1);
                    headers = newHeaders.ToArray();

                    line = string.Join(",", headers);


                    writer.WriteLine(line); // write the original header to the temp file.
                }

                while ((line = reader.ReadLine()) != null)
                {
                    lineNumber++;

                    var columns = splitExpression.Split(line).Where(s => s != delimiter).ToArray();

                    //if (columns[0] == "123153" || columns[0] == "Subtotal 123153")
                    //{
                    //    columns[1] = "Max Arnold and Sons LLC";
                    //    var lstRemove = new List<string>(columns);
                    //    lstRemove.RemoveAt(2);
                    //    columns = lstRemove.ToArray();
                    //}


                    //remove columns 7 and 8
                    try
                    {
                        List<string> list = new List<string>(columns);
                        list.RemoveAt(list.Count - 1);
                        list.RemoveAt(1);
                        columns = list.ToArray();
                    }
                    catch { }

                    string lineToWrite = "";

                    try
                    {
                        lineToWrite = string.Join(delimiter, columns);
                    }
                    catch { }


                    if (lineToWrite.Length > 0)
                    {
                        writer.WriteLine(lineToWrite);
                    }


                }
            }

            File.Delete(sourcePath);
            File.Move(tempPath, sourcePath);
            Process.Start(sourcePath);
        }
    }
}
