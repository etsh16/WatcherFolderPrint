using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Management;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WatcherFolderPrint
{
    public partial class Form1 : Form
    {
        readonly string FileExt = "*.html";
        List<string> Printers = new List<string>();
            FileSystemWatcher watcher1;
            FileSystemWatcher watcher2;
            FileSystemWatcher watcher3;
            FileSystemWatcher watcher4;
        string printer1 = null;
        string printer2 = null;
        string printer3 = null;
        string printer4 = null;
        public Form1()
        {
            InitializeComponent();

        }
        private void Form1_Load(object sender, EventArgs e)
        {
            CMP_Printer1.ValueMember = "Name";
            CMP_Printer1.DisplayMember = "Name";

            CMP_Printer2.ValueMember = "Name";
            CMP_Printer2.DisplayMember = "Name";

            CMP_Printer3.ValueMember = "Name";
            CMP_Printer3.DisplayMember = "Name";

            CMP_Printer4.ValueMember = "Name";
            CMP_Printer4.DisplayMember = "Name";

            LoadPrinters();
        }

        void LoadSettings()
        {
            TXT_Folder1.Text = Properties.Settings.Default.Path1;
            TXT_Folder2.Text = Properties.Settings.Default.Path2;
            TXT_Folder3.Text = Properties.Settings.Default.Path3;
            TXT_Folder4.Text = Properties.Settings.Default.Path4;
           CMP_Printer1.SelectedText= Properties.Settings.Default.Printer1;
            CMP_Printer2.SelectedText = Properties.Settings.Default.Printer2 ;
            CMP_Printer3.SelectedText = Properties.Settings.Default.Printer3 ;
            CMP_Printer4.SelectedText = Properties.Settings.Default.Printer4 ;

        }
        void SaveSettings()
        {
              Properties.Settings.Default.Path1= TXT_Folder1.Text; 
              Properties.Settings.Default.Path2 = TXT_Folder2.Text;
              Properties.Settings.Default.Path3 = TXT_Folder3.Text;
              Properties.Settings.Default.Path4 = TXT_Folder4.Text;
            Properties.Settings.Default.Printer1 = CMP_Printer1.SelectedText;
            Properties.Settings.Default.Printer2 = CMP_Printer2.SelectedText;
            Properties.Settings.Default.Printer3 = CMP_Printer3.SelectedText;
            Properties.Settings.Default.Printer4 = CMP_Printer4.SelectedText;
            Properties.Settings.Default.Save();
        }
        void LoadPrinters()
        {
            Printers.Clear();

            foreach (string printer in System.Drawing.Printing.PrinterSettings.InstalledPrinters)
            {
                Printers.Add(printer);
            }
            CMP_Printer1.DataSource = new List<string>(Printers);
            CMP_Printer2.DataSource = new List<string>(Printers);
            CMP_Printer3.DataSource = new List<string>(Printers);
            CMP_Printer4.DataSource = new List<string>(Printers);
        }
        void GetFolder(TextBox TXT_Path)
        {
            using (var fbd = new FolderBrowserDialog())
            {
                DialogResult result = fbd.ShowDialog();

                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    TXT_Path.Text = fbd.SelectedPath;
                }
            }
        }
        FileSystemWatcher StartWatcher( out FileSystemWatcher watcher,string path,  FileSystemEventHandler fun)
        {
            // Create a new FileSystemWatcher and set its properties.
             watcher = new FileSystemWatcher();
            try
            {
                
                watcher.Path = path;
                watcher.NotifyFilter = NotifyFilters.FileName | NotifyFilters.CreationTime | NotifyFilters.LastWrite;
                watcher.Filter = FileExt;
                watcher.IncludeSubdirectories = true;
                watcher.Created += new FileSystemEventHandler(fun);
                watcher.EnableRaisingEvents = true;
            }
            catch (IOException e)
            {
                Console.WriteLine("A Exception Occurred :" + e);
            }
            catch (Exception oe)
            {
                Console.WriteLine("An Exception Occurred :" + oe);
            }
            return watcher;
        }
        void StopWatcher( FileSystemWatcher watcher, Button btn_start, Button btn_stop, Button btn_browse, ComboBox cmb_printer)
        {
            if (watcher!=null)
            {
                watcher.Dispose();
            }
            btn_start.Enabled = true;
            btn_browse.Enabled = true;
            cmb_printer.Enabled = true;
            btn_stop.Enabled = false;
        }
        void StartWatcher(out FileSystemWatcher watch,  FileSystemEventHandler onChanged , TextBox TXT_Folder, Button btn_start, Button btn_stop, Button btn_browse,ComboBox cmb_printer)
        {
            if (!string.IsNullOrEmpty(TXT_Folder.Text))
            {
                StartWatcher(out watch, TXT_Folder.Text, onChanged);
                btn_start.Enabled = false;
                btn_browse.Enabled = false;
                cmb_printer.Enabled = false;
                btn_stop.Enabled = true;
            }
            else
            {
                watch = null;   
            }
        }
        void PrintHandle(FileSystemEventArgs e,string printer_name)
        {
            try
            {
            Thread th = new Thread(() =>
            {
            this.Invoke((MethodInvoker)(() => {
                SetDefaultPrinter(printer_name);
                PrintHelpPage(e.FullPath);
            }));
            });
            th.IsBackground = true; // 
            th.Start();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }


        public void OnChanged1(object source, FileSystemEventArgs e)
        {
            if (! string.IsNullOrEmpty(printer1))
            {
                PrintHandle(e, printer1);
            }
        }
        public void OnChanged2(object source, FileSystemEventArgs e)
        {
            if (!string.IsNullOrEmpty(printer2))
            {
                PrintHandle(e, printer2);
            }
        }
        public void OnChanged3(object source, FileSystemEventArgs e)
        {
            if (!string.IsNullOrEmpty(printer3))
            {
                PrintHandle(e, printer3);
            }

        }
        public void OnChanged4(object source, FileSystemEventArgs e)
        {
            if (!string.IsNullOrEmpty(printer4))
            {
                PrintHandle(e, printer4);
                File.Delete(e.FullPath);
            }
        }


        public static bool SetDefaultPrinter(string defaultPrinter)
        {
            using (ManagementObjectSearcher objectSearcher = new ManagementObjectSearcher("SELECT * FROM Win32_Printer"))
            {
                using (ManagementObjectCollection objectCollection = objectSearcher.Get())
                {
                    foreach (ManagementObject mo in objectCollection)
                    {
                        if (string.Compare(mo["Name"].ToString(), defaultPrinter, true) == 0)
                        {
                            mo.InvokeMethod("SetDefaultPrinter", null, null);
                            return true;
                        }
                    }
                }
            }
            return false;
        }
        private void PrintHelpPage(string PathStr)
        {
           
                    WebBrowser webBrowserForPrinting = new WebBrowser();

                    // Add an event handler that prints the document after it loads.
                    webBrowserForPrinting.DocumentCompleted +=
                        new WebBrowserDocumentCompletedEventHandler(PrintDocument);

                    // Set the Url property to load the document.
                    webBrowserForPrinting.Url = new Uri(PathStr);
              
        }
        private void PrintDocument(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            // Print the document now that it is fully loaded.
            ((WebBrowser)sender).Print();
            // Dispose the WebBrowser now that the task is complete. 
            ((WebBrowser)sender).Dispose();
        }


        private void BTN_Browse1_Click(object sender, EventArgs e)
        {
            GetFolder(TXT_Folder1);
        }
        private void BTN_Browse2_Click(object sender, EventArgs e)
        {
            GetFolder(TXT_Folder2);
        }
        private void BTN_Browse3_Click(object sender, EventArgs e)
        {
            GetFolder(TXT_Folder3);
        }
        private void BTN_Browse4_Click(object sender, EventArgs e)
        {
            GetFolder(TXT_Folder4);
        }

        private void BTN_Start1_Click(object sender, EventArgs e)
        {
            printer1 = CMP_Printer1.Text;
            StartWatcher(out watcher1,OnChanged1, TXT_Folder1, BTN_Start1, BTN_Stop1, BTN_Browse1, CMP_Printer1);
        }
        private void BTN_Start2_Click(object sender, EventArgs e)
        {
            printer2 = CMP_Printer2.Text;
            StartWatcher(out watcher2, OnChanged2, TXT_Folder2, BTN_Start2, BTN_Stop2, BTN_Browse2, CMP_Printer2);
        }
        private void BTN_Start3_Click(object sender, EventArgs e)
        {
            printer3 = CMP_Printer3.Text;
            StartWatcher(out watcher3, OnChanged3, TXT_Folder3,  BTN_Start3, BTN_Stop3, BTN_Browse3, CMP_Printer3);
        }
        private void BTN_Start4_Click(object sender, EventArgs e)
        {
            printer4 = CMP_Printer4.Text;
            StartWatcher(out watcher4, OnChanged4, TXT_Folder4,  BTN_Start4, BTN_Stop4, BTN_Browse4, CMP_Printer4);
        }

        private void BTN_Stop1_Click(object sender, EventArgs e)
        {
            printer1 = null;
            StopWatcher( watcher1, BTN_Start1, BTN_Stop1, BTN_Browse1, CMP_Printer1); 
        }
        private void BTN_Stop2_Click(object sender, EventArgs e)
        {
            printer2 = null;
            StopWatcher( watcher2, BTN_Start2, BTN_Stop2, BTN_Browse2, CMP_Printer2);
        }
        private void BTN_Stop3_Click(object sender, EventArgs e)
        {
            printer3 = null;
            StopWatcher( watcher3, BTN_Start3, BTN_Stop3, BTN_Browse3, CMP_Printer3);
        }
        private void BTN_Stop4_Click(object sender, EventArgs e)
        {
            printer4 = null;
            StopWatcher( watcher4, BTN_Start4, BTN_Stop4, BTN_Browse4, CMP_Printer4);
        }

        private void saveSettingsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveSettings();
        }
        private void loadSettingsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LoadSettings();
        }
        private void refreshPrintersToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LoadPrinters();
        }
        private void resetToolStripMenuItem_Click(object sender, EventArgs e)
        {
            BTN_Stop1_Click(null, null);
            BTN_Stop2_Click(null, null);
            BTN_Stop3_Click(null, null);
            BTN_Stop4_Click(null, null);


            TXT_Folder1.Clear();
            TXT_Folder2.Clear();
            TXT_Folder3.Clear();
            TXT_Folder4.Clear();

            CMP_Printer1.SelectedText = "";
            CMP_Printer2.SelectedText = "";
            CMP_Printer3.SelectedText = "";
            CMP_Printer4.SelectedText = "";
        }
    }
}
