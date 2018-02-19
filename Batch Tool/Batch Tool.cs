
using System;
using System.Xml;
using System.Xml.Linq;
using System.Diagnostics;
using System.IO;
using System.IO.Packaging;
using Microsoft.Office.Interop.Visio;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Collections;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing.Printing;
using System.Runtime.InteropServices;

using PdfSharp;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using PdfSharp.Pdf.Content;

using iTextSharp;
using iTextSharp.text;
using iTextSharp.text.pdf;

using Microsoft.WindowsAPICodePack.Dialogs;

using Visio = Microsoft.Office.Interop.Visio;




//Stealing keyboard focus
//https://social.msdn.microsoft.com/Forums/office/en-US/7ec7ee04-9408-4205-a16e-741e624a8ab7/excel-automation-application-steals-focus?forum=exceldev
//http://csharp.wekeepcoding.com/article/23656762/How+to+keep+Excel+interop+from+stealing+focus+while+inserting+images


namespace $safeprojectname$
{
    
    public partial class Form1 : Form
    {
        private BackgroundWorker bw;

        browsePopulate browsePop = new browsePopulate();

        transplant tr = new transplant();

        

        public int progressInc = 0;
        public int progressCur = 0;
        public int indivProg = 0;
        public int indivInc = 0;
        
        public int totalAction = 0;
        public static string whatStatus = "0";
        public bool titlePage = false;
        public bool howToRead = false;
        public bool revPage = false;
        public bool backPage = false;
        public bool tableOfContents = false;
        public bool cloudTriangle = false;
        public bool sheetNumCheck = false;
        public bool drawNumCheck = false;
        public bool dsgnByChkBy = false;
        public bool sheetMargins = false;
        public bool printIndiv = false;
        public int printComb = 0;
        public static bool escape = false;


        public Form1()
        {
            InitializeComponent();
            InitializeOpenFileDialog();

            notifyIcon1.Visible = false;

            toolStripStatusLabel1.Text = "Click 'Browse Files' to begin";
                       
            if (Properties.Settings.Default.PDFPath == "" || Properties.Settings.Default.PDFPath == null)
            {
                label3.Visible = true;
                radioButton10.Enabled = false;
                radioButton11.Enabled = false;
            }
           
            treeView1.Nodes.Clear();
            button3.Enabled = false;

            this.bw = new BackgroundWorker();
            this.bw.DoWork += new DoWorkEventHandler(bw_DoWork);
            this.bw.ProgressChanged += new ProgressChangedEventHandler(bw_ProgressChanged);
            this.bw.WorkerSupportsCancellation = true;
            //this.bw.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bw_RunWorkerCompleted);
            this.bw.WorkerReportsProgress = true;

            
        }

        [DllImport("user32.dll")]
        static extern void FlashWindow(IntPtr a, bool b);
        IntPtr handle = Process.GetCurrentProcess().MainWindowHandle;

        private void message(string message)
        {
            balloon notify = new balloon();
            notify.data(message);
            notify.Show();
            
            FlashWindow(Handle, true);
            
            
        } //NOTIFICATION

        private void InitializeOpenFileDialog()
        {
            //Filter Visio files only
            this.openFileDialog1.Filter = "visio files (*.vsd; *.vsdx;)|*.vsd; *.vsdx;|All files (*.*)|*.*";

            //Enable multi-file select
            this.openFileDialog1.Multiselect = true;
            this.openFileDialog1.Title = "Visio File Select";
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void Form1_Load_1(object sender, EventArgs e)
        {
            this.Focus();
            
            if (Properties.Settings.Default.Notification == false)
            {
                message("Ready");
            }
        }

        private void folderBrowserDialog1_HelpRequest(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {          
            browsePop.fileFetch(ref treeView1, ref toolStripStatusLabel1);
            if (treeView1.Nodes != null)
            {
                button3.Enabled = true;
                button3.BackColor = System.Drawing.Color.Green;
                button3.ForeColor = System.Drawing.Color.White;

            }
        } //FILE FETCH

        private void button2_Click(object sender, EventArgs e)
        {
            if (treeView1.Nodes.Count != 0)
            {
                DialogResult dialogResult = MessageBox.Show("Clear Files?", "Clear Files?", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (dialogResult == DialogResult.Yes)
                {
                    toolStripStatusLabel1.Text = "Clearing memory";
                    treeView1.Nodes.Clear();
                    toolStripStatusLabel1.Text = "Click 'Browse Files' to begin";

                    button3.ForeColor = System.Drawing.Color.Black;
                    button3.BackColor = System.Drawing.Color.White;
                    button3.Enabled = false;
                }
                if (dialogResult == DialogResult.No)
                {
                    return;
                }
            }

        } //CLEAR ALL FILES

        protected void button3_Click(object sender, EventArgs e) 
        {
            if(button3.Text == "STOP")
            {
                toolStripStatusLabel1.Text = "Stopping actions";
                escape = true;
                bw.CancelAsync();
                
                bw.ReportProgress(0);
                
            }

            else
            {
                escape = false;
                button3.Text = "STOP";
                button3.BackColor = System.Drawing.Color.Red;
                button3.ForeColor = System.Drawing.Color.White;

                //tabControl1.Enabled = false;
                button1.Enabled = false;
                button2.Enabled = false;
                //button3.Enabled = false;
                button4.Enabled = false;
                button6.Enabled = false;

                toolStripStatusLabel1.Text = "Calculating actions";

                countAction();

                if (totalAction == 0)
                {
                    DialogResult dialogResult = MessageBox.Show("Execution was terminated because no actions were selected.\nPlease select an action.", "No Actions Selected", 
                        MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                
                    toolStripStatusLabel1.Text = "Select an action to preform";
                    toolStripProgressBar1.Value = 0;
                    button1.Enabled = true;
                    button2.Enabled = true;
                    button3.Enabled = true;
                    button4.Enabled = true;
                    button6.Enabled = true;

                    button3.BackColor = System.Drawing.Color.Green;
                    button3.Text = "EXECUTE";
                    button3.ForeColor = System.Drawing.Color.White;
                    return;
                }
                else
                {
                    toolStripProgressBar1.Value = 0;
                    toolStripStatusLabel1.Text = "Executing";
                    tabControl1.Enabled = false;
                    bw.RunWorkerAsync();

                }
            }
        } //EXECUTE

        private void button4_Click(object sender, EventArgs e)
        {
            browsePop.folderFetch(ref treeView1, ref toolStripStatusLabel1);
            if (treeView1.Nodes != null)
            {
                button3.Enabled = true;
                button3.BackColor = System.Drawing.Color.Green;
                button3.ForeColor = System.Drawing.Color.White;
            }
        } //FOLDER FETCH

        private void button5_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("End current session?", "End Session?", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dialogResult == DialogResult.Yes)
            {
                escape = true;
                bw.CancelAsync();
                Close();
            }
            else if (dialogResult == DialogResult.No)
            {

            }

        } //END SESSION

        private void button6_Click(object sender, EventArgs e)
        {
            int i = 0;
            
            if (treeView1.Nodes.Count != 0)
            {
                string name = treeView1.SelectedNode.ToString();
                Console.WriteLine(name);
                List<String> directories = browsePop.directoryName.OfType<string>().ToList();
                while (browsePop.directoryName[i] != null)
                {
                    if (browsePop.directoryName[i] == name)
                    {
                        directories.Remove(i.ToString());

                    }
                    i++;
                }
                treeView1.SelectedNode.Remove();

                browsePop.directoryName = directories.ToArray();
            }
        } //REMOVE SELECTED FILES

        public void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Setting form2 = new Setting();
            form2.Show();
            if (Properties.Settings.Default.PDFPath == "" || Properties.Settings.Default.PDFPath == null)
            {
                label3.Visible = true;
                radioButton10.Enabled = false;
                radioButton11.Enabled = false;
            }
            else
            {
                label3.Visible = false;
                radioButton10.Enabled = true;
                radioButton11.Enabled = true;
            }

        } //PREFERECES

        private void helpToolStripMenuItem_Click(object sender, EventArgs e)
        {
            help Help = new help();
            Help.Show();
        } //HELP

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            whatStatus = "0";
        } //WHATSTATUS

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            whatStatus = "PRELIMINARY";
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            whatStatus = "MANUFACTURING";
        }

        private void radioButton4_CheckedChanged(object sender, EventArgs e)
        {
            whatStatus = "AS SHIPPED";
        }

        private void radioButton5_CheckedChanged(object sender, EventArgs e)
        {
            whatStatus = "FIELD CHANGES";
        }

        private void radioButton6_CheckedChanged(object sender, EventArgs e)
        {
            whatStatus = "AS INSTALLED";
        }

        private void radioButton7_CheckedChanged(object sender, EventArgs e)
        {
            whatStatus = "DELETE IF PRESENT";
        }

        private void radioButton8_CheckedChanged(object sender, EventArgs e)
        {
            printComb = 0;
        } //PRINTING

        private void radioButton11_CheckedChanged(object sender, EventArgs e)
        {
            printComb = 1;
        } //for all

        private void radioButton10_CheckedChanged(object sender, EventArgs e)
        {
            printComb = 2;
        } //by book

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if(checkBox1.Checked == true)
            {
                if (File.Exists(Properties.Settings.Default.SupportFilePath + "\\TITLE-PAGE.VSD"))
                {
                    titlePage = true;
                    
                }
                else
                {
                    checkBox1.Checked = false;
                    titlePage = true;
                    DialogResult dialogResult = MessageBox.Show("The Title page supporting files are missing.\n" +
                    "Please verify if the files exist and the correct filepath was referenced.", "Missing Files",
                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                    Setting form2 = new Setting();
                    form2.Show();
                }
            }
            else
            {
                titlePage = false;
            }
        } //PAGE ADDITIONS

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked == true)
            {
                if (File.Exists(Properties.Settings.Default.SupportFilePath + "\\012.VSD") && File.Exists(Properties.Settings.Default.SupportFilePath + "\\013.VSD")
                    && File.Exists(Properties.Settings.Default.SupportFilePath + "\\014.VSD") && File.Exists(Properties.Settings.Default.SupportFilePath + "\\015.VSD"))
                {
                    howToRead = true;
                    
                }
                else
                {

                    checkBox2.Checked = false;
                    howToRead = false;
                    DialogResult dialogResult = MessageBox.Show("The How-to-Read supporting files are missing.\n" +
                    "Please verify if the files exist and the correct filepath was referenced.", "Missing Files",
                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                    Setting form2 = new Setting();
                    form2.Show();
                }
                
            }
            else
            {
                howToRead = false;
            }
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox3.Checked == true)
            {
                if (File.Exists(Properties.Settings.Default.SupportFilePath + "\\REV-SHEET.vsd"))
                {
                    revPage = true;
                }
                else
                {
                    checkBox3.Checked = false;
                    revPage = false;
                    DialogResult dialogResult = MessageBox.Show("The Revision page supporting file is missing.\n" +
                    "Please verify if the file exists and the correct filepath was referenced.", "Missing File",
                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                    Setting form2 = new Setting();
                    form2.Show();
                }
                
            }
            else
            {
                revPage = false;
            }
        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox4.Checked == true)
            {
                if (File.Exists(Properties.Settings.Default.SupportFilePath + "\\BACK-PAGE.vsd"))
                {
                    backPage = true;
                }
                else
                {
                    checkBox4.Checked = false;
                    backPage = false;
                    DialogResult dialogResult = MessageBox.Show("The Back page supporting file is missing.\n" +
                    "Please verify if the file exists and the correct filepath was referenced.", "Missing File",
                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                    Setting form2 = new Setting();
                    form2.Show();
                }
                
            }
            else
            {
                backPage = false;
            }
        }

        private void checkBox5_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox5.Checked == true)
            {
                cloudTriangle = true;
            }
            else
            {
                cloudTriangle = false;
            }
        } //ACTIONS

        private void checkBox6_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox6.Checked == true)
            {
                if(File.Exists(Properties.Settings.Default.SupportFilePath + "\\CONTENT.vsd"))
                {
                    tableOfContents = true;
                }
                else
                {
                    checkBox6.Checked = false;
                    tableOfContents = false;
                    DialogResult dialogResult = MessageBox.Show("The Table of Contents supporting file is missing.\n" +
                    "Please verify if the file exists and the correct filepath was referenced.", "Missing File",
                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                    Setting form2 = new Setting();
                    form2.Show();
                }
                
            }
            else
            {
                tableOfContents = false;
            }
        } 

        private void checkBox7_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox7.Checked == true)
            {
                drawNumCheck = true;
            }
            else
            {
                drawNumCheck = false;
            }
        }

        private void checkBox8_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox8.Checked == true)
            {
                dsgnByChkBy = true;
            }
            else
            {
                dsgnByChkBy = false;
            }
        }

        private void checkBox9_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox9.Checked == true)
            {
                sheetNumCheck = true;
            }
            else
            {
                sheetNumCheck = false;
            }
        }

        private void checkBox10_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox10.Checked == true)
            {
                sheetMargins = true;
            }
            else
            {
                sheetMargins = false;
            }
        }

        //BACKGROUND WORKER CLASSES

        private void bw_DoWork(object sender, DoWorkEventArgs e)
        {
            bool done = false;
            BackgroundWorker worker = (BackgroundWorker)sender;

            progressInc = 100 / (totalAction);
            progressCur = 0;

            while (escape != true && done == false)
            {
                if (Properties.Settings.Default.Backups == true)
                {
                    backup();
                    progressCur = progressCur + progressInc;
                    toolStripProgressBar2.GetCurrentParent().Invoke(new MethodInvoker(delegate
                    {
                        //toolStripProgressBar2.Value = progressCur;
                    }));

                }
                if ((titlePage || backPage || revPage || howToRead) != false)
                {
                    
                    insertDoc();
                    progressCur = progressCur + progressInc;
                    toolStripProgressBar2.GetCurrentParent().Invoke(new MethodInvoker(delegate
                    {
                        toolStripProgressBar2.Value = progressCur;
                    }));
                    
                }  

                if (whatStatus != "0")
                {
                    
                    changeStatus();
                    progressCur = progressCur + progressInc;
                    toolStripProgressBar2.GetCurrentParent().Invoke(new MethodInvoker(delegate
                    {
                        toolStripProgressBar2.Value = progressCur;
                    }));
                }

                if(cloudTriangle == true)
                {
                    
                    removeObjects();
                    progressCur = progressCur + progressInc;
                    toolStripProgressBar2.GetCurrentParent().Invoke(new MethodInvoker(delegate
                    {
                        toolStripProgressBar2.Value = progressCur;
                    }));
                }
                
                if(tableOfContents == true)
                {
                    
                    tableOfCont();
                    progressCur = progressCur + progressInc;
                    toolStripProgressBar2.GetCurrentParent().Invoke(new MethodInvoker(delegate
                    {
                        toolStripProgressBar2.Value = progressCur;
                    }));
                }

                if(drawNumCheck == true)
                {
                    
                    drawNoCheck();
                    progressCur = progressCur + progressInc;
                    toolStripProgressBar2.GetCurrentParent().Invoke(new MethodInvoker(delegate
                    {
                        toolStripProgressBar2.Value = progressCur;
                    }));
                }

                if(dsgnByChkBy == true)
                {
                    
                    namesCheck();
                    progressCur = progressCur + progressInc;
                    toolStripProgressBar2.GetCurrentParent().Invoke(new MethodInvoker(delegate
                    {
                        toolStripProgressBar2.Value = progressCur;
                    }));
                }

                if(sheetNumCheck == true)
                {
                    
                    sheetCheck();
                    progressCur = progressCur + progressInc;
                    toolStripProgressBar2.GetCurrentParent().Invoke(new MethodInvoker(delegate
                    {
                        toolStripProgressBar2.Value = progressCur;
                    }));
                }

                if(sheetMargins == true)
                {
                    
                    marginCheck();
                    progressCur = progressCur + progressInc;
                    toolStripProgressBar2.GetCurrentParent().Invoke(new MethodInvoker(delegate
                    {
                        toolStripProgressBar2.Value = progressCur;
                    }));
                }

                if (printComb > 0)
                {

                    printingDoc(printComb);
                    progressCur = progressCur + progressInc;
                    toolStripProgressBar2.GetCurrentParent().Invoke(new MethodInvoker(delegate
                    {
                        toolStripProgressBar2.Value = progressCur;
                    }));
                }

                if (Properties.Settings.Default.Notification == false && escape != true)
                {
                    
                    toolStripProgressBar2.GetCurrentParent().Invoke(new MethodInvoker(delegate
                    {
                        message("Task completed");
                        toolStripProgressBar2.Value = 0;
                        toolStripStatusLabel1.Text = "Task complete";
                        bw.ReportProgress(0);

                        if (WindowState == FormWindowState.Minimized)
                        {
                            WindowState = FormWindowState.Normal;
                        }
                        BringToFront();
                    }));

                    //bw.ReportProgress(0, 0);
                }

                button3.Invoke((MethodInvoker)delegate
                {
                    button3.BackColor = System.Drawing.Color.Green;
                    button3.Text = "EXECUTE";
                    button3.ForeColor = System.Drawing.Color.White;
                    tabControl1.Enabled = true;
                    toolStripProgressBar2.Value = 0;
                    bw.ReportProgress(0);

                    button1.Enabled = true;
                    button2.Enabled = true;
                    button4.Enabled = true;
                    button6.Enabled = true;

                    

                });
                done = true;

                totalAction = 0;

            }

            if (escape == true && done == false)
            {
                toolStripProgressBar2.GetCurrentParent().Invoke(new MethodInvoker(delegate
                {
                    button3.BackColor = System.Drawing.Color.Green;
                    button3.Text = "EXECUTE";
                    button3.ForeColor = System.Drawing.Color.White;

                    button1.Enabled = true;
                    button2.Enabled = true;

                    button4.Enabled = true;
                    button6.Enabled = true;
                    tabControl1.Enabled = true;

                    toolStripProgressBar2.Value = 0;
                    
                    bw.ReportProgress(0);
                }));

                //bw.ReportProgress(0, 0);
            }

            escape = false;
        }

        private void bw_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            toolStripProgressBar1.Value = e.ProgressPercentage;
            //toolStripProgressBar2.Value = (int)e.UserState;
        }

        //COMMAND CLASSES

        //https://stackoverflow.com/questions/31896952/how-to-programmatically-print-to-pdf-file-without-prompting-for-filename-in-c-sh

        private void insertDoc()
        {
            var visioChecks = new visioChecks();
            string[] index = new string[20];
            int indCount = 0;
            int i = 0;
            int j = 0;
            int population = 0;

            Console.WriteLine("Inserting documents...");
            toolStripStatusLabel1.Text = "Inserting documents";

            string supLoc = Properties.Settings.Default.SupportFilePath + "\\";

            foreach (TreeNode node in treeView1.Nodes)
            {
                index[indCount] = node.Text;
                Console.WriteLine(node);
                indCount++;
            }
            indivInc = 0;
            indivProg = 0;
            indivInc = 100 / indCount;

            Invoke(new MethodInvoker(delegate
            {
                this.Focus();
            }));
            System.Threading.Thread.Sleep(2000);

            toolStripProgressBar1.GetCurrentParent().Invoke(new MethodInvoker(delegate
            {
                toolStripProgressBar1.Style = ProgressBarStyle.Marquee;
            }));

            if ((titlePage || backPage || revPage || howToRead) != false)
            {
                while (j != indCount && escape == false)
                {
                    Console.WriteLine("Current node: " + j + " of " + indCount);
                    Console.WriteLine("folder name: " + index[j]);
                    
                    Console.WriteLine("copy path: " + supLoc);
                    string path = browsePop.directoryName[j];
                    int largeNum = tr.findFileRange(path);
                    
                    tr.Inserts(supLoc, path, largeNum, titlePage, backPage, revPage, howToRead, ref toolStripStatusLabel1);
                    j++;
                    Console.WriteLine(j);
                    indivProg = indivProg + indivInc;
                    bw.ReportProgress(indivProg);
                }
                visioChecks.endSession();
            }
            toolStripProgressBar1.GetCurrentParent().Invoke(new MethodInvoker(delegate
            {
                toolStripProgressBar1.Style = ProgressBarStyle.Continuous;
            }));
            

        }

        private void sheetCheck()
        {
            var visioChecks = new visioChecks();
            string[] index = new string[20];
            int indCount = 0;
            int i = 0;
            int j = 0;
            int population = 0;

            Console.WriteLine("Checking sheet numbers...");
            toolStripStatusLabel1.Text = "Checking sheet numbers";

            foreach (TreeNode node in treeView1.Nodes)
            {
                index[indCount] = node.Text;
                Console.WriteLine(node);
                indCount++;
            }

            while (j != indCount && escape == false)
            {
                Console.WriteLine("Current node: " + j + " of " + browsePop.currentNodes);
                Console.WriteLine("folder name: " + browsePop.directoryName[j]);

                population = treeView1.Nodes[j].GetNodeCount(true);
                indivInc = 0;
                indivProg = 0;
                indivInc = 100 / population;

                Invoke(new MethodInvoker(delegate
                {
                    this.Focus();
                }));
                System.Threading.Thread.Sleep(2000);

                while (i < population && escape == false)
                {
                    indivProg = indivProg + indivInc;
                    bw.ReportProgress(indivProg);
                    Console.WriteLine("population value: " + population);
                    Console.WriteLine("file name: " + treeView1.Nodes[j].Nodes[i].Text);
                    string path = browsePop.directoryName[j] + "\\" + treeView1.Nodes[j].Nodes[i].Text;

                    string name = System.IO.Path.GetFileNameWithoutExtension(path);
                    visioChecks.sheetCheck(path, name);
                    i++;
                }
                j++;
                i = 0;
                bw.ReportProgress(100);
                visioChecks.endSession();

            }
            bw.ReportProgress(0);

        }

        private void backup()
        {
            string[] index = new string[20];
            int indCount = 0;
            int i = 0;
            int j = 0;
            int population = 0;
            
            Console.WriteLine("Creating backups...");
            toolStripStatusLabel1.Text = "Creating backups";

            foreach (TreeNode node in treeView1.Nodes)
            {
                index[indCount] = node.Text;
                Console.WriteLine(node);
                indCount++;
            }



            while (j != indCount && escape == false)
            {
                bw.ReportProgress(0);
                Console.WriteLine("Current node: " + j + " of " + browsePop.currentNodes);
                Console.WriteLine("folder name: " + browsePop.directoryName[j]);

                population = treeView1.Nodes[j].GetNodeCount(true);
                indivInc = 0;
                indivProg = 0;
                indivInc = 100 / population;

                Invoke(new MethodInvoker(delegate
                {
                    this.Focus();
                }));
                System.Threading.Thread.Sleep(2000);

                string timenow = DateTime.Now.ToString("yyyyMMddHHmm");
                string dest = browsePop.directoryName[j] + "\\backup" + timenow;

                bool exists = System.IO.Directory.Exists(dest);
                if (!exists)
                {
                    System.IO.Directory.CreateDirectory(dest);
                }

                while (i < population && escape == false)
                {
                    string cur = browsePop.directoryName[j] + "\\" + treeView1.Nodes[j].Nodes[i].Text;
                    Console.WriteLine(cur + " and " + dest + "\\" + treeView1.Nodes[j].Nodes[i].Text);
                    indivProg = indivProg + indivInc;
                    bw.ReportProgress(indivProg);
                    Console.WriteLine("population value: " + population);
                    Console.WriteLine("file name: " + treeView1.Nodes[j].Nodes[i].Text);
                    File.Copy(cur, dest + "\\" + treeView1.Nodes[j].Nodes[i].Text, true);
                    File.SetAttributes(dest + "\\" + treeView1.Nodes[j].Nodes[i].Text, FileAttributes.Normal);
                    i++;
                }
                j++;
                i = 0;
                bw.ReportProgress(100);

            }
            
            bw.ReportProgress(0);

        }

        private void drawNoCheck() 
        {
            var visioChecks = new visioChecks();
            var question = new Form3();
            string[] index = new string[20];
            int indCount = 0;
            int i = 0;
            int j = 0;
            int population = 0;

            Console.WriteLine("Checking drawing numbers...");
            toolStripStatusLabel1.Text = "Checking drawing numbers";

            foreach (TreeNode node in treeView1.Nodes)
            {
                index[indCount] = node.Text;
                Console.WriteLine(node);
                indCount++;
            }
           

            while (j != indCount && escape == false)
            {
                string suggNo;
                Console.WriteLine("Current node: " + j + " of " + indCount);
                Console.WriteLine("folder name: " + index[indCount - 1]);

                if (index[indCount - 1].Contains('_'))
                {
                    suggNo = index[indCount - 1].Substring(0, index[indCount - 1].IndexOf('_'));
                }
                else
                {
                    suggNo = "Enter Here";
                }

                question.data(index[indCount - 1], suggNo);
                DialogResult result = question.ShowDialog();

                population = treeView1.Nodes[j].GetNodeCount(true);
                indivInc = 0;
                indivProg = 0;
                indivInc = 100 / population;

                Invoke(new MethodInvoker(delegate
                {
                    this.Focus();
                }));
                System.Threading.Thread.Sleep(2000);

                if (question.result == 1)
                {
                    j++;
                    i = 0;
                    Console.WriteLine("Terminated...");
                    return;
                }

                else if (question.result == 2)
                {
                    suggNo = question.drawNumber;
                    //question.Close();
                    Console.WriteLine("Proceeding...");

                    while (i < population && escape == false)
                    {
                        indivProg = indivProg + indivInc;
                        bw.ReportProgress(indivProg);

                        Console.WriteLine("population value: " + population);
                        Console.WriteLine("file name: " + treeView1.Nodes[j].Nodes[i].Text);
                        string path = browsePop.directoryName[j] + "\\" + treeView1.Nodes[j].Nodes[i].Text;
                        visioChecks.drawNoCheck(path, suggNo);
                        i++;
                    }
                    j++;
                    i = 0;
                    visioChecks.endSession();
                    bw.ReportProgress(0);
                }
                
            }
            bw.ReportProgress(0);

        }

        private void namesCheck()
        {
            var visioChecks = new visioChecks();
            string[] index = new string[20];
            int indCount = 0;
            int i = 0;
            int j = 0;
            int population = 0;
            indivProg = indivProg + indivInc;
            bw.ReportProgress(indivProg);

            Console.WriteLine("Checking names and dates...");
            toolStripStatusLabel1.Text = "Checking names and dates";

            foreach (TreeNode node in treeView1.Nodes)
            {
                index[indCount] = node.Text;
                Console.WriteLine(node);
                indCount++;
            }

            while (j != indCount && escape == false)
            {
                Console.WriteLine("Current node: " + j + " of " + browsePop.currentNodes);
                Console.WriteLine("folder name: " + browsePop.directoryName[j]);

                population = treeView1.Nodes[j].GetNodeCount(true);
                indivInc = 0;
                indivProg = 0;
                indivInc = 100 / population;

                Invoke(new MethodInvoker(delegate
                {
                    this.Focus();
                }));
                System.Threading.Thread.Sleep(2000);

                while (i < population && escape == false)
                {
                    indivProg = indivProg + indivInc;
                    bw.ReportProgress(indivProg);

                    Console.WriteLine("population value: " + population);
                    Console.WriteLine("file name: " + treeView1.Nodes[j].Nodes[i].Text);
                    string path = browsePop.directoryName[j] + "\\" + treeView1.Nodes[j].Nodes[i].Text;
                    visioChecks.designCheck(path);
                    i++;
                }
                j++;
                i = 0;
                bw.ReportProgress(100);
                
            }
            visioChecks.endSession();
            bw.ReportProgress(0);

        }

        private void marginCheck()
        {
            var visio07Edit = new visio07Edit();
            string[] index = new string[20];
            int indCount = 0;
            int i = 0;
            int j = 0;
            int population = 0;

            Console.WriteLine("Checking margins...");
            toolStripStatusLabel1.Text = "Checking margins";

            foreach (TreeNode node in treeView1.Nodes)
            {
                index[indCount] = node.Text;
                Console.WriteLine(node);
                indCount++;
            }
            

            while (j != indCount && escape == false)
            {
                

                Console.WriteLine("Current node: " + j + " of " + indCount);
                Console.WriteLine("folder name: " + index[indCount - 1]);

                population = treeView1.Nodes[j].GetNodeCount(true);
                indivInc = 0;
                indivProg = 0;
                indivInc = 100 / population;

                Invoke(new MethodInvoker(delegate
                {
                    this.Focus();
                }));
                System.Threading.Thread.Sleep(2000);

                while (i < population && escape == false)
                {
                    indivProg = indivProg + indivInc;
                    bw.ReportProgress(indivProg);

                    Console.WriteLine("population value: " + population);
                    Console.WriteLine("file name: " + treeView1.Nodes[j].Nodes[i].Text);
                    string path = browsePop.directoryName[j] + "\\" + treeView1.Nodes[j].Nodes[i].Text;
                    visio07Edit.visioMargins(path);
                    i++;
                }
                j++;
                bw.ReportProgress(100);
                i = 0;
            }
            visio07Edit.endSession();
            bw.ReportProgress(0);

        }

        private void changeStatus()
        {
            var visio07Edit = new visio07Edit();
            string[] index = new string[20];
            int indCount = 0;
            int i = 0;
            int j = 0;
            int population = 0;

            Console.WriteLine("Changing status...");
            toolStripStatusLabel1.Text = "Changing drawing status";

            foreach (TreeNode node in treeView1.Nodes)
            {
                index[indCount] = node.Text;
                Console.WriteLine(node);
                indCount++;
            }
            

            while (j != indCount && escape == false)
            {
                Console.WriteLine("Current node: " + j + " of " + indCount);
                Console.WriteLine("folder name: " + index[indCount - 1]);

                population = treeView1.Nodes[j].GetNodeCount(true);
                indivInc = 0;
                indivProg = 0;
                indivInc = 100 / population;

                Invoke(new MethodInvoker(delegate
                {
                    this.Focus();
                }));
                System.Threading.Thread.Sleep(2000);

                while (i < population && escape == false)
                {
                    indivProg = indivProg + indivInc;
                    bw.ReportProgress(indivProg);

                    Console.WriteLine("population value: " + population);
                    Console.WriteLine("file name: " + treeView1.Nodes[j].Nodes[i].Text);
                    string path = browsePop.directoryName[j] + "\\" + treeView1.Nodes[j].Nodes[i].Text;
                    visio07Edit.visioDrawingStat(path, whatStatus);
                    i++;
                }
                j++;
                bw.ReportProgress(100);
                
                i = 0;
            }
            visio07Edit.endSession();
            bw.ReportProgress(0);

        }

        private void removeObjects()
        {
            var visio07Edit = new visio07Edit();
            string[] index = new string[20];
            int indCount = 0;
            int i = 0;
            int j = 0;
            int population = 0;

            Console.WriteLine("Removing objects...");
            toolStripStatusLabel1.Text = "Removing clouds and triangles";

            foreach (TreeNode node in treeView1.Nodes)
            {
                index[indCount] = node.Text;
                Console.WriteLine(node);
                indCount++;
            }
            

            while (j != indCount && escape == false)
            {
                Console.WriteLine("Current node: " + j + " of " + indCount);
                Console.WriteLine("folder name: " + index[indCount - 1]);

                population = treeView1.Nodes[j].GetNodeCount(true);
                indivInc = 0;
                indivProg = 0;
                indivInc = 100 / population;

                Invoke(new MethodInvoker(delegate
                {
                    this.Focus();
                }));
                System.Threading.Thread.Sleep(2000);

                while (i < population && escape == false)
                {
                    indivProg = indivProg + indivInc;
                    bw.ReportProgress(indivProg);

                    Console.WriteLine("population value: " + population);
                    Console.WriteLine("file name: " + treeView1.Nodes[j].Nodes[i].Text);
                    string path = browsePop.directoryName[j] + "\\" + treeView1.Nodes[j].Nodes[i].Text;
                    visio07Edit.visioShapeRemove(path);
                    i++;
                }
                j++;
                bw.ReportProgress(100);
                i = 0;
            }
            visio07Edit.endSession();
            bw.ReportProgress(0);
        }

        private void printingDoc(int printComb)
        {
            string[] index = new string[20];
            int indCount = 0;
            int i = 0;
            int j = 0;
            int population = 0;

            Setting form = new Setting();
            printer printer = new printer();

            Console.WriteLine("Creating individual PDFs...");
            toolStripStatusLabel1.Text = "Creating PDFs of folders";

            foreach (TreeNode node in treeView1.Nodes)
            {
                index[indCount] = node.Text;
                Console.WriteLine(node);
                indCount++;
            }
            


            while (j != indCount && escape == false)
            {
                Console.WriteLine("Current node: " + j + " of " + indCount);
                Console.WriteLine("folder name: " + index[indCount - 1]);
                string dir = browsePop.directoryName[j];
                population = Directory.GetFiles(browsePop.directoryName[j], "*", SearchOption.TopDirectoryOnly).Length;
                indivInc = 0;
                indivProg = 0;
                indivInc = 100 / population;

                            

                foreach (string file in Directory.GetFiles(browsePop.directoryName[j], "*", SearchOption.TopDirectoryOnly))
                {
                    FileInfo f = new FileInfo(file);
                    string CurFolder = f.Directory.Name;

                    if (System.IO.Path.GetFileName(file).Contains("~") || (!System.IO.Path.GetExtension(file).ToUpper().Contains("VSD")))
                    {
                        indivProg = indivProg + indivInc;
                        Console.WriteLine(indivInc + " " + indivProg);
                        bw.ReportProgress(indivProg);

                        Console.WriteLine("HI");
                        //skip the odd files
                    }
                    else
                    {
                        Invoke(new MethodInvoker(delegate
                        {
                            this.Focus();
                        }));
                        System.Threading.Thread.Sleep(2000);

                        indivProg = indivProg + indivInc;
                        Console.WriteLine(indivInc + " " + indivProg);
                        bw.ReportProgress(indivProg);

                        Console.WriteLine("population value: " + population);
                        Console.WriteLine("file name: " + file);
                        
                        printer.exporting(file, file, index[j], j);
                    }
                }
                printer.compress(index[j]);
                bw.ReportProgress(100);
                j++;
                
                i = 0;
            }

            toolStripProgressBar1.GetCurrentParent().Invoke(new MethodInvoker(delegate
            {
                toolStripProgressBar1.Style = ProgressBarStyle.Marquee;
            }));
            if (printComb == 1) //for all
            {
                printer.combineSet(indCount);
            }
            else if(printComb == 2) //by book
            {
                
            }
            toolStripProgressBar1.GetCurrentParent().Invoke(new MethodInvoker(delegate
            {
                toolStripProgressBar1.Style = ProgressBarStyle.Continuous;
            }));
            bw.ReportProgress(0);
            printer.endSession();
            
        }

        private void tableOfCont()
        {
            var question = new DBCB();
            var visio07Edit = new visio07Edit();
            string[] index = new string[20];
            int indCount = 0;
            int j = 0;
            int population = 0;
            int totalPop = 0;
            visio07Edit.i = 0;


            Console.WriteLine("Working on Table of Contents...");
            toolStripStatusLabel1.Text = "Working on Table of Contents";

            foreach (TreeNode node in treeView1.Nodes)
            {
                index[indCount] = node.Text;
                Console.WriteLine(node);
                indCount++;
            }


            toolStripProgressBar1.GetCurrentParent().Invoke(new MethodInvoker(delegate
            {
                toolStripProgressBar1.Style = ProgressBarStyle.Marquee;
            }));


            while (j != indCount && escape == false)
            {
                visio07Edit.initialize();
                
                Console.WriteLine("Current node: " + j + " of " + browsePop.currentNodes);
                Console.WriteLine("folder name: " + browsePop.directoryName[j]);

                string[] folder = Directory.GetFiles(browsePop.directoryName[j] + "\\");
                population = 0;

                question.result = 0;
                while (question.result == 0 && escape == false)
                {
                    DialogResult result = question.ShowDialog();
                }
                question.result = 0;

                Console.WriteLine("Hello" + question.result);
                if (question.result == 1)
                {
                    question.dbdate = "";
                    question.cbname = "";

                    question.dbname = "";
                    question.cbname = "";
                }

                else if (question.result == 2)
                {
                    question.cbdate = question.cbdate.ToUpper();
                    question.dbdate = question.dbdate.ToUpper();
                }

                Invoke(new MethodInvoker(delegate
                {
                    this.Focus();
                }));

                System.Threading.Thread.Sleep(2000);

                visio07Edit.buildTOC(ref toolStripStatusLabel1, browsePop.directoryName[j] + "\\", j, totalPop, question.dbname, question.cbname, question.dbdate, question.cbdate, escape);
                Console.WriteLine(browsePop.directoryName[indCount]);
                
                   
                foreach (string filename in folder)
                {
                    if (System.IO.Path.GetExtension(filename) == ".VSD" || System.IO.Path.GetExtension(filename) == ".vsd")
                    {
                        visio07Edit.gatherTOC(ref toolStripStatusLabel1, filename, j, population);
                        totalPop = population;
                    }
                    else
                    {

                    }
                    population++;

                }
                
                visio07Edit.fillTOC(ref toolStripStatusLabel1, browsePop.directoryName[j] + "\\", j, totalPop);
                j++;
                bw.ReportProgress(100);
                population = 0;
                visio07Edit.i++;
                bw.ReportProgress(0);
                
            }

            if(j == indCount || escape == true)
            {
                visio07Edit.endSession();
            }
            
            toolStripProgressBar1.GetCurrentParent().Invoke(new MethodInvoker(delegate
            {
                toolStripProgressBar1.Value = 0;
                toolStripProgressBar1.Style = ProgressBarStyle.Continuous;
            }));
            
        }

        private void countAction()
        {
            //Determining amount of work to do
            totalAction = 0;

            //if (Properties.Settings.Default.Backups == true) totalAction++;

            //Drawing Status
            if (whatStatus != "0") totalAction++;
            toolStripProgressBar1.Value = 25;

            Console.WriteLine(totalAction);

            //Drawing Inserts
            if (titlePage != false) totalAction++;
            if (backPage != false) totalAction++;
            if (revPage != false) totalAction++;
            if (howToRead != false) totalAction++;
            toolStripProgressBar1.Value = 50;

            Console.WriteLine(totalAction);

            //Drawing Actions
            if (cloudTriangle != false) totalAction++;
            if (tableOfContents != false) totalAction++;
            if (drawNumCheck != false) totalAction++;
            if (dsgnByChkBy != false) totalAction++;
            if (sheetNumCheck != false) totalAction++;
            if (sheetMargins != false) totalAction++;
            toolStripProgressBar1.Value = 75;

            Console.WriteLine(totalAction);

            //Print Actions
            if (printComb != 0) totalAction++;
            toolStripProgressBar1.Value = 100;

            Console.WriteLine(totalAction);
        }

        private void menuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (Properties.Settings.Default.PDFPath != "" || Properties.Settings.Default.PDFPath != null)
            {
                label3.Visible = false;
                radioButton10.Enabled = true;
                radioButton11.Enabled = true;
                
            }
            else
            {
                label3.Visible = true;
                radioButton10.Enabled = false;
                radioButton11.Enabled = false;
            }
        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

    }

    //For Visio 2003-2017
    //Collects files to edit, displays in treeView
    public class browsePopulate
    {
        public int[] fileAmount = new int[15];
        public string[][] fileName = new string[15][];
        public string[] folderName = new string[15];
        public string[] directoryName = new string[30];

        public int currentNodes = 0;
        public int totalFiles = 0;

        public static Form1 gui = new Form1();

        //Fetch files
        public void fileFetch(ref TreeView treeBuilder, ref ToolStripStatusLabel Label)
        {
            bool breakFlag = false;
            string[] index = new string[20];
            string[] number = new string[20];
            int indCount = 0;
            int search = 0;

            //Open file browser
            DialogResult result = gui.openFileDialog1.ShowDialog();

            //Count number of folders in treeview
            foreach(TreeNode node in treeBuilder.Nodes)
            {
                index[indCount] = node.Text;
                Console.WriteLine(node + " ok");
                indCount++;
            }

            //File selection area
            if (result == DialogResult.OK)
            {
                //Get list of files
                foreach (string file in gui.openFileDialog1.FileNames)
                {
                    breakFlag = false;
                    FileInfo f = new FileInfo(file);
                    string CurFolder = f.Directory.Name;

                    //Ignore hidden, weird .vsd files
                    if (System.IO.Path.GetFileName(file).Contains("~") || (!System.IO.Path.GetExtension(file).ToUpper().Contains("VSD")))
                    {
                        Console.WriteLine("Skipping File");
                    }
                    else
                    {
                        //Check to see if file was already added in a folder or not
                        while (!breakFlag && search < indCount + 1)
                        {
                            if (index[search] == CurFolder)
                            {
                                Console.WriteLine("found match at " + index[search] + CurFolder);

                                Console.WriteLine("add " + search);
                                treeBuilder.Nodes[search].Nodes.Add(System.IO.Path.GetFileName(file));
                                breakFlag = true;
                                break;
                            }
                            else if (index[search] != CurFolder && index[search] != null)
                            {
                                Console.WriteLine("skipping " + search);
                                search++;
                                breakFlag = false;
                            }
                            else if (index[search] != CurFolder && index[search] == null)
                            {
                                directoryName[search] = System.IO.Path.GetDirectoryName(file);
                                Console.WriteLine("found open at " + search + index[search]);
                                treeBuilder.Nodes.Add(CurFolder);
                                index[search] = CurFolder;
                                treeBuilder.Nodes[search].Nodes.Add(System.IO.Path.GetFileName(file));
                                breakFlag = true;
                                break;
                            }
                        }
                    }
                }
            }
            treeBuilder.ExpandAll();
        }

        //Fetch entire folder
        public void folderFetch(ref TreeView treeBuilder, ref ToolStripStatusLabel Label)
        {
            bool breakFlag = false;
            string[] index = new string[20];
            string[] number = new string[20];
            int indCount = 0;
            int search = 0;

            Microsoft.WindowsAPICodePack.Dialogs.CommonOpenFileDialog FolderDi = new Microsoft.WindowsAPICodePack.Dialogs.CommonOpenFileDialog();

            //Open folder browser
            FolderDi.IsFolderPicker = true;

            //Count current folders in treeview
            foreach (TreeNode node in treeBuilder.Nodes)
            {
                index[indCount] = node.Text;
                Console.WriteLine(node + " ok");
                indCount++;
            }

            if (FolderDi.ShowDialog() == CommonFileDialogResult.Ok)
            {
                //Get all files in folder
                foreach (string file in Directory.GetFiles(FolderDi.FileName))
                {
                    breakFlag = false;
                    FileInfo f = new FileInfo(file);
                    string CurFolder = f.Directory.Name;

                    //Skip odd hidden .vsd files
                    if (System.IO.Path.GetFileName(file).Contains("~") || (!System.IO.Path.GetExtension(file).ToUpper().Contains("VSD")))
                    {
                        Console.WriteLine("Skipping File");
                    }
                    else
                    {
                        //Check for already inputted files
                        while (!breakFlag && search < indCount + 1)
                        {
                            if (index[search] == CurFolder)
                            {
                                Console.WriteLine("found match at " + index[search] + CurFolder);

                                Console.WriteLine("add " + search);
                                treeBuilder.Nodes[search].Nodes.Add(System.IO.Path.GetFileName(file));
                                breakFlag = true;
                                break;
                            }
                            else if (index[search] != CurFolder && index[search] != null)
                            {
                                Console.WriteLine("skipping " + search);
                                search++;
                                breakFlag = false;
                            }
                            else if (index[search] != CurFolder && index[search] == null)
                            {
                                directoryName[search] = System.IO.Path.GetDirectoryName(file);
                                Console.WriteLine("found open at " + search + index[search]);
                                treeBuilder.Nodes.Add(CurFolder);
                                index[search] = CurFolder;
                                treeBuilder.Nodes[search].Nodes.Add(System.IO.Path.GetFileName(file));
                                breakFlag = true;
                                break;
                            }
                        }
                    }
                }
            }
            treeBuilder.ExpandAll();
        }
    }

    //For Visio 2003-2017
    //Copies files into project folder(s)
    public class transplant
    {
        Setting form2 = new Setting();

        //Insert files        
        public void Inserts(string beginLoc, string destLoc, int largestNum, bool titlePage, bool backPage, bool revPage, bool howToPage, ref ToolStripStatusLabel Label)
        {
            string suggNo;
            var visioChecks = new visioChecks();
            Console.WriteLine("Inserts");
            var question = new Form3();
            Console.WriteLine(beginLoc + " >> " + destLoc);

            FileInfo f = new FileInfo(destLoc);
            string curFol = System.IO.Path.GetFileName(destLoc);

            //Automatically suggest current folder
            if (curFol.Contains('_'))
            {
                suggNo = curFol.Substring(0, curFol.IndexOf('_'));
            }
            else
            {
                suggNo = "Enter Here";
            }

            destLoc = destLoc + "\\";

            //Open current folder prompt
            question.data(curFol, suggNo);
            DialogResult result = question.ShowDialog();

            string supLoc = "";

            //title page insert
            if (titlePage == true)
            {
                if (question.result == 1)
                {
                    Console.WriteLine("No drawing numbers Updated...");
                    Label.Text = "Inserting title page";
                    supLoc = beginLoc + "TITLE-PAGE.vsd";

                    Console.WriteLine(beginLoc);
                    Console.WriteLine(supLoc);

                    File.Copy(supLoc, destLoc + "001.vsd", true);
                    File.SetAttributes(destLoc, FileAttributes.Normal);
                    supLoc = "";
                }


                if (question.result == 2)
                {
                    Label.Text = "Inserting title page";
                    supLoc = beginLoc + "TITLE-PAGE.vsd";
                    File.Copy(supLoc, destLoc + "001.vsd", true);
                    File.SetAttributes(destLoc, FileAttributes.Normal);
                    supLoc = "";

                    suggNo = question.drawNumber;
                    //question.Close();
                    Console.WriteLine("Proceeding to update drawing numbers...");
                    string path = destLoc + "001.vsd";
                    visioChecks.drawNoCheck(path, suggNo);
                }
                
            }

            //how to page insert
            if (howToPage == true)
            {
                if (question.result == 1)
                {
                    Console.WriteLine("No drawing numbers Updated...");
                    Label.Text = "Inserting How-to-Read pages";
                    supLoc = beginLoc + "012.vsd";
                    File.Copy(supLoc, destLoc + "012.vsd", true);
                    File.SetAttributes(destLoc, FileAttributes.Normal);
                    supLoc = beginLoc + "013.vsd";
                    File.Copy(supLoc, destLoc + "013.vsd", true);
                    File.SetAttributes(destLoc, FileAttributes.Normal);
                    supLoc = beginLoc + "014.vsd";
                    File.Copy(supLoc, destLoc + "014.vsd", true);
                    File.SetAttributes(destLoc, FileAttributes.Normal);
                    supLoc = beginLoc + "015.vsd";
                    File.Copy(supLoc, destLoc + "015.vsd", true);
                    File.SetAttributes(destLoc, FileAttributes.Normal);
                }


                if (question.result == 2)
                {
                    Label.Text = "Inserting How-to-Read pages";
                    Console.WriteLine("Proceeding to update drawing numbers...");
                    supLoc = beginLoc + "012.vsd";
                    File.Copy(supLoc, destLoc + "012.vsd", true);
                    File.SetAttributes(destLoc, FileAttributes.Normal);
                    string path = destLoc + "012.vsd";
                    visioChecks.drawNoCheck(path, suggNo);
                    supLoc = beginLoc + "013.vsd";
                    File.Copy(supLoc, destLoc + "013.vsd", true);
                    File.SetAttributes(destLoc, FileAttributes.Normal);
                    path = destLoc + "013.vsd";
                    visioChecks.drawNoCheck(path, suggNo);
                    supLoc = beginLoc + "014.vsd";
                    File.Copy(supLoc, destLoc + "014.vsd", true);
                    File.SetAttributes(destLoc, FileAttributes.Normal);
                    path = destLoc + "013.vsd";
                    visioChecks.drawNoCheck(path, suggNo);
                    supLoc = beginLoc + "015.vsd";
                    File.Copy(supLoc, destLoc + "015.vsd", true);
                    File.SetAttributes(destLoc, FileAttributes.Normal);
                    path = destLoc + "013.vsd";
                    visioChecks.drawNoCheck(path, suggNo);
                    suggNo = question.drawNumber;
                    //question.Close();
                    
                }
            }

            //revision page insert
            if (revPage == true)
            {
                if (question.result == 1)
                {
                    Console.WriteLine("No drawing numbers Updated...");
                    Label.Text = "Inserting Revision page";
                    supLoc = beginLoc + "REV-SHEET.vsd";
                    Console.WriteLine(destLoc + largestNum + "96.vsd");
                    File.Copy(supLoc, destLoc + largestNum + "96.vsd", true);
                    File.SetAttributes(destLoc, FileAttributes.Normal);
                    supLoc = "";
                }


                if (question.result == 2)
                {
                    Label.Text = "Inserting Revision page";
                    supLoc = beginLoc + "REV-SHEET.vsd";
                    Console.WriteLine(destLoc + largestNum + "96.vsd");
                    File.Copy(supLoc, destLoc + largestNum + "96.vsd", true);
                    File.SetAttributes(destLoc, FileAttributes.Normal);
                    supLoc = "";

                    suggNo = question.drawNumber;
                    //question.Close();
                    Console.WriteLine("Proceeding to update drawing numbers...");
                    string path = destLoc + largestNum + "96.vsd";
                    visioChecks.drawNoCheck(path, suggNo);
                    
                }
            }

            //last page insert

            if (backPage == true)
            {
                if (question.result == 1)
                {
                    Console.WriteLine("No drawing numbers Updated...");
                    Label.Text = "Inserting Back Page";
                    supLoc = beginLoc + "BACK-PAGE.vsd";
                    File.Copy(supLoc, destLoc + largestNum + "99.vsd", true);
                    File.SetAttributes(destLoc, FileAttributes.Normal);
                    supLoc = "";
                }


                if (question.result == 2)
                {
                    Label.Text = "Inserting Back Page";
                    supLoc = beginLoc + "BACK-PAGE.vsd";
                    File.Copy(supLoc, destLoc + largestNum + "99.vsd", true);
                    File.SetAttributes(destLoc, FileAttributes.Normal);
                    supLoc = "";

                    suggNo = question.drawNumber;
                    //question.Close();
                    Console.WriteLine("Proceeding to update drawing numbers...");
                    string path = destLoc + largestNum + "99.vsd";
                    visioChecks.drawNoCheck(path, suggNo);
                    
                }
                
            }

            
            
            beginLoc = "";
        }

        // figure out last file number (e.g. 699.vsd -> 699)
        public int findFileRange(string loc)
        {
            string file = "";
            int Num = 0;
            int largestNum = 0;
            string [] files = Directory.GetFiles(loc);
            
            foreach (string filename in files)
            {
                Console.WriteLine(System.IO.Path.GetExtension(filename));
                if (System.IO.Path.GetExtension(filename) == ".VSD" || System.IO.Path.GetExtension(filename) == ".vsd")
                {
                    file = System.IO.Path.GetFileNameWithoutExtension(filename).Substring(0, 1);
                    Num = Convert.ToInt32(file);
                    Console.WriteLine(Num);

                    if (Num >= largestNum)
                    {
                        largestNum = Num;
                        Console.WriteLine(largestNum);
                    }
                }
                else
                {

                }

            }
            
            return largestNum;

        }
    }

    //Printer actions for Visio 2003-2007
    //ExportAsFixedFormat details stripped from MSDN
    public class printer
    {
        Form1 main = new Form1();
        public static int i = 0;
        private static string[] pdfFileName = new string[15];

        Visio.InvisibleApp vApp = new Visio.InvisibleApp();


        //combine pdf files
        private void combineFile(string file, string addition)
        {
            try
            {
                
                PdfSharp.Pdf.PdfDocument main = PdfSharp.Pdf.IO.PdfReader.Open(file, PdfDocumentOpenMode.Modify);
                PdfSharp.Pdf.PdfDocument concat = PdfSharp.Pdf.IO.PdfReader.Open(addition, PdfDocumentOpenMode.Import);
                main.AddPage(concat.Pages[0]);

                main.Save(file);
                main.Close();
                
                File.Delete(addition);
            }
            catch (Exception ex)
            {
                DialogResult dialogResult = MessageBox.Show((ex.Message + "     " + ex.InnerException), "Error", MessageBoxButtons.OK, MessageBoxIcon.Hand);

            }
        }

        //combine completed sets
        public void combineSet(int currentNodes)
        {
            int j = 0;
            Console.WriteLine("Merging all...");
            //PdfSharp.Pdf.PdfDocument main = PdfReader.Open(pdfFileName[0][0], PdfDocumentOpenMode.Modify);
            Console.WriteLine(pdfFileName[0]);
            Console.WriteLine(currentNodes - 1);
            Console.WriteLine(j);

            try
            {
                while (j != (currentNodes - 1))
                {
                    j++;
                    //PdfSharp.Pdf.PdfDocument concat = PdfReader.Open(pdfFileName[j][0], PdfDocumentOpenMode.Import);
                    //main.AddPage(concat.Pages[0]);
                    Console.WriteLine(pdfFileName[j]);

                    //main.Save(pdfFileName[0][0]);
                    File.Delete(pdfFileName[j]);
                }
                //main.Save(Properties.Settings.Default.PDFPath + "\\$safeprojectname$ Output.pdf");

                File.Delete(pdfFileName[0]);
            }
            catch (Exception ex)
            {
                DialogResult dialogResult = MessageBox.Show((ex.Message + "     " + ex.InnerException), "Error", MessageBoxButtons.OK, MessageBoxIcon.Hand);

            }
        }

        //attempt compression
        public void compress(string folder)
        {
            iTextSharp.text.pdf.PdfReader reader = new iTextSharp.text.pdf.PdfReader(Properties.Settings.Default.PDFPath + "\\" + folder + ".pdf");
            iTextSharp.text.pdf.PdfStamper stamper = new PdfStamper(reader, new FileStream(Properties.Settings.Default.PDFPath + "\\temp.pdf", FileMode.Create), PdfWriter.VERSION_1_5);


            int pageNum = reader.NumberOfPages;
            Console.WriteLine(pageNum);
            for (int i = 1; i <= pageNum; i++)
            {
                reader.SetPageContent(i, reader.GetPageContent(i));
            }
            stamper.SetFullCompression();
            stamper.Close();
            File.Delete(Properties.Settings.Default.PDFPath + "\\temp.pdf");
        }

        //export to pdf
        public void exporting(string filePath, string fileName, string folder, int j)
        {
            string extensionless;
            try
            {

                if (File.Exists(Properties.Settings.Default.PDFPath + "\\" + folder + ".pdf") == false)
                {
                    vApp.ScreenUpdating = Convert.ToInt16(false);
                    vApp.EventsEnabled = Convert.ToInt16(false);

                    Visio.Document vDoc = vApp.Application.Documents.Open(filePath);
                    Console.WriteLine("New File");
                    
                    pdfFileName[j] = (Properties.Settings.Default.PDFPath + "\\" + folder + ".pdf").ToString();


                    //vDoc.ExportAsFixedFormat(VisFixedFormatTypes.visFixedFormatXPS, pdfFileName[j][population], VisDocExIntent.visDocExIntentScreen, VisPrintOutRange.visPrintAll, IncludeBackground: true, ColorAsBlack: true);
                    vDoc.ExportAsFixedFormat(VisFixedFormatTypes.visFixedFormatPDF, pdfFileName[j], VisDocExIntent.visDocExIntentScreen, VisPrintOutRange.visPrintAll, IncludeDocumentProperties: false, IncludeBackground: true, ColorAsBlack: true);

                    vDoc.Saved = true; //Forces document to think it has saved itself(it has not)
                    vDoc.Close();

                    Console.WriteLine(pdfFileName[j][0]);
                }

                else if (File.Exists(Properties.Settings.Default.PDFPath + "\\" + folder + ".pdf") == true)
                {
                    pdfFileName[j] = Properties.Settings.Default.PDFPath + "\\" + folder + ".pdf";
                    vApp.ScreenUpdating = Convert.ToInt16(false);
                    vApp.EventsEnabled = Convert.ToInt16(false);
                    Console.WriteLine(filePath);
                    Visio.Document vDoc = vApp.Application.Documents.Open(filePath);
                    Console.WriteLine("Append File");
                    extensionless = System.IO.Path.GetFileNameWithoutExtension(filePath);
                    Console.WriteLine(extensionless);
                    string pdfAdd = (Properties.Settings.Default.PDFPath + "\\" + extensionless + ".pdf").ToString();


                    //vDoc.ExportAsFixedFormat(VisFixedFormatTypes.visFixedFormatXPS, pdfFileName[j], VisDocExIntent.visDocExIntentScreen, VisPrintOutRange.visPrintAll, IncludeBackground: true, ColorAsBlack: true);
                    vDoc.ExportAsFixedFormat(VisFixedFormatTypes.visFixedFormatPDF, pdfAdd, VisDocExIntent.visDocExIntentScreen, VisPrintOutRange.visPrintAll, IncludeDocumentProperties: false, IncludeBackground: true, ColorAsBlack: true);

                    vDoc.Saved = true; //Forces document to think it has saved itself(it has not)
                    vDoc.Close();

                    combineFile(pdfFileName[j], pdfAdd);
                }


                else
                {

                }

                
            }
            catch (Exception ex)
            {
                DialogResult dialogResult = MessageBox.Show((ex.Message + "     " + ex.InnerException), "Error", MessageBoxButtons.OK, MessageBoxIcon.Hand);

            }

        }

        //Clear com objects
        public void endSession()
        {
            vApp.Quit();
            Marshal.ReleaseComObject(vApp);

            var process = Process.GetProcessesByName("COM .exe").FirstOrDefault();
            if (process != null)
            {
                process.Kill();
            }
            Marshal.FinalReleaseComObject(vApp);
            vApp = null;
        }
    }

    //For Visio 2003-2007, currently being used
    //Performs drawing and sheet number checks
    public class visioChecks
    {
        Visio.InvisibleApp vApp = new Visio.InvisibleApp();

        //Sheet number check
        public void sheetCheck(string filePath, string name)
        {
            
            try
            {
                vApp.ScreenUpdating = Convert.ToInt16(false);
                vApp.Visible = false;
                vApp.EventsEnabled = Convert.ToInt16(false);

                Visio.Document vDoc = vApp.Application.Documents.Open(filePath);

                foreach (Visio.Page thisPage in vDoc.Pages)
                {
                    foreach (Visio.Shape thisShape in thisPage.Shapes)
                    {
                        if (thisShape.Name == "SHEET_NO_1" || thisShape.Name == "SHEET_NO_2")
                        {
                            if (thisShape.Text == name)
                            {

                            }
                            else
                            {
                                thisShape.Text = name;
                            }

                        }
                        else
                        {
                            //Keep looking
                        }
                    }
                }
                vDoc.Save();
                vDoc.Close();
            }
            catch (Exception ex)
            {
                DialogResult dialogResult = MessageBox.Show((ex.Message + "     " + ex.InnerException), "Error", MessageBoxButtons.OK, MessageBoxIcon.Hand);
            }
        }

        //Crawing number check
        public void drawNoCheck(string filePath, string drawNumber)
        {
            try
            {
                vApp.ScreenUpdating = Convert.ToInt16(false);
                vApp.Visible = false;
                vApp.EventsEnabled = Convert.ToInt16(false);

                Visio.Document vDoc = vApp.Application.Documents.Open(filePath);

                foreach (Visio.Page thisPage in vDoc.Pages)
                {
                    foreach (Visio.Shape thisShape in thisPage.Shapes)
                    {
                        if (thisShape.Name == "DRAW_NO_1" || thisShape.Name == "DRAW_NO_2")
                        {
                            if (thisShape.Text == drawNumber)
                            {

                            }
                            else
                            {
                                thisShape.Text = drawNumber;
                            }

                        }
                        else
                        {
                            //Keep looking
                        }
                    }
                }
                vDoc.Save();
                vDoc.Close();
            }
            catch (Exception ex)
            {
                DialogResult dialogResult = MessageBox.Show((ex.Message + "     " + ex.InnerException), "Error", MessageBoxButtons.OK, MessageBoxIcon.Hand);
            }

        }

        //Design by, check by dates and names check
        public void designCheck(string filePath)
        {
            bool flag = false;
            var question = new Form5();
            var dquestion = new Form6();
            string text = "";
            string type = "";
            string objType = "";
            string name = System.IO.Path.GetFileName(filePath);

            try
            {
                vApp.ScreenUpdating = Convert.ToInt16(false);
                vApp.Visible = false;
                vApp.EventsEnabled = Convert.ToInt16(false);

                Visio.Document vDoc = vApp.Application.Documents.Open(filePath);

                filePath = System.IO.Path.GetFileName(filePath);

                foreach (Visio.Page thisPage in vDoc.Pages)
                {
                    foreach (Visio.Shape thisShape in thisPage.Shapes)
                    {

                        if (thisShape.Name == "CHECKED_NAME" || thisShape.Name == "DESIGNED_NAME")
                        {
                            objType = thisShape.Name.Substring(0, thisShape.Name.IndexOf('_')).ToLower() + " by " +
                                thisShape.Name.Substring(thisShape.Name.LastIndexOf('_') + 1).ToLower();

                            foreach (string line in File.ReadLines(Environment.CurrentDirectory + "\\data\\names.txt"))
                            {
                                if(line != thisShape.Text)
                                {
                                    flag = false;
                                }

                                else if(line == thisShape.Text)
                                {
                                    flag = true;
                                    break;
                                }
                            }

                            if(flag == false)
                            {
                                if (dquestion.repeat == false || question.origname != thisShape.Text)
                                {
                                    if (thisShape.Text == null || thisShape.Text == " " || thisShape.Text == "")
                                    {
                                        text = "There is no " + objType + " recorded in " + filePath + ".";
                                        type = "one";
                                    }
                                    else
                                    {
                                        text = "The " + objType + " " + thisShape.Text + " in " + filePath + " does not exist in your names list.";
                                        type = "it";
                                    }
                                    DialogResult dialogResult = MessageBox.Show(text + " \n" +
                                    "Would you like to add " + type + "?", "Wrong Name", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
                                    if (dialogResult == DialogResult.Yes)
                                    {
                                        if (type == "one")
                                        {
                                            question.data(thisShape.Text);
                                            while (question.result == 0)
                                            {
                                                DialogResult result = question.ShowDialog();
                                            }
                                            question.result = 0;
                                            thisShape.Text = question.name.ToUpper();
                                        }
                                        else if(type == "it")
                                        {
                                            using (System.IO.StreamWriter textfile = System.IO.File.AppendText(Environment.CurrentDirectory + "\\data\\names.txt"))
                                            {
                                                textfile.WriteLine(thisShape.Text);
                                                textfile.Close();
                                            }
                                        }

                                    }
                                    else if (dialogResult == DialogResult.No)
                                    {
                                        
                                    }
                                }
                                else
                                {
                                    thisShape.Text = question.name.ToUpper();
                                }
                            }
                            else
                            {

                            }
                        }
                        else
                        {
                            
                        }
                    }
                }
                vDoc.Save();

                
                foreach (Visio.Page thisPage in vDoc.Pages)
                {
                    foreach (Visio.Shape thisShape in thisPage.Shapes)
                    {
                        if (thisShape.Name == "CHECKED_DATE" || thisShape.Name == "DESIGNED_DATE")
                        {
                            objType = thisShape.Name.Substring(0, thisShape.Name.IndexOf('_')).ToLower() + " on " +
                                thisShape.Name.Substring(thisShape.Name.LastIndexOf('_') + 1).ToLower();

                            string format = thisShape.Text;
                            Console.WriteLine(format);
                            if (!format.Contains("-"))
                            {
                                if (dquestion.repeat == false || dquestion.origdate != thisShape.Text)
                                {
                                    if(thisShape.Text == null || thisShape.Text == " " || thisShape.Text == "" )
                                    {
                                        text = "There is no " + objType + " recorded in " + filePath + ".";
                                        type = "add one";
                                    }
                                    else
                                    {
                                        text = "The " + objType + " " + thisShape.Text + " in " + filePath + " does not match the correct format.";
                                        type = "change it";
                                    }
                                    DialogResult dialogResult = MessageBox.Show(text + "\n" +
                                        "Would you like to " + type + "?", "Wrong Date", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
                                    if (dialogResult == DialogResult.Yes)
                                    {
                                        dquestion.data(thisShape.Text);
                                        DialogResult result = dquestion.ShowDialog();
                                        thisShape.Text = dquestion.date.ToUpper();  
                                    }
                                    else if (dialogResult == DialogResult.No)
                                    {

                                    }
                                }
                                else
                                {
                                    thisShape.Text = dquestion.date.ToUpper();
                                }
                            }
                        }
                        else
                        {
                            
                        }
                    }
                }
                vDoc.Save();
                vDoc.Close();
            }
            catch (Exception ex)
            {
                DialogResult dialogResult = MessageBox.Show((ex.Message + "     " + ex.InnerException), "Error", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                
            }
        }

        //Clear com objects
        public void endSession()
        {
            vApp.Quit();
            Marshal.ReleaseComObject(vApp);

            var process = Process.GetProcessesByName("COM .exe").FirstOrDefault();
            if(process != null)
            {
                process.Kill();
            }
            Marshal.FinalReleaseComObject(vApp);
            vApp = null;
        }
    }

    //For Visio 2003-2007, currently being used
    public class visio07Edit
    {

        public int i = 0;
        public string[][] pageNum = new string[16][];
        public string[][] revision = new string[16][]; 
        public string[][] title = new string[16][];

        public void initialize()
        {
            pageNum[i] = new string[300];
            revision[i] = new string[300];
            title[i] = new string[300];
        }

        //public static string whatStatus;
        //Creates invisible instance of Visio
        Visio.InvisibleApp vApp = new Visio.InvisibleApp();
        
        //Change visio drawing status
        public void visioDrawingStat(string filePath, string whatStatus)
        {
            bool shapeExist = false;

            vApp.ScreenUpdating = Convert.ToInt16(false);
            vApp.EventsEnabled = Convert.ToInt16(false);
            try
            {
                Visio.Document vDoc = vApp.Application.Documents.Open(filePath);

                foreach (Visio.Page thisPage in vDoc.Pages)
                {
                    foreach (Visio.Shape thisShape in thisPage.Shapes)
                    {
                        if (thisShape.Name == "DRWG_STAT")
                        {
                            Console.WriteLine("Found");
                            thisShape.Text = whatStatus;
                            shapeExist = true;
                            break;
                        }
                        else
                        {
                            shapeExist = false;
                        }

                    }

                    if (shapeExist == false)
                    {
                        //No DRWG_STAT textbox? Create DRWG_STAT textbox!
                        int undoScopeID1;
                        Console.WriteLine("Create");
                        undoScopeID1 = vDoc.BeginUndoScope("Add Text Shape");
                        Visio.Shape newShape = thisPage.DrawRectangle(4.895669, 1.395341, 10.895669, 1.562008);

                        //Shape properties
                        //Coordinates X/Y
                        newShape.get_CellsSRC((short)VisSectionIndices.visSectionObject, (short)VisRowIndices.visRowXFormOut,
                            (short)VisCellIndices.visXFormPinX).FormulaU = "56mm";
                        newShape.get_CellsSRC((short)VisSectionIndices.visSectionObject, (short)VisRowIndices.visRowXFormOut,
                            (short)VisCellIndices.visXFormPinY).FormulaU = "31mm";

                        //Width and Height
                        newShape.get_CellsSRC((short)VisSectionIndices.visSectionObject, (short)VisRowIndices.visRowXFormOut,
                            (short)VisCellIndices.visXFormHeight).FormulaU = "4.2mm";
                        newShape.get_CellsSRC((short)VisSectionIndices.visSectionObject, (short)VisRowIndices.visRowXFormOut,
                            (short)VisCellIndices.visXFormWidth).FormulaU = "60mm";

                        //Align left
                        newShape.get_CellsSRC((short)VisSectionIndices.visSectionParagraph, 0,
                            (short)VisCellIndices.visHorzAlign).FormulaU = "0";
                        Console.WriteLine("Create>>");
                        newShape.Text = whatStatus;
                        newShape.Name = "DRWG_STAT";
                        newShape.LineStyle = "None";
                        newShape.FillStyle = "None";
                        vDoc.EndUndoScope(undoScopeID1, true);
                        Console.WriteLine("<<Create");
                    }
                }
                vDoc.Save();
                vDoc.Close();
            }
            catch (Exception ex)
            {
                DialogResult dialogResult = MessageBox.Show((ex.Message + "     " + ex.InnerException), "Error", MessageBoxButtons.OK, MessageBoxIcon.Hand);

            }
        }

        //Correct margins
        public void visioMargins(string filePath)
        {
            bool torf = false;
            vApp.ScreenUpdating = Convert.ToInt16(false);
            vApp.EventsEnabled = Convert.ToInt16(false);
            try
            {
                Visio.Document vDoc = vApp.Application.Documents.Open(filePath);

                int undoScopeID1;

                undoScopeID1 = vDoc.BeginUndoScope("Page Setup");
                foreach (Visio.Page thisPage in vDoc.Pages)
                {
                    thisPage.Background = Convert.ToInt16(torf);
                    thisPage.BackPage = "";
                    thisPage.PageSheet.get_CellsSRC((short)VisSectionIndices.visSectionObject,
                        (short)VisRowIndices.visRowPrintProperties, (short)VisCellIndices.visPrintPropertiesLeftMargin).FormulaU = "0 mm";
                    thisPage.PageSheet.get_CellsSRC((short)VisSectionIndices.visSectionObject,
                        (short)VisRowIndices.visRowPrintProperties, (short)VisCellIndices.visPrintPropertiesRightMargin).FormulaU = "0 mm";
                    thisPage.PageSheet.get_CellsSRC((short)VisSectionIndices.visSectionObject,
                        (short)VisRowIndices.visRowPrintProperties, (short)VisCellIndices.visPrintPropertiesTopMargin).FormulaU = "0 mm";
                    thisPage.PageSheet.get_CellsSRC((short)VisSectionIndices.visSectionObject,
                        (short)VisRowIndices.visRowPrintProperties, (short)VisCellIndices.visPrintPropertiesBottomMargin).FormulaU = "0 mm";
                    thisPage.PageSheet.get_CellsSRC((short)VisSectionIndices.visSectionObject,
                        (short)VisRowIndices.visRowPrintProperties, (short)VisCellIndices.visPrintPropertiesPaperSource).FormulaU = "1";
                }
                vDoc.EndUndoScope(undoScopeID1, true);
                vDoc.Save();
                vDoc.Close();
            }
            catch (Exception ex)
            {
                DialogResult dialogResult = MessageBox.Show((ex.Message + "     " + ex.InnerException), "Error", MessageBoxButtons.OK, MessageBoxIcon.Hand);

            }
        }

        //Remove clouds and triangles
        public void visioShapeRemove(string filePath)
        {
            int count = 0;
            int deleted = 0;
            string[] deleteList = new string[80];

            vApp.ScreenUpdating = Convert.ToInt16(false);
            vApp.EventsEnabled = Convert.ToInt16(false);
            try
            {
                Visio.Document vDoc = vApp.Application.Documents.Open(filePath);


                foreach (Visio.Page thisPage in vDoc.Pages)
                {
                    foreach (Visio.Shape thisShape in thisPage.Shapes)
                    {

                        if (thisShape.NameU.Contains("Triangle") || thisShape.Name.Contains("Revision cloud"))
                        {
                            count++;
                        }
                        else
                        {
                        }
                    }
                }
                if (count > 0)
                {
                    foreach (Visio.Page thisPage in vDoc.Pages)
                    {
                        while (deleted != count)
                        {
                            foreach (Visio.Shape thisShape in thisPage.Shapes)
                            {

                                if (thisShape.NameU.Contains("Triangle") || thisShape.Name.Contains("Revision cloud"))
                                {
                                    thisShape.Delete();
                                    deleted++;
                                }
                                else
                                {
                                }
                            }
                        }
                    }
                }
                Console.WriteLine("Task done");
                vDoc.Save();
                vDoc.Close();
            }
            catch (Exception ex)
            {
                DialogResult dialogResult = MessageBox.Show((ex.Message + "     " + ex.InnerException), "Error", MessageBoxButtons.OK, MessageBoxIcon.Hand);

            }
        }

        //Gather information for Table of Contents
        public void gatherTOC(ref ToolStripStatusLabel Label, string fileName, int folder, int population)
        {
            
            int[] rev = new int[3];
            

            Label.Text = "Gathering data";

            try
            {
                vApp.ScreenUpdating = Convert.ToInt16(false);
                vApp.Visible = false;
                vApp.EventsEnabled = Convert.ToInt16(false);
                
                Visio.Document vDoc = vApp.Application.Documents.Open(fileName);
                vApp.ScreenUpdating = Convert.ToInt16(false);
                vApp.Visible = false;
                vApp.EventsEnabled = 0;
                vApp.AlertResponse = 0;
                       
                

                foreach (Visio.Page thisPage in vDoc.Pages)
                {

                    foreach (Visio.Shape thisShape in thisPage.Shapes)
                    {
                        if (thisShape.Name == "REV_1" && thisShape.Text != "")
                        {
                            rev[0] = Convert.ToInt32(thisShape.Text);
                        }

                        if (thisShape.Name == "REV_2" && thisShape.Text != "")
                        {
                            rev[1] = Convert.ToInt32(thisShape.Text);
                        }

                        if (thisShape.Name == "REV_3" && thisShape.Text != "")
                        {
                            rev[2] = Convert.ToInt32(thisShape.Text);
                        }

                        if (thisShape.Name == "PAGE_TITLE" && thisShape.Text != "")
                        {
                            Console.WriteLine(thisShape.Text);
                            title[folder][population] = thisShape.Text;
                        }
                    }
                }

                int max = rev.Max();
                Console.WriteLine(rev[0] + " " + rev[1] + " " + rev[2]);

                if (max != 0)
                {
                    revision[folder][population] = max.ToString();
                }
                else
                {
                    revision[folder][population] = "";
                }

                pageNum[folder][population] = System.IO.Path.GetFileNameWithoutExtension(fileName);
                Console.WriteLine("Page: " + pageNum[folder][population]);
                Console.WriteLine(revision[folder][population]);
                vDoc.Saved = true;
                vDoc.Close();
                
            }
            catch (Exception ex)
            {
                DialogResult dialogResult = MessageBox.Show((ex.Message + "     " + ex.InnerException), "Error", MessageBoxButtons.OK, MessageBoxIcon.Hand);

            }

        }

        //Fill the Table of Contents
        public void fillTOC(ref ToolStripStatusLabel Label, string destLoc, int folder, int population)
        {
           
            int filelist = 0;
            int num = 0;
            string titleno = "";

            vApp.ScreenUpdating = Convert.ToInt16(false);
            vApp.Visible = false;
            vApp.EventsEnabled = Convert.ToInt16(false);
            
            Visio.Document vDoc;
            
            Label.Text = "Populating Table of Contents";
            try
            {
                while (filelist != population + 1)
                {

                    double lastPage = Convert.ToInt32(pageNum[folder][filelist]);

                    double data = Math.Ceiling(Convert.ToDouble(pageNum[folder][filelist]) / 100) + 1;
                    string fname = data.ToString();

                    Console.WriteLine(destLoc + fname.PadLeft(3, '0') + ".vsd");
                    vDoc = vApp.Application.Documents.Open(destLoc + fname.PadLeft(3, '0') + ".vsd");


                    Console.WriteLine(pageNum[folder][filelist]);

                    if ((Convert.ToInt32(pageNum[folder][filelist]) % 100) == 0)
                    {
                        num = 100;
                        titleno = num.ToString();
                    }

                    if ((Convert.ToInt32(pageNum[folder][filelist]) % 100) != 0)
                    {
                        num = (Convert.ToInt32(pageNum[folder][filelist]) % 100);
                        titleno = num.ToString();
                        titleno = titleno.PadLeft(3, '0');
                    }

                    foreach (Visio.Page thisPage in vDoc.Pages)
                    {
                        foreach (Visio.Shape thisShape in thisPage.Shapes)
                        {
                            if (thisShape.Name == "REV_NO_" + titleno)
                            {
                                thisShape.Text = revision[folder][filelist];
                                Console.WriteLine(thisShape.Name + " " + thisShape.Text);
                            }

                            if (thisShape.Name == "TITLE_" + titleno)
                            {
                                thisShape.Text = title[folder][filelist];
                                Console.WriteLine(thisShape.Name + " " + thisShape.Text);
                            }
                        }
                    }

                    vDoc.Save();
                    vDoc.Close();
                    filelist++;
                }
            }
            catch (Exception ex)
            {
                DialogResult dialogResult = MessageBox.Show((ex.Message + "     " + ex.InnerException), "Error", MessageBoxButtons.OK, MessageBoxIcon.Hand);

            }

        }

        //Insert and prepare Table of Contents for filling
        public void buildTOC(ref ToolStripStatusLabel Label, string destLoc, int folder, int population, string dbname, string cbname, string dbdate, string cbdate, bool escape)
        {
            
            Form1 es = new Form1();

            double i = 0;
            int count = 0;
            int page = 2;
            string pNum = "";
            int range = 0;
            int num = 1;
            string padding = "";
            Console.WriteLine("Building Table of Contents...");
            int filenum = 0;
            int largest = 0;
            
            string fileName = "";
            string supLoc = Properties.Settings.Default.SupportFilePath + "\\CONTENT.vsd";
            Label.Text = "Checking to see if Table of Contents files exist";

            

            try
            {
                string[] folderNums = Directory.GetFiles(destLoc);
                
                    foreach (string filename in folderNums)
                    {
                        if (System.IO.Path.GetFileName(filename).Contains("~") || (!System.IO.Path.GetExtension(filename).ToUpper().Contains("VSD")))
                        {

                            //skip the odd files
                        }
                        else
                        {
                            Console.WriteLine(filename);
                            Console.WriteLine(System.IO.Path.GetFileNameWithoutExtension(filename));
                            filenum = Convert.ToInt32(System.IO.Path.GetFileNameWithoutExtension(filename));
                        }

                        if (filenum > largest)
                        {
                            largest = filenum;
                        }
                    }

                    double lastPage = largest;

                    i = Math.Ceiling((lastPage / 100));

                    if (i == 0)
                    {
                        i = 1;
                    }
                    if (i >= 2)
                    {
                        i = i;
                    }
                

                while (count != i && Form1.escape == false)
                {
                    if (page < 10)
                    {
                        pNum = "00" + page.ToString();
                        fileName = destLoc + pNum + ".vsd";
                    }
                    else if (page >= 10)
                    {
                        pNum = "0" + page.ToString();
                        fileName = destLoc + pNum + ".vsd";
                    }

                    if (File.Exists(fileName) == false)
                    {
                        num = 1;
                        Label.Text = "Creating " + pNum + ".vsd";

                        Console.WriteLine("TOC file for " + pNum + " doesn't exist. Creating TOC file...");
                        Console.WriteLine("Copying " + supLoc + " to " + fileName + "\n");

                        File.Copy(supLoc, fileName, true);

                        File.SetAttributes(destLoc, FileAttributes.Normal);

                        vApp.ScreenUpdating = Convert.ToInt16(false);
                        vApp.Visible = false;
                        vApp.EventsEnabled = Convert.ToInt16(false);

                        Visio.Document vDoc = vApp.Application.Documents.Open(fileName);

                        while (num < 101 && Form1.escape == false)
                        {
                            range = page - 2;

                            if (num < 10)
                            {
                                padding = range.ToString() + "0" + num.ToString();
                            }
                            else if (num >= 10 && num <= 99)
                            {
                                padding = range + num.ToString();
                            }
                            else if (num == 100)
                            {
                                padding = ((range + 1) * 100).ToString();
                            }
                            string data = num.ToString();

                            foreach (Visio.Page thisPage in vDoc.Pages)
                            {
                                foreach (Visio.Shape thisShape in thisPage.Shapes)
                                {
                                    if (thisShape.Name == ("PAGE_NO_" + data.PadLeft(3, '0')) && thisShape.Text == "")
                                    {

                                        thisShape.Text = padding;

                                    }
                                    if (thisShape.Name == "SHEET_NO_1" || thisShape.Name == "SHEET_NO_2")
                                    {
                                        thisShape.Text = pNum;
                                    }

                                    if (thisShape.Name == "CHECKED_NAME")
                                    {
                                        thisShape.Text = cbname;
                                    }

                                    if (thisShape.Name == "DESIGNED_NAME")
                                    {
                                        Console.WriteLine("ding " + dbname);
                                        thisShape.Text = dbname;
                                    }

                                    if (thisShape.Name == "CHECKED_DATE")
                                    {
                                        thisShape.Text = cbdate;
                                    }

                                    if (thisShape.Name == "DESIGNED_DATE")
                                    {
                                        thisShape.Text = dbdate;
                                    }
                                }
                            }
                            num++;
                        }
                        vDoc.Save();
                        vDoc.Close();
                    }
                    page++;
                    count++;
                }
                largest = 0;
            }
            catch (Exception ex)
            {
                DialogResult dialogResult = MessageBox.Show((ex.Message + "     " + ex.InnerException), "Error", MessageBoxButtons.OK, MessageBoxIcon.Hand);

            }
        }

        //Clear com objects
        public void endSession()
        {
            vApp.Quit();
            Marshal.ReleaseComObject(vApp);

            var process = Process.GetProcessesByName("COM .exe").FirstOrDefault();
            if (process != null)
            {
                process.Kill();
            }
            Marshal.FinalReleaseComObject(vApp);
            vApp = null;
        }
    }

    //For Visio 2010 and newer, not being used at the moment
    //Code stripped from MSDN
    public class visioXEdit
    {
        public void commandVisioChange(string fileName)
        {
            using (Package visioPackage = OpenPackage(fileName,
                Environment.SpecialFolder.Desktop))
            {
                //IteratePackageParts(visioPackage);

                // Get a reference to the Visio Document part contained in the file package.
                PackagePart documentPart = GetPackagePart(visioPackage,
                    "http://schemas.microsoft.com/visio/2010/relationships/document");

                // Get a reference to the collection of pages in the document, 
                // and then to the first page in the document.
                PackagePart pagesPart = GetPackagePart(visioPackage, documentPart,
                    "http://schemas.microsoft.com/visio/2010/relationships/pages");
                PackagePart pagePart = GetPackagePart(visioPackage, pagesPart,
                    "http://schemas.microsoft.com/visio/2010/relationships/page");

                // Open the XML from the Page Contents part.
                XDocument pageXML = GetXMLFromPart(pagePart);

                /// Get all of the shapes from the page by getting
                // all of the Shape elements from the pageXML document.
                IEnumerable<XElement> shapesXML = GetXElementsByName(pageXML, "Shape");

                // Select a Shape element from the shapes on the page by 
                // its name. You can modify this code to select elements
                // by other attributes and their values.
                XElement startEndShapeXML =
                    GetXElementByAttribute(shapesXML, "NameU", "Start/End");

                // Query the XML for the shape to get the Text element, and
                // return the first Text element node.

                IEnumerable<XElement> textElements = from element in startEndShapeXML.Elements()
                                                     where element.Name.LocalName == "Text"             
                                                     select element;
                XElement textElement = textElements.ElementAt(0);

                // Change the shape text, leaving the <cp> element alone.
                textElement.LastNode.ReplaceWith("Start process");

                SaveXDocumentToPart(pagePart, pageXML);
            }
        }


        private static Package OpenPackage(string fileName,
            Environment.SpecialFolder folder)
        {
            Package visioPackage = null;

            // Get a reference to the location 
            // where the Visio file is stored.
            string directoryPath = System.Environment.GetFolderPath(
                folder);
            DirectoryInfo dirInfo = new DirectoryInfo(directoryPath);

            // Get the Visio file from the location.
            FileInfo[] fileInfos = dirInfo.GetFiles(fileName);
            if (fileInfos.Count() > 0)
            {
                FileInfo fileInfo = fileInfos[0];
                string filePathName = fileInfo.FullName;

                // Open the Visio file as a package with
                // read/write file access.
                visioPackage = Package.Open(
                    filePathName,
                    FileMode.Open,
                    FileAccess.ReadWrite);
            }
            // Return the Visio file as a package.
            return visioPackage;
        }

        private static void IteratePackageParts(Package filePackage)
        {
            // Get all of the package parts contained in the package
            // and then write the URI and content type of each one to the console.
            PackagePartCollection packageParts = filePackage.GetParts();
            foreach (PackagePart part in packageParts)
            {
                Console.WriteLine("Package part URI: {0}", part.Uri);
                Console.WriteLine("Content type: {0}", part.ContentType.ToString());
            }
        }

        private static PackagePart GetPackagePart(Package filePackage,
            string relationship)
        {
            // Use the namespace that describes the relationship 
            // to get the relationship.
            PackageRelationship packageRel =
                filePackage.GetRelationshipsByType(relationship).FirstOrDefault();
            PackagePart part = null;

            // If the Visio file package contains this type of relationship with 
            // one of its parts, return that part.
            if (packageRel != null)
            {
                // Clean up the URI using a helper class and then get the part.
                Uri docUri = PackUriHelper.ResolvePartUri(
                    new Uri("/", UriKind.Relative), packageRel.TargetUri);
                part = filePackage.GetPart(docUri);
            }
            return part;
        }

        private static PackagePart GetPackagePart(Package filePackage,
            PackagePart sourcePart, string relationship)
        {
            // This gets only the first PackagePart that shares the relationship
            // with the PackagePart passed in as an argument. You can modify the code
            // here to return a different PackageRelationship from the collection.
            PackageRelationship packageRel =
                sourcePart.GetRelationshipsByType(relationship).FirstOrDefault();
            PackagePart relatedPart = null;

            if (packageRel != null)
            {

                // Use the PackUriHelper class to determine the URI of PackagePart
                // that has the specified relationship to the PackagePart passed in
                // as an argument.
                Uri partUri = PackUriHelper.ResolvePartUri(
                    sourcePart.Uri, packageRel.TargetUri);
                relatedPart = filePackage.GetPart(partUri);
            }
            return relatedPart;
        }

        private static XDocument GetXMLFromPart(PackagePart packagePart)
        {
            XDocument partXml = null;

            // Open the packagePart as a stream and then 
            // open the stream in an XDocument object.
            Stream partStream = packagePart.GetStream();
            partXml = XDocument.Load(partStream);
            return partXml;
        }

        private static IEnumerable<XElement> GetXElementsByName(
            XDocument packagePart, string elementType)
        {
            // Construct a LINQ query that selects elements by their element type.
            IEnumerable<XElement> elements =
                from element in packagePart.Descendants()
                where element.Name.LocalName == ""
                select element;

            // Return the selected elements to the calling code.
            return elements.DefaultIfEmpty(null);
        }

        private static XElement GetXElementByAttribute(IEnumerable<XElement> elements,
            string attributeName, string attributeValue)
        {
            // Construct a LINQ query that selects elements from a group
            // of elements by the value of a specific attribute.
            IEnumerable<XElement> selectedElements =
                from el in elements
                where el.Attribute(attributeName).Value == attributeValue
                select el;

            // If there aren't any elements of the specified type
            // with the specified attribute value in the document,
            // return null to the calling code.
            return selectedElements.DefaultIfEmpty(null).FirstOrDefault();
        }

        private static void SaveXDocumentToPart(PackagePart packagePart,
            XDocument partXML)
        {
            // Create a new XmlWriterSettings object to 
            // define the characteristics for the XmlWriter
            XmlWriterSettings partWriterSettings = new XmlWriterSettings();
            partWriterSettings.Encoding = Encoding.UTF8;

            // Create a new XmlWriter and then write the XML
            // back to the document part.
            XmlWriter partWriter = XmlWriter.Create(packagePart.GetStream(),
                partWriterSettings);
            partXML.WriteTo(partWriter);

            // Flush and close the XmlWriter.
            partWriter.Flush();
            partWriter.Close();
        }
    }  
};