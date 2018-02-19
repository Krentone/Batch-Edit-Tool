using System;
using System.IO;
using System.Diagnostics;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace $safeprojectname$
{
    public partial class Setting : Form
    {
        List<string> names = new List<string>();

        public Setting()
        {
            Properties.Settings.Default.Reload();

            string root = Environment.CurrentDirectory;

            bool exists = System.IO.Directory.Exists(root + "\\data");
            if (!exists)
            {
                System.IO.Directory.CreateDirectory(root + "\\data");
                File.Create(root + "\\data\\names.txt");
            }
            
            InitializeComponent();
            
           // Properties.Settings.Default.PDFPath;
          //  Properties.Settings.Default.SupportFilePath;

            Console.WriteLine("hi: " + Properties.Settings.Default.PDFPath);

            foreach(string line in File.ReadLines(root + "\\data\\names.txt"))
            {
                names.Add(line);
                listBox1.Items.Add(line);
            }
            
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
           // Properties.Settings.Default.DontShow = this.checkBox1.Checked;
           // Properties.Settings.Default.Save();
        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.Backups = checkBox4.Checked;
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.Notification = checkBox2.Checked;
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.BrowserAutoStart = checkBox3.Checked;
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            int result;

            DialogResult dialogResult = MessageBox.Show("All changes made will be discarded if you\n" +
                "cancel this. Continue cancelling?", "Cancel?", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
            if (dialogResult == DialogResult.Yes)
            {
                textBox1.Text = Properties.Settings.Default.PDFPath;
                textBox2.Text = Properties.Settings.Default.SupportFilePath;
                
                result = 1;
                Close();
            }
            else if (dialogResult == DialogResult.No)
            {
                result = 0;
            }
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void folderBrowserDialog1_HelpRequest(object sender, EventArgs e)
        {

        }

        private void label2_Click_1(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.PDFPath = textBox1.Text;
        }

        private void textBox1_Click(object sender, MouseEventArgs e)
        {
            
        }

        private void textBox2_Click(object sender, MouseEventArgs e)
        {
            
        }

        private void button5_Click(object sender, EventArgs e)
        {
            DialogResult result = folderBrowserDialog1.ShowDialog();
            if(result == DialogResult.OK)
            {
                textBox1.Text = folderBrowserDialog1.SelectedPath;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            DialogResult result = folderBrowserDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {
               textBox2.Text = folderBrowserDialog1.SelectedPath;
            }
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (listBox1.Items.Count != 0) {
                string remove = listBox1.SelectedItem.ToString();
                listBox1.Items.Remove(remove);
                names.Remove(remove);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (textBox3.Text != "" || textBox3.Text != " ")
            {
                names.Add(textBox3.Text.ToUpper());
                listBox1.Items.Add(textBox3.Text.ToUpper());
                textBox3.Clear();
                textBox3.Focus();
            }
        }



        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.SupportFilePath = textBox2.Text;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            
            Properties.Settings.Default.SupportFilePath = textBox2.Text;
            Properties.Settings.Default.PDFPath = textBox1.Text;
            Properties.Settings.Default.Notification = checkBox2.Checked;
            Properties.Settings.Default.Backups = checkBox4.Checked;
            Properties.Settings.Default.BrowserAutoStart = checkBox3.Checked;

            Properties.Settings.Default.Save();

            Console.WriteLine("saving");
            
            using (System.IO.StreamWriter textfile = new System.IO.StreamWriter(Environment.CurrentDirectory + "\\data\\names.txt"))
            {
                foreach (string item in names)
                {
                    textfile.WriteLine(item);
                }
            }


            Form1 form1 = new Form1();
            if (Properties.Settings.Default.PDFPath == "" || Properties.Settings.Default.PDFPath == null)
            {
                form1.label3.Visible = true;
                form1.radioButton10.Enabled = false;
                form1.radioButton11.Enabled = false;
            }
            else
            {
                form1.label3.Visible = false;
                form1.radioButton10.Enabled = true;
                form1.radioButton11.Enabled = true;
            }

            Close();
        }

        

        private void Form2_FormClosing(object sender, FormClosingEventArgs e)
        {
            using (System.IO.StreamWriter textfile = new System.IO.StreamWriter(Environment.CurrentDirectory + "\\data\\names.txt"))
            {
                foreach (string item in names)
                {
                    textfile.WriteLine(item);
                }
            } 
        }

        

        
    }
}
