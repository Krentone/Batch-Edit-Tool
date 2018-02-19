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
    public partial class Form5 : Form
    {
        public int result = 0;
        public string name;
        public string origname;
        public bool repeat;

        public Form5()
        {
            InitializeComponent();
            foreach (string line in File.ReadLines(Environment.CurrentDirectory + "\\data\\names.txt"))
            {
                comboBox1.Items.Add(line);
            }
            repeat = false;
        }

        private const int CP_NOCLOSE_BUTTON = 0x200;
        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams myCp = base.CreateParams;
                myCp.ClassStyle = myCp.ClassStyle | CP_NOCLOSE_BUTTON;
                return myCp;
            }
        }

        public void data(string name)
        {
            Console.WriteLine("Showing");
            label3.Text = name;
            origname = name;
            repeat = false;
        }
        
        private void No_Click(object sender, EventArgs e)
        {
            
            DialogResult dialogResult = MessageBox.Show("You must change this name, or else it will not be correct.\n"
                +  "Continue cancelling?", "Cancel?", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
            if (dialogResult == DialogResult.Yes)
            {
                result = 1;
                name = origname;
                Close();
            }
            else if (dialogResult == DialogResult.No)
            {
                result = 0;
            }
        }

        private void Yes_Click(object sender, EventArgs e)
        {
            try
            {
                result = 1;
                name = comboBox1.SelectedItem.ToString();
                Close();
            }
            catch
            {
                result = 0;
                DialogResult dialogResult = MessageBox.Show("Please select a name.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {
                repeat = true;
            }
            else
            {
                repeat = false;
            }
        }
    }
}