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
    public partial class Form6 : Form
    {
        public int result = 0;
        public string date;
        public bool repeat;
        public string origdate;

        public Form6()
        {
            InitializeComponent();
            repeat = false;
            dateTimePicker1.CustomFormat = "dd-MMM-yyyy";
            dateTimePicker1.Format = DateTimePickerFormat.Custom;
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

        public void data(string date)
        {
            Console.WriteLine("Showing");
            label3.Text = date;
            origdate = date;
            repeat = false;
        }

        private void Yes_Click(object sender, EventArgs e)
        {
            result = 2;
            date = dateTimePicker1.Value.ToString("dd-MMM-yy");
            Close();
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if(checkBox1.Checked == true)
            {
                repeat = true;
            }
            else
            {
                repeat = false;
            }
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            date = dateTimePicker1.Value.ToString("dd-MMM-yyyy");
            Console.WriteLine(date);
        }

        private void No_Click_1(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("You must change this date, or else it will not be correct.\n"
                + "Continue cancelling?", "Cancel?", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
            if (dialogResult == DialogResult.Yes)
            {
                date = origdate;
                result = 1;
                Close();
            }
            else if (dialogResult == DialogResult.No)
            {
                result = 0;
            }
        }

        private void Yes_Click_1(object sender, EventArgs e)
        {

        }
    }
}