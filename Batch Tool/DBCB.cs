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
    public partial class DBCB : Form
    {
        public int result = 0;
        public string dbname;
        public string cbname;
        public string dbdate;
        public string cbdate;
        public string origname;
        public bool repeat;

        public DBCB()
        {
            InitializeComponent();
            foreach (string line in File.ReadLines(Environment.CurrentDirectory + "\\data\\names.txt"))
            {
                comboBox1.Items.Add(line);
                comboBox2.Items.Add(line);
                Console.WriteLine(line);
            }
            dateTimePicker1.CustomFormat = "dd-MMM-yyyy";
            dateTimePicker1.Format = DateTimePickerFormat.Custom;
            dateTimePicker2.CustomFormat = "dd-MMM-yyyy";
            dateTimePicker2.Format = DateTimePickerFormat.Custom;
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

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void DBCB_Load(object sender, EventArgs e)
        {

        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            cbdate = dateTimePicker1.Value.ToString("dd-MMM-yyyy");
            Console.WriteLine(cbdate);
        }

        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {
            dbdate = dateTimePicker2.Value.ToString("dd-MMM-yyyy");
            Console.WriteLine(dbdate);
        }

        private void No_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("Names and dates will not be inserted if you cancel this.\n"
                + "Continue cancelling?", "Cancel?", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
            if (dialogResult == DialogResult.Yes)
            {
                result = 1;
                Close();
            }
            else if (dialogResult == DialogResult.No)
            {
                result = 0;
            }
        }

        private void Yes_Click(object sender, EventArgs e)
        {
            
            
        }

        private void Yes_Click_1(object sender, EventArgs e)
        {
            
            try
            {
                result = 2;
                dbname = comboBox2.SelectedItem.ToString();
                cbname = comboBox1.SelectedItem.ToString();

                dbdate = dateTimePicker2.Value.ToString("dd-MMM-yy");
                cbdate = dateTimePicker1.Value.ToString("dd-MMM-yy");
                Console.WriteLine(cbdate);
                Console.WriteLine(dbdate);
                Console.WriteLine(cbname);
                Console.WriteLine(dbname);
                Close();
            }
            catch
            {
                result = 0;
                DialogResult dialogResult = MessageBox.Show("Please select a design by and check by name.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
        }
    }
}
