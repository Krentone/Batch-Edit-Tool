using System;
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
    public partial class Form3 : Form
    {
        public int result;
        public string drawNumber;
        public bool drawCancelled;
        public string suggestedDrawNumber;
        public string currentFolder;

        public Form3()
        {
            InitializeComponent();
            
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

        public void data(string folder, string number)
        {
            Console.WriteLine("Showing");
            currentFolder = folder;
            suggestedDrawNumber = number;
            textBox1.Text = suggestedDrawNumber;
            textBox1.CharacterCasing = CharacterCasing.Upper;
            label2.Text = "This will be applied to the files in " + currentFolder;
        }

        private void No_Click(object sender, EventArgs e)
        {
            
            DialogResult dialogResult = MessageBox.Show("The drawing number check will not be performed on\n" + currentFolder +
                " if you cancel this. Continue cancelling?", "Cancel?" , MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
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
            result = 2;
            drawNumber = textBox1.Text.ToUpper();
            Close();
        }

        private void Form3_Load(object sender, EventArgs e)
        {
            this.textBox1.CharacterCasing = CharacterCasing.Upper;
        }
    }
}
