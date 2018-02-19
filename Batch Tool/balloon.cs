using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Timers;

namespace $safeprojectname$
{
    public partial class balloon : Form
    {
        private static System.Timers.Timer timer;
        private static System.Timers.Timer atimer;

        public balloon()
        {
            InitializeComponent();
            Rectangle workingArea = Screen.GetWorkingArea(this);
            this.Location = new Point((workingArea.Right - 15) - Size.Width,
                                      (workingArea.Bottom - 15) - Size.Height);

        }

        public void data(string message)
        {
            label2.Text = message;
        }

        //notification balloon
        private void balloon_Load(object sender, EventArgs e)
        {
            //comment out the next two lines if debugging
            System.Media.SoundPlayer player = new System.Media.SoundPlayer(Environment.CurrentDirectory + "\\Resources\\notify.wav");
            player.Play();
            FadeIn(this, 25);
            timer = new System.Timers.Timer(3000);
            timer.Elapsed += OnTimedEvent;
            timer.Enabled = true;
            
            GC.KeepAlive(timer);
        }

        private void OnTimedEvent(object source, ElapsedEventArgs e)
        {
           
            timer.Dispose();
            if (IsHandleCreated)
            {
                Invoke((MethodInvoker)delegate ()
                {
                FadeOut(this, 25);

                });
            }
            else
            {

            }
        }

        //Fade in and out balloon
        //STACK EXCHANGE https://stackoverflow.com/questions/12497826/better-algorithm-to-fade-a-winform
        private async void FadeIn(Form o, int interval)
        {
            o.Opacity = 0;
            //Object is not fully invisible. Fade it in
            while (o.Opacity < 0.85)
            {
                await Task.Delay(interval);
                o.Opacity += 0.05;
            }
            o.Opacity = 0.85; //make fully visible       
        }

        
        private async void FadeOut(Form o, int interval)
        {
            //Object is fully visible. Fade it out
            while (o.Opacity > 0.0)
            {
                await Task.Delay(interval);
                o.Opacity -= 0.05;
            }
            o.Opacity = 0; //make fully invisible   
            this.Close();
        }

        //Make baloon disappear if clicked on, open main application
        protected override void OnClick(EventArgs e)
        {
            timer.Dispose();
            this.Close();
            base.OnClick(e);
        }
    }
}
