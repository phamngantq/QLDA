using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace formQLmain
{
    public partial class frmUser : Form
    {
        public frmUser()
        {
            InitializeComponent();
        }
        public bool Expandmenu = false; // khai báo biến Expand2 
        private void menutimer_Tick(object sender, EventArgs e)
        {
            // trượt dọc đã làm được 
            if (Expandmenu == false)
            {
                panelMenu.Width += 25;
                if (panelMenu.Width >= panelMenu.MaximumSize.Width)
                {

                    menutimer.Stop();
                    Expandmenu = true;
                   
                    
                }
            }
            else
            {
                panelMenu.Width -= 25;
                if (panelMenu.Width <= panelMenu.MinimumSize.Width)
                {
                    menutimer.Stop();
                    Expandmenu = false;
                    
                    
                }

            }
        }

        private void pictureBox13_Click(object sender, EventArgs e)
        {
            menutimer.Start();
            Console.WriteLine(Expandmenu);
            
            Console.ReadLine();
        }

        private void pictureBox4_Click(object sender, EventArgs e)
        {
            menutimer.Start();
            Console.WriteLine(Expandmenu);
            
            Console.ReadLine();
        }

        private void pictureBox10_Click(object sender, EventArgs e)
        {
            menutimer.Start();
            Console.WriteLine(Expandmenu);
            
            Console.ReadLine();
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            menutimer.Start();
            Console.WriteLine(Expandmenu);
           
            Console.ReadLine();
        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {
            menutimer.Start();
            Console.WriteLine(Expandmenu);
           
            Console.ReadLine();
        }

        private void frmUser_Load(object sender, EventArgs e)
        {

        }

        private void btnLove_Click(object sender, EventArgs e)
        {
            frmDAYT f = new frmDAYT();
            f.Show();
            this.Hide();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            frmTracuu f = new frmTracuu();
            f.Show();
            this.Hide();
        }

        private void btnKhamPha_Click(object sender, EventArgs e)
        {
            frmTracuu f = new frmTracuu();
            f.Show();
            this.Hide();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            frmHome f = new frmHome();
            f.Show();
            this.Hide();
        }
    }
}
