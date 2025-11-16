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
    public partial class frmChitiet : Form


    {
        public frmChitiet()
        {
            InitializeComponent();
        }

        public bool Expand = false; // khai báo biến Expand 
        public bool Expand2 = false; // khai báo biến Expand2 
        public bool Expandmenu = false; // khai báo biến Expand2 

        // ✅ Constructor có 8 tham số – bắt buộc phải có để dòng bạn gọi chạy được
        public frmChitiet(
            string tenDeTai,
                      string tenSinhVien,
                      string lop,
                      string gvhd,
                      string tuKhoa,
                      string namBaoVe,
                      string tomTat,
                      string fileBaoCao,
                      string slide,
                      string lyLich
                    )
        {
            InitializeComponent();

            // Gán dữ liệu lên control
            txtTenDA.Text = tenDeTai;
            txtTenSV.Text = tenSinhVien;
            txtLop.Text = lop;
            txtGVHD.Text = gvhd;
            txtKeyword.Text = tuKhoa;
            dtpickNambaove.Text = namBaoVe;
            txtTomtat.Text = tomTat;
            txtFileBC.Text = fileBaoCao;
            txtSlide.Text = slide;
            txtliLich.Text = lyLich;
        }
    


        private void btnDrop2_Click(object sender, EventArgs e)
        {
            
        }

        private void guna2Button1_Click(object sender, EventArgs e)
        {
            frmTracuu f = new frmTracuu();
            f.Show();
            this.Hide();

        }

        private void frmChitiet_Load(object sender, EventArgs e)
        {

        }

        private void guna2GroupBox1_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
           
        }

        private void button4_Click(object sender, EventArgs e)
        {
            frmHome f = new frmHome();
            f.Show();
            this.Hide();
        
    }

        

        private void btnTraCuu_Click(object sender, EventArgs e)
        {
            frmTracuu f = new frmTracuu();
            f.Show();
            this.Hide();
        }

        private void btnDAYT_Click(object sender, EventArgs e)
        {
            frmDAYT f = new frmDAYT();
            f.Show();
            this.Hide();
        }

        private void txtTenSV_TextChanged(object sender, EventArgs e)
        {

        }

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

        private void pictureBox1_Click(object sender, EventArgs e)
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

        private void pictureBox5_Click(object sender, EventArgs e)
        {
            menutimer.Start();
            Console.WriteLine(Expandmenu);

            Console.ReadLine();
        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {
            menutimer.Start();
            Console.WriteLine(Expandmenu);

            Console.ReadLine();
        }

        private void pictureBox6_Click(object sender, EventArgs e)
        {
            menutimer.Start();
            Console.WriteLine(Expandmenu);

            Console.ReadLine();
        }

        private void pictureBox9_Click(object sender, EventArgs e)
        {
            menutimer.Start();
            Console.WriteLine(Expandmenu);

            Console.ReadLine();
        }

        private void btnLove_Click(object sender, EventArgs e)
        {
            frmTracuu f = new frmTracuu();
            f.Show();
            this.Hide();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            frmTracuu f = new frmTracuu();
            f.Show();
            this.Hide();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            frmHome f = new frmHome();  
            f.Show();
            this.Hide();
        }
    }
}
