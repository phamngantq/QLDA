using DevExpress.XtraPrinting;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices.Expando;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace formQLmain
{
    public partial class frmTLKT : Form
    {
        
        public frmTLKT()
        {
            InitializeComponent();
        }
        public bool Expand = false; // khai báo biến Expand 
        public bool Expand2 = false; // khai báo biến Expand2 
        public bool Expandmenu = false; // khai báo biến Expand2 
        private void btndropQLDL_Click(object sender, EventArgs e)
        {

            QLDLdrop.Start();
            Console.WriteLine(Expand);
            Console.ReadLine();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            frmQLSV f = new frmQLSV();
            f.Show();
            this.Hide();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            frmQLTK f = new frmQLTK();
            f.Show();
            this.Hide();
        }

        private void btnDrop2_Click(object sender, EventArgs e)
        {
            QLDAdrop.Start();
            Console.WriteLine(Expand2);
            Console.ReadLine();
        }

        private void btnDSDA_Click(object sender, EventArgs e)
        {
            frmDSDA f = new frmDSDA();
            f.Show();
            this.Hide();
        }

        private void btnTLKT_Click(object sender, EventArgs e)
        {
           
        }

        private void QLDAdrop_Tick(object sender, EventArgs e)
        {
            if (Expand2 == false)
            {
                dropdown2.Height += 15;
                if (dropdown2.Height >= dropdown2.MaximumSize.Height)
                {

                    QLDAdrop.Stop();
                    Expand2 = true;
                }
            }
            else
            {
                dropdown2.Height -= 15;
                if (dropdown2.Height <= dropdown2.MinimumSize.Height)
                {

                    QLDAdrop.Stop();
                    Expand2 = false;
                }
            }
        }

        private void QLDLdrop_Tick(object sender, EventArgs e)
        {
            if (Expand == false)
            {
                dropdown.Height += 15;
                if (dropdown.Height >= dropdown2.MaximumSize.Height)
                {

                    QLDLdrop.Stop();
                    Expand = true;
                }
            }
            else
            {
                dropdown.Height -= 15;
                if (dropdown.Height <= dropdown2.MinimumSize.Height)
                {

                    QLDLdrop.Stop();
                    Expand = false;
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            FrmQLmain f = new FrmQLmain();
            f.Show();
            this.Hide();
        }
        private ModifyTaiLieuKemTheo repo;
        private string maTLCu = ""; // lưu mã TLBC gốc để cho phép đổi khóa khi sửa
        private void frmTLKT_Load(object sender, EventArgs e)
        {
            
        }
        
       

        // ================== ĐỒNG BỘ GRID → TEXTBOX ==================
       

       
        private void button4_Click(object sender, EventArgs e)
        {
            frmHome f = new frmHome();
            f.Show();
            this.Hide();
        }

        private void pictureBox14_Click(object sender, EventArgs e)
        {
            menutimer.Start();
            Console.WriteLine(Expandmenu);
            pictureBox12.Visible = true;
            Console.ReadLine();
        }

        private void pictureBox4_Click(object sender, EventArgs e)
        {
            menutimer.Start();
            Console.WriteLine(Expandmenu);
            pictureBox12.Visible = true;
            Console.ReadLine();
        }

        private void pictureBox10_Click(object sender, EventArgs e)
        {
            menutimer.Start();
            Console.WriteLine(Expandmenu);
            pictureBox12.Visible = true;
            Console.ReadLine();
        }

        private void pictureBox9_Click(object sender, EventArgs e)
        {
            menutimer.Start();
            Console.WriteLine(Expandmenu);
            pictureBox12.Visible = true;
            Console.ReadLine();
        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {
            menutimer.Start();
            Console.WriteLine(Expandmenu);
            pictureBox12.Visible = true;
            Console.ReadLine();
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            menutimer.Start();
            Console.WriteLine(Expandmenu);
            pictureBox12.Visible = true;
            Console.ReadLine();
        }

        private void pictureBox11_Click(object sender, EventArgs e)
        {
            menutimer.Start();
            Console.WriteLine(Expandmenu);
            pictureBox12.Visible = true;
            Console.ReadLine();
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
                    pictureBox12.Visible = true;
                    panelALL.Left = 245;
                }
            }
            else
            {
                panelMenu.Width -= 25;
                if (panelMenu.Width <= panelMenu.MinimumSize.Width)
                {
                    menutimer.Stop();
                    Expandmenu = false;
                    pictureBox12.Visible = false;
                    panelALL.Left = 73;
                }

            }
        }

        private void pictureBox6_Click(object sender, EventArgs e)
        {

        }

        private void pictureBox8_Click(object sender, EventArgs e)
        {

        }

        private void button12_Click(object sender, EventArgs e)
        {
            frmBaocao f = new frmBaocao();
            f.Show();
            this.Hide();
        }

        private void btnRefresh_Click(object sender, EventArgs e)
        {

        }
    }
}
