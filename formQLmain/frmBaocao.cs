using DevExpress.CodeParser;
using DevExpress.DataAccess.Sql;
using DevExpress.XtraCharts.Native;
using DevExpress.XtraReports.UI;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace formQLmain
{
    public partial class frmBaocao : Form
    {
        SqlConnection conn = new SqlConnection("Data Source=LAPTOP-D4IEITM3\\SQLEXPRESS02;Initial Catalog=DOAN;User ID=sa;Password=Sa@12345;TrustServerCertificate=True");
        SqlDataAdapter da = new SqlDataAdapter();
        SqlCommand cmd = new SqlCommand();
        DataTable dt = new DataTable();
        string sql, constr;
        public bool Expand = false; // khai báo biến Expand 
        public bool Expand2 = false; // khai báo biến Expand2 
        public bool Expandmenu = false; // khai báo biến Expand2 

        // THÊM: Biến lưu vai trò người dùng
        private string _userRole;

        public frmBaocao()
        { InitializeComponent(); }
        public frmBaocao(string userRole)
        {
            InitializeComponent();

            _userRole = userRole; // Gán vai trò người dùng
        }

        public void NapCT()
        {
           


        }

        private void grdBaocao_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void panel10_Paint(object sender, PaintEventArgs e)
        {

        }

        private void pnlgrid_Paint(object sender, PaintEventArgs e)
        {

        }

        private void comTruong_SelectedIndexChanged(object sender, EventArgs e)
        {
            sql = "Select distinct " + comTruong.Text + " FROM DOAN DA JOIN SINHVIEN SV ON DA.MASINHVIEN = SV.MASINHVIEN JOIN GVHD ON GVHD.MAGVHD=DA.MAGVHD ";
            da = new SqlDataAdapter(sql, conn);
            DataTable dt1 = new DataTable();
            da.Fill(dt1);
            //comGT.Items.Clear();
            comGT.DataSource = dt1;
            comGT.DisplayMember = comTruong.Text; //trường hiện ra
            comGT.ValueMember = comTruong.Text;// trường để lấy 

        }

        private void btnFillter_Click(object sender, EventArgs e)
        {
            sql = @"SELECT 
    DA.TENDETAI,
    SV.HOTEN,
    SV.CHUYENNGANH,
    SV.KHOA,
    GVHD.GVHD,
    YEAR(DA.NAMBAOVE) AS N'NĂM'
FROM DOAN DA
JOIN SINHVIEN SV   ON DA.MASINHVIEN = SV.MASINHVIEN JOIN GVHD ON GVHD.MAGVHD=DA.MAGVHD WHERE " + comTruong.Text + "= N'" + comGT.Text + "'";// Đảm bảo mọi dữ liệu có tiếng việt vẫn lọc được
            da = new SqlDataAdapter(sql, conn);
            dt = new DataTable();
            da.Fill(dt);
            grdBaocao.DataSource = dt;
            grdBaocao.Refresh();
        }

        private void btnRefresh_Click(object sender, EventArgs e)
        {
            try
            {

                sql = @"SELECT 
    DA.TENDETAI,
    SV.HOTEN,
    SV.CHUYENNGANH,
    SV.KHOA,
    GVHD.GVHD,
    YEAR(DA.NAMBAOVE) AS N'NĂM'
FROM DOAN DA
JOIN SINHVIEN SV   ON DA.MASINHVIEN = SV.MASINHVIEN
JOIN GVHD ON GVHD.MAGVHD=DA.MAGVHD"; 
                da = new SqlDataAdapter(sql, conn);
                dt = new DataTable();
                dt.Clear();
                da.Fill(dt);
                grdBaocao.DataSource = dt;
                grdBaocao.Refresh();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi làm mới dữ liệu: " + ex.Message);
            }
           


        }

        private void btnInBC_Click(object sender, EventArgs e)
        {
            //PHÂN QUYỀN: Kiểm tra nếu là GIANGVIEN
            if (_userRole.Equals("GIANGVIEN", StringComparison.OrdinalIgnoreCase))
            {
                MessageBox.Show("Tài khoản Giảng viên không có quyền hạn In báo cáo.", "Không Có Quyền", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return; // Ngừng thực thi và không tiến hành in


            }

            rptDoann rpt = new rptDoann();
            sql = "SELECT DA.TENDETAI, SV.HOTEN, SV.CHUYENNGANH, SV.KHOA, GVHD.GVHD, YEAR(DA.NAMBAOVE) AS N'NĂM' FROM DOAN DA JOIN SINHVIEN SV ON DA.MASINHVIEN = SV.MASINHVIEN JOIN GVHD ON GVHD.MAGVHD = DA.MAGVHD" +
                " where " + comTruong.Text + " = N'" + comGT.Text + " '";

            da = new SqlDataAdapter(sql, conn);
            DataTable rdt = new DataTable();
            da.Fill(rdt);
            rpt.rptNgayIn.Text = string.Format("MIS, ngày {0} tháng {1} năm {2}", DateTime.Now.Day, DateTime.Now.Month, DateTime.Now.Year);
            rpt.xrgroupby.Text = "Điều kiện lọc dữ liệu: " + comTruong.Text + ": " + comGT.Text;
            rpt.DataSource = rdt;
            rpt.ShowPreview();
        }

        private void button13_Click(object sender, EventArgs e)
        {
            OpenHTML.OpenDefault();

        }

        private void btnLogout_Click(object sender, EventArgs e)
        {
            frmHome f = new frmHome();
            f.Show();
            this.Hide();
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

        private void pictureBox1_Click(object sender, EventArgs e)
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

        private void pictureBox2_Click(object sender, EventArgs e)
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

        private void pictureBox9_Click(object sender, EventArgs e)
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

        private void btnFirst_Click(object sender, EventArgs e)
        {
            grdBaocao.ClearSelection();
            grdBaocao.CurrentCell = grdBaocao[0, 0];
        }

        private void btnPre_Click(object sender, EventArgs e)
        {
            int i = grdBaocao.CurrentRow.Index;
            if (i > 0)
            {
                grdBaocao.CurrentCell = grdBaocao[0, i - 1];
               
            }
        }

        private void btnNext_Click(object sender, EventArgs e)
        {
            int i = grdBaocao.CurrentRow.Index;
            if (i < grdBaocao.Rows.Count - 1)
            {
                grdBaocao.CurrentCell = grdBaocao[0, i + 1];
                NapCT();
            }
        }

        private void btnEnd_Click(object sender, EventArgs e)
        {
            int i = grdBaocao.Rows.Count - 1;
            grdBaocao.CurrentCell = grdBaocao[0, i - 1];
        }

        private void button13_Click_1(object sender, EventArgs e)
        {
            OpenHTML.OpenDefault();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            FrmQLmain f = new FrmQLmain();
            f.Show();
            this.Hide();
        }

        private void btndropQLDL_Click(object sender, EventArgs e)
        {

        }

        private void button12_Click(object sender, EventArgs e)
        {
           
        }

        private void btnDSDA_Click(object sender, EventArgs e)
        {
            frmDSDA f = new frmDSDA();
            f.Show();
            this.Hide();

        }

        private void btnTracuu_Click(object sender, EventArgs e)
        {
            frmTLTK f = new frmTLTK();
            f.Show();
            this.Hide();
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

        private void button4_Click(object sender, EventArgs e)
        {
            frmHome f = new frmHome();
            f.Show();
            this.Hide();
        }

        private void frmBaocao_Load(object sender, EventArgs e)
        {
            //if (_userRole.Equals("GIANGVIEN", StringComparison.OrdinalIgnoreCase))
            //{
            //    MessageBox.Show("Tài khoản Giảng viên không có quyền hạn In báo cáo.", "Không Có Quyền", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    return; // Ngừng thực thi và không tiến hành in


            //}


            sql = @"SELECT 
    DA.TENDETAI,
    SV.HOTEN,
    SV.CHUYENNGANH,
    SV.KHOA,
    GVHD.GVHD,
    YEAR(DA.NAMBAOVE) AS N'NĂM'
   FROM DOAN DA
JOIN SINHVIEN SV   ON DA.MASINHVIEN = SV.MASINHVIEN 
JOIN GVHD ON GVHD.MAGVHD=DA.MAGVHD";
            conn.Open();
            da = new SqlDataAdapter(sql, conn);
            da.Fill(dt);
            grdBaocao.DataSource = dt;
            grdBaocao.Refresh();


        }
    }
}
