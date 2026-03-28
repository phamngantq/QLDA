using DevExpress.XtraCharts.Native;
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
    public partial class frmTLTK : Form
    {
        
        public frmTLTK()
        {
            InitializeComponent();
        }
        public bool Expand = false; // khai báo biến Expand 
        public bool Expand2 = false; // khai báo biến Expand2 
        public bool Expandmenu = false; // khai báo biến Expand2 
        SqlConnection conn = new SqlConnection("Data Source=LAPTOP-D4IEITM3\\SQLEXPRESS02;Initial Catalog=DOAN;User ID=sa;Password=Sa@12345;TrustServerCertificate=True");
        SqlDataAdapter da = new SqlDataAdapter();
        SqlCommand cmd = new SqlCommand();
        DataTable dt = new DataTable();
        string sql, constr;


        private void btndropQLDL_Click(object sender, EventArgs e)
        {

            QLDLdrop.Start();
            Console.WriteLine(Expand);
            Console.ReadLine();
        }

        public void NapCT()
        {
            int i = grdTKQL.CurrentRow.Index;
            txtTenDeTai.Text = grdTKQL.Rows[i].Cells[0].Value.ToString();
            txtTenSinhVien.Text = grdTKQL.Rows[i].Cells[1].Value.ToString();
            txtChuyenNganh.Text = grdTKQL.Rows[i].Cells[2].Value.ToString();
            txtKhoa.Text = grdTKQL.Rows[i].Cells[3].Value.ToString();
            txtGVHD.Text = grdTKQL.Rows[i].Cells[4].Value.ToString();
            dtpickNambaove.Text = grdTKQL.Rows[i].Cells[5].Value.ToString();
            txtTomTat.Text = grdTKQL.Rows[i].Cells[6].Value.ToString();



        }


        private DataTable TimKiemTheoTuKhoa(string keyword)
        {
            string selectSql = @"
SELECT 
    DA.TENDETAI,
    SV.HOTEN,
    SV.CHUYENNGANH,
    SV.KHOA,
    GVHD.GVHD,
    DA.NAMBAOVE,
    DA.TOMTAT
FROM DOAN DA
JOIN SINHVIEN SV ON DA.MASINHVIEN = SV.MASINHVIEN
JOIN GVHD GVHD ON DA.MAGVHD = GVHD.MAGVHD
";

            DataTable dtResult = new DataTable();

            try
            {
                // Nếu không có từ khóa => lấy toàn bộ
                if (string.IsNullOrWhiteSpace(keyword))
                {
                    using (SqlDataAdapter adapter = new SqlDataAdapter(selectSql, conn))
                    {
                        adapter.Fill(dtResult);
                    }
                }
                else
                {
                    string whereClause = @"
WHERE DA.TENDETAI    LIKE @kw
   OR SV.HOTEN       LIKE @kw
   OR SV.CHUYENNGANH LIKE @kw
   OR SV.KHOA        LIKE @kw
   OR GVHD.GVHD        LIKE @kw
";
                    string finalSql = selectSql + whereClause;

                    using (SqlDataAdapter adapter = new SqlDataAdapter(finalSql, conn))
                    {
                        adapter.SelectCommand.Parameters.AddWithValue("@kw", "%" + keyword.Trim() + "%");
                        adapter.Fill(dtResult);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi tìm kiếm: " + ex.Message);
            }

            return dtResult;
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
            sql = @"SELECT 
    DA.TENDETAI,
    SV.HOTEN,
    SV.CHUYENNGANH,
    SV.KHOA,
    GVHD.GVHD,
    DA.NAMBAOVE ,
    DA.TOMTAT
   FROM DOAN DA
JOIN SINHVIEN SV   ON DA.MASINHVIEN = SV.MASINHVIEN 
JOIN GVHD ON GVHD.MAGVHD=DA.MAGVHD";
            conn.Open();
            da = new SqlDataAdapter(sql, conn);
            da.Fill(dt);
            grdTKQL.DataSource = dt;
            grdTKQL.Refresh();
            NapCT();
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
            if (Expand == true)
            {
                dropdown.Height -= 15;
                if (dropdown.Height <= dropdown.MinimumSize.Height)
                {

                    QLDLdrop.Stop();
                    Expand = false;
                }
            }
            if (Expand2 == true)
            {
                dropdown2.Height -= 15;
                if (dropdown2.Height <= dropdown2.MinimumSize.Height)
                {

                    QLDLdrop.Stop();
                    Expand2 = false;
                }
            }
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

        private void btnFirst_Click(object sender, EventArgs e)
        {
            grdTKQL.ClearSelection();
            grdTKQL.CurrentCell = grdTKQL[0, 0];
            NapCT();
        }

        private void btnNext_Click(object sender, EventArgs e)
        {
            int i = grdTKQL.CurrentRow.Index;
            if (i < grdTKQL.Rows.Count - 1)
            {
                grdTKQL.CurrentCell = grdTKQL[0, i + 1];
                NapCT();
            }
        }

        private void grdTKQL_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            NapCT();// khi click vào ô nào đó thì NapCT() sẽ được gọi

        }

        private void btnPre_Click(object sender, EventArgs e)
        {
            int i = grdTKQL.CurrentRow.Index;
            if (i > 0)
            {
                grdTKQL.CurrentCell = grdTKQL[0, i - 1];
                NapCT();
            }
        }

        private void btnEnd_Click(object sender, EventArgs e)
        {
            int i = grdTKQL.Rows.Count - 1;
            grdTKQL.CurrentCell = grdTKQL[0, i - 1];
            NapCT();
        }

        private void comTruong_SelectedIndexChanged(object sender, EventArgs e)
        {
            sql = "Select distinct " + comTruong.Text + " FROM DOAN DA JOIN SINHVIEN SV ON DA.MASINHVIEN = SV.MASINHVIEN JOIN GVHD  ON DA.MAGVHD = GVHD.MAGVHD ";
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
            sql = " SELECT  DA.TENDETAI, SV.HOTEN, SV.CHUYENNGANH, SV.KHOA, GVHD.GVHD, DA.NAMBAOVE, DA.TOMTAT FROM DOAN DA JOIN SINHVIEN SV ON DA.MASINHVIEN = SV.MASINHVIEN JOIN GVHD  ON DA.MAGVHD = GVHD.MAGVHD WHERE " + comTruong.Text + "= N'" + comGT.Text + "'";// Đảm bảo mọi dữ liệu có tiếng việt vẫn lọc được
            da = new SqlDataAdapter(sql, conn);
            dt = new DataTable();
            da.Fill(dt);
            grdTKQL.DataSource = dt;
            grdTKQL.Refresh();
            NapCT();
        }

        private void panelALL_Paint(object sender, PaintEventArgs e)
        {

        }

        private void btnTimKiem_Click(object sender, EventArgs e)
        {
            TimKiemTheoTuKhoa(txt_TimKiem.Text);

            string kw = txt_TimKiem.Text; // textbox chứa từ khóa
            DataTable dt = TimKiemTheoTuKhoa(kw);
            grdTKQL.DataSource = dt;
            grdTKQL.Refresh();

        }

        private void button5_Click(object sender, EventArgs e)
        {
            try
            {

                sql = @" SELECT 
    DA.TENDETAI,
    SV.HOTEN,
    SV.CHUYENNGANH,
    SV.KHOA,
    GVHD.GVHD,
    DA.NAMBAOVE,
    DA.TOMTAT
FROM DOAN DA
JOIN SINHVIEN SV   ON DA.MASINHVIEN = SV.MASINHVIEN
JOIN GVHD  ON DA.MAGVHD = GVHD.MAGVHD";
                da = new SqlDataAdapter(sql, conn);
                dt = new DataTable();
                dt.Clear();
                da.Fill(dt);
                grdTKQL.DataSource = dt;
                grdTKQL.Refresh();
                NapCT();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi làm mới dữ liệu: " + ex.Message);
            }

            //2️⃣ Reset toàn bộ phần lọc và tìm kiếm
            //comTruong.SelectedIndex = -1;
            //comGT.SelectedIndex = -1;
            //comTruong.SelectedIndex = -1;  // Bỏ chọn tên trường
            //comGT.DataSource = null;       // Xóa dữ liệu trong combo giá trị
            //comGT.Text = "";               // Làm trống text hiển thị

            txt_TimKiem.Clear();   //  Xóa ô tìm kiếm về rỗng
        }

        private void button13_Click(object sender, EventArgs e)
        {
            OpenHTML.OpenDefault();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            menutimer.Start();
            Console.WriteLine(Expandmenu);
            pictureBox12.Visible = true;
            Console.ReadLine();
        }

        private void grdTracuu_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
