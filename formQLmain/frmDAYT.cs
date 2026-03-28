using DevExpress.CodeParser;
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
    public partial class frmDAYT : Form
    {
        //private DataTable dtYeuThich;
        SqlConnection conn = new SqlConnection("Data Source=LAPTOP-D4IEITM3\\SQLEXPRESS02;Initial Catalog=DOAN;User ID=sa;Password=Sa@12345;TrustServerCertificate=True");
        SqlDataAdapter da = new SqlDataAdapter();
        SqlCommand cmd = new SqlCommand();
        DataTable dt = new DataTable();
        string sql, constr;
        public bool Expandmenu = false; // khai báo biến Expand2 
        public frmDAYT()
               

        {
            InitializeComponent();
            //KhoiTaoBang();

        }

        public frmDAYT(
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

            
        }




        public void NapCT()
        {
            int i = grdDAYT.CurrentRow.Index;
            txtTenDeTai.Text = grdDAYT.Rows[i].Cells[0].Value.ToString();
            txtTenSinhVien.Text = grdDAYT.Rows[i].Cells[1].Value.ToString();
            txtChuyenNganh.Text = grdDAYT.Rows[i].Cells[2].Value.ToString();
            txtKhoa.Text = grdDAYT.Rows[i].Cells[3].Value.ToString();
            txtGVHD.Text = grdDAYT.Rows[i].Cells[4].Value.ToString();
            dtpickNambaove.Text = grdDAYT.Rows[i].Cells[5].Value.ToString();
            txtTomTat.Text = grdDAYT.Rows[i].Cells[6].Value.ToString();



        }


        
        private void button1_Click(object sender, EventArgs e)
        {
            frmUser f = new frmUser();
            f.Show();
            this.Hide();
        }

        

        private void button4_Click(object sender, EventArgs e)
        {
            frmHome f = new frmHome();
            f.Show();
            this.Hide();
        }

        private void btndropQLDL_Click(object sender, EventArgs e)
        {

        }

        private void grdDoAn_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            NapCT();
        }

        private void pnlgrid_Paint(object sender, PaintEventArgs e)
        {

        }

        private void btnEnd_Click(object sender, EventArgs e)
        {
            int i = grdDAYT.Rows.Count - 1;
            grdDAYT.CurrentCell = grdDAYT[0, i-1];
            NapCT();
        }

        private void btnFirst_Click(object sender, EventArgs e)
        {
            grdDAYT.ClearSelection();
            grdDAYT.CurrentCell = grdDAYT[0, 0];
            NapCT();
        }

        private void btnPre_Click(object sender, EventArgs e)
        {
            int i = grdDAYT.CurrentRow.Index;
            if (i > 0)
            {
                grdDAYT.CurrentCell = grdDAYT[0, i - 1];
                NapCT();
            }
        }

        private void btnNext_Click(object sender, EventArgs e)
        {
            int i = grdDAYT.CurrentRow.Index;
            if (i < grdDAYT.Rows.Count - 1)
            {
                grdDAYT.CurrentCell = grdDAYT[0, i + 1];
                NapCT();
            }
        }

        private void btnDetail_Click(object sender, EventArgs e)
        {

            if (grdDAYT.CurrentRow == null)
            {
                MessageBox.Show("Vui lòng chọn một đồ án trong danh sách!");
                return;
            }

            int i = grdDAYT.CurrentRow.Index;

            // Lấy dữ liệu đầy đủ từ dòng đang chọn
            string tenDeTai = grdDAYT.Rows[i].Cells["TENDETAI"].Value.ToString();
            string tenSinhVien = grdDAYT.Rows[i].Cells["HOTEN"].Value.ToString();
            string chuyenNganh = grdDAYT.Rows[i].Cells["CHUYENNGANH"].Value.ToString();
            string khoa = grdDAYT.Rows[i].Cells["KHOA"].Value.ToString();
            string gvhd = grdDAYT.Rows[i].Cells["GVHD"].Value.ToString();
            string namBaoVe = grdDAYT.Rows[i].Cells["NAMBAOVE"].Value.ToString();
            string tomTat = grdDAYT.Rows[i].Cells["TOMTAT"].Value.ToString();

            // Các cột bị ẩn vẫn có thể lấy được dữ liệu
            string fileBaoCao = grdDAYT.Rows[i].Cells["FILEBC"].Value.ToString();
            string slide = grdDAYT.Rows[i].Cells["SLIDE"].Value.ToString();
            string lyLich = grdDAYT.Rows[i].Cells["LY_LICH"].Value.ToString();
            string lop = grdDAYT.Rows[i].Cells["LOP"].Value.ToString();
            string tuKhoa = grdDAYT.Rows[i].Cells["TUKHOA"].Value.ToString();
            string maDa = grdDAYT.Rows[i].Cells["MADOAN"].Value.ToString();




            // Mở form chi tiết và truyền dữ liệu qua
            frmChitiet f = new frmChitiet(
                tenDeTai,
                tenSinhVien,
                lop,
                gvhd,
                tuKhoa,
                namBaoVe,
                tomTat,
                fileBaoCao,
                slide,
                lyLich
            );

            f.Show();
            this.Hide();
        }

        private void btnMyfavorite_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Bạn có chắc chắn muốn xóa bản ghi hiện thời?", "Xác nhận yêu cầu",
MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                if (grdDAYT.CurrentRow != null)
                {
                    int i = grdDAYT.CurrentRow.Index; //lấy i ở dòng hiện thời
                    string maDa = grdDAYT.Rows[i].Cells["MADOAN"].Value.ToString();




                    // đảm bảo ko xóa nhầm
                    sql = "DELETE FROM DOAN_YEUTHICH WHERE MADOAN = N'" + maDa+ "'"; //xóa trên cở sở dữ liệu
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.ExecuteNonQuery();
                    grdDAYT.Rows.RemoveAt(i); // xóa ở ô lưới
                    NapCT();

                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            frmTracuu f = new frmTracuu();
            f.Show();
            this.Hide();
        }

        private void button3_Click(object sender, EventArgs e)
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

        private void pictureBox7_Click(object sender, EventArgs e)
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

        private void pictureBox1_Click(object sender, EventArgs e)
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

        private void pictureBox10_Click(object sender, EventArgs e)
        {
            menutimer.Start();
            Console.WriteLine(Expandmenu);
            pictureBox12.Visible = true;
            Console.ReadLine();
        }

        private void button13_Click(object sender, EventArgs e)
        {
            OpenHTML.OpenDefault();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            menutimer.Start();
            Console.WriteLine(Expandmenu);
            pictureBox12.Visible = true;
            Console.ReadLine();
        }

        private void frmDAYT_Load(object sender, EventArgs e)
        {
            sql = @"SELECT
            DA.TENDETAI,
            SV.HOTEN,
            SV.CHUYENNGANH,
            SV.KHOA,
            GVHD.GVHD,
            DA.NAMBAOVE,
            DA.TOMTAT,
	        SV.LOP,
	        STRING_AGG(TK.TUKHOA, ', ') AS TUKHOA,
            TL.FILEBC,
            TL.SLIDE,
            TL.LY_LICH,
            DA.MADOAN,
            SV.MASINHVIEN
            FROM DOAN_YEUTHICH Y
            JOIN DOAN DA          ON DA.MADOAN = Y.MADOAN
            JOIN SINHVIEN SV      ON SV.MASINHVIEN = DA.MASINHVIEN  
            JOIN GVHD ON GVHD.MAGVHD = DA.MAGVHD
            JOIN TAILIEUBC TL     ON TL.MATAILIEUBC = DA.MATAILIEUBC
            JOIN TUKHOA_DOAN TKDA ON TKDA.MADOAN = DA.MADOAN
            JOIN TUKHOA TK        ON TK.MATUKHOA = TKDA.MATUKHOA
            WHERE Y.MASINHVIEN = 'SV001' 
            GROUP BY 
            DA.TENDETAI,
            SV.HOTEN,
            SV.CHUYENNGANH,
            SV.KHOA,
            GVHD.GVHD,
            DA.NAMBAOVE,
            DA.TOMTAT,
            SV.LOP,
            TL.FILEBC,
            TL.SLIDE,
            TL.LY_LICH,
            DA.MADOAN,
            SV.MASINHVIEN";
            
            conn.Open();
            da = new SqlDataAdapter(sql, conn);
            da.Fill(dt);
            grdDAYT.DataSource = dt;
            grdDAYT.Refresh();
            NapCT();






            if (grdDAYT.Columns.Contains("FILEBC"))
                grdDAYT.Columns["FILEBC"].Visible = false;
            if (grdDAYT.Columns.Contains("SLIDE"))
                grdDAYT.Columns["SLIDE"].Visible = false;
            if (grdDAYT.Columns.Contains("LY_LICH"))
                grdDAYT.Columns["LY_LICH"].Visible = false;
            if (grdDAYT.Columns.Contains("LOP"))
                grdDAYT.Columns["LOP"].Visible = false;
            if (grdDAYT.Columns.Contains("TUKHOA"))
                grdDAYT.Columns["TUKHOA"].Visible = false;
            if (grdDAYT.Columns.Contains("MADOAN"))
                grdDAYT.Columns["MADOAN"].Visible = false;
            if(grdDAYT.Columns.Contains("MASINHVIEN"))
                grdDAYT.Columns["MASINHVIEN"].Visible = false;



        }
    }
}
