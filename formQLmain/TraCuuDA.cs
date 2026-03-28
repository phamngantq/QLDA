//using DevExpress.XtraRichEdit.Import.Html;
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
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Menu;


namespace formQLmain
{

    public partial class frmTracuu : Form
    {
        SqlConnection conn=new SqlConnection("Data Source=LAPTOP-D4IEITM3\\SQLEXPRESS02;Initial Catalog=DOAN;User ID=sa;Password=Sa@12345;TrustServerCertificate=True");
        SqlDataAdapter da = new SqlDataAdapter();
        SqlCommand cmd = new SqlCommand();
        DataTable dt = new DataTable();
        string sql, constr;
        public bool Expand = false; // khai báo biến Expand 
        public bool Expand2 = false; // khai báo biến Expand2 
        public bool Expandmenu = false; // khai báo biến Expand2 


        //private frmDAYT _frmDayt;

        public frmTracuu()
        {
            InitializeComponent();
        }

        public void NapCT()
        {
            int i = grdTracuu.CurrentRow.Index;
            txtTenDeTai.Text = grdTracuu.Rows[i].Cells[0].Value.ToString();
            txtTenSinhVien.Text = grdTracuu.Rows[i].Cells[1].Value.ToString();
            txtChuyenNganh.Text = grdTracuu.Rows[i].Cells[2].Value.ToString();
            txtKhoa.Text = grdTracuu.Rows[i].Cells[3].Value.ToString();
            txtGVHD.Text = grdTracuu.Rows[i].Cells[4].Value.ToString();
            dtpickNambaove.Text = grdTracuu.Rows[i].Cells[5].Value.ToString();
            txtTomTat.Text = grdTracuu.Rows[i].Cells[6].Value.ToString();
            


        }


        


        private void TimKiemTheoTuKhoa(string keyword)
        {

            // SELECT + JOIN (không có GROUP BY)
            string selectSql = @"
SELECT 
    DA.TENDETAI,
    SV.HOTEN,
    SV.CHUYENNGANH,
    SV.KHOA,
    GVHD.GVHD,
    DA.NAMBAOVE,
    DA.TOMTAT,
    TLBC.FILEBC,
    TLBC.SLIDE,
    TLBC.LY_LICH,
    SV.LOP,
    STRING_AGG(TK.TUKHOA, ', ') AS TUKHOA,
    SV.MASINHVIEN,
    DA.MADOAN
FROM DOAN DA
JOIN SINHVIEN SV   ON DA.MASINHVIEN = SV.MASINHVIEN
JOIN GVHD  ON DA.MAGVHD = GVHD.MAGVHD
JOIN TAILIEUBC TLBC ON DA.MATAILIEUBC = TLBC.MATAILIEUBC
JOIN TUKHOA_DOAN TKDA ON TKDA.MADOAN = DA.MADOAN
JOIN TUKHOA TK ON TK.MATUKHOA = TKDA.MATUKHOA";

            // GROUP BY (cuối cùng) - giữ nguyên các cột không aggregate
            string groupBy = @"
GROUP BY 
    DA.TENDETAI,
    SV.HOTEN,
    SV.CHUYENNGANH,
    SV.KHOA,
    GVHD.GVHD,
    DA.NAMBAOVE,
    DA.TOMTAT,
    TLBC.FILEBC,
    TLBC.SLIDE,
    TLBC.LY_LICH,
    SV.LOP,
    SV.MASINHVIEN,
    DA.MADOAN
";

            DataTable dtLocal = new DataTable();

            if (string.IsNullOrWhiteSpace(keyword))
            {
                // không filter: SELECT + GROUP BY
                sql = selectSql + groupBy;
                da = new SqlDataAdapter(sql, conn);
            }
            else
            {
                // có filter: thêm WHERE trước GROUP BY (và dùng parameter)
                sql = selectSql + @"
WHERE DA.TENDETAI    LIKE @kw
   OR SV.HOTEN       LIKE @kw
   OR SV.CHUYENNGANH LIKE @kw
   OR SV.KHOA        LIKE @kw
   OR GVHD.GVHD        LIKE @kw
";
                sql += groupBy;

                da = new SqlDataAdapter(sql, conn);
                da.SelectCommand.Parameters.Clear();
                da.SelectCommand.Parameters.AddWithValue("@kw", "%" + keyword.Trim() + "%");
            }

            // Fill DataTable an toàn (không dùng conn.Open() toàn cục ở đây)
            dt = new DataTable();
            da.Fill(dt);

            grdTracuu.DataSource = dt;
            grdTracuu.Refresh();

            if (dt.Rows.Count > 0)
                NapCT();   // nhớ NapCT đã an toàn với currentRow null
        }

        



        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void frmTracuu_Load(object sender, EventArgs e)
        {
            sql = @"SELECT 
    DA.TENDETAI,
    SV.HOTEN,
    SV.CHUYENNGANH,
    SV.KHOA,
    GVHD.GVHD,
    DA.NAMBAOVE,
    DA.TOMTAT,
    TLBC.FILEBC,
    TLBC.SLIDE,
    TLBC.LY_LICH,
    SV.LOP,
    STRING_AGG(TK.TUKHOA, ', ') AS TUKHOA,
    SV.MASINHVIEN,
    DA.MADOAN
FROM DOAN DA
JOIN SINHVIEN SV   ON DA.MASINHVIEN = SV.MASINHVIEN
JOIN GVHD  ON DA.MAGVHD = GVHD.MAGVHD
JOIN TAILIEUBC TLBC ON DA.MATAILIEUBC = TLBC.MATAILIEUBC
JOIN TUKHOA_DOAN TKDA ON TKDA.MADOAN = DA.MADOAN
JOIN TUKHOA TK ON TK.MATUKHOA = TKDA.MATUKHOA
GROUP BY 
    DA.TENDETAI,
    SV.HOTEN,
    SV.CHUYENNGANH,
    SV.KHOA,
    GVHD.GVHD,
    DA.NAMBAOVE,
    DA.TOMTAT,
    TLBC.FILEBC,
    TLBC.SLIDE,
    TLBC.LY_LICH,
    SV.LOP,
    SV.MASINHVIEN,
    DA.MADOAN ";
            conn.Open();
            da = new SqlDataAdapter(sql, conn);
            da.Fill(dt);
            grdTracuu.DataSource = dt;
            grdTracuu.Refresh();
            NapCT();
            TimKiemTheoTuKhoa(""); // load toàn bộ dữ liệu lúc ban đầu

            

           
            if (grdTracuu.Columns.Contains("FILEBC"))
                grdTracuu.Columns["FILEBC"].Visible = false;
            if (grdTracuu.Columns.Contains("SLIDE"))
                grdTracuu.Columns["SLIDE"].Visible = false;
            if (grdTracuu.Columns.Contains("LY_LICH"))
                grdTracuu.Columns["LY_LICH"].Visible = false;
            if (grdTracuu.Columns.Contains("LOP"))
                grdTracuu.Columns["LOP"].Visible = false;
            if (grdTracuu.Columns.Contains("TUKHOA"))
                grdTracuu.Columns["TUKHOA"].Visible = false;
            if (grdTracuu.Columns.Contains("MASINHVIEN"))
                grdTracuu.Columns["MASINHVIEN"].Visible = false;
            if (grdTracuu.Columns.Contains("MADOAN"))
                grdTracuu.Columns["MADOAN"].Visible = false;



        }

        

      
   

        private void grdDoAn_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            NapCT();// khi click vào ô nào đó thì NapCT() sẽ được gọi


        }

        private void btnRefresh_Click(object sender, EventArgs e)
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
    DA.TOMTAT,
    TLBC.FILEBC,
    TLBC.SLIDE,
    TLBC.LY_LICH,
    SV.LOP,
    STRING_AGG(TK.TUKHOA, ', ') AS TUKHOA,
    SV.MASINHVIEN,
    DA.MADOAN
FROM DOAN DA
JOIN SINHVIEN SV   ON DA.MASINHVIEN = SV.MASINHVIEN
JOIN GVHD  ON DA.MAGVHD = GVHD.MAGVHD
JOIN TAILIEUBC TLBC ON DA.MATAILIEUBC = TLBC.MATAILIEUBC
JOIN TUKHOA_DOAN TKDA ON TKDA.MADOAN = DA.MADOAN
JOIN TUKHOA TK ON TK.MATUKHOA = TKDA.MATUKHOA
GROUP BY 
    DA.TENDETAI,
    SV.HOTEN,
    SV.CHUYENNGANH,
    SV.KHOA,
    GVHD.GVHD,
    DA.NAMBAOVE,
    DA.TOMTAT,
    TLBC.FILEBC,
    TLBC.SLIDE,
    TLBC.LY_LICH,
    SV.LOP,
    SV.MASINHVIEN,
    DA.MADOAN";
                da = new SqlDataAdapter(sql, conn);
            dt = new DataTable();
            dt.Clear();
            da.Fill(dt);
            grdTracuu.DataSource = dt;
            grdTracuu.Refresh();
            NapCT(); }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi làm mới dữ liệu: " + ex.Message);
            }
            comTruong.SelectedIndex = -1;
            comGT.SelectedIndex = -1;
            // 2️⃣ Reset toàn bộ phần lọc và tìm kiếm
            comTruong.SelectedIndex = -1;  // Bỏ chọn tên trường
            comGT.DataSource = null;       // Xóa dữ liệu trong combo giá trị
            comGT.Text = "";               // Làm trống text hiển thị

            txtTukhoa.Clear();             //  Xóa ô tìm kiếm về rỗng






        }

        private void pnlgrid_Paint(object sender, PaintEventArgs e)
        {

        }

        private void btnDrop2_Click(object sender, EventArgs e)
        {
               
        }

        private void guna2Button2_Click(object sender, EventArgs e)
        {
            

            if (grdTracuu.CurrentRow == null)
            {
                MessageBox.Show("Vui lòng chọn một đồ án trong danh sách!");
                return;
            }

            int i = grdTracuu.CurrentRow.Index;

            // Lấy dữ liệu đầy đủ từ dòng đang chọn
            string tenDeTai = grdTracuu.Rows[i].Cells["TENDETAI"].Value.ToString();
            string tenSinhVien = grdTracuu.Rows[i].Cells["HOTEN"].Value.ToString();
            string chuyenNganh = grdTracuu.Rows[i].Cells["CHUYENNGANH"].Value.ToString();
            string khoa = grdTracuu.Rows[i].Cells["KHOA"].Value.ToString();
            string gvhd = grdTracuu.Rows[i].Cells["GVHD"].Value.ToString();
            string namBaoVe = grdTracuu.Rows[i].Cells["NAMBAOVE"].Value.ToString();
            string tomTat = grdTracuu.Rows[i].Cells["TOMTAT"].Value.ToString();

            // Các cột bị ẩn vẫn có thể lấy được dữ liệu
            string fileBaoCao = grdTracuu.Rows[i].Cells["FILEBC"].Value.ToString();
            string slide = grdTracuu.Rows[i].Cells["SLIDE"].Value.ToString();
            string lyLich = grdTracuu.Rows[i].Cells["LY_LICH"].Value.ToString();
            string lop = grdTracuu.Rows[i].Cells["LOP"].Value.ToString();
            string tuKhoa = grdTracuu.Rows[i].Cells["TUKHOA"].Value.ToString();


            

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

        


        

        private void comTruong_SelectedIndexChanged(object sender, EventArgs e)
        {

            try
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
            catch
            {

            }



        }
        

        private void btnFillter_Click(object sender, EventArgs e)
        {
            sql = " SELECT  DA.TENDETAI, SV.HOTEN, SV.CHUYENNGANH, SV.KHOA, GVHD.GVHD, DA.NAMBAOVE, DA.TOMTAT FROM DOAN DA JOIN SINHVIEN SV ON DA.MASINHVIEN = SV.MASINHVIEN JOIN GVHD  ON DA.MAGVHD = GVHD.MAGVHD WHERE " + comTruong.Text + "= N'" + comGT.Text + "'";// Đảm bảo mọi dữ liệu có tiếng việt vẫn lọc được
            da = new SqlDataAdapter(sql, conn);
            dt = new DataTable();
            da.Fill(dt);
            grdTracuu.DataSource = dt;
            grdTracuu.Refresh();
            NapCT();
        }

        private void btnFirst_Click(object sender, EventArgs e)
        {
            grdTracuu.ClearSelection();
            grdTracuu.CurrentCell = grdTracuu[0, 0];
            NapCT();
        }

        private void btnEnd_Click(object sender, EventArgs e)
        {
            int i = grdTracuu.Rows.Count - 1;
            grdTracuu.CurrentCell = grdTracuu[0, i-1];
            NapCT();
        }

        private void btnPre_Click(object sender, EventArgs e)
        {
            int i = grdTracuu.CurrentRow.Index;
            if (i > 0)
            {
                grdTracuu.CurrentCell = grdTracuu[0, i - 1];
                NapCT();
            }
        }

        private void btnNext_Click(object sender, EventArgs e)
        {
            int i = grdTracuu.CurrentRow.Index;
            if (i < grdTracuu.Rows.Count - 1)
            {
                grdTracuu.CurrentCell = grdTracuu[0, i + 1];
                NapCT();
            }
        }

        private void btnTimKiem_Click(object sender, EventArgs e)
        {
            TimKiemTheoTuKhoa(txtTukhoa.Text);



        }

        private void button4_Click(object sender, EventArgs e)
        {
            frmHome f = new frmHome();
            f.Show();
            this.Hide();
        }

        private void txtTenSinhVien_TextChanged(object sender, EventArgs e)
        {

        }

        private void txtKhoa_TextChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            
        }

        private void guna2ImageButton1_Click(object sender, EventArgs e)
        {




            if (grdTracuu.CurrentRow == null)
            {
                MessageBox.Show("Vui lòng chọn một đồ án!");
                return;
            }

            int i = grdTracuu.CurrentRow.Index;
            string maSinhVien = "SV001"; // có thể truyền từ form đăng nhập
            string maDoAn = grdTracuu.Rows[i].Cells["MADOAN"].Value.ToString();

            using (SqlConnection conn = new SqlConnection("Data Source=LAPTOP-D4IEITM3\\SQLEXPRESS02;Initial Catalog=DOAN;User ID=sa;Password=Sa@12345;TrustServerCertificate=True"))
            {
                conn.Open();

                // Kiểm tra trùng
                string sqlCheck = "SELECT COUNT(*) FROM DOAN_YEUTHICH WHERE MASINHVIEN=@masv AND MADOAN=@madoan";
                SqlCommand cmdCheck = new SqlCommand(sqlCheck, conn);
                cmdCheck.Parameters.AddWithValue("@masv", maSinhVien);
                cmdCheck.Parameters.AddWithValue("@madoan", maDoAn);

                int count = (int)cmdCheck.ExecuteScalar();
                if (count > 0)
                {
                    MessageBox.Show("Đồ án này đã có trong danh sách yêu thích!");
                    return;
                }

                // Thêm mới
                string sqlInsert = "INSERT INTO DOAN_YEUTHICH (MASINHVIEN, MADOAN) VALUES (@masv, @madoan)";
                SqlCommand cmd = new SqlCommand(sqlInsert, conn);
                cmd.Parameters.AddWithValue("@masv", maSinhVien);
                cmd.Parameters.AddWithValue("@madoan", maDoAn);
                cmd.ExecuteNonQuery();

                MessageBox.Show("Đã thêm vào danh sách yêu thích!");
            }
        }








        
        

        private void btnDSDA_Click(object sender, EventArgs e)
        {
            
        }

        private void btnTLKT_Click(object sender, EventArgs e)
        {
            
        }

        

       

        private void flowLayoutPanel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void panel4_Paint(object sender, PaintEventArgs e)
        {

        }

        private void panelMenu_Paint(object sender, PaintEventArgs e)
        {

        }

        private void dropdown2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void panel6_Paint(object sender, PaintEventArgs e)
        {

        }

        private void panel7_Paint(object sender, PaintEventArgs e)
        {

        }

        private void panel8_Paint(object sender, PaintEventArgs e)
        {

        }

        private void trogiup_Paint(object sender, PaintEventArgs e)
        {

        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void panel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {

        }

        private void panel3_Paint(object sender, PaintEventArgs e)
        {

        }

        private void button4_Click_1(object sender, EventArgs e)
        {

        }

        private void flowLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void panel9_Paint(object sender, PaintEventArgs e)
        {

        }

        private void panel11_Paint(object sender, PaintEventArgs e)
        {

        }

        private void comGT_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void frmEnd_Click(object sender, EventArgs e)
        {

        }

        private void panel10_Paint(object sender, PaintEventArgs e)
        {

        }

        private void label10_Click(object sender, EventArgs e)
        {

        }

        private void txtChuyenNganh_TextChanged(object sender, EventArgs e)
        {

        }

        private void label8_Click_1(object sender, EventArgs e)
        {

        }

        private void txtTomTat_TextChanged(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void label9_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void dtpickNambaove_ValueChanged(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void txtGVHD_TextChanged(object sender, EventArgs e)
        {

        }

        private void txtTenDeTai_TextChanged(object sender, EventArgs e)
        {

        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void label11_Click(object sender, EventArgs e)
        {

        }

        private void txtTukhoa_TextChanged(object sender, EventArgs e)
        {

        }

        private void toolTipMyfavorite_Popup(object sender, PopupEventArgs e)
        {

        }

        private void toolTipRefresh_Popup(object sender, PopupEventArgs e)
        {

        }

        private void toolTipResearch_Popup(object sender, PopupEventArgs e)
        {

        }

        private void pictureBox4_Click(object sender, EventArgs e)
        {

        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {

        }

        private void pictureBox7_Click(object sender, EventArgs e)
        {

        }

        private void pictureBox8_Click(object sender, EventArgs e)
        {

        }

        private void pictureBox10_Click(object sender, EventArgs e)
        {

        }

        private void pictureBox5_Click(object sender, EventArgs e)
        {

        }

        private void pictureBox6_Click(object sender, EventArgs e)
        {

        }

        private void pictureBox11_Click(object sender, EventArgs e)
        {

        }

        private void pictureBox12_Click(object sender, EventArgs e)
        {

        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }

        private void pictureBox13_Click(object sender, EventArgs e)
        {

        }

        private void frmTracuu_KeyDown(object sender, KeyEventArgs e)
        {
            // Kiểm tra xem nút nào đã được nhấn
            switch (e.KeyCode)
            {
                case Keys.Up:
                    // Hành động cho nút "Trước" (Previous)
                    btnPre_Click(null, null);
                    e.Handled = true; // Ngăn không cho sự kiện lan truyền thêm
                    break;

                case Keys.Down:
                    // Hành động cho nút "Sau" (Next)
                    btnNext_Click(null, null);
                    e.Handled = true;
                    break;

                case Keys.Home:
                    // Tùy chọn: Dùng phím Home cho nút "Đầu"
                    btnFirst_Click(null, null);
                    e.Handled = true;
                    break;

                case Keys.End:
                    // Tùy chọn: Dùng phím End cho nút "Cuối"
                    btnEnd_Click(null, null);
                    e.Handled = true;
                    break;
            }
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

        private void pictureBox7_Click_1(object sender, EventArgs e)
        {
            menutimer.Start();
            Console.WriteLine(Expandmenu);
            pictureBox12.Visible = true;
            Console.ReadLine();
        }

        private void pictureBox8_Click_1(object sender, EventArgs e)
        {
            menutimer.Start();
            Console.WriteLine(Expandmenu);
            pictureBox12.Visible = true;
            Console.ReadLine();
        }

        private void pictureBox3_Click_1(object sender, EventArgs e)
        {
            menutimer.Start();
            Console.WriteLine(Expandmenu);
            pictureBox12.Visible = true;
            Console.ReadLine();
        }

        private void pictureBox10_Click_1(object sender, EventArgs e)
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

        private void button2_Click(object sender, EventArgs e)
        {
            frmTracuu f = new frmTracuu();
            f.Show();
            this.Hide();
        }

        private void btnLove_Click(object sender, EventArgs e)
        {
            frmDAYT f = new frmDAYT();
            f.Show();
            this.Hide();
        }

        private void button4_Click_2(object sender, EventArgs e)
        {
            frmUser f = new frmUser();
            f.Show();
            this.Hide();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            frmHome f = new frmHome();
            f.Show();
            this.Hide();
        }

        private void panelALL_Paint(object sender, PaintEventArgs e)
        {

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

        private void panelMenu_Paint_1(object sender, PaintEventArgs e)
        {
            
        }

        private void label8_Click(object sender, EventArgs e)
        {

        }
    }
}
