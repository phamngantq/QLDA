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
    public partial class frmDSDA : Form
    {
        private string _userRole; // Biến lưu vai trò


        SqlConnection conn = new SqlConnection("Data Source=LAPTOP-D4IEITM3\\SQLEXPRESS02;Initial Catalog=DOAN;User ID=sa;Password=Sa@12345;TrustServerCertificate=True");
        SqlDataAdapter da = new SqlDataAdapter();
        SqlCommand cmd = new SqlCommand();
        DataTable dt = new DataTable();
        string sql, constr;
        public frmDSDA(string userRole)
        {
            InitializeComponent();
            _userRole = userRole;
        }
        public frmDSDA()
        {
            InitializeComponent();
            _userRole = string.Empty;
            
        }

        public void NapCT()
        {
            int i = grdDoAn.CurrentRow.Index;
            txtbox_MaDoAn.Text = grdDoAn.Rows[i].Cells[0].Value.ToString();
            txtbox_TenDeTai.Text = grdDoAn.Rows[i].Cells[1].Value.ToString();
            txtbox_MaSV.Text = grdDoAn.Rows[i].Cells[2].Value.ToString();
            txtbox_GVHD.Text = grdDoAn.Rows[i].Cells[3].Value.ToString();
            dtpick_Nambaove.Text = grdDoAn.Rows[i].Cells[4].Value.ToString();
            txtbox_MaTLBC.Text = grdDoAn.Rows[i].Cells[5].Value.ToString();
            txtbox_TomTat.Text = grdDoAn.Rows[i].Cells[6].Value.ToString();
            txtbox_MaTLBC2.Text = grdDoAn.Rows[i].Cells[7].Value.ToString();
            txtbox_FileBC.Text = grdDoAn.Rows[i].Cells[8].Value.ToString();
            txtbox_Slide.Text = grdDoAn.Rows[i].Cells[9].Value.ToString();
            txtbox_LyLich.Text = grdDoAn.Rows[i].Cells[10].Value.ToString();
            txtbox_MaTK.Text = grdDoAn.Rows[i].Cells[11].Value.ToString();
            txtbox_TK.Text = grdDoAn.Rows[i].Cells[12].Value.ToString();
        }






            //txtbox_MaDoAn.Text = row.Cells["MADOAN"].Value?.ToString() ?? "";
            //txtbox_TenDeTai.Text = row.Cells["TENDETAI"].Value?.ToString() ?? "";
            //txtbox_MaSV.Text = row.Cells["MASINHVIEN"].Value?.ToString() ?? "";
            //txtbox_GVHD.Text = row.Cells["GVHD"].Value?.ToString() ?? "";
            //txtbox_MaTLBC.Text = row.Cells["MATAILIEUBC"].Value?.ToString() ?? "";
            //txtbox_TomTat.Text = row.Cells["TOMTAT"].Value?.ToString() ?? "";
            ////
            //txtbox_MaTLBC2.Text = row.Cells["MATAILIEUBC"].Value?.ToString() ?? "";
            //txtbox_FileBC.Text = row.Cells["FILEBC"].Value?.ToString() ?? "";
            //txtbox_Slide.Text = row.Cells["SLIDE"].Value?.ToString() ?? "";
            //txtbox_LyLich.Text = row.Cells["LY_LICH"].Value?.ToString() ?? "";
            //txtbox_MaTK.Text = row.Cells["MATUKHOA"].Value?.ToString() ?? "";
            //txtbox_TK.Text = row.Cells["TUKHOA"].Value?.ToString() ?? "";
            //}

        public bool Expand = false; // khai báo biến Expand 
        public bool Expand2 = false; // khai báo biến Expand2 
        public bool Expandmenu = false; // khai báo biến Expand2 
        private void btnDrop2_Click(object sender, EventArgs e)
        {
            QLDAdrop.Start();
            Console.WriteLine(Expand2);
            Console.ReadLine();
        }

        private void btnDSDA_Click(object sender, EventArgs e)
        {
            
        }

        private void btnTLKT_Click(object sender, EventArgs e)
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

        private void btndropQLDL_Click(object sender, EventArgs e)
        {
            QLDLdrop.Start();
            Console.WriteLine(Expand);
            Console.ReadLine();
        }
        private ModifyDanhSachDoAn repo;
        private string maDoAnCu = "";   // lưu MADOAN gốc để cho phép đổi khóa khi sửa
        private string maTLCu = ""; // lưu mã TLBC gốc để cho phép đổi khóa khi sửa
        // ================ LOAD FORM ================
        // ================ LOAD FORM ================
        private void frmDSDA_Load(object sender, EventArgs e) /// from load 
        {
            repo = new ModifyDanhSachDoAn();

            try
            {
                grdDoAn.AutoGenerateColumns = true; // nếu không tự add cột
                grdDoAn.DataSource = repo.getAllDoAn();
                grdDoAn.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                grdDoAn.MultiSelect = false;

                grdDoAn.CellClick += grdDoAn_CellClick;
                grdDoAn.SelectionChanged += grdDoAn_SelectionChanged;

                // DateTimePicker cho phép null bằng checkbox
                dtpick_Nambaove.ShowCheckBox = true;

                if (grdDoAn.Rows.Count > 0)
                {
                    grdDoAn.Rows[0].Selected = true;
                    SyncFromGrid();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi tải dữ liệu: " + ex.Message,
                                "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
 
        // ================ THÊM ================
        private void btnThem_Click_1(object sender, EventArgs e)
        {
            var da = new DanhSachDoAn(
                txtbox_MaDoAn.Text.Trim(),
                txtbox_TenDeTai.Text.Trim(),
                txtbox_MaSV.Text.Trim(),
                txtbox_GVHD.Text.Trim(),
                dtpick_Nambaove.Checked ? dtpick_Nambaove.Value.Date : (DateTime?)null,
                txtbox_MaTLBC.Text.Trim(),
                txtbox_TomTat.Text.Trim(),
                txtbox_MaTLBC2.Text.Trim(),
                txtbox_FileBC.Text.Trim(),
                txtbox_Slide.Text.Trim(),
                txtbox_LyLich.Text.Trim(),
                txtbox_MaTK.Text.Trim(),
                txtbox_TK.Text.Trim()

            );

            string err;
            if (repo.insert(da, out err))
            {
                MessageBox.Show("Thêm đồ án thành công.", "Thông báo",
                                MessageBoxButtons.OK, MessageBoxIcon.Information);
                grdDoAn.DataSource = repo.getAllDoAn();
                ClearInputs();
            }
            else
            {
                MessageBox.Show("Thêm thất bại.\nLỗi: " + err,
                                "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        // ================ XÓA ================
        private void btnXoa_Click_1(object sender, EventArgs e)
        {
        }

        // ================ SỬA (có thể đổi MADOAN) ================
        private void btnSua_Click_1(object sender, EventArgs e) // Sửa cũ 
        {
            
        }
        // ================== ĐỒNG BỘ GRID → CONTROL ==================
        private void grdDoAn_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0) SyncFromGrid();
        }

        private void grdDoAn_SelectionChanged(object sender, EventArgs e)
        {
            SyncFromGrid();
        }




      private void SyncFromGrid()
        {
            if (grdDoAn.CurrentRow == null || grdDoAn.CurrentRow.IsNewRow) return;

            var row = grdDoAn.CurrentRow;

            txtbox_MaDoAn.Text = row.Cells["MADOAN"].Value?.ToString() ?? "";
            txtbox_TenDeTai.Text = row.Cells["TENDETAI"].Value?.ToString() ?? "";
            txtbox_MaSV.Text = row.Cells["MASINHVIEN"].Value?.ToString() ?? "";
            txtbox_GVHD.Text = row.Cells["MAGVHD"].Value?.ToString() ?? "";
            txtbox_MaTLBC.Text = row.Cells["MATAILIEUBC"].Value?.ToString() ?? "";
            txtbox_TomTat.Text = row.Cells["TOMTAT"].Value?.ToString() ?? "";
            //
            txtbox_MaTLBC2.Text = row.Cells["MATAILIEUBC"].Value?.ToString() ?? "";
            txtbox_FileBC.Text = row.Cells["FILEBC"].Value?.ToString() ?? "";
            txtbox_Slide.Text = row.Cells["SLIDE"].Value?.ToString() ?? "";
            txtbox_LyLich.Text = row.Cells["LY_LICH"].Value?.ToString() ?? "";
            txtbox_MaTK.Text = row.Cells["MATUKHOA"].Value?.ToString() ?? "";
            txtbox_TK.Text = row.Cells["TUKHOA"].Value?.ToString() ?? "";

            // DateTimePicker: set null/giá trị
            var cell = row.Cells["NAMBAOVE"].Value;
            if (cell == null || cell == DBNull.Value)
            {
                dtpick_Nambaove.Checked = false;
            }
            else
            {
                dtpick_Nambaove.Checked = true;
                dtpick_Nambaove.Value = Convert.ToDateTime(cell);
            }

            maDoAnCu = txtbox_MaDoAn.Text;



            //int i = grdDoAn.CurrentRow.Index;
            //txtbox_MaDoAn.Text = grdDoAn.Rows[i].Cells[0].Value.ToString();
            //txtbox_TenDeTai.Text = grdDoAn.Rows[i].Cells[1].Value.ToString();
            //txtbox_MaSV.Text = grdDoAn.Rows[i].Cells[2].Value.ToString();
            //txtbox_GVHD.Text = grdDoAn.Rows[i].Cells[3].Value.ToString();
            //dtpick_Nambaove.Text = grdDoAn.Rows[i].Cells[4].Value.ToString();
            //txtbox_MaTLBC.Text = grdDoAn.Rows[i].Cells[5].Value.ToString();
            //txtbox_TomTat.Text = grdDoAn.Rows[i].Cells[6].Value.ToString();
            //txtbox_MaTLBC2.Text = grdDoAn.Rows[i].Cells[7].Value.ToString();
            //txtbox_FileBC.Text = grdDoAn.Rows[i].Cells[8].Value.ToString();
            //txtbox_Slide.Text = grdDoAn.Rows[i].Cells[9].Value.ToString();
            //txtbox_LyLich.Text = grdDoAn.Rows[i].Cells[10].Value.ToString();
            //txtbox_MaTK.Text = grdDoAn.Rows[i].Cells[11].Value.ToString();
            //txtbox_TK.Text = grdDoAn.Rows[i].Cells[12].Value.ToString();


            //// DateTimePicker: set null/giá trị
            //var cell = grdDoAn.Rows[i].Cells["NAMBAOVE"].Value;
            //if (cell == null || cell == DBNull.Value)
            //{
            //    dtpick_Nambaove.Checked = false;
            //}
            //else
            //{
            //    dtpick_Nambaove.Checked = true;
            //    dtpick_Nambaove.Value = Convert.ToDateTime(cell);
            //}

            //maDoAnCu = txtbox_MaDoAn.Text;


            //int i = grdDoAn.CurrentRow.Index;
            //txtbox_MaDoAn.Text = grdDoAn.Rows[i].Cells[0].Value.ToString();
            //txtbox_TenDeTai.Text = grdDoAn.Rows[i].Cells[1].Value.ToString();
            //txtbox_MaSV.Text = grdDoAn.Rows[i].Cells[2].Value.ToString();
            //txtbox_GVHD.Text = grdDoAn.Rows[i].Cells[3].Value.ToString();
            //txtbox_MaTLBC.Text = grdDoAn.Rows[i].Cells[4].Value.ToString();
            //txtbox_TomTat.Text = grdDoAn.Rows[i].Cells[5].Value.ToString();
            //txtbox_MaTLBC2.Text = grdDoAn.Rows[i].Cells[6].Value.ToString();
            //txtbox_FileBC.Text = grdDoAn.Rows[i].Cells[7].Value.ToString();
            //txtbox_Slide.Text = grdDoAn.Rows[i].Cells[8].Value.ToString();
            //txtbox_LyLich.Text = grdDoAn.Rows[i].Cells[9].Value.ToString();
            //txtbox_MaTK.Text = grdDoAn.Rows[i].Cells[10].Value.ToString();
            //txtbox_TK.Text = grdDoAn.Rows[i].Cells[11].Value.ToString();

        }

        // ================== TIỆN ÍCH ==================
        private void ClearInputs()
        {
            txtbox_MaDoAn.Clear();
            txtbox_TenDeTai.Clear();
            txtbox_MaSV.Clear();
            txtbox_GVHD.Clear();
            txtbox_MaTLBC.Clear();
            txtbox_MaTLBC2.Clear();
            txtbox_TomTat.Clear();
            dtpick_Nambaove.Checked = false;
            txtbox_MaTLBC.Clear();
            txtbox_FileBC.Clear();
            txtbox_Slide.Clear();
            txtbox_LyLich.Clear();
            txtbox_MaTK.Clear();
            txtbox_TK.Clear();

            maDoAnCu = "";
        }

        //private void btnTimKiem_Click(object sender, EventArgs e)
        //{
        //    string kw = textbox_TimKiem.Text.Trim();
        //    if (string.IsNullOrEmpty(kw))
        //    {
        //        grdDoAn.DataSource = repo.getAllDoAn();
        //        return;
        //    }
        //    grdDoAn.DataSource = repo.searchDoAn(kw);
        //}

        //private void btnRefresh_Click(object sender, EventArgs e)
        //{
        //    textbox_TimKiem.Clear();
        //    grdDoAn.DataSource = repo.getAllDoAn();
        //}

        //private void textbox_TimKiem_KeyDown(object sender, KeyEventArgs e)
        //{
        //    if (e.KeyCode == Keys.Enter)
        //    {
        //        e.SuppressKeyPress = true; // tránh tiếng 'ding'
        //        btnTimKiem_Click(sender, e);
        //    }
        //}

        private void grdDoAn_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void panel10_Paint(object sender, PaintEventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e) // nút thêm giả 
        {
            int i = grdDoAn.Rows.Count - 1; //#########
            grdDoAn.CurrentCell = grdDoAn[0, i]; // nhảy đến dòng i (tức là dòng cuối  dòng cuối 
            MessageBox.Show("Bạn hãy nhập thông tin  vào các ô thông tin, sau đó nhấn nút Lưu");
            txtbox_MaDoAn.Focus();    // chuyển con trỏ đến textbox mã nhóm 
            ClearInputs();
            //#########################################################3
            btnLuu.Visible = true; // hiện nút Lưu lên 
            // nút lưu ở đây là nút thêm thật 
            // thêm mới là thêm giả 
            label_Thongbao.Visible = true;
            btn_Thoat.Visible = true;
        }

        private void btnLuu_Click(object sender, EventArgs e)
        {
            //
            string maSV = txtbox_MaSV.Text.Trim();

            // Nếu để trống
            if (string.IsNullOrEmpty(maSV))
            {
                MessageBox.Show("Vui lòng nhập mã sinh viên.", "Cảnh báo",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            var da = new DanhSachDoAn(
                null,                               // MADOAN -> để null
                txtbox_TenDeTai.Text.Trim(),        // Tên đề tài
                txtbox_MaSV.Text.Trim(),            // Mã sinh viên
                txtbox_GVHD.Text.Trim(),            // GVHD
                dtpick_Nambaove.Checked ? dtpick_Nambaove.Value.Date : (DateTime?)null,

                null,                               // MATAILIEUBC (1) -> để null
                txtbox_TomTat.Text.Trim(),          // Tóm tắt

                null,                               // MATAILIEUBC (2) -> bỏ
                txtbox_FileBC.Text.Trim(),          // FILE BC
                txtbox_Slide.Text.Trim(),           // Slide
                txtbox_LyLich.Text.Trim(),          // Lý lịch

                null,                               // MATUKHOA -> để null
                txtbox_TK.Text.Trim()               // Từ khóa
            );

            string err;
            if (repo.insert(da, out err))
            {
                MessageBox.Show("Thêm đồ án thành công.", "Thông báo",
                                MessageBoxButtons.OK, MessageBoxIcon.Information);

                grdDoAn.DataSource = repo.getAllDoAn();
                ClearInputs();
            }
            else
            {
                MessageBox.Show("Thêm thất bại.\nLỗi: " + err,
                                "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnSua_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(maDoAnCu) && grdDoAn.CurrentRow != null)
                maDoAnCu = grdDoAn.CurrentRow.Cells["MADOAN"].Value?.ToString();

            if (string.IsNullOrWhiteSpace(maDoAnCu))
            {
                MessageBox.Show("Vui lòng chọn đồ án cần sửa từ bảng.",
                                "Thiếu thông tin", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var da = new DanhSachDoAn(
                txtbox_MaDoAn.Text.Trim(),
                txtbox_TenDeTai.Text.Trim(),
                txtbox_MaSV.Text.Trim(),
                txtbox_GVHD.Text.Trim(),
                dtpick_Nambaove.Checked ? dtpick_Nambaove.Value.Date : (DateTime?)null,
                txtbox_MaTLBC.Text.Trim(),
                txtbox_TomTat.Text.Trim(),
                txtbox_MaTLBC2.Text.Trim(),
                txtbox_FileBC.Text.Trim(),
                txtbox_Slide.Text.Trim(),
                txtbox_LyLich.Text.Trim(),
                txtbox_MaTK.Text.Trim(),
                txtbox_TK.Text.Trim()
            );

            string err;
            if (repo.update(da, maDoAnCu, out err))
            {
                MessageBox.Show("Cập nhật đồ án thành công.", "Thông báo",
                                MessageBoxButtons.OK, MessageBoxIcon.Information);
                grdDoAn.DataSource = repo.getAllDoAn();
                maDoAnCu = da.MaDoAn; // cập nhật lại mã gốc nếu tiếp tục sửa
            }
            else
            {
                MessageBox.Show("Cập nhật thất bại.\n" + err,
                                "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnXoa_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtbox_MaDoAn.Text))
            {
                MessageBox.Show("Bạn chưa chọn đồ án để xóa!",
                                "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var confirm = MessageBox.Show(
                $"Bạn có chắc muốn xóa đồ án {txtbox_MaDoAn.Text.Trim()} không?",
                "Xác nhận",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (confirm == DialogResult.No)
                return;

            // Tạo object DanhSachDoAn chỉ chứa thông tin cần thiết cho xóa toàn bộ
            DanhSachDoAn da = new DanhSachDoAn()
            {
                MaDoAn = txtbox_MaDoAn.Text.Trim(),
                MaTaiLieuBC = txtbox_MaTLBC.Text.Trim(),
                MaTuKhoa = txtbox_MaTK.Text.Trim()
            };

            string err;
            if (repo.delete(da, out err))
            {
                MessageBox.Show("Xóa thành công!", "Thông báo",
                                MessageBoxButtons.OK, MessageBoxIcon.Information);

                grdDoAn.DataSource = repo.getAllDoAn();
                ClearInputs();
            }
            else
            {
                MessageBox.Show("Xóa thất bại.\nLỗi: " + err,
                                "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button5_Click_1(object sender, EventArgs e)
        {
            sql = @"SELECT DA.MADOAN,DA.TENDETAI,DA.MASINHVIEN, DA.GVHD, DA.NAMBAOVE, DA.TOMTAT,DA.MATAILIEUBC,TLBC.MATAILIEUBC,TLBC.FILEBC,
TLBC.SLIDE,TLBC.LY_LICH,TK.MATUKHOA,TK.TUKHOA FROM DOAN DA 
JOIN TAILIEUBC TLBC ON DA.MATAILIEUBC=TLBC.MATAILIEUBC JOIN  TUKHOA_DOAN TKDA ON TKDA.MADOAN=DA.MADOAN JOIN TUKHOA TK ON TK.MATUKHOA=TKDA.MATUKHOA   
WHERE " + comTruong.Text + "= N'" + comGT.Text + "'";// Đảm bảo mọi dữ liệu có tiếng việt vẫn lọc được
            da = new SqlDataAdapter(sql, conn);
            dt = new DataTable();
            da.Fill(dt);
            grdDoAn.DataSource = dt;
            grdDoAn.Refresh();
            //SyncFromGrid();
            NapCT();
        }

        private void pnlgrid_Paint(object sender, PaintEventArgs e)
        {

        }

        private void button6_Click(object sender, EventArgs e)
        {
            sql = @" SELECT DA.MADOAN,DA.TENDETAI,DA.MASINHVIEN, DA.GVHD, DA.NAMBAOVE, DA.TOMTAT,DA.MATAILIEUBC,TLBC.MATAILIEUBC,TLBC.FILEBC,
TLBC.SLIDE,TLBC.LY_LICH,TK.MATUKHOA,TK.TUKHOA FROM DOAN DA 
JOIN TAILIEUBC TLBC ON DA.MATAILIEUBC=TLBC.MATAILIEUBC JOIN  TUKHOA_DOAN TKDA ON TKDA.MADOAN=DA.MADOAN JOIN TUKHOA TK ON TK.MATUKHOA=TKDA.MATUKHOA   ";
            da = new SqlDataAdapter(sql, conn);
            dt = new DataTable();
            dt.Clear();
            da.Fill(dt);
            grdDoAn.DataSource = dt;
            grdDoAn.Refresh();
            NapCT();

            comTruong.SelectedIndex = -1;
            comGT.SelectedIndex = -1;
            grdDoAn.DataSource = repo.getAllDoAn(); ;



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

        private void button7_Click(object sender, EventArgs e)
        {
            SyncFromGrid();
            MessageBox.Show("Bạn hãy nhập thông tin cần thay đổi vào các ô thông tin, sau đó nhấn nút Cập nhật");
            //txtbox_MaSV.Focus();    // chuyển con trỏ đến textbox mã nhóm 
            btnCapNhat.Visible = true; // hiện nút cập nhật lên
            label_CapNhat.Visible = true;
            btn_Thoat2.Visible = true;
        }

        private void btn_Thoat_Click(object sender, EventArgs e)
        {
            grdDoAn.CurrentCell = grdDoAn[0, 2]; // nhảy đến dòng 2
            SyncFromGrid();
            btnLuu.Visible = false; // hiện nút Lưu lên 
            label_Thongbao.Visible = false;
            btn_Thoat.Visible = false;
        }

        private void btn_Thoat2_Click(object sender, EventArgs e)
        {
            SyncFromGrid();
            btnCapNhat.Visible = false; // ẩn nút cập nhật  
            label_CapNhat.Visible = false;
            btn_Thoat2.Visible = false;
        }

        private void textbox_TimKiem2_TextChanged(object sender, EventArgs e)
        {

        }

        private void panelALL_Paint(object sender, PaintEventArgs e)
        {

        }

        private void textbox_TimKiem2_Click(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            frmHome f = new frmHome();
            f.Show();
            this.Hide();
        }

        private void btnRefresh_Click(object sender, EventArgs e)
        {
            textbox_TimKiem.Clear();
            grdDoAn.DataSource = repo.getAllDoAn();
        }

        private void btnTimKiem_Click(object sender, EventArgs e)
        {
            string kw = textbox_TimKiem.Text.Trim();
            if (string.IsNullOrEmpty(kw))
            {
                grdDoAn.DataSource = repo.getAllDoAn();
                return;
            }
            grdDoAn.DataSource = repo.searchDoAn(kw);
        }

        private void textbox_TimKiem_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true; // tránh tiếng 'ding'
                btnTimKiem_Click(sender, e);
            }
        }

        private void panel10_Paint_1(object sender, PaintEventArgs e)
        {

        }

        private void frmDSDA_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                SendKeys.Send("{TAB}");  // đây là lệnh gửi phím tab
                e.Handled = true;
            }
        }

        private void button13_Click(object sender, EventArgs e)
        {
            OpenHTML.OpenDefault();
        }

        private void button6_Click_1(object sender, EventArgs e)
        {
            if (_userRole.Equals("GIANGVIEN", StringComparison.OrdinalIgnoreCase))
            {
                MessageBox.Show("Tài khoản Giảng viên không có quyền lập báo cáo.", "Không Có Quyền", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return; // Ngừng thực thi và không tiến hành in


            }
            frmBaocao f = new frmBaocao();
            f.Show();
            this.Hide();
        }

        private void button5_Click_2(object sender, EventArgs e)
        {
            menutimer.Start();
            Console.WriteLine(Expandmenu);
            pictureBox12.Visible = true;
            Console.ReadLine();
        }

        

        //private void textbox_TimKiem_KeyDown_1(object sender, KeyEventArgs e)
        //{
        //    if (e.KeyCode == Keys.Enter)
        //    {
        //        e.SuppressKeyPress = true; // tránh tiếng 'ding'
        //        btnTimKiem_Click(sender, e);
        //    }
        //}

        private void comTruong_SelectedIndexChanged(object sender, EventArgs e)
        {
            sql = "Select distinct " + comTruong.Text + " FROM DOAN DA JOIN TAILIEUBC TLBC ON DA.MATAILIEUBC=TLBC.MATAILIEUBC JOIN  TUKHOA_DOAN TKDA ON TKDA.MADOAN=DA.MADOAN JOIN TUKHOA TK ON TK.MATUKHOA=TKDA.MATUKHOA  ";
            da = new SqlDataAdapter(sql, conn);
            DataTable dt1 = new DataTable();
            da.Fill(dt1);
            //comGT.Items.Clear();
            comGT.DataSource = dt1;
            comGT.DisplayMember = comTruong.Text; //trường hiện ra
            comGT.ValueMember = comTruong.Text;// trường để lấy 
        }
    }
}
//ghi chú