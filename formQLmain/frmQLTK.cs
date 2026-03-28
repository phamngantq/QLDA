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
    public partial class frmQLTK : Form
    {

        public frmQLTK()
        {
            InitializeComponent();
        }
        public bool Expand = false; // khai báo biến Expand 
        public bool Expand2 = false; // khai báo biến Expand2 
        public bool Expandmenu = false; // khai báo biến Expand2 

        // select = btndropQLDL
        // selectDropdown = QLDLdrop



        private void btnDSDA_Click(object sender, EventArgs e)
        {

            
            frmDSDA f = new frmDSDA();
            f.Show();
            this.Hide();
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

        }
        private ModifyTaiKhoan tkRepo;
        private string maTKCu = ""; // lưu mã TK gốc để cho phép đổi khóa
        private void frmQLTK_Load(object sender, EventArgs e) // quá là vip
        {
            tkRepo = new ModifyTaiKhoan();

            try
            {
                grdTaiKhoan.AutoGenerateColumns = true;     // nếu bạn không tự add cột
                grdTaiKhoan.DataSource = tkRepo.getAllTaiKhoan();
                grdTaiKhoan.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                grdTaiKhoan.MultiSelect = false;

                // ➜ Gắn sự kiện để đổ dữ liệu sang textbox
                grdTaiKhoan.CellClick += grdTaiKhoan_CellClick;
                grdTaiKhoan.SelectionChanged += grdTaiKhoan_SelectionChanged;

                // ➜ Chọn dòng đầu và sync dữ liệu sang textbox
                if (grdTaiKhoan.Rows.Count > 0)
                {
                    grdTaiKhoan.Rows[0].Selected = true;
                    SyncFromGrid(); // hàm đồng bộ bạn đã/ sẽ thêm
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi kết nối cơ sở dữ liệu: " + ex.Message,
                    "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnDrop2_Click(object sender, EventArgs e)
        {
            QLDAdrop.Start();
            Console.WriteLine(Expand2);
            Console.ReadLine();
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

        private void btndropQLDL_Click(object sender, EventArgs e)
        {

            QLDLdrop.Start();
            Console.WriteLine(Expand);
            Console.ReadLine();
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

        private void pnlgrid_Paint(object sender, PaintEventArgs e)
        {

        }
        private DateTime? ParseNullableDate(string text) //Thêm hàm parse ngày nullable
        {
            if (string.IsNullOrWhiteSpace(text)) return null;
            if (DateTime.TryParse(text, out var dt)) return dt;
            // Nếu muốn báo định dạng sai:
            // MessageBox.Show("Ngày không hợp lệ: " + text);
            return null;
        }
        private void ClearInputs() //Thêm hàm dọn ô nhập
        {
            txtbox_MaTK.Clear();
            txtbox_Email.Clear();
            txtbox_MatKhau.Clear();
            txtbox_NgayCap.Clear();
            txtbox_NgayCapNhat.Clear();
            txtbox_VaiTro.Clear();
        }
        private void btnThem_Click(object sender, EventArgs e)
        {
           
        }

        private void btnSua_Click(object sender, EventArgs e)
        {
            
        }

        private void btnXoa_Click(object sender, EventArgs e)
        {
            string maTK = txtbox_MaTK.Text.Trim();
            if (string.IsNullOrWhiteSpace(maTK) && grdTaiKhoan.CurrentRow != null)
                maTK = grdTaiKhoan.CurrentRow.Cells["MATK"].Value?.ToString();

            if (string.IsNullOrWhiteSpace(maTK))
            {
                MessageBox.Show("Vui lòng chọn tài khoản để xóa.",
                    "Thiếu thông tin", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var confirm = MessageBox.Show($"Bạn có chắc muốn xóa tài khoản: {maTK}?",
                "Xác nhận xóa", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (confirm != DialogResult.Yes) return;

            string err;
            if (tkRepo.delete(maTK, out err))
            {
                MessageBox.Show("Xóa tài khoản thành công.", "Thông báo",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                grdTaiKhoan.DataSource = tkRepo.getAllTaiKhoan();
                ClearInputs();
            }
            else
            {
                MessageBox.Show("Xóa thất bại.\nLỗi: " + err,
                    "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void grdTaiKhoan_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0) SyncFromGrid();
        }

        private void grdTaiKhoan_SelectionChanged(object sender, EventArgs e)
        {
            SyncFromGrid();
        }

        private void SyncFromGrid()
        {
            if (grdTaiKhoan.CurrentRow == null || grdTaiKhoan.CurrentRow.IsNewRow) return;

            var row = grdTaiKhoan.CurrentRow;
            txtbox_MaTK.Text = row.Cells["MATK"].Value?.ToString() ?? "";
            txtbox_Email.Text = row.Cells["EMAIL"].Value?.ToString() ?? "";
            txtbox_MatKhau.Text = row.Cells["MATKHAU"].Value?.ToString() ?? "";
            txtbox_VaiTro.Text = row.Cells["VAITRO"].Value?.ToString() ?? "";
            txtbox_NgayCap.Text = (row.Cells["NGAYCAP"].Value == null || row.Cells["NGAYCAP"].Value == DBNull.Value)
                                      ? "" : Convert.ToDateTime(row.Cells["NGAYCAP"].Value).ToString("yyyy-MM-dd HH:mm:ss");
            txtbox_NgayCapNhat.Text = (row.Cells["NGAYCAPNHAT"].Value == null || row.Cells["NGAYCAPNHAT"].Value == DBNull.Value)
                                      ? "" : Convert.ToDateTime(row.Cells["NGAYCAPNHAT"].Value).ToString("yyyy-MM-dd HH:mm:ss");

            maTKCu = txtbox_MaTK.Text; // lưu mã gốc để dùng khi Sửa
        }

        private void btnTimKiem_Click(object sender, EventArgs e)
        {
            string kw = textbox_TimKiem.Text.Trim();
            if (string.IsNullOrEmpty(kw))
            {
                grdTaiKhoan.DataSource = tkRepo.getAllTaiKhoan();
                return;
            }
            grdTaiKhoan.DataSource = tkRepo.searchTaiKhoan(kw);
        }

        private void textbox_TimKiem_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true;
                btnTimKiem_Click(sender, e);
            }
        }

        private void btnRefresh_Click(object sender, EventArgs e)
        {
            textbox_TimKiem.Clear();
            grdTaiKhoan.DataSource = tkRepo.getAllTaiKhoan();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            int i = grdTaiKhoan.Rows.Count - 1; //#########
            grdTaiKhoan.CurrentCell = grdTaiKhoan[0, i]; // nhảy đến dòng i (tức là dòng cuối  dòng cuối 
            MessageBox.Show("Bạn hãy nhập thông tin  vào các ô thông tin, sau đó nhấn nút Lưu");
            txtbox_Email.Focus();    // chuyển con trỏ đến textbox mã nhóm 
            ClearInputs();
            //#########################################################3
            btnLuu.Visible = true; // hiện nút Lưu lên 
            // nút lưu ở đây là nút thêm thật 
            // thêm mới là thêm giả 
        }

        private void pictureBox14_Click(object sender, EventArgs e)
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

        private void btnThem_Click_1(object sender, EventArgs e)
        {
            {
                var tk = new TaiKhoan(
                    txtbox_MaTK.Text.Trim(),
                    txtbox_Email.Text.Trim(),
                    txtbox_MatKhau.Text.Trim(),
                    ParseNullableDate(txtbox_NgayCap.Text),
                    ParseNullableDate(txtbox_NgayCapNhat.Text),
                    txtbox_VaiTro.Text.Trim()
                );

                string err;
                if (tkRepo.insert(tk, out err))
                {
                    MessageBox.Show("Thêm tài khoản thành công.", "Thông báo",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    grdTaiKhoan.DataSource = tkRepo.getAllTaiKhoan();
                    ClearInputs();
                }
                else
                {
                    MessageBox.Show("Thêm tài khoản thất bại.\nLỗi: " + err,
                        "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void btnSua_Click_1(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(maTKCu) && grdTaiKhoan.CurrentRow != null)
                maTKCu = grdTaiKhoan.CurrentRow.Cells["MATK"].Value?.ToString();

            if (string.IsNullOrWhiteSpace(maTKCu))
            {
                MessageBox.Show("Vui lòng chọn tài khoản cần sửa từ bảng.",
                    "Thiếu thông tin", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var tk = new TaiKhoan(
                txtbox_MaTK.Text.Trim(),
                txtbox_Email.Text.Trim(),
                txtbox_MatKhau.Text.Trim(),
                ParseNullableDate(txtbox_NgayCap.Text),
                ParseNullableDate(txtbox_NgayCapNhat.Text),
                txtbox_VaiTro.Text.Trim()
            );

            string err;
            if (tkRepo.update(tk, maTKCu, out err))
            {
                MessageBox.Show("Cập nhật tài khoản thành công.", "Thông báo",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);

                grdTaiKhoan.DataSource = tkRepo.getAllTaiKhoan();
                maTKCu = tk.MaTK; // cập nhật mã gốc nếu tiếp tục sửa
            }
            else
            {
                MessageBox.Show("Cập nhật thất bại.\n" + err,
                    "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            frmHome f = new frmHome();
            f.Show();
            this.Hide();
        }

        private void btnTimKiem_Click_1(object sender, EventArgs e)
        {
            string kw = textbox_TimKiem.Text.Trim();
            if (string.IsNullOrEmpty(kw))
            {
                grdTaiKhoan.DataSource = tkRepo.getAllTaiKhoan();
                return;
            }
            grdTaiKhoan.DataSource = tkRepo.searchTaiKhoan(kw);
        }

        private void textbox_TimKiem_KeyDown_1(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true;
                btnTimKiem_Click(sender, e);
            }
        }

        private void btnRefresh_Click_1(object sender, EventArgs e)
        {
            textbox_TimKiem.Clear();
            grdTaiKhoan.DataSource = tkRepo.getAllTaiKhoan();
        }

        private void btnThemMoi_Click(object sender, EventArgs e)
        {
            //####################################################3
            int i = grdTaiKhoan.Rows.Count - 1; //#########
            grdTaiKhoan.CurrentCell = grdTaiKhoan[0, i]; // nhảy đến dòng i (tức là dòng cuối  dòng cuối 
            MessageBox.Show("Bạn hãy nhập thông tin  vào các ô thông tin, sau đó nhấn nút Lưu");
            txtbox_MaTK.Focus();    // chuyển con trỏ đến textbox mã nhóm 
            ClearInputs();
            //#########################################################3
            btnLuu.Visible = true; // hiện nút Lưu lên 
            label_Thongbao.Visible = true;
            btn_Thoat.Visible = true;

            // nút lưu ở đây là nút thêm thật 
            // thêm mới là thêm giả 
        }

        private void button8_Click(object sender, EventArgs e)
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
            grdTaiKhoan.CurrentCell = grdTaiKhoan[0, 2]; // nhảy đến dòng 2
            SyncFromGrid();
            btnLuu.Visible = false; // hiện nút Lưu lên 
            label_Thongbao.Visible = false;
            btn_Thoat.Visible = false;
        }

        private void btn_Thoat2_Click(object sender, EventArgs e)
        {
            SyncFromGrid();
            btnCapNhat.Visible = false; // hiện nút Lưu lên 
            label_CapNhat.Visible = false;
            btn_Thoat2.Visible = false;
        }

        private void frmQLTK_KeyPress(object sender, KeyPressEventArgs e)
        {
            //Nếu được bấm trên bàn phím là Enter thì chuyển sang điều khiển tiếp theo, gửi vào bộ xử lí là phím tab
            if (e.KeyChar == (char)Keys.Enter) //nếu nút được bấm là Enter
            {
                SendKeys.Send("{TAB}");
                e.Handled = true;
            }
        }

        private void button13_Click(object sender, EventArgs e)
        {
            OpenHTML.OpenDefault();
        }

        private void panelMenu_Paint(object sender, PaintEventArgs e)
        {

        }

        private void button7_Click(object sender, EventArgs e)
        {
            menutimer.Start();
            Console.WriteLine(Expandmenu);
            pictureBox12.Visible = true;
            Console.ReadLine();
        }

        private void button12_Click(object sender, EventArgs e)
        {
            frmBaocao f = new frmBaocao();
            f.Show();
            this.Hide();
        }
    }
}
