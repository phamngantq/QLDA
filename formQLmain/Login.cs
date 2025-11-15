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
    public partial class FrmLogin : Form
    {
        string loginMode; // Biến lưu kiểu đăng nhập (USER hoặc ADMIN)
        bool isPasswordVisible = false;


        public FrmLogin(string mode)
        {
            InitializeComponent();
            loginMode = mode;
        }
        private void TogglePassword()
        {
            // Đảo trạng thái
            isPasswordVisible = !isPasswordVisible;

            // Áp dụng trạng thái
            txtPassword.UseSystemPasswordChar = !isPasswordVisible;

            // Đổi nút
            btnHien.Visible = isPasswordVisible;  // nút "Hiện"
            btnAn.Visible = !isPasswordVisible;  // nút "Ẩn"
        }

        private void btnLogin_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtEmail.Text) && string.IsNullOrWhiteSpace(txtPassword.Text))

            {
                MessageBox.Show("Vui lòng nhập đầy đủ thông tin");
                txtEmail.Focus(); // Đưa con trỏ về ô trống
                return; // Ngừng không cho chạy tiếp
            }


            if (string.IsNullOrWhiteSpace(txtEmail.Text))
            {
                MessageBox.Show("Vui lòng nhập Email!");
                txtEmail.Focus(); // Đưa con trỏ về ô trống
                return; // Ngừng không cho chạy tiếp
            }

            if (string.IsNullOrWhiteSpace(txtPassword.Text))
            {
                MessageBox.Show("Vui lòng nhập Mật khẩu!");
                txtPassword.Focus();
                return;
            }

            //Tạo một đối tượng chứa kết nối
            SqlConnection conn = new SqlConnection("Data Source=LAPTOP-D4IEITM3\\SQLEXPRESS02;Initial Catalog=DOAN;User ID=sa;Password=Sa@12345;TrustServerCertificate=True"); //khởi tạo đối tượng kết nối
                                                                                                                                                                               //tạo chuỗi truy vấn 
            string sql = " ";

            if (loginMode == "ADMIN")
            {
                sql = "SELECT * FROM TAIKHOAN WHERE EMAIL='" + txtEmail.Text +
                      "' AND MATKHAU='" + txtPassword.Text +
                      "' AND VAITRO='QUANLY'";
            }
            else // loginMode == "USER"
            {
                sql = "SELECT * FROM TAIKHOAN WHERE EMAIL='" + txtEmail.Text +
                      "' AND MATKHAU='" + txtPassword.Text +
                      "' AND (VAITRO='SINHVIEN' OR VAITRO='GIANGVIEN')";
            }
            // Mở kết nối Sql
            conn.Open();
            //Thực hiện thao tác với csdl chứa câu truy vấn sql và biết dùng kết nối conn để thực thi
            SqlCommand cmd = new SqlCommand(sql, conn);

            SqlDataAdapter da = new SqlDataAdapter(cmd);
            DataSet ds = new DataSet();
            da.Fill(ds); // thực thi cmd và đổ (fill) kết quả trả về vào DataSet ds.
            conn.Close();
            int count = ds.Tables[0].Rows.Count;  //kiểm tra số dòng trả về trong bảng đầu tiên của DataSet ds.
            if (count == 1)
            {
                MessageBox.Show("Đăng nhập thành công!");
                string role = ds.Tables[0].Rows[0]["VAITRO"].ToString();

                if (role == "SINHVIEN")
                {
                    frmTracuu f = new frmTracuu();
                    f.Show();
                }
                else if (role == "GIANGVIEN")
                {
                    frmTracuu f = new frmTracuu();
                    f.Show();
                }
                else if (role == "QUANLY")
                {
                    var f = new formQLmain.FrmQLmain();
                    f.Show();
                }

                this.Hide();
            }
            else
            {
                MessageBox.Show("Email hoặc mật khẩu không chính xác. Hãy thử lại");
            }
        }

        private void FrmLogin_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            frmHome f = new frmHome();
            f.Show();
            this.Hide(); // Đóng form hiện tại
        }

        private void guna2Panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void btnAn_Click(object sender, EventArgs e)
        {
            
            TogglePassword();

        }

        private void btnHien_Click(object sender, EventArgs e)
        {

            

            TogglePassword();

        }

        private void FrmLogin_KeyPress(object sender, KeyPressEventArgs e)
        {

            //Nếu được bấm trên bàn phím là Enter thì chuyển sang điều khiển tiếp theo, gửi vào bộ xử lí là phím tab
            if (e.KeyChar == (char)Keys.Enter) //nếu nút được bấm là Enter
            {
                SendKeys.Send("{TAB}");
                e.Handled = true;
            }


            }
    }
}
