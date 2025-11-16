using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;

namespace formQLmain
{
    public partial class FrmQLmain : Form
    {
        private string _userRole; // Biến lưu vai trò
        public FrmQLmain(string role)
        //public FrmQLmain()

        {

            InitializeComponent();
            _userRole = role; // Lưu vai trò từ FrmLogin
        }

        // Hàm tạo mặc định
        public FrmQLmain()
        {
            InitializeComponent();
            _userRole = string.Empty;
        }
        public bool Expand = false; // khai báo biến Expand 
        public bool Expand2 = false; // khai báo biến Expand2 
        public bool Expandmenu = false; // khai báo biến Expand2 

        // select = btndropQLDL
        // selectDropdown = QLDLdrop
        private void Select_Click_1(object sender, EventArgs e) // đổi tên thành btndropQLDL
        {
            QLDLdrop.Start();
            Console.WriteLine(Expand);
            Console.ReadLine();
        }

        private void selectDrowdown_Tick_1(object sender, EventArgs e) // đổi tên thành QLDLdrop 
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
        private readonly DashboardRepo _repo = new DashboardRepo();
        private void FrmQLmain_Load(object sender, EventArgs e)
        {
            try
            {
                LoadDashboard();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Không thể nạp dashboard.\n" + ex.Message, "Lỗi",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        //==============================



        private void btnDrop2_Click(object sender, EventArgs e)
        {
            QLDAdrop.Start();
            Console.WriteLine(Expand2);
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

        private void button8_Click(object sender, EventArgs e)
        {

        }

        private void dropdown_Paint(object sender, PaintEventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            frmQLTK f = new frmQLTK();
            f.Show();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            frmTracuu f = new frmTracuu();
            f.Show();
        }

        private void pictureBox11_Click(object sender, EventArgs e) // đồng hồ của trượt dọc 
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

        private void button1_Click(object sender, EventArgs e)
        {
         //   frmQLhome f = new frmQLhome();
          //  f.Show();
        }

       

        private void pictureBox3_Click(object sender, EventArgs e)
        {

        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void pictureBox12_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            frmQLSV f = new frmQLSV();
            f.Show();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            frmDSDA f = new frmDSDA();
            f.Show();
        }

        private void button4_Click(object sender, EventArgs e)
        {

        }

        private void panel11_Paint(object sender, PaintEventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void chartYearly_Click(object sender, EventArgs e)
        {

        }

        private void btnRefreshDash_Click(object sender, EventArgs e)
        {
            try
            {
                LoadDashboard();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Làm mới thất bại.\n" + ex.Message, "Lỗi",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        // ================== CORE ==================
        private void LoadDashboard()
        {
            // 1) 4 con số tổng
            var totals = _repo.GetTotals();
            lblTotalDoAn.Text = totals.DoAn.ToString("N0");
            lblTotalSinhVien.Text = totals.SinhVien.ToString("N0");
            lblTotalTaiKhoan.Text = totals.TaiKhoan.ToString("N0");
            lblTotalTaiLieu.Text = totals.TaiLieu.ToString("N0");

            // 2) Chart: đề án theo năm (cột)
            BindChartYearly();

            // 3) Chart: đề án theo khoa (donut/pie)
            BindChartByKhoa();

            // 4) Grid: tài khoản mới
            BindRecentAccounts();
        }

        private void BindChartYearly()
        {
            var dt = _repo.GetProjectsByYear(years: 5);

            chartYearly.Series.Clear();
            chartYearly.ChartAreas.Clear();
            chartYearly.Titles.Clear();

            var area = new ChartArea("area1");
            area.AxisX.Interval = 1;
            area.AxisX.MajorGrid.Enabled = false;
            area.AxisY.MajorGrid.LineDashStyle = ChartDashStyle.Dot;
            chartYearly.ChartAreas.Add(area);

            var series = new Series("Số đề án");
            series.ChartType = SeriesChartType.Column;
            series.IsValueShownAsLabel = true;
            series.LabelFormat = "N0";

            foreach (DataRow r in dt.Rows)
            {
                var nam = r["Nam"]?.ToString();
                var sl = Convert.ToInt32(r["SoLuong"]);
                series.Points.AddXY(nam, sl);
            }

            chartYearly.Series.Add(series);
            chartYearly.Titles.Add("Đề án theo năm");
        }

        private void BindChartByKhoa()
        {
            var dt = _repo.GetProjectsByKhoa();

            chartByKhoa.Series.Clear();
            chartByKhoa.ChartAreas.Clear();
            chartByKhoa.Titles.Clear();
            chartByKhoa.Legends.Clear();

            var area = new ChartArea("area1");
            area.Area3DStyle.Enable3D = false;
            chartByKhoa.ChartAreas.Add(area);

            var legend = new Legend("legend1");
            legend.Docking = Docking.Right;
            chartByKhoa.Legends.Add(legend);

            var series = new Series("Theo khoa");
            series.ChartType = SeriesChartType.Doughnut;   // đổi Pie nếu bạn thích
            series.IsValueShownAsLabel = true;
            series.Label = "#VAL (#PERCENT{P1})";          // Hiển thị số + % trên lát
            series.LegendText = "#VALX";                   // Hiển thị tên khoa ở legend

            foreach (DataRow r in dt.Rows)
            {
                var khoa = r["Khoa"]?.ToString();
                var sl = Convert.ToInt32(r["SoLuong"]);
                series.Points.AddXY(khoa, sl);
            }

            chartByKhoa.Series.Add(series);
            chartByKhoa.Titles.Add("Phân bổ đề án theo khoa");
        }

        private void BindRecentAccounts()
        {
            var dt = _repo.GetRecentAccounts(10);
            grdRecent.DataSource = dt;

            // Tùy chọn: định dạng cột
            if (grdRecent.Columns.Contains("NGAYCAP"))
            {
                grdRecent.Columns["NGAYCAP"].HeaderText = "Ngày cấp";
            }
            if (grdRecent.Columns.Contains("VAITRO"))
            {
                grdRecent.Columns["VAITRO"].HeaderText = "Vai trò";
            }
        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void panel5_Click(object sender, EventArgs e)
        {

        }

        private void pictureBox7_Click(object sender, EventArgs e)
        {

        }

        private void pictureBox13_Click(object sender, EventArgs e)
        {
            menutimer.Start();
            Console.WriteLine(Expandmenu);
            pictureBox12.Visible = true;
            Console.ReadLine();
            
            
        }

        private void pictureBox12_Click_1(object sender, EventArgs e)
        {

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

        private void pictureBox2_Click(object sender, EventArgs e)
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

        private void button12_Click(object sender, EventArgs e)
        {
            frmBaocao f = new frmBaocao(_userRole);
            f.Show();
            this.Hide();
        }
    }
    }

