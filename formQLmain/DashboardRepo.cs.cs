using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace formQLmain
{
    // DTO nhỏ để trả về 4 số tổng
    internal class DashboardTotals
    {
        public int DoAn { get; set; }
        public int SinhVien { get; set; }
        public int TaiKhoan { get; set; }
        public int TaiLieu { get; set; }
    }

    internal class DashboardRepo
    {
        // --------- Tổng số bản ghi cho 4 bảng chính ----------
        public DashboardTotals GetTotals()
        {
            var t = new DashboardTotals();

            using (var conn = Connection.getConnection())
            {
                conn.Open();

                using (var cmd = new SqlCommand("SELECT COUNT(*) FROM DOAN", conn))
                    t.DoAn = (int)cmd.ExecuteScalar();

                using (var cmd = new SqlCommand("SELECT COUNT(*) FROM SINHVIEN", conn))
                    t.SinhVien = (int)cmd.ExecuteScalar();

                using (var cmd = new SqlCommand("SELECT COUNT(*) FROM TAIKHOAN", conn))
                    t.TaiKhoan = (int)cmd.ExecuteScalar();

                using (var cmd = new SqlCommand("SELECT COUNT(*) FROM TAILIEUBC", conn))
                    t.TaiLieu = (int)cmd.ExecuteScalar();
            }

            return t;
        }

        // --------- Biểu đồ đề án theo N năm gần nhất ----------
        // Lưu ý: nếu cột NAMBAOVE là NVARCHAR, dùng TRY_CONVERT để lấy YEAR.
        public DataTable GetProjectsByYear(int years = 5)
        {
            var dt = new DataTable();

            string sql = @"
                SELECT  YEAR(TRY_CONVERT(date, NAMBAOVE)) AS Nam,
                        COUNT(*) AS SoLuong
                FROM DOAN
                WHERE TRY_CONVERT(date, NAMBAOVE) IS NOT NULL
                  AND YEAR(TRY_CONVERT(date, NAMBAOVE)) >= YEAR(GETDATE()) - @years + 1
                GROUP BY YEAR(TRY_CONVERT(date, NAMBAOVE))
                ORDER BY Nam;";

            using (var conn = Connection.getConnection())
            using (var cmd = new SqlCommand(sql, conn))
            {
                cmd.Parameters.Add("@years", SqlDbType.Int).Value = years;
                using (var da = new SqlDataAdapter(cmd))
                {
                    da.Fill(dt);
                }
            }

            return dt;
        }

        // --------- Pie/Donut theo khoa (join SINHVIEN) ----------
        public DataTable GetProjectsByKhoa()
        {
            var dt = new DataTable();

            string sql = @"
                SELECT  sv.KHOA       AS Khoa,
                        COUNT(*)      AS SoLuong
                FROM DOAN d
                INNER JOIN SINHVIEN sv ON d.MASINHVIEN = sv.MASINHVIEN
                GROUP BY sv.KHOA
                ORDER BY SoLuong DESC;";

            using (var conn = Connection.getConnection())
            using (var cmd = new SqlCommand(sql, conn))
            using (var da = new SqlDataAdapter(cmd))
            {
                da.Fill(dt);
            }

            return dt;
        }

        // --------- 10 tài khoản mới cấp ----------
        // Nếu NGAYCAP là NVARCHAR, bạn có thể thay ORDER BY bằng TRY_CONVERT(date, NGAYCAP) DESC
        public DataTable GetRecentAccounts(int take = 10)
        {
            var dt = new DataTable();

            string sql = @"
                SELECT TOP (@take)
                       MATK, EMAIL, VAITRO, NGAYCAP
                FROM TAIKHOAN
                ORDER BY 
                    CASE WHEN TRY_CONVERT(datetime, NGAYCAP, 105) IS NOT NULL 
                         THEN TRY_CONVERT(datetime, NGAYCAP, 105) 
                         ELSE TRY_CONVERT(datetime, NGAYCAP) 
                    END DESC, NGAYCAP DESC;";

            using (var conn = Connection.getConnection())
            using (var cmd = new SqlCommand(sql, conn))
            {
                cmd.Parameters.Add("@take", SqlDbType.Int).Value = take;
                using (var da = new SqlDataAdapter(cmd))
                {
                    da.Fill(dt);
                }
            }

            return dt;
        }
        public DataTable GetRecentDoAn(int take = 10)
        {
            var dt = new DataTable();

            string sql = @"
        SELECT TOP (@take)
               MADOAN, TENDETAI, MASINHVIEN, NAMBAOVE
        FROM DOAN
        ORDER BY 
            CASE WHEN TRY_CONVERT(datetime, NAMBAOVE, 105) IS NOT NULL 
                 THEN TRY_CONVERT(datetime, NAMBAOVE, 105) 
                 ELSE TRY_CONVERT(datetime, NAMBAOVE) 
            END DESC, NAMBAOVE DESC;";

            using (var conn = Connection.getConnection())
            using (var cmd = new SqlCommand(sql, conn))
            {
                cmd.Parameters.Add("@take", SqlDbType.Int).Value = take;

                using (var da = new SqlDataAdapter(cmd))
                {
                    da.Fill(dt);
                }
            }

            return dt;
        }
    }
}
