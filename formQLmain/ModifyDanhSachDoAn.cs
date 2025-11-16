using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace formQLmain
{
    class ModifyDanhSachDoAn
    {
        SqlDataAdapter _da;
        SqlCommand _cmd;

        // Đọc toàn bộ
        public DataTable getAllDoAn()
        {
            DataTable dt = new DataTable();
            string sql = @"SELECT DA.MADOAN,DA.TENDETAI,DA.MASINHVIEN, GVHD.MAGVHD, DA.NAMBAOVE, DA.TOMTAT,DA.MATAILIEUBC,TLBC.MATAILIEUBC,TLBC.FILEBC, TLBC.SLIDE,TLBC.LY_LICH,TK.MATUKHOA,TK.TUKHOA FROM DOAN DA 
                   JOIN TAILIEUBC TLBC ON DA.MATAILIEUBC=TLBC.MATAILIEUBC JOIN  TUKHOA_DOAN TKDA ON TKDA.MADOAN=DA.MADOAN JOIN TUKHOA TK ON TK.MATUKHOA=TKDA.MATUKHOA JOIN GVHD ON GVHD.MAGVHD=DA.MAGVHD  ;";
            using (SqlConnection conn = Connection.getConnection())
            {
                _da = new SqlDataAdapter(sql, conn);
                _da.Fill(dt);
            }
            return dt;
        }
        // ==============hàm check mã sv đã tồn tại chưa =============
        private bool CheckMaSV(string maSV, SqlConnection conn)
        {
            string sql = "SELECT COUNT(*) FROM SINHVIEN WHERE MASINHVIEN = @maSV";
            using (SqlCommand cmd = new SqlCommand(sql, conn))
            {
                cmd.Parameters.AddWithValue("@maSV", maSV);
                int count = (int)cmd.ExecuteScalar();
                return count > 0;
            }
        }
        // ==============hàm check mã GVHD đã tồn tại chưa =============
        private bool CheckMaGV(string gvhd, SqlConnection conn)
        {
            string sql = "SELECT COUNT(*) FROM GVHD WHERE MAGVHD = @gvhd";
            using (SqlCommand cmd = new SqlCommand(sql, conn))
            {
                cmd.Parameters.AddWithValue("@gvhd", gvhd);
                int count = (int)cmd.ExecuteScalar();
                return count > 0;
            }
        }

        // ============hàm insert =====================
        public bool insert(DanhSachDoAn da, out string error)
        {
            error = "";
            SqlConnection conn = Connection.getConnection();

            try
            {
                conn.Open();

                // == gọi quả hàm check mã vào đây 
                if (!CheckMaSV(da.MaSinhVien, conn))
                {
                    error = "Mã sinh viên này không tồn tại.";
                    return false;
                }

                // gọi hàm check mã gvhd
                // == gọi quả hàm check mã vào đây 
                if (!CheckMaGV(da.GVHD, conn))
                {
                    error = "Mã giảng viên này không tồn tại.";
                    return false;
                }
                // ------------------------------------------------------
                SqlTransaction tran = conn.BeginTransaction();

               
                // 1. Insert TAILIEUBC
                string sqlTLBC = @"
            INSERT INTO TAILIEUBC (FILEBC, SLIDE, LY_LICH)
            VALUES (@fileBC, @slide, @lylich);

            SELECT TOP 1 MATAILIEUBC 
            FROM TAILIEUBC 
            WHERE FILEBC = @fileBC 
              AND SLIDE = @slide 
              AND LY_LICH = @lylich
            ORDER BY MATAILIEUBC DESC;
        ";

                SqlCommand cmd = new SqlCommand(sqlTLBC, conn, tran);
                cmd.Parameters.AddWithValue("@fileBC", da.FileBC ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@slide", da.Slide ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@lylich", da.LyLich ?? (object)DBNull.Value);

                string maTLBC = cmd.ExecuteScalar().ToString();

                // -------------------------------
                // 2. Insert TUKHOA
                // -------------------------------
                string sqlTK = @"
            INSERT INTO TUKHOA (TUKHOA)
            VALUES (@tuKhoa);

            SELECT TOP 1 MATUKHOA
            FROM TUKHOA
            WHERE TUKHOA = @tuKhoa
            ORDER BY MATUKHOA DESC;
        ";

                cmd = new SqlCommand(sqlTK, conn, tran);
                cmd.Parameters.AddWithValue("@tuKhoa", da.TuKhoa ?? (object)DBNull.Value);

                string maTK = cmd.ExecuteScalar().ToString();

                // -------------------------------
                // 3. Insert DOAN
                // -------------------------------
                string sqlDA = @"
            INSERT INTO DOAN (TENDETAI, MASINHVIEN, MAGVHD, NAMBAOVE, TOMTAT, MATAILIEUBC)
            VALUES (@ten, @maSV, @gvhd, @namBV, @tomtat, @maTLBC);

            SELECT TOP 1 MADOAN
            FROM DOAN
            WHERE TENDETAI = @ten 
              AND MASINHVIEN = @maSV
              AND MATAILIEUBC = @maTLBC
            ORDER BY MADOAN DESC;
        ";

                cmd = new SqlCommand(sqlDA, conn, tran);
                cmd.Parameters.AddWithValue("@ten", da.TenDeTai ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@maSV", da.MaSinhVien ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@gvhd", da.GVHD ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@namBV", da.NamBaoVe ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@tomtat", da.TomTat ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@maTLBC", maTLBC);

                string maDA = cmd.ExecuteScalar().ToString();

                // -------------------------------
                // 4. Insert bảng TUKHOA_DOAN
                // -------------------------------
                string sqlTG = @"INSERT INTO TUKHOA_DOAN (MADOAN, MATUKHOA)
                         VALUES (@maDA, @maTK)";

                cmd = new SqlCommand(sqlTG, conn, tran);
                cmd.Parameters.AddWithValue("@maDA", maDA);
                cmd.Parameters.AddWithValue("@maTK", maTK);
                cmd.ExecuteNonQuery();

                // -------------------------------
                // DONE
                // -------------------------------
                tran.Commit();
                return true;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return false;
            }
            finally
            {
                conn.Close();
            }
        }

        // Cập nhật (cho phép đổi MADOAN thông qua maDoAnCu)
        public bool update(DanhSachDoAn da, string maDoAnCu, out string error)
        {
            error = "";
            SqlConnection conn = Connection.getConnection();

            try
            {
                conn.Open();
                SqlTransaction tran = conn.BeginTransaction();

                //------------------------------------------------------------------
                // 1. UPDATE TAILIEUBC
                //------------------------------------------------------------------
                string sqlTLBC = @"
            UPDATE TAILIEUBC
            SET FILEBC = @fileBC,
                SLIDE = @slide,
                LY_LICH = @lylich
            WHERE MATAILIEUBC = @maTLBC;
        ";

                SqlCommand cmd = new SqlCommand(sqlTLBC, conn, tran);
                cmd.Parameters.AddWithValue("@fileBC", da.FileBC ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@slide", da.Slide ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@lylich", da.LyLich ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@maTLBC", da.MaTaiLieuBC);
                cmd.ExecuteNonQuery();


                //------------------------------------------------------------------
                // 2. UPDATE TUKHOA
                //------------------------------------------------------------------
                string sqlTK = @"
            UPDATE TUKHOA
            SET TUKHOA = @tk
            WHERE MATUKHOA = @maTK;
        ";

                cmd = new SqlCommand(sqlTK, conn, tran);
                cmd.Parameters.AddWithValue("@tk", da.TuKhoa ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@maTK", da.MaTuKhoa);
                cmd.ExecuteNonQuery();


                //------------------------------------------------------------------
                // 3. UPDATE DOAN
                //------------------------------------------------------------------
                string sqlDA = @"
            UPDATE DOAN
            SET TENDETAI = @ten,
                MASINHVIEN = @maSV,
                MAGVHD = @gvhd,
                NAMBAOVE = @namBV,
                TOMTAT = @tomtat,
                MATAILIEUBC = @maTLBC
            WHERE MADOAN = @maDA;
        ";

                cmd = new SqlCommand(sqlDA, conn, tran);
                cmd.Parameters.AddWithValue("@ten", da.TenDeTai ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@maSV", da.MaSinhVien ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@gvhd", da.GVHD ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@namBV", da.NamBaoVe ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@tomtat", da.TomTat ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@maTLBC", da.MaTaiLieuBC);
                cmd.Parameters.AddWithValue("@maDA", da.MaDoAn);
                cmd.ExecuteNonQuery();


                //------------------------------------------------------------------
                // 4. UPDATE TUKHOA_DOAN – luôn update MATUKHOA nếu đổi từ khóa
                //------------------------------------------------------------------
                string sqlTG = @"
            UPDATE TUKHOA_DOAN
            SET MATUKHOA = @maTK
            WHERE MADOAN = @maDA;
        ";

                cmd = new SqlCommand(sqlTG, conn, tran);
                cmd.Parameters.AddWithValue("@maTK", da.MaTuKhoa);
                cmd.Parameters.AddWithValue("@maDA", da.MaDoAn);
                cmd.ExecuteNonQuery();

                tran.Commit();
                return true;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return false;
            }
            finally
            {
                conn.Close();
            }
        }

        // Xóa
        public bool delete(DanhSachDoAn da, out string error)
        {
            error = "";
            SqlConnection conn = Connection.getConnection();

            try
            {
                conn.Open();
                SqlTransaction tran = conn.BeginTransaction();

                // 1. Xóa trung gian
                SqlCommand cmd = new SqlCommand(
                    "DELETE FROM TUKHOA_DOAN WHERE MADOAN = @maDA",
                    conn, tran);
                cmd.Parameters.AddWithValue("@maDA", da.MaDoAn);
                cmd.ExecuteNonQuery();

                // 2. Xóa DOAN
                cmd = new SqlCommand(
                    "DELETE FROM DOAN WHERE MADOAN = @maDA",
                    conn, tran);
                cmd.Parameters.AddWithValue("@maDA", da.MaDoAn);
                cmd.ExecuteNonQuery();

                // 3. Xóa TAILIEUBC
                cmd = new SqlCommand(
                    "DELETE FROM TAILIEUBC WHERE MATAILIEUBC = @maTL",
                    conn, tran);
                cmd.Parameters.AddWithValue("@maTL", da.MaTaiLieuBC);
                cmd.ExecuteNonQuery();

                // 4. Xóa TUKHOA
                cmd = new SqlCommand(
                    "DELETE FROM TUKHOA WHERE MATUKHOA = @maTK",
                    conn, tran);
                cmd.Parameters.AddWithValue("@maTK", da.MaTuKhoa);
                cmd.ExecuteNonQuery();

                tran.Commit();
                return true;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return false;
            }
            finally
            {
                conn.Close();
            }
        }

        // (tuỳ chọn) kiểm tra trùng mã
        public bool exists(string maDoAn)
        {
            string sql = "SELECT COUNT(1) FROM DOAN WHERE MADOAN = @maDA";
            using (SqlConnection conn = Connection.getConnection())
            using (SqlCommand cmd = new SqlCommand(sql, conn))
            {
                cmd.Parameters.Add("@maDA", SqlDbType.VarChar).Value = maDoAn;
                conn.Open();
                return Convert.ToInt32(cmd.ExecuteScalar()) > 0;
            }
        }
        // Tìm theo mọi cột chính + ngày bảo vệ (yyyy-MM-dd)
        public DataTable searchDoAn(string keyword)
        {
            DataTable dt = new DataTable();
            string sql = @"
                SELECT DA.MADOAN,DA.TENDETAI,DA.MASINHVIEN, GVHD.MAGVHD, DA.NAMBAOVE, DA.TOMTAT,DA.MATAILIEUBC,TLBC.MATAILIEUBC,TLBC.FILEBC, TLBC.SLIDE,TLBC.LY_LICH,TK.MATUKHOA,TK.TUKHOA FROM DOAN DA 
                   JOIN TAILIEUBC TLBC ON DA.MATAILIEUBC=TLBC.MATAILIEUBC JOIN  TUKHOA_DOAN TKDA ON TKDA.MADOAN=DA.MADOAN JOIN TUKHOA TK ON TK.MATUKHOA=TKDA.MATUKHOA JOIN GVHD ON GVHD.MAGVHD=DA.MAGVHD 
                WHERE MADOAN      LIKE @kw
                   OR TENDETAI    LIKE @kw
                   OR MASINHVIEN  LIKE @kw
                   OR MAGVHD        LIKE @kw
                   OR TOMTAT      LIKE @kw
                   OR MATAILIEUBC LIKE @kw
                   OR CONVERT(varchar(10), NAMBAOVE, 120) LIKE @kw
                   OR MATAILIEUBC LIKE @KW";

            using (SqlConnection conn = Connection.getConnection())
            using (SqlCommand cmd = new SqlCommand(sql, conn))
            {
                cmd.Parameters.Add("@kw", SqlDbType.NVarChar).Value = "%" + keyword + "%";
                using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                {
                    da.Fill(dt);
                }
            }
            return dt;
        }
    }
}
