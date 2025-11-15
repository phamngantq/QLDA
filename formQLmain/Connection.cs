using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;

namespace formQLmain
{
    class Connection
    {
        //private static string stringConnection = @"Data Source=DESKTOP-QQ88INT\\SQLEXPRESS;Initial Catalog=QLDA;Integrated Security=True;TrustServerCertificate=True";
        // private static string stringConnection = @"Data Source=LAPTOP-D4IEITM3\SQLEXPRESS02;Initial Catalog=DOAN;User ID=sa;Password=Sa@12345;Encrypt=True;TrustServerCertificate=True";
        private static string stringConnection = @"Data Source=LAPTOP-D4IEITM3\SQLEXPRESS02;Initial Catalog=DOAN;User ID=sa; password=Sa@12345;TrustServerCertificate=True";
        //private static string stringConnection = @"Data Source=DESKTOP-QQ88INT\SQLEXPRESS;Initial Catalog=DOAN1211;Integrated Security=True;TrustServerCertificate=True;Encrypt=False";

        public static  SqlConnection getConnection()  // sqlConnection là ổ khóa
        {
            return new SqlConnection(stringConnection); // stringconnection là chìa 
        } 
    }
}
