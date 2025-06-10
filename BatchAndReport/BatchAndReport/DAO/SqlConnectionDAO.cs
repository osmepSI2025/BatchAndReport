using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;

namespace BatchAndReport.DAO
{
    public class SqlConnectionDAO
    {
        private readonly IConfiguration _configuration;

        public SqlConnectionDAO(IConfiguration configuration)
        {
            _configuration = configuration;
        }

        // เลือก connection ตามชื่อ (ที่กำหนดไว้ใน appsettings.json)
        public SqlConnection GetConnection(string connectionName = "DefaultConnection")
        {
            var connectionString = _configuration.GetConnectionString(connectionName);
            return new SqlConnection(connectionString);
        }
    }
}