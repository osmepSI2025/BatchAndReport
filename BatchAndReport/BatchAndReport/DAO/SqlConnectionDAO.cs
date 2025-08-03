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

        public SqlConnection GetConnectionK2Econctract(string connectionName = "K2DBContext_EContract")
        {
            var connectionString = _configuration.GetConnectionString(connectionName);
            return new SqlConnection(connectionString);
        }
        public SqlConnection GetConnectionWorkflow(string connectionName = "K2DBContext_Workflow")
        {
            var connectionString = _configuration.GetConnectionString(connectionName);
            return new SqlConnection(connectionString);
        }
        public SqlConnection GetConnectionHR(string connectionName = "K2DBContext")
        {
            var connectionString = _configuration.GetConnectionString(connectionName);
            return new SqlConnection(connectionString);
        }
        public SqlConnection GetConnectionK2DBContext_SME(string connectionName = "K2DBContext_SME")
        {
            var connectionString = _configuration.GetConnectionString(connectionName);
            return new SqlConnection(connectionString);
        }
        
    }
}