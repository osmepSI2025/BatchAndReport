using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;

namespace BatchAndReport.DAO
{
    public class SqlConnectionDAO
    {
        private readonly string _connectionString;

        public SqlConnectionDAO(IConfiguration configuration)
        {
            // "DefaultConnection" should match your connection string name in appsettings.json
            _connectionString = configuration.GetConnectionString("DefaultConnection");
        }

        public SqlConnection GetConnection()
        {
            return new SqlConnection(_connectionString);
        }
    }
}