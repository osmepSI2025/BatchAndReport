using Microsoft.Data.SqlClient;
using System.Collections.Generic;
using System.Threading.Tasks;
using BatchAndReport.Entities;

namespace BatchAndReport.DAO
{
    public class MscheduledJobDAO
    {
        private readonly SqlConnectionDAO _connectionDAO;

        public MscheduledJobDAO(SqlConnectionDAO connectionDAO)
        {
            _connectionDAO = connectionDAO;
        }

        // CREATE
        public async Task<int> AddJobAsync(MscheduledJob job)
        {
            using var conn = _connectionDAO.GetConnection();
            using var cmd = new SqlCommand(
                @"INSERT INTO MscheduledJob (JobName, RunHour, RunMinute, IsActive) 
                  VALUES (@name, @hour, @minute, @active)", conn);
            cmd.Parameters.AddWithValue("@name", job.JobName ?? (object)DBNull.Value);
            cmd.Parameters.AddWithValue("@hour", job.RunHour ?? (object)DBNull.Value);
            cmd.Parameters.AddWithValue("@minute", job.RunMinute ?? (object)DBNull.Value);
            cmd.Parameters.AddWithValue("@active", job.IsActive ?? (object)DBNull.Value);
            await conn.OpenAsync();
            return await cmd.ExecuteNonQueryAsync();
        }

        // READ
        public async Task<List<MscheduledJob>> GetAllJobsAsync()
        {
            var list = new List<MscheduledJob>();
            using var conn = _connectionDAO.GetConnection();
            using var cmd = new SqlCommand(
                "SELECT Id, JobName, RunHour, RunMinute, IsActive FROM MscheduledJob", conn);
            await conn.OpenAsync();
            using var reader = await cmd.ExecuteReaderAsync();
            while (await reader.ReadAsync())
            {
                list.Add(new MscheduledJob
                {
                    Id = reader.GetInt32(0),
                    JobName = reader.IsDBNull(1) ? null : reader.GetString(1),
                    RunHour = reader.IsDBNull(2) ? null : reader.GetInt16(2),
                    RunMinute = reader.IsDBNull(3) ? null : reader.GetInt16(3),
                    IsActive = reader.IsDBNull(4) ? null : reader.GetBoolean(4)
                });
            }
            return list;
        }

        // UPDATE
        public async Task<int> UpdateJobAsync(MscheduledJob job)
        {
            using var conn = _connectionDAO.GetConnection();
            using var cmd = new SqlCommand(
                @"UPDATE MscheduledJob 
                  SET JobName=@name, RunHour=@hour, RunMinute=@minute, IsActive=@active 
                  WHERE Id=@id", conn);
            cmd.Parameters.AddWithValue("@name", job.JobName ?? (object)DBNull.Value);
            cmd.Parameters.AddWithValue("@hour", job.RunHour ?? (object)DBNull.Value);
            cmd.Parameters.AddWithValue("@minute", job.RunMinute ?? (object)DBNull.Value);
            cmd.Parameters.AddWithValue("@active", job.IsActive ?? (object)DBNull.Value);
            cmd.Parameters.AddWithValue("@id", job.Id);
            await conn.OpenAsync();
            return await cmd.ExecuteNonQueryAsync();
        }

        // DELETE
        public async Task<int> DeleteJobAsync(int id)
        {
            using var conn = _connectionDAO.GetConnection();
            using var cmd = new SqlCommand("DELETE FROM MscheduledJob WHERE Id=@id", conn);
            cmd.Parameters.AddWithValue("@id", id);
            await conn.OpenAsync();
            return await cmd.ExecuteNonQueryAsync();
        }
    }
}