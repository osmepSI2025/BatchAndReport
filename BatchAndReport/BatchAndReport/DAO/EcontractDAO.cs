using BatchAndReport.Entities;
using BatchAndReport.Models;
using Microsoft.Data.SqlClient;
using Microsoft.EntityFrameworkCore;
using System.Collections.Generic;
using System.Data;
using System.Text;
using System.Threading.Tasks;

namespace BatchAndReport.DAO
{
    public class EContractDAO // Fixed spelling error: Changed "EcontractDAO" to "EContractDAO"  
    {
        private readonly K2DBContext_EContract _context;
        private readonly SqlConnectionDAO _connectionDAO;
        private readonly string _fallbackPassword;
        public EContractDAO(K2DBContext_EContract context, SqlConnectionDAO connectionDAO, IConfiguration configuration) // Fixed spelling error: Changed "EcontractDAO" to "EContractDAO"  
        {
            _context = context;
            _connectionDAO = connectionDAO;
            _fallbackPassword = configuration["Password:PaswordPDF"] ?? string.Empty;
        }

        public async Task InsertOrUpdateEmployeeContractsAsync(List<MEmployeeContractModels> contracts)
        {
            foreach (var emp in contracts)
            {
                var existingEmp = await _context.EmployeeContracts
                    .FirstOrDefaultAsync(e => e.EmployeeId == emp.EmployeeId);

                if (existingEmp != null)
                {
                    // UPDATE  
                    existingEmp.ContractFlag = emp.ContractFlag;
                    existingEmp.EmployeeCode = emp.EmployeeCode;
                    existingEmp.NameTh = emp.NameTh;
                    existingEmp.NameEn = emp.NameEn;
                    existingEmp.FirstNameTh = emp.FirstNameTh;
                    existingEmp.FirstNameEn = emp.FirstNameEn;
                    existingEmp.LastNameTh = emp.LastNameTh;
                    existingEmp.LastNameEn = emp.LastNameEn;
                    existingEmp.Email = emp.Email;
                    existingEmp.Mobile = emp.Mobile;
                    existingEmp.EmploymentDate = emp.EmploymentDate;
                    existingEmp.TerminationDate = emp.TerminationDate;
                    existingEmp.EmployeeType = emp.EmployeeType;
                    existingEmp.EmployeeStatus = emp.EmployeeStatus;
                    existingEmp.SupervisorId = emp.SupervisorId;
                    existingEmp.CompanyId = emp.CompanyId;
                    existingEmp.BusinessUnitId = emp.BusinessUnitId;
                    existingEmp.PositionId = emp.PositionId;
                    existingEmp.Salary = emp.Salary;
                    existingEmp.IdCard = emp.IdCard;
                    existingEmp.PassportNo = emp.PassportNo;
                    existingEmp.Address = emp.Address;

                    _context.EmployeeContracts.Update(existingEmp);
                }
                else
                {
                    // INSERT  
                    var newEmp = new EmployeeContract
                    {
                        ContractFlag = emp.ContractFlag,
                        EmployeeId = emp.EmployeeId,
                        EmployeeCode = emp.EmployeeCode,
                        NameTh = emp.NameTh,
                        NameEn = emp.NameEn,
                        FirstNameTh = emp.FirstNameTh,
                        FirstNameEn = emp.FirstNameEn,
                        LastNameTh = emp.LastNameTh,
                        LastNameEn = emp.LastNameEn,
                        Email = emp.Email,
                        Mobile = emp.Mobile,
                        EmploymentDate = emp.EmploymentDate,
                        TerminationDate = emp.TerminationDate,
                        EmployeeType = emp.EmployeeType,
                        EmployeeStatus = emp.EmployeeStatus,
                        SupervisorId = emp.SupervisorId,
                        CompanyId = emp.CompanyId,
                        BusinessUnitId = emp.BusinessUnitId,
                        PositionId = emp.PositionId,
                        Salary = emp.Salary,
                        IdCard = emp.IdCard,
                        PassportNo = emp.PassportNo,
                        Address = emp.Address
                    };

                    await _context.EmployeeContracts.AddAsync(newEmp);
                }
            }

            await _context.SaveChangesAsync();
        }
        public async Task InsertOrUpdatePartyContractsAsync(List<MContractPartyModels> parties)
        {
            foreach (var party in parties)
            {
                var existingParty = await _context.ContractParties
                    .FirstOrDefaultAsync(p => p.RegIden == party.RegIden);

                if (existingParty != null)
                {
                    // UPDATE
                    existingParty.ContractPartyName = party.ContractPartyName;
                    existingParty.RegType = party.RegType;
                    existingParty.RegDetail = party.RegDetail;
                    existingParty.AddressNo = party.AddressNo;
                    existingParty.SubDistrict = party.SubDistrict;
                    existingParty.District = party.District;
                    existingParty.Province = party.Province;
                    existingParty.PostalCode = party.PostalCode;
                    existingParty.FlagActive = party.FlagActive?.StartsWith("ยัง") == true ? "Y" : "N";

                    _context.ContractParties.Update(existingParty);
                }
                else
                {
                    // INSERT
                    var newParty = new ContractParty
                    {
                        ContractPartyName = party.ContractPartyName,
                        RegType = party.RegType,
                        RegIden = party.RegIden,
                        RegDetail = party.RegDetail,
                        AddressNo = party.AddressNo,
                        SubDistrict = party.SubDistrict,
                        District = party.District,
                        Province = party.Province,
                        PostalCode = party.PostalCode,
                        FlagActive = party.FlagActive?.StartsWith("ยัง") == true ? "Y" : "N"
                    };

                    await _context.ContractParties.AddAsync(newParty);
                }
            }

            await _context.SaveChangesAsync();
        }
        public async Task<List<MContractPartyModels>> SyncAllContractPartiesAsync(List<MContractPartyModels> externalParties)
        {
            var resultList = new List<MContractPartyModels>();

            foreach (var party in externalParties)
            {
                var existing = await _context.ContractParties
                    .FirstOrDefaultAsync(p => p.RegIden == party.RegIden);

                if (existing != null)
                {
                    // UPDATE if any field changed
                    existing.ContractPartyName = party.ContractPartyName;
                    existing.RegType = party.RegType;
                    existing.RegDetail = party.RegDetail;
                    existing.AddressNo = party.AddressNo;
                    existing.SubDistrict = party.SubDistrict;
                    existing.District = party.District;
                    existing.Province = party.Province;
                    existing.PostalCode = party.PostalCode;
                    existing.FlagActive = party.FlagActive?.StartsWith("ยัง") == true ? "Y" : "N";

                    _context.ContractParties.Update(existing);
                    resultList.Add(party); // track updated
                }
                else
                {
                    // INSERT
                    var newParty = new ContractParty
                    {
                        ContractPartyName = party.ContractPartyName,
                        RegType = party.RegType,
                        RegIden = party.RegIden,
                        RegDetail = party.RegDetail,
                        AddressNo = party.AddressNo,
                        SubDistrict = party.SubDistrict,
                        District = party.District,
                        Province = party.Province,
                        PostalCode = party.PostalCode,
                        FlagActive = party.FlagActive?.StartsWith("ยัง") == true ? "Y" : "N"
                    };

                    await _context.ContractParties.AddAsync(newParty);
                    resultList.Add(party); // track inserted
                }
            }

            await _context.SaveChangesAsync();
            return resultList;
        }
        public async Task<List<E_ConReport_RelatedDocumentsModels>> GetRelatedDocumentsAsync(string? id = "0", string TypeContract = "")
        {
            var result = new List<E_ConReport_RelatedDocumentsModels>();
            try
            {
                await using var connection = _connectionDAO.GetConnectionK2Econctract();
                await using var command = new SqlCommand(@"
            SELECT 
                Document_ID,
                Contract_ID,
                Contract_Type,
                DocumentTitle,
                Required_Flag,
                FilePath,
                PageAmount,
                Flag_Delete,
                File_Name,
                File_Location
            FROM RelatedDocuments
            WHERE Contract_ID = @Contract_ID AND Contract_Type = @Contract_Type", connection);

                command.Parameters.AddWithValue("@Contract_ID", id ?? "0");
                command.Parameters.AddWithValue("@Contract_Type", TypeContract ?? "0");
                await connection.OpenAsync();

                using var reader = await command.ExecuteReaderAsync();
                while (await reader.ReadAsync())
                {
                    result.Add(new E_ConReport_RelatedDocumentsModels
                    {
                        Document_ID = reader.IsDBNull(0) ? 0 : reader.GetInt32(0),
                        Contract_ID = reader.IsDBNull(1) ? 0 : reader.GetInt32(1),
                        Contract_Type = reader.IsDBNull(2) ? null : reader.GetString(2),
                        DocumentTitle = reader.IsDBNull(3) ? null : reader.GetString(3),
                        Required_Flag = reader.IsDBNull(4) ? null : reader.GetString(4),
                        FilePath = reader.IsDBNull(5) ? null : reader.GetString(5),
                        PageAmount = reader.IsDBNull(6) ? 0 : reader.GetInt32(6),
                        Flag_Delete = reader.IsDBNull(7) ? null : reader.GetString(7),
                        File_Name = reader.IsDBNull(8) ? null : reader.GetString(8),
                        File_Location = reader.IsDBNull(9) ? null : reader.GetString(9)
                    });
                }
            }
            catch (Exception ex)
            {
                // Optionally log the exception here
            }
            return result;
        }

        public async Task<string> GetProjectByProjectCodeAsync(string projectCode)
        {
            var dbConn = _context.Database.GetDbConnection();

            await using var cmd = dbConn.CreateCommand();
            cmd.CommandText = "dbo.SP_Get_All_Contract_API";
            cmd.CommandType = CommandType.StoredProcedure;

            var p = cmd.CreateParameter();
            p.ParameterName = "@ProjectCode";
            p.DbType = DbType.String;
            p.Size = 10;
            p.Value = string.IsNullOrWhiteSpace(projectCode) ? (object)DBNull.Value : projectCode;
            cmd.Parameters.Add(p);

            // เพิ่ม timeout เผื่อผลลัพธ์ใหญ่
            cmd.CommandTimeout = 300;

            var shouldClose = dbConn.State != ConnectionState.Open;
            if (shouldClose) await dbConn.OpenAsync();

            try
            {
                var sb = new StringBuilder(1024 * 64);

                await using var reader = await cmd.ExecuteReaderAsync(CommandBehavior.SequentialAccess);
                while (await reader.ReadAsync())
                {
                    // สมมติว่า SP คืนคอลัมน์เดียวเป็นชิ้นส่วนของ JSON
                    // (ถ้ามากกว่าหนึ่งคอลัมน์ให้เปลี่ยน index ตามจริง)
                    if (!reader.IsDBNull(0))
                        sb.Append(reader.GetString(0));
                }

                var json = sb.ToString();

                if (string.IsNullOrWhiteSpace(json) ||
                    !(json.TrimStart().StartsWith("{") || json.TrimStart().StartsWith("[")))
                {
                    json = "{\"responseCode\":\"500\",\"responseMsg\":\"No or invalid JSON returned from SP_SME_PROJECT_API_BY_YEAR\",\"data\":[]}";
                }

                return json;
            }
            finally
            {
                if (shouldClose) await dbConn.CloseAsync();
            }
        }
        public async Task<string?> GetPdfPasswordByEmpIdAsync(string? empId, CancellationToken ct = default)
        {

            var pwd = await _context.ContractFilePasswords
                .Where(x => x.EmpId == empId)
                .Select(x => x.Password)
                .FirstOrDefaultAsync(ct);

            // ไม่พบใน DB -> ใช้ fallback จาก config
            return string.IsNullOrWhiteSpace(pwd) ? _fallbackPassword : pwd;
        }
        public async Task<List<ImportContractModels>> FindImportContractsAsync(
    string? contractNumber, CancellationToken ct = default)
        {
            var list = new List<ImportContractModels>();

            const string sql = @"
SELECT
    ic.[ContractNumber],
    ic.[ProjectName],
    ic.[ContractParty],
    ic.[Domicile],
    ic.[Start_Date],
    ic.[End_Date],
    ic.[Status],
    ic.[Amount],
    ic.[Installment],
    ic.[Owner],
    ic.[Contract_Storage],
    t.[WFTypeNameTH] AS [ContractType],
    ic.[InstallmentNo],
    ic.[PaymentDate],
    ic.[InstallmentAmount]
FROM [E-Contract].[dbo].[Import_Contract] AS ic
LEFT JOIN [E-Contract].[dbo].[EContract_WF_Type] AS t
    ON ic.[ContractType] = t.[WFTypeCode]
WHERE (@ContractNumber IS NULL OR ic.[ContractNumber] LIKE '%' + @ContractNumber + '%')
ORDER BY ic.[ContractNumber];";

            await using var conn = _connectionDAO.GetConnectionK2Econctract(); // DB E-Contract
            await conn.OpenAsync(ct);

            await using var cmd = new SqlCommand(sql, conn)
            {
                CommandType = CommandType.Text
            };

            cmd.Parameters.Add("@ContractNumber", SqlDbType.NVarChar, 50).Value =
                string.IsNullOrWhiteSpace(contractNumber) ? (object)DBNull.Value : contractNumber.Trim();

            using var rd = await cmd.ExecuteReaderAsync(ct);

            // เตรียม ordinal ล่วงหน้าเพื่อความเร็ว/กันพลาดพิมพ์ชื่อคอลัมน์
            int oContractNumber = rd.GetOrdinal("ContractNumber");
            int oProjectName = rd.GetOrdinal("ProjectName");
            int oContractParty = rd.GetOrdinal("ContractParty");
            int oDomicile = rd.GetOrdinal("Domicile");
            int oStartDate = rd.GetOrdinal("Start_Date");
            int oEndDate = rd.GetOrdinal("End_Date");
            int oStatus = rd.GetOrdinal("Status");
            int oAmount = rd.GetOrdinal("Amount");
            int oInstallment = rd.GetOrdinal("Installment");
            int oOwner = rd.GetOrdinal("Owner");
            int oContractStorage = rd.GetOrdinal("Contract_Storage");
            int oContractType = rd.GetOrdinal("ContractType");
            int oInstallmentNo = rd.GetOrdinal("InstallmentNo");
            int oPaymentDate = rd.GetOrdinal("PaymentDate");
            int oInstallmentAmt = rd.GetOrdinal("InstallmentAmount");

            while (await rd.ReadAsync(ct))
            {
                list.Add(new ImportContractModels
                {
                    ContractNumber = rd.IsDBNull(oContractNumber) ? null : rd.GetString(oContractNumber),
                    ProjectName = rd.IsDBNull(oProjectName) ? null : rd.GetString(oProjectName),
                    ContractParty = rd.IsDBNull(oContractParty) ? null : rd.GetString(oContractParty),
                    Domicile = rd.IsDBNull(oDomicile) ? null : rd.GetString(oDomicile),
                    StartDate = rd.IsDBNull(oStartDate) ? (DateTime?)null : rd.GetDateTime(oStartDate),
                    EndDate = rd.IsDBNull(oEndDate) ? (DateTime?)null : rd.GetDateTime(oEndDate),
                    Status = rd.IsDBNull(oStatus) ? null : rd.GetString(oStatus),
                    Amount = rd.IsDBNull(oAmount) ? (decimal?)null : Convert.ToDecimal(rd[oAmount]),
                    Installment = rd.IsDBNull(oInstallment) ? (int?)null : Convert.ToInt32(rd[oInstallment]),
                    Owner = rd.IsDBNull(oOwner) ? null : rd.GetString(oOwner),
                    ContractStorage = rd.IsDBNull(oContractStorage) ? null : rd.GetString(oContractStorage),
                    ContractType = rd.IsDBNull(oContractType) ? null : rd.GetString(oContractType),
                    InstallmentNo = rd.IsDBNull(oInstallmentNo) ? (int?)null : Convert.ToInt32(rd[oInstallmentNo]),
                    PaymentDate = rd.IsDBNull(oPaymentDate) ? (DateTime?)null : rd.GetDateTime(oPaymentDate),
                    InstallmentAmount = rd.IsDBNull(oInstallmentAmt) ? (decimal?)null : Convert.ToDecimal(rd[oInstallmentAmt])
                });
            }

            return list;
        }

    }
}