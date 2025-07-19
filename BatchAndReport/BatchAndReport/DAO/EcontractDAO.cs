using BatchAndReport.Entities;
using BatchAndReport.Models;
using Microsoft.EntityFrameworkCore;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace BatchAndReport.DAO
{
    public class EContractDAO // Fixed spelling error: Changed "EcontractDAO" to "EContractDAO"  
    {
        private readonly K2DBContext_EContract _context;

        public EContractDAO(K2DBContext_EContract context) // Fixed spelling error: Changed "EcontractDAO" to "EContractDAO"  
        {
            _context = context;
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
                        PassportNo = emp.PassportNo
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

    }
}