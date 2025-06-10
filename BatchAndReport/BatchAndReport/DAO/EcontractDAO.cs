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
    }
}