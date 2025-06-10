using BatchAndReport.Entities;
using BatchAndReport.Models;
using Microsoft.Data.SqlClient;
using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.EntityFrameworkCore;

namespace BatchAndReport.DAO
{
    public class HrDAO
    {
        private readonly SqlConnectionDAO _connectionDAO;
        private readonly K2DBContext _k2context;

        public HrDAO(SqlConnectionDAO connectionDAO, K2DBContext k2context)
        {
            _connectionDAO = connectionDAO;
            _k2context = k2context;
        }

        // CREATE OR UPDATE
        public async Task InsertOrUpdateEmployeesAsync(List<MEmployeeModels> employees)
        {
            foreach (var emp in employees)
            {
                var existingEmp = await _k2context.Employees
                    .FirstOrDefaultAsync(e => e.EmployeeId == emp.EmployeeId);

                if (existingEmp != null)
                {
                    // UPDATE
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

                    _k2context.Employees.Update(existingEmp);
                }
                else
                {
                    // INSERT
                    var newEmp = new Employee
                    {
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
                        PositionId = emp.PositionId
                    };

                    await _k2context.Employees.AddAsync(newEmp);
                }
            }

            await _k2context.SaveChangesAsync();
        }

        public async Task InsertOrUpdateEmployeesMovementAsync(List<MEmployeeMovementModels> employees)
        {
            foreach (var emp in employees)
            {
                var existingEmpMovement = await _k2context.EmployeeMovements
                    .FirstOrDefaultAsync(e => e.EmployeeId == emp.EmployeeId && e.PositionId == emp.PositionId);

                if (existingEmpMovement != null)
                {
                    // UPDATE
                    existingEmpMovement.EmployeeId = emp.EmployeeId;
                    existingEmpMovement.EffectiveDate = emp.EffectiveDate;
                    existingEmpMovement.MovementTypeId = emp.MovementTypeId;
                    existingEmpMovement.MovementReasonId = emp.MovementReasonId;
                    existingEmpMovement.EmployeeCode = emp.EmployeeCode;
                    existingEmpMovement.Employment = emp.Employment;
                    existingEmpMovement.EmployeeStatus = emp.EmployeeStatus;
                    existingEmpMovement.EmployeeTypeId = emp.EmployeeTypeId;
                    existingEmpMovement.PayrollGroupId = emp.PayrollGroupId;
                    existingEmpMovement.CompanyId = emp.CompanyId;
                    existingEmpMovement.BusinessUnitId = emp.BusinessUnitId;
                    existingEmpMovement.PositionId = emp.PositionId;
                    existingEmpMovement.WorkLocationId = emp.WorkLocationId;
                    existingEmpMovement.CalendarGroupId = emp.CalendarGroupId;
                    existingEmpMovement.JobTitleId = emp.JobTitleId;
                    existingEmpMovement.JobLevelId = emp.JobLevelId;
                    existingEmpMovement.JobGradeId = emp.JobGradeId;
                    existingEmpMovement.ContractStartDate = emp.ContractStartDate;
                    existingEmpMovement.ContractEndDate = emp.ContractEndDate;
                    existingEmpMovement.RenewContractCount = emp.RenewContractCount;
                    existingEmpMovement.ProbationDate = emp.ProbationDate;
                    existingEmpMovement.ProbationDuration = emp.ProbationDuration;
                    existingEmpMovement.ProbationResult = emp.ProbationResult;
                    existingEmpMovement.ProbationExtend = emp.ProbationExtend;
                    existingEmpMovement.EmploymentDate = emp.EmploymentDate;
                    existingEmpMovement.JoinDate = emp.JoinDate;
                    existingEmpMovement.OccupationDate = emp.OccupationDate;
                    existingEmpMovement.TerminationDate = emp.TerminationDate;
                    existingEmpMovement.TerminationReason = emp.TerminationReason;
                    existingEmpMovement.TerminationSSO = emp.TerminationSSO;
                    existingEmpMovement.IsBlacklist = emp.IsBlacklist;
                    existingEmpMovement.PaymentDate = emp.PaymentDate;
                    existingEmpMovement.Remark = emp.Remark;
                    existingEmpMovement.ServiceYearAdjust = emp.ServiceYearAdjust;
                    existingEmpMovement.SupervisorCode = emp.SupervisorCode;
                    existingEmpMovement.StandardWorkHoursID = emp.StandardWorkHoursID;
                    existingEmpMovement.WorkOperation = emp.WorkOperation;

                    _k2context.EmployeeMovements.Update(existingEmpMovement);
                }
                else
                {
                    // INSERT
                    var newEmp = new EmployeeMovement
                    {
                        Id = emp.Id,
                        EmployeeId = emp.EmployeeId,
                        EffectiveDate = emp.EffectiveDate,
                        MovementTypeId = emp.MovementTypeId,
                        MovementReasonId = emp.MovementReasonId,
                        EmployeeCode = emp.EmployeeCode,
                        Employment = emp.Employment,
                        EmployeeStatus = emp.EmployeeStatus,
                        EmployeeTypeId = emp.EmployeeTypeId,
                        PayrollGroupId = emp.PayrollGroupId,
                        CompanyId = emp.CompanyId,
                        BusinessUnitId = emp.BusinessUnitId,
                        PositionId = emp.PositionId,
                        WorkLocationId = emp.WorkLocationId,
                        CalendarGroupId = emp.CalendarGroupId,
                        JobTitleId = emp.JobTitleId,
                        JobLevelId = emp.JobLevelId,
                        JobGradeId = emp.JobGradeId,
                        ContractStartDate = emp.ContractStartDate,
                        ContractEndDate = emp.ContractEndDate,
                        RenewContractCount = emp.RenewContractCount,
                        ProbationDate = emp.ProbationDate,
                        ProbationDuration = emp.ProbationDuration,
                        ProbationResult = emp.ProbationResult,
                        ProbationExtend = emp.ProbationExtend,
                        EmploymentDate = emp.EmploymentDate,
                        JoinDate = emp.JoinDate,
                        OccupationDate = emp.OccupationDate,
                        TerminationDate = emp.TerminationDate,
                        TerminationReason = emp.TerminationReason,
                        TerminationSSO = emp.TerminationSSO,
                        IsBlacklist = emp.IsBlacklist,
                        PaymentDate = emp.PaymentDate,
                        Remark = emp.Remark,
                        ServiceYearAdjust = emp.ServiceYearAdjust,
                        SupervisorCode = emp.SupervisorCode,
                        StandardWorkHoursID = emp.StandardWorkHoursID,
                        WorkOperation = emp.WorkOperation
                    };

                    await _k2context.EmployeeMovements.AddAsync(newEmp);
                }
            }

            await _k2context.SaveChangesAsync();
        }

        // CREATE OR UPDATE
        public async Task InsertOrUpdatePositionAsync(List<MPositionModels> positions)
        {

            foreach (var pos in positions)
            {
                var existingPos = await _k2context.Positions
                    .FirstOrDefaultAsync(e => e.PositionId == pos.PositionId);

                if (existingPos != null)
                {
                    // UPDATE
                    existingPos.ProjectCode = pos.ProjectCode;
                    existingPos.PositionId = pos.PositionId;
                    existingPos.TypeCode = pos.TypeCode;
                    existingPos.Module = pos.Module;
                    existingPos.NameTh = pos.NameTh;
                    existingPos.NameEn = pos.NameEn;
                    existingPos.DescriptionTh = pos.DescriptionTh;
                    existingPos.DescriptionEn = pos.DescriptionEn;

                    _k2context.Positions.Update(existingPos);
                }
                else
                {
                    // INSERT
                    var newPos = new Position
                    {
                        ProjectCode = pos.ProjectCode,
                        PositionId = pos.PositionId,
                        TypeCode = pos.TypeCode,
                        Module = pos.Module,
                        NameTh = pos.NameTh,
                        NameEn = pos.NameEn,
                        DescriptionTh = pos.DescriptionTh,
                        DescriptionEn = pos.DescriptionEn,
                    };

                    await _k2context.Positions.AddAsync(newPos);
                }
            }

            await _k2context.SaveChangesAsync();
        }

        // CREATE OR UPDATE
        public async Task InsertOrUpdateBusinessUnitAsync(List<MBusinessUnitModels> businessUnits)
        {

            foreach (var bus in businessUnits)
            {
                var existingBus = await _k2context.BusinessUnits
                    .FirstOrDefaultAsync(e => e.BusinessUnitId == bus.BusinessUnitId);

                if (existingBus != null)
                {
                    // UPDATE
                    existingBus.BusinessUnitId = bus.BusinessUnitId;
                    existingBus.BusinessUnitCode = bus.BusinessUnitCode;
                    existingBus.BusinessUnitLevel = bus.BusinessUnitLevel;
                    existingBus.ParentId = bus.ParentId;
                    existingBus.CompanyId = bus.CompanyId;
                    existingBus.EffectiveDate = bus.EffectiveDate;
                    existingBus.NameTh = bus.NameTh;
                    existingBus.NameEn = bus.NameEn;
                    existingBus.AbbreviationTh = bus.AbbreviationTh;
                    existingBus.AbbreviationEn = bus.AbbreviationEn;
                    existingBus.DescriptionTh = bus.DescriptionTh;
                    existingBus.DescriptionEn = bus.DescriptionEn;
                    existingBus.CreateDate = bus.CreateDate;

                    _k2context.BusinessUnits.Update(existingBus);
                }
                else
                {
                    // INSERT
                    var newBus = new BusinessUnit
                    {
                        BusinessUnitId = bus.BusinessUnitId,
                        BusinessUnitCode = bus.BusinessUnitCode,
                        BusinessUnitLevel = bus.BusinessUnitLevel,
                        ParentId = bus.ParentId,
                        CompanyId = bus.CompanyId,
                        EffectiveDate = bus.EffectiveDate,
                        NameTh = bus.NameTh,
                        NameEn = bus.NameEn,
                        AbbreviationTh = bus.AbbreviationTh,
                        AbbreviationEn = bus.AbbreviationEn,
                        DescriptionTh = bus.DescriptionTh,
                        DescriptionEn = bus.DescriptionEn,
                        CreateDate = bus.CreateDate
                    };

                    await _k2context.BusinessUnits.AddAsync(newBus);
                }
            }

            await _k2context.SaveChangesAsync();
        }

    }
}