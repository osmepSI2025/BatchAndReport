using System;
using System.Collections.Generic;

namespace BatchAndReport.Entities
{
    public class Employee
    {
        public int Id { get; set; }

        public string? EmployeeId { get; set; }

        public string? EmployeeCode { get; set; }

        public string? NameTh { get; set; }

        public string? NameEn { get; set; }

        public string? FirstNameTh { get; set; }

        public string? FirstNameEn { get; set; }

        public string? LastNameTh { get; set; }

        public string? LastNameEn { get; set; }

        public string? Email { get; set; }

        public string? Mobile { get; set; }

        public DateTime? EmploymentDate { get; set; }

        public DateTime? TerminationDate { get; set; }

        public string? EmployeeType { get; set; }

        public string? EmployeeStatus { get; set; }

        public string? SupervisorId { get; set; }

        public string? CompanyId { get; set; }

        public string? BusinessUnitId { get; set; }

        public string? PositionId { get; set; }
    }

   

}
