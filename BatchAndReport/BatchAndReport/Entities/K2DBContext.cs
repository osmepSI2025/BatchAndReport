using BatchAndReport.Models;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
namespace BatchAndReport.Entities;


public partial class K2DBContext : DbContext
{
    public K2DBContext()
    {
    }

    public K2DBContext(DbContextOptions<K2DBContext> options)
        : base(options)
    {
    }


    public virtual DbSet<Employee> Employees { get; set; }
    public virtual DbSet<EmployeeMovement> EmployeeMovements { get; set; }
    public virtual DbSet<Position> Positions { get; set; }
    public virtual DbSet<BusinessUnit> BusinessUnits { get; set; }

    protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
#warning To protect potentially sensitive information in your connection string, you should move it out of source code. You can avoid scaffolding the connection string by using the Name= syntax to read it from configuration - see https://go.microsoft.com/fwlink/?linkid=2131148. For more guidance on storing connection strings, see http://go.microsoft.com/fwlink/?LinkId=723263.
        => optionsBuilder.UseSqlServer("Server=192.168.9.156;Database=HR;User Id=sa;Password=Osmep@2025;TrustServerCertificate=True;");

    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {

        modelBuilder.Entity<Employee>(entity =>
{
            entity.ToTable("Employee");

            entity.HasKey(e => e.Id);

            entity.Property(e => e.Id)
                .HasColumnName("Id");

            entity.Property(e => e.EmployeeId)
                .HasColumnName("EmployeeId")
                .HasMaxLength(50)
                .IsRequired()
                .UseCollation("Thai_CI_AS");

            entity.Property(e => e.EmployeeCode)
                .HasColumnName("EmployeeCode")
                .HasMaxLength(50)
                .UseCollation("Thai_CI_AS");

            entity.Property(e => e.NameTh)
                .HasColumnName("NameTh")
                .HasMaxLength(100)
                .UseCollation("Thai_CI_AS");

            entity.Property(e => e.NameEn)
                .HasColumnName("NameEn")
                .HasMaxLength(100)
                .UseCollation("Thai_CI_AS");

            entity.Property(e => e.FirstNameTh)
                .HasColumnName("FirstNameTh")
                .HasMaxLength(100)
                .UseCollation("Thai_CI_AS");

            entity.Property(e => e.FirstNameEn)
                .HasColumnName("FirstNameEn")
                .HasMaxLength(100)
                .UseCollation("Thai_CI_AS");

            entity.Property(e => e.LastNameTh)
                .HasColumnName("LastNameTh")
                .HasMaxLength(100)
                .UseCollation("Thai_CI_AS");

            entity.Property(e => e.LastNameEn)
                .HasColumnName("LastNameEn")
                .HasMaxLength(100)
                .UseCollation("Thai_CI_AS");

            entity.Property(e => e.Email)
                .HasColumnName("Email")
                .HasMaxLength(100)
                .UseCollation("Thai_CI_AS");

            entity.Property(e => e.Mobile)
                .HasColumnName("Mobile")
                .HasMaxLength(20)
                .UseCollation("Thai_CI_AS");

            entity.Property(e => e.EmploymentDate)
                .HasColumnName("EmploymentDate");

            entity.Property(e => e.TerminationDate)
                .HasColumnName("TerminationDate");

            entity.Property(e => e.EmployeeType)
                .HasColumnName("EmployeeType")
                .HasMaxLength(50)
                .UseCollation("Thai_CI_AS");

            entity.Property(e => e.EmployeeStatus)
                .HasColumnName("EmployeeStatus")
                .HasMaxLength(50)
                .UseCollation("Thai_CI_AS");

            entity.Property(e => e.SupervisorId)
                .HasColumnName("SupervisorId")
                .HasMaxLength(50)
                .UseCollation("Thai_CI_AS");

            entity.Property(e => e.CompanyId)
                .HasColumnName("CompanyId")
                .HasMaxLength(50)
                .UseCollation("Thai_CI_AS");

            entity.Property(e => e.BusinessUnitId)
                .HasColumnName("BusinessUnitId")
                .HasMaxLength(50)
                .UseCollation("Thai_CI_AS");

            entity.Property(e => e.PositionId)
                .HasColumnName("PositionId")
                .HasMaxLength(50)
                .UseCollation("Thai_CI_AS");

            // Fix for CS0305: Correctly specify the generic type argument for HasOne
            entity.HasOne(e => e.Position)
                .WithMany()
                .HasForeignKey(e => e.PositionId)
                .HasPrincipalKey(p => p.PositionId);
        });
        modelBuilder.Entity<EmployeeMovement>(entity =>
        {
            entity.ToTable("EmployeeMovements");

            entity.HasKey(e => e.Id);

            entity.Property(e => e.Id)
                .HasColumnName("Id")
                .ValueGeneratedOnAdd();

            entity.Property(e => e.EmployeeId)
                .HasMaxLength(50)
                .IsUnicode(false);

            entity.Property(e => e.EffectiveDate);

            entity.Property(e => e.MovementTypeId)
                .HasMaxLength(50)
                .IsUnicode(false);

            entity.Property(e => e.MovementReasonId)
                .HasMaxLength(50)
                .IsUnicode(false);

            entity.Property(e => e.EmployeeCode)
                .HasMaxLength(50)
                .IsUnicode(false);

            entity.Property(e => e.Employment)
                .HasMaxLength(50)
                .IsUnicode(false);

            entity.Property(e => e.EmployeeStatus)
                .HasMaxLength(1)
                .IsUnicode(false);

            entity.Property(e => e.EmployeeTypeId)
                .HasMaxLength(50)
                .IsUnicode(false);

            entity.Property(e => e.PayrollGroupId)
                .HasMaxLength(50)
                .IsUnicode(false);

            entity.Property(e => e.CompanyId)
                .HasMaxLength(50)
                .IsUnicode(false);

            entity.Property(e => e.BusinessUnitId)
                .HasMaxLength(50)
                .IsUnicode(false);

            entity.Property(e => e.PositionId)
                .HasMaxLength(50)
                .IsUnicode(false);

            entity.Property(e => e.WorkLocationId)
                .HasMaxLength(50)
                .IsUnicode(false);

            entity.Property(e => e.CalendarGroupId)
                .HasMaxLength(50)
                .IsUnicode(false);

            entity.Property(e => e.JobTitleId)
                .HasMaxLength(50)
                .IsUnicode(false);

            entity.Property(e => e.JobLevelId)
                .HasMaxLength(50)
                .IsUnicode(false);

            entity.Property(e => e.JobGradeId)
                .HasMaxLength(50)
                .IsUnicode(false);

            entity.Property(e => e.ContractStartDate);

            entity.Property(e => e.ContractEndDate);

            entity.Property(e => e.RenewContractCount);

            entity.Property(e => e.ProbationDate);

            entity.Property(e => e.ProbationDuration);

            entity.Property(e => e.ProbationResult)
                .HasMaxLength(50)
                .IsUnicode(false);

            entity.Property(e => e.ProbationExtend)
                .HasMaxLength(50)
                .IsUnicode(false);

            entity.Property(e => e.EmploymentDate);

            entity.Property(e => e.JoinDate);

            entity.Property(e => e.OccupationDate);

            entity.Property(e => e.TerminationDate);

            entity.Property(e => e.TerminationReason)
                .HasMaxLength(100)
                .IsUnicode(false);

            entity.Property(e => e.TerminationSSO)
                .HasMaxLength(50)
                .IsUnicode(false);

            entity.Property(e => e.IsBlacklist)
                .HasMaxLength(10)
                .IsUnicode(false);

            entity.Property(e => e.PaymentDate);

            entity.Property(e => e.Remark)
                .HasMaxLength(255)
                .IsUnicode(false);

            entity.Property(e => e.ServiceYearAdjust);

            entity.Property(e => e.SupervisorCode)
                .HasMaxLength(50)
                .IsUnicode(false);

            entity.Property(e => e.StandardWorkHoursID)
                .HasMaxLength(50)
                .IsUnicode(false);

            entity.Property(e => e.WorkOperation)
                .HasMaxLength(50)
                .IsUnicode(false);
        });
        modelBuilder.Entity<Position>(entity =>
        {
            entity.ToTable("Position");

            entity.HasKey(e => e.Id);

            entity.Property(e => e.Id)
                .HasColumnName("Id");

            entity.Property(e => e.ProjectCode)
                .HasColumnName("ProjectCode")
                .HasMaxLength(50)
                .UseCollation("Thai_CI_AS");

            entity.Property(e => e.PositionId)
                .HasColumnName("PositionId")
                .HasMaxLength(50)
                .UseCollation("Thai_CI_AS");

            entity.Property(e => e.TypeCode)
                .HasColumnName("TypeCode")
                .HasMaxLength(50)
                .UseCollation("Thai_CI_AS");

            entity.Property(e => e.Module)
                .HasColumnName("Module")
                .HasMaxLength(50)
                .UseCollation("Thai_CI_AS");

            entity.Property(e => e.NameTh)
                .HasColumnName("NameTh")
                .HasMaxLength(255)
                .UseCollation("Thai_CI_AS");

            entity.Property(e => e.NameEn)
                .HasColumnName("NameEn")
                .HasMaxLength(255)
                .UseCollation("Thai_CI_AS");

            entity.Property(e => e.DescriptionTh)
                .HasColumnName("DescriptionTh")
                .HasColumnType("text")
                .UseCollation("Thai_CI_AS");

            entity.Property(e => e.DescriptionEn)
                .HasColumnName("DescriptionEn")
                .HasColumnType("text")
                .UseCollation("Thai_CI_AS");
        });
        modelBuilder.Entity<BusinessUnit>(entity =>
        {
            entity.ToTable("BusinessUnits");

            entity.HasKey(e => e.Id);

            entity.Property(e => e.Id)
                .HasColumnName("Id");

            entity.Property(e => e.BusinessUnitId)
                .HasColumnName("BusinessUnitId")
                .HasMaxLength(50)
                .UseCollation("Thai_CI_AS");

            entity.Property(e => e.BusinessUnitCode)
                .HasColumnName("BusinessUnitCode")
                .HasMaxLength(50)
                .UseCollation("Thai_CI_AS");

            entity.Property(e => e.BusinessUnitLevel)
                .HasColumnName("BusinessUnitLevel");

            entity.Property(e => e.ParentId)
                .HasColumnName("ParentId")
                .HasMaxLength(50)
                .UseCollation("Thai_CI_AS");

            entity.Property(e => e.CompanyId)
                .HasColumnName("CompanyId")
                .HasMaxLength(50)
                .UseCollation("Thai_CI_AS");

            entity.Property(e => e.EffectiveDate)
                .HasColumnName("EffectiveDate");

            entity.Property(e => e.NameTh)
                .HasColumnName("NameTh")
                .HasMaxLength(500)
                .UseCollation("Thai_CI_AS");

            entity.Property(e => e.NameEn)
                .HasColumnName("NameEn")
                .HasMaxLength(500)
                .UseCollation("Thai_CI_AS");

            entity.Property(e => e.AbbreviationTh)
                .HasColumnName("AbbreviationTh")
                .HasMaxLength(100)
                .UseCollation("Thai_CI_AS");

            entity.Property(e => e.AbbreviationEn)
                .HasColumnName("AbbreviationEn")
                .HasMaxLength(100)
                .UseCollation("Thai_CI_AS");

            entity.Property(e => e.DescriptionTh)
                .HasColumnName("DescriptionTh")
                .UseCollation("Thai_CI_AS");

            entity.Property(e => e.DescriptionEn)
                .HasColumnName("DescriptionEn")
                .UseCollation("Thai_CI_AS");

            entity.Property(e => e.CreateDate)
                .HasColumnName("CreateDate");
        });


        OnModelCreatingPartial(modelBuilder);
    }

    partial void OnModelCreatingPartial(ModelBuilder modelBuilder);
}
