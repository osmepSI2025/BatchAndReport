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

    protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
#warning To protect potentially sensitive information in your connection string, you should move it out of source code. You can avoid scaffolding the connection string by using the Name= syntax to read it from configuration - see https://go.microsoft.com/fwlink/?linkid=2131148. For more guidance on storing connection strings, see http://go.microsoft.com/fwlink/?LinkId=723263.
        => optionsBuilder.UseSqlServer("Server=doublep-cloud.servehttp.com;Database=K2_HR;User Id=Dev;Password=J01nSP0pD0ub!EP;TrustServerCertificate=True;");

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
        });


        OnModelCreatingPartial(modelBuilder);
    }

    partial void OnModelCreatingPartial(ModelBuilder modelBuilder);
}
