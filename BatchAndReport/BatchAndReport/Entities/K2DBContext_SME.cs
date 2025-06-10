using BatchAndReport.Models;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
namespace BatchAndReport.Entities;


public partial class K2DBContext_SME : DbContext
{
    public K2DBContext_SME()
    {
    }

    public K2DBContext_SME(DbContextOptions<K2DBContext_SME> options)
        : base(options)
    {
    }


    public virtual DbSet<ProjectMaster> ProjectMasters { get; set; }
    public virtual DbSet<ProjectYear> ProjectYears { get; set; }

    protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
#warning To protect potentially sensitive information in your connection string, you should move it out of source code. You can avoid scaffolding the connection string by using the Name= syntax to read it from configuration - see https://go.microsoft.com/fwlink/?linkid=2131148. For more guidance on storing connection strings, see http://go.microsoft.com/fwlink/?LinkId=723263.
        => optionsBuilder.UseSqlServer("Server=192.168.9.156;Database=SME;User Id=sa;Password=Osmep@2025;TrustServerCertificate=True;");

    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {

        modelBuilder.Entity<ProjectMaster>(entity =>
        {
            entity.ToTable("SME_PROJECT_MASTER");

            entity.HasKey(e => e.ProjectMasterId);

            entity.Property(e => e.ProjectMasterId)
                .HasColumnName("PROJECT_MASTER_ID")
                .IsRequired();

            entity.Property(e => e.ProjectName)
                .HasColumnName("PROJECT_NAME")
                .HasMaxLength(1)
                .IsRequired()
                .UseCollation("Thai_CI_AS");

            entity.Property(e => e.BudgetAmount)
                .HasColumnName("BUDGET_AMOUNT")
                .HasColumnType("decimal(18, 0)")
                .IsRequired();

            entity.Property(e => e.Issue)
                .HasColumnName("ISSUE")
                .HasMaxLength(255)
                .UseCollation("Thai_CI_AS");

            entity.Property(e => e.Strategy)
                .HasColumnName("STRATEGY")
                .HasMaxLength(255)
                .UseCollation("Thai_CI_AS");

            entity.Property(e => e.FiscalYear)
                .HasColumnName("FISCAL_YEAR")
                .HasMaxLength(4)
                .UseCollation("Thai_CI_AS");

            entity.Property(e => e.KeyId)
                .HasColumnName("KeyId");
        });

        modelBuilder.Entity<ProjectYear>(entity =>
        {
            entity.ToTable("SME_PROJECT_FISCAL_YEAR");

            entity.HasKey(e => e.FISCAL_YEAR_DESC);

            entity.Property(e => e.FISCAL_YEAR_DESC)
                .HasColumnName("FISCAL_YEAR_DESC")
                .HasMaxLength(4)
                .IsRequired()
                .UseCollation("Thai_CI_AS");

            entity.Property(e => e.START_DATE)
                .HasColumnName("START_DATE")
                .HasColumnType("date")
                .IsRequired();

            entity.Property(e => e.END_DATE)
                .HasColumnName("END_DATE")
                .HasColumnType("date")
                .IsRequired();

            entity.Property(e => e.CREATE_DATE)
                .HasColumnName("CREATE_DATE")
                .HasColumnType("date")
                .IsRequired();

            entity.Property(e => e.CREATE_BY)
                .HasColumnName("CREATE_BY")
                .HasMaxLength(10)
                .IsRequired()
                .UseCollation("Thai_CI_AS");

            entity.Property(e => e.ACTIVE_FLAG)
                .HasColumnName("ACTIVE_FLAG")
                .HasMaxLength(1)
                .IsRequired()
                .UseCollation("Thai_CI_AS");

            entity.Property(e => e.UPDATE_DATE)
                .HasColumnName("UPDATE_DATE")
                .HasColumnType("date");

            entity.Property(e => e.UPDATE_BY)
                .HasColumnName("UPDATE_BY")
                .HasMaxLength(10)
                .UseCollation("Thai_CI_AS");
        });




        OnModelCreatingPartial(modelBuilder);
    }

    partial void OnModelCreatingPartial(ModelBuilder modelBuilder);
}
