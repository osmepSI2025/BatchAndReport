using BatchAndReport.Models;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
namespace BatchAndReport.Entities;


public partial class BatchDBContext : DbContext
{
    public BatchDBContext()
    {
    }

    public BatchDBContext(DbContextOptions<BatchDBContext> options)
        : base(options)
    {
    }

    public virtual DbSet<MApiInformation> MApiInformations { get; set; }

    public virtual DbSet<MscheduledJob> MscheduledJobs { get; set; }

    protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
#warning To protect potentially sensitive information in your connection string, you should move it out of source code. You can avoid scaffolding the connection string by using the Name= syntax to read it from configuration - see https://go.microsoft.com/fwlink/?linkid=2131148. For more guidance on storing connection strings, see http://go.microsoft.com/fwlink/?LinkId=723263.
        => optionsBuilder.UseSqlServer("Server=192.168.9.155;Database=SME_BatchTask;User Id=sa;Password=Osmep@2025;TrustServerCertificate=True;");

    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {
        modelBuilder.Entity<MApiInformation>(entity =>
        {
            entity.ToTable("M_ApiInformation");

            entity.Property(e => e.Id).HasColumnName("id");
            entity.Property(e => e.AccessToken).UseCollation("Thai_CI_AS");
            entity.Property(e => e.ApiKey)
                .HasMaxLength(150)
                .UseCollation("Thai_CI_AS");
            entity.Property(e => e.AuthorizationType)
                .HasMaxLength(50)
                .UseCollation("Thai_CI_AS");
            entity.Property(e => e.Bearer)
                .UseCollation("Thai_CI_AS")
                .HasColumnType("ntext");
            entity.Property(e => e.ContentType)
                .HasMaxLength(150)
                .UseCollation("Thai_CI_AS");
            entity.Property(e => e.CreateDate).HasColumnType("datetime");
            entity.Property(e => e.MethodType)
                .HasMaxLength(50)
                .UseCollation("Thai_CI_AS");
            entity.Property(e => e.Password)
                .HasMaxLength(150)
                .UseCollation("Thai_CI_AS");
            entity.Property(e => e.ServiceNameCode)
                .HasMaxLength(250)
                .UseCollation("Thai_CI_AS");
            entity.Property(e => e.ServiceNameTh)
                .HasMaxLength(250)
                .UseCollation("Thai_CI_AS");
            entity.Property(e => e.UpdateDate).HasColumnType("datetime");
            entity.Property(e => e.Urldevelopment)
                .UseCollation("Thai_CI_AS")
                .HasColumnName("URLDevelopment");
            entity.Property(e => e.Urlproduction)
                .UseCollation("Thai_CI_AS")
                .HasColumnName("URLProduction");
            entity.Property(e => e.Username)
                .HasMaxLength(50)
                .UseCollation("Thai_CI_AS");
        });

        modelBuilder.Entity<MscheduledJob>(entity =>
        {
            entity.ToTable("MScheduledJobs");

            entity.Property(e => e.Id).HasColumnName("id");
            entity.Property(e => e.JobName)
                .HasMaxLength(150)
                .UseCollation("Thai_CI_AS");
        });


        OnModelCreatingPartial(modelBuilder);
    }

    partial void OnModelCreatingPartial(ModelBuilder modelBuilder);
}
