using System;
using System.Collections.Generic;
using System.Configuration;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;

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

    public virtual DbSet<MscheduledJob> MscheduledJobs { get; set; }

    protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
    {
        if (!optionsBuilder.IsConfigured)
        {
            var configuration = new ConfigurationBuilder()
                .SetBasePath(AppDomain.CurrentDomain.BaseDirectory)
                .AddJsonFile("appsettings.json")
                .Build();

            var connectionString = configuration.GetConnectionString("DefaultConnection");
            optionsBuilder.UseSqlServer(connectionString);
        }
    }

    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {
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
