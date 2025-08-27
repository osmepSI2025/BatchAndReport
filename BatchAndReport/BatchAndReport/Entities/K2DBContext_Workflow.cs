using BatchAndReport.Models;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;

namespace BatchAndReport.Entities
{
    public partial class K2DBContext_Workflow : DbContext
    {
        public K2DBContext_Workflow() { }

        public K2DBContext_Workflow(DbContextOptions<K2DBContext_Workflow> options)
            : base(options) { }

        public virtual DbSet<AnnualProcessReview> AnnualProcessReviews { get; set; }
        public virtual DbSet<ProjectFiscalYear> ProjectFiscalYears { get; set; }
        public virtual DbSet<AnnualProcessReviewDetail> AnnualProcessReviewDetails { get; set; }
        public virtual DbSet<ProcessReviewType> ProcessReviewTypes { get; set; }
        public virtual DbSet<AnnualProcessReviewHistory> AnnualProcessReviewHistories { get; set; }
        public virtual DbSet<PlanCategory> PlanCategories { get; set; }
        public virtual DbSet<PlanCategoriesDetail> PlanCategoriesDetails { get; set; }
        public virtual DbSet<ProcessMasterDetail> ProcessMasterDetails { get; set; }
        public virtual DbSet<TempProcessMasterDetail> TempProcessMasterDetails { get; set; }
        public virtual DbSet<SubProcessReviewApproval> SubProcessReviewApprovals { get; set; }
        public virtual DbSet<WfTaskList> WfTaskLists { get; set; }
        public virtual DbSet<Evaluation> Evaluations { get; set; }
        public virtual DbSet<SubProcessMasterHistory> SubProcessMasterHistories { get; set; }
        public virtual DbSet<SubProcessMaster> SubProcessMasters { get; set; }
        public virtual DbSet<SubProcessControlDetail> SubProcessControlDetails { get; set; }
        public virtual DbSet<WfLookup> WFLookup { get; set; }

        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
#warning To protect potentially sensitive information in your connection string, you should move it out of source code.
            if (!optionsBuilder.IsConfigured)
            {
                optionsBuilder.UseSqlServer("Server=192.168.9.156;Database=SME;User Id=sa;Password=Osmep@2025;TrustServerCertificate=True;MultipleActiveResultSets=True;");
            }
        }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<AnnualProcessReview>(entity =>
            {
                entity.ToTable("ANNUAL_PROCESS_REVIEW");

                entity.HasKey(e => e.AnnualProcessReviewId);

                entity.Property(e => e.AnnualProcessReviewId).HasColumnName("ANNUAL_PROCESS_REVIEW_ID").IsRequired();
                entity.Property(e => e.ProcessReviewDetail).HasColumnName("PROCESS_REVIEW_DETAIL");
                entity.Property(e => e.ProcessBackground).HasColumnName("PROCESS_BACKGROUND");
                entity.Property(e => e.OwnerBusinessUnitId).HasColumnName("OWNER_BusinessUnitId").HasMaxLength(30);
                entity.Property(e => e.StatusCode).HasColumnName("STATUS_CODE").HasMaxLength(10);
                entity.Property(e => e.Detail).HasColumnName("DETAIL");
                entity.Property(e => e.CreatedDateTime).HasColumnName("CREATED_DATETIME").HasColumnType("datetime");
                entity.Property(e => e.UpdatedDateTime).HasColumnName("UPDATED_DATETIME").HasColumnType("datetime");
                entity.Property(e => e.CreatedBy).HasColumnName("CREATED_BY").HasMaxLength(50);
                entity.Property(e => e.UpdatedBy).HasColumnName("UPDATED_BY").HasMaxLength(50);
                entity.Property(e => e.FiscalYearId).HasColumnName("FISCAL_YEAR_ID");
                entity.Property(e => e.IsDeleted).HasColumnName("IS_DELETED");
                entity.Property(e => e.IsDraft).HasColumnName("IS_DRAFT");
                entity.Property(e => e.ApproveRemark).HasColumnName("APPROVE_REMARK");

                entity.HasOne(d => d.ProjectFiscalYear)
                    .WithMany(p => p.AnnualProcessReviews)
                    .HasForeignKey(d => d.FiscalYearId);
            });

            modelBuilder.Entity<AnnualProcessReviewDetail>(entity =>
            {
                entity.ToTable("ANNUAL_PROCESS_REVIEW_DETAIL");

                entity.HasKey(e => e.AnnualProcessReviewDetailId);

                entity.Property(e => e.AnnualProcessReviewDetailId).HasColumnName("ANNUAL_PROCESS_REVIEW_DETAIL_ID");
                entity.Property(e => e.AnnualProcessReviewId).HasColumnName("ANNUAL_PROCESS_REVIEW_ID");
                entity.Property(e => e.PrevProcessMasterId).HasColumnName("PREV_PROCESS_MASTER_ID");
                entity.Property(e => e.PrevProcessGroupCode).HasColumnName("PREV_PROCESS_GROUP_CODE").HasMaxLength(10);
                entity.Property(e => e.PrevProcessCode).HasColumnName("PREV_PROCESS_CODE").HasMaxLength(10);
                entity.Property(e => e.PrevProcessName).HasColumnName("PREV_PROCESS_NAME").HasMaxLength(100);
                entity.Property(e => e.ProcessGroupCode).HasColumnName("PROCESS_GROUP_CODE").HasMaxLength(10);
                entity.Property(e => e.ProcessCode).HasColumnName("PROCESS_CODE").HasMaxLength(10);
                entity.Property(e => e.ProcessName).HasColumnName("PROCESS_NAME").HasMaxLength(100);
                entity.Property(e => e.IsWiFilePath).HasColumnName("IS_WI_FILE_PATH").HasMaxLength(100);
                entity.Property(e => e.ProcessReviewTypeId).HasColumnName("PROCESS_REVIEW_TYPE_ID");
                entity.Property(e => e.FileUpload).HasColumnName("FILE_UPLOAD");
                entity.Property(e => e.IsWorkflow).HasColumnName("IS_WORKFLOW");
                entity.Property(e => e.IsCgdControlProcess).HasColumnName("IS_CGD_CONTROL_PROCESS");
                entity.Property(e => e.IsDeleted).HasColumnName("IS_DELETED");
                entity.Property(e => e.IsWi).HasColumnName("IS_WI");
                entity.Property(e => e.PrevIsWorkflow).HasColumnName("PREV_IS_WORKFLOW");
                entity.Property(e => e.PrevAnnualProcessReviewDetailId).HasColumnName("PREV_ANNUAL_PROCESS_REVIEW_DETAIL_ID");
            });

            modelBuilder.Entity<AnnualProcessReviewHistory>(entity =>
            {
                entity.ToTable("ANNUAL_PROCESS_REVIEW_HISTORY");

                entity.HasKey(e => e.AnnualProcessReviewHistoryId);

                entity.Property(e => e.AnnualProcessReviewHistoryId).HasColumnName("ANNUAL_PROCESS_REVIEW_HISTORY_ID");
                entity.Property(e => e.AnnualProcessReviewId).HasColumnName("ANNUAL_PROCESS_REVIEW_ID");
                entity.Property(e => e.Datetime).HasColumnName("DATETIME").HasColumnType("datetime");
                entity.Property(e => e.StatusCode).HasColumnName("STATUS_CODE").HasMaxLength(10);
                entity.Property(e => e.EmployeeId).HasColumnName("EMPLOYEE_ID").HasMaxLength(50);
                entity.Property(e => e.CreatedDateTime).HasColumnName("CREATED_DATETIME").HasColumnType("datetime");
                entity.Property(e => e.UpdatedDateTime).HasColumnName("UPDATED_DATETIME").HasColumnType("datetime");
                entity.Property(e => e.CreatedBy).HasColumnName("CREATED_BY").HasMaxLength(50);
                entity.Property(e => e.UpdatedBy).HasColumnName("UPDATED_BY").HasMaxLength(50);
                entity.Property(e => e.IsDeleted).HasColumnName("IS_DELETED");
            });

            modelBuilder.Entity<ProjectFiscalYear>(entity =>
            {
                entity.ToTable("PROJECT_FISCAL_YEAR");

                entity.HasKey(e => e.FiscalYearId);

                entity.Property(e => e.FiscalYearId).HasColumnName("FISCAL_YEAR_ID");
                entity.Property(e => e.StartDate).HasColumnName("START_DATE").HasColumnType("date");
                entity.Property(e => e.EndDate).HasColumnName("END_DATE").HasColumnType("date");
                entity.Property(e => e.FiscalYearDesc).HasColumnName("FISCAL_YEAR_DESC").HasMaxLength(50);
                entity.Property(e => e.CreateDate).HasColumnName("CREATE_DATE").HasColumnType("datetime");
                entity.Property(e => e.CreateBy).HasColumnName("CREATE_BY").HasMaxLength(50);
                entity.Property(e => e.UpdateDate).HasColumnName("UPDATE_DATE").HasColumnType("datetime");
                entity.Property(e => e.UpdateBy).HasColumnName("UPDATE_BY").HasMaxLength(50);
                entity.Property(e => e.ActiveFlag).HasColumnName("ACTIVE_FLAG");
                entity.Property(e => e.StartEndDateDesc).HasColumnName("START_END_DATE_DESC").HasMaxLength(50);

                entity.HasMany(p => p.AnnualProcessReviews)
                    .WithOne(d => d.ProjectFiscalYear)
                    .HasForeignKey(d => d.FiscalYearId);
            });

            modelBuilder.Entity<ProcessReviewType>(entity =>
            {
                entity.ToTable("PROCESS_REVIEW_TYPE");

                entity.HasKey(e => e.ProcessReviewTypeId);

                entity.Property(e => e.ProcessReviewTypeId).HasColumnName("PROCESS_REVIEW_TYPE_ID");
                entity.Property(e => e.ProcessReviewTypeName).HasColumnName("PROCESS_REVIEW_TYPE_NAME").HasMaxLength(100);
                entity.Property(e => e.CreatedDateTime).HasColumnName("CREATED_DATETIME").HasColumnType("datetime");
                entity.Property(e => e.UpdatedDateTime).HasColumnName("UPDATED_DATETIME").HasColumnType("datetime");
                entity.Property(e => e.CreatedBy).HasColumnName("CREATED_BY").HasMaxLength(50);
                entity.Property(e => e.UpdatedBy).HasColumnName("UPDATED_BY").HasMaxLength(50);
                entity.Property(e => e.IsActive).HasColumnName("IS_ACTIVE");
                entity.Property(e => e.IsDeleted).HasColumnName("IS_DELETED");
            });
            modelBuilder.Entity<PlanCategory>(entity =>
            {
                entity.ToTable("PLAN_CATEGORIES");

                entity.HasKey(e => e.PlanCategoriesId);

                entity.Property(e => e.PlanCategoriesId).HasColumnName("PLAN_CATEGORIES_ID");
                entity.Property(e => e.PlanCategoriesName).HasColumnName("PLAN_CATEGORIES_NAME").HasMaxLength(200);
                entity.Property(e => e.CreatedDateTime).HasColumnName("CREATED_DATETIME");
                entity.Property(e => e.UpdatedDateTime).HasColumnName("UPDATED_DATETIME");
                entity.Property(e => e.CreatedBy).HasColumnName("CREATED_BY").HasMaxLength(50);
                entity.Property(e => e.UpdatedBy).HasColumnName("UPDATED_BY").HasMaxLength(50);
                entity.Property(e => e.IsActive).HasColumnName("IS_ACTIVE");
                entity.Property(e => e.IsDeleted).HasColumnName("IS_DELETED");
            });
            modelBuilder.Entity<PlanCategoriesDetail>(entity =>
            {
                entity.ToTable("PLAN_CATEGORIES_DETAIL");

                entity.HasKey(e => e.PlanCategoriesDetailId);

                entity.Property(e => e.PlanCategoriesDetailId).HasColumnName("PLAN_CATEGORIES_DETAIL_ID");
                entity.Property(e => e.PlanCategoriesId).HasColumnName("PLAN_CATEGORIES_ID");
                entity.Property(e => e.BusinessUnitId).HasColumnName("BusinessUnitId").HasMaxLength(30);
                entity.Property(e => e.CreatedDateTime).HasColumnName("CREATED_DATETIME");
                entity.Property(e => e.UpdatedDateTime).HasColumnName("UPDATED_DATETIME");
                entity.Property(e => e.CreatedBy).HasColumnName("CREATED_BY").HasMaxLength(50);
                entity.Property(e => e.UpdatedBy).HasColumnName("UPDATED_BY").HasMaxLength(50);
                entity.Property(e => e.IsActive).HasColumnName("IS_ACTIVE");
                entity.Property(e => e.IsDeleted).HasColumnName("IS_DELETED");
                entity.Property(e => e.Objective).HasColumnName("OBJECTIVE");

                entity.HasOne(d => d.PlanCategory)
                      .WithMany(p => p.PlanCategoriesDetails)
                      .HasForeignKey(d => d.PlanCategoriesId);
            });
            modelBuilder.Entity<ProcessMasterDetail>(entity =>
            {
                entity.ToTable("PROCESS_MASTER_DETAIL");

                entity.HasKey(e => e.ProcessMasterDetailId);

                entity.Property(e => e.ProcessMasterDetailId)
                    .HasColumnName("PROCESS_MASTER_DETAIL_ID");

                entity.Property(e => e.ProcessMasterId)
                    .HasColumnName("PROCESS_MASTER_ID");

                entity.Property(e => e.ProcessTypeCode)
                    .HasColumnName("PROCESS_TYPE_CODE")
                    .HasMaxLength(10);

                entity.Property(e => e.ProcessGroupCode)
                    .HasColumnName("PROCESS_GROUP_CODE")
                    .HasMaxLength(10);

                entity.Property(e => e.ProcessGroupName)
                    .HasColumnName("PROCESS_GROUP_NAME")
                    .HasMaxLength(100);

                entity.Property(e => e.CreatedDateTime)
                    .HasColumnName("CREATED_DATETIME");

                entity.Property(e => e.UpdatedDateTime)
                    .HasColumnName("UPDATED_DATETIME");

                entity.Property(e => e.CreatedBy)
                    .HasColumnName("CREATED_BY")
                    .HasMaxLength(50);

                entity.Property(e => e.UpdatedBy)
                    .HasColumnName("UPDATED_BY")
                    .HasMaxLength(50);

                entity.Property(e => e.FiscalYearId)
                    .HasColumnName("FISCAL_YEAR_ID");

                entity.Property(e => e.IsDeleted)
                    .HasColumnName("IS_DELETED");

            });
            modelBuilder.Entity<SubProcessReviewApproval>(entity =>
            {
                entity.ToTable("SUB_PROCESS_REVIEW_APPROVAL");

                entity.HasKey(e => e.SubProcessReviewApprovalId);

                entity.Property(e => e.SubProcessReviewApprovalId).HasColumnName("SUB_PROCESS_REVIEW_APPROVAL_ID");
                entity.Property(e => e.SubProcessMasterId).HasColumnName("SUB_PROCESS_MASTER_ID");
                entity.Property(e => e.ApprovalTypeCode).HasColumnName("APPROVAL_TYPE_CODE").HasMaxLength(50);
                entity.Property(e => e.EmployeePositionId).HasColumnName("EMPLOYEE_PositionId").HasMaxLength(30);
                entity.Property(e => e.EmployeeId).HasColumnName("EMPLOYEE_ID").HasMaxLength(50);
                entity.Property(e => e.CreatedDateTime).HasColumnName("CREATED_DATETIME");
                entity.Property(e => e.UpdatedDateTime).HasColumnName("UPDATED_DATETIME");
                entity.Property(e => e.CreatedBy).HasColumnName("CREATED_BY").HasMaxLength(50);
                entity.Property(e => e.UpdatedBy).HasColumnName("UPDATED_BY").HasMaxLength(50);
                entity.Property(e => e.IsDeleted).HasColumnName("IS_DELETED");
            });
            modelBuilder.Entity<WfTaskList>(entity =>
            {
                entity.ToTable("WF_WFTaskList");

                entity.HasKey(e => e.WfTaskListId);

                entity.Property(e => e.WfTaskListId).HasColumnName("WFTaskListID");
                entity.Property(e => e.WfId).HasColumnName("WF_ID");
                entity.Property(e => e.Status).HasColumnName("STATUS").HasMaxLength(10);
                entity.Property(e => e.RequestId).HasColumnName("Request_ID");
                entity.Property(e => e.WfType).HasColumnName("WF_TYPE").HasMaxLength(100);
                entity.Property(e => e.CreateBy).HasColumnName("CREATEBY").HasMaxLength(100);
                entity.Property(e => e.UpdateBy).HasColumnName("UPDATEBY").HasMaxLength(100);
                entity.Property(e => e.CreateDate).HasColumnName("CREATEDATE");
                entity.Property(e => e.LastUpdate).HasColumnName("LASTUPDATE");
                entity.Property(e => e.CompleteOn).HasColumnName("COMPLETEON");
                entity.Property(e => e.Owner).HasColumnName("OWNER");
            });
            modelBuilder.Entity<Evaluation>(entity =>
            {
                entity.ToTable("EVALUATION");

                entity.HasKey(e => e.EvaluationId);

                entity.Property(e => e.EvaluationId).HasColumnName("EVALUATION_ID");
                entity.Property(e => e.SubProcessMasterId).HasColumnName("SUB_PROCESS_MASTER_ID");
                entity.Property(e => e.EvaluationDesc).HasColumnName("EVALUATION_DESC").HasMaxLength(300);
                entity.Property(e => e.PerformanceResult).HasColumnName("PERFORMANCE_RESULT").HasMaxLength(300);
                entity.Property(e => e.CreatedDateTime).HasColumnName("CREATED_DATETIME");
                entity.Property(e => e.UpdatedDateTime).HasColumnName("UPDATED_DATETIME");
                entity.Property(e => e.CreatedBy).HasColumnName("CREATED_BY").HasMaxLength(50);
                entity.Property(e => e.UpdatedBy).HasColumnName("UPDATED_BY").HasMaxLength(50);
                entity.Property(e => e.IsDeleted).HasColumnName("IS_DELETED");
            });
            modelBuilder.Entity<SubProcessMasterHistory>(entity =>
            {
                entity.ToTable("SUB_PROCESS_MASTER_HISTORY");

                entity.HasKey(e => e.SubProcessMasterHistoryId);

                entity.Property(e => e.SubProcessMasterHistoryId).HasColumnName("SUB_PROCESS_MASTER_HISTORY_ID");
                entity.Property(e => e.SubProcessMasterId).HasColumnName("SUB_PROCESS_MASTER_ID");
                entity.Property(e => e.ProcessMasterHistoryType).HasColumnName("PROCESS_MASTER_HISTORY_TYPE").HasMaxLength(10);
                entity.Property(e => e.EditDetail).HasColumnName("EDIT_DETAIL");
                entity.Property(e => e.DateTime).HasColumnName("DATETIME");
                entity.Property(e => e.StatusCode).HasColumnName("STATUS_CODE").HasMaxLength(10);
                entity.Property(e => e.EmployeeId).HasColumnName("EMPLOYEE_ID").HasMaxLength(50);
                entity.Property(e => e.CreatedDateTime).HasColumnName("CREATED_DATETIME");
                entity.Property(e => e.UpdatedDateTime).HasColumnName("UPDATED_DATETIME");
                entity.Property(e => e.CreatedBy).HasColumnName("CREATED_BY").HasMaxLength(50);
                entity.Property(e => e.UpdatedBy).HasColumnName("UPDATED_BY").HasMaxLength(50);
                entity.Property(e => e.IsDeleted).HasColumnName("IS_DELETED");
            });
            modelBuilder.Entity<SubProcessMaster>(entity =>
            {
                entity.ToTable("SUB_PROCESS_MASTER");

                entity.HasKey(e => e.SubProcessMasterId);

                entity.Property(e => e.SubProcessMasterId).HasColumnName("SUB_PROCESS_MASTER_ID");
                entity.Property(e => e.ProcessGroupCode).HasColumnName("PROCESS_GROUP_CODE").HasMaxLength(10);
                entity.Property(e => e.ProcessGroupName).HasColumnName("PROCESS_GROUP_NAME").HasMaxLength(100);
                entity.Property(e => e.ProcessCode).HasColumnName("PROCESS_CODE").HasMaxLength(10);
                entity.Property(e => e.ProcessName).HasColumnName("PROCESS_NAME").HasMaxLength(100);
                entity.Property(e => e.IsWorkflow).HasColumnName("IS_WORKFLOW");
                entity.Property(e => e.IsDigital).HasColumnName("IS_DIGITAL");
                entity.Property(e => e.IsCreateWorkflow).HasColumnName("IS_CREATE_WORKFLOW");
                entity.Property(e => e.ProcessTypeCode).HasColumnName("PROCESS_TYPE_CODE").HasMaxLength(10);
                entity.Property(e => e.DiagramAttachFile).HasColumnName("DIAGRAM_ATTACH_FILE");
                entity.Property(e => e.ProcessAttachFile).HasColumnName("PROCESS_ATTACH_FILE");
                entity.Property(e => e.ApprovalReviewDetail).HasColumnName("APPROVAL_REVIEW_DETAIL");
                entity.Property(e => e.StatusCode).HasColumnName("STATUS_CODE").HasMaxLength(10);
                entity.Property(e => e.EvaluationStatus).HasColumnName("EVALUATION_STATUS").HasMaxLength(10);
                entity.Property(e => e.EvaluationReviewRemark).HasColumnName("EVALUATION_REVIEW_REMARK");
                entity.Property(e => e.CreatedDateTime).HasColumnName("CREATED_DATETIME");
                entity.Property(e => e.UpdatedDateTime).HasColumnName("UPDATED_DATETIME");
                entity.Property(e => e.CreatedBy).HasColumnName("CREATED_BY").HasMaxLength(50);
                entity.Property(e => e.UpdatedBy).HasColumnName("UPDATED_BY").HasMaxLength(50);
                entity.Property(e => e.IsDeleted).HasColumnName("IS_DELETED");
                entity.Property(e => e.FiscalYearId).HasColumnName("FISCAL_YEAR_ID");
                entity.Property(e => e.ProcessMasterId).HasColumnName("PROCESS_MASTER_ID");
                entity.Property(e => e.ProcessDay).HasColumnName("PROCESS_DAY");
            });
            modelBuilder.Entity<SubProcessControlDetail>(entity =>
            {
                entity.ToTable("SUB_PROCESS_CONTROL_DETAIL");

                entity.HasKey(e => e.SubProcessControlDetailId);

                entity.Property(e => e.SubProcessControlDetailId).HasColumnName("SUB_PROCESS_CONTROL_DETAIL_ID");
                entity.Property(e => e.SubProcessMasterId).HasColumnName("SUB_PROCESS_MASTER_ID");
                entity.Property(e => e.ProcessControlCode).HasColumnName("PROCESS_CONTROL_CODE").HasMaxLength(10);
                entity.Property(e => e.ProcessControlActivity).HasColumnName("PROCESS_CONTROL_ACTIVITY").HasMaxLength(200);
                entity.Property(e => e.ProcessControlDetail).HasColumnName("PROCESS_CONTROL_DETAIL");
                entity.Property(e => e.ProcessControlDay).HasColumnName("PROCESS_CONTROL_DAY");
                entity.Property(e => e.CreatedDateTime).HasColumnName("CREATED_DATETIME");
                entity.Property(e => e.UpdatedDateTime).HasColumnName("UPDATED_DATETIME");
                entity.Property(e => e.CreatedBy).HasColumnName("CREATED_BY").HasMaxLength(50);
                entity.Property(e => e.UpdatedBy).HasColumnName("UPDATED_BY").HasMaxLength(50);
                entity.Property(e => e.IsDeleted).HasColumnName("IS_DELETED");
            });
            modelBuilder.Entity<TempProcessMasterDetail>(entity =>
            {
                entity.ToTable("TMP_PROCESS_MASTER_DETAIL");

                entity.HasKey(e => e.ProcessMasterDetailId);

                entity.Property(e => e.ProcessMasterDetailId)
                    .HasColumnName("TMP_PROCESS_MASTER_DETAIL_ID");

                entity.Property(e => e.ProcessMasterId)
                    .HasColumnName("TMP_PROCESS_MASTER_ID");

                entity.Property(e => e.ProcessTypeCode)
                    .HasColumnName("PROCESS_TYPE_CODE")
                    .HasMaxLength(10);

                entity.Property(e => e.ProcessGroupCode)
                    .HasColumnName("PROCESS_GROUP_CODE")
                    .HasMaxLength(10);

                entity.Property(e => e.ProcessGroupName)
                    .HasColumnName("PROCESS_GROUP_NAME")
                    .HasMaxLength(100);

                entity.Property(e => e.CreatedDateTime)
                    .HasColumnName("CREATED_DATETIME");

              

                entity.Property(e => e.CreatedBy)
                    .HasColumnName("CREATED_BY")
                    .HasMaxLength(50);

              

                entity.Property(e => e.FiscalYearId)
                    .HasColumnName("FISCAL_YEAR_ID");


                entity.Property(e => e.USER_PROCESS_REVIEW_NAME)
                  .HasColumnName("USER_PROCESS_REVIEW_NAME")
                  .HasMaxLength(500); 
            });

            modelBuilder.Entity<WfLookup>(entity =>
            {
                entity.HasKey(e => e.Id).HasName("PK_Workflow_Lookup");

                entity.ToTable("WF_Lookup");

                entity.Property(e => e.CreateBy).HasMaxLength(50);
                entity.Property(e => e.CreateDate).HasColumnType("datetime");
                entity.Property(e => e.FlagDelete)
                    .HasMaxLength(10)
                    .IsFixedLength();
                entity.Property(e => e.LookupCode).HasMaxLength(50);
                entity.Property(e => e.LookupType).HasMaxLength(50);
                entity.Property(e => e.LookupValue).HasMaxLength(50);
                entity.Property(e => e.UpdateBy).HasMaxLength(50);
                entity.Property(e => e.UpdateDate).HasColumnType("datetime");
            });

            OnModelCreatingPartial(modelBuilder);
        }


        partial void OnModelCreatingPartial(ModelBuilder modelBuilder);
    }
}
