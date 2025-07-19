using Microsoft.EntityFrameworkCore;

namespace BatchAndReport.Entities
{
    public class K2DBContext_EContract : DbContext
    {
        public K2DBContext_EContract() { }

        public K2DBContext_EContract(DbContextOptions<K2DBContext_EContract> options)
            : base(options)
        {
        }

        public virtual DbSet<EmployeeContract> EmployeeContracts { get; set; }
        public virtual DbSet<EmployeeProfile> EmployeeProfiles { get; set; }
        public virtual DbSet<ContractParty> ContractParties { get; set; }

        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
#warning To protect potentially sensitive information in your connection string, you should move it out of source code. You can avoid scaffolding the connection string by using the Name= syntax to read it from configuration - see https://go.microsoft.com/fwlink/?linkid=2131148. For more guidance on storing connection strings, see http://go.microsoft.com/fwlink/?LinkId=723263.
        => optionsBuilder.UseSqlServer("Server=192.168.9.156;Database=E-Contract;User Id=sa;Password=Osmep@2025;TrustServerCertificate=True;");
        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<EmployeeContract>(entity =>
            {
                entity.ToTable("EMPLOYEE_CONTRACT");

                entity.HasKey(e => e.Id);

                entity.Property(e => e.Id)
                    .HasColumnName("Id")
                    .IsRequired();

                entity.Property(e => e.ContractFlag)
                    .HasColumnName("Contract_Flag")
                    .IsRequired();

                entity.Property(e => e.EmployeeId)
                    .HasColumnName("Employee_Id")
                    .HasMaxLength(100)
                    .IsRequired();

                entity.Property(e => e.EmployeeCode)
                    .HasColumnName("Employee_Code")
                    .HasMaxLength(50);

                entity.Property(e => e.NameTh)
                    .HasColumnName("Name_Th")
                    .HasMaxLength(255);

                entity.Property(e => e.NameEn)
                    .HasColumnName("Name_En")
                    .HasMaxLength(255);

                entity.Property(e => e.FirstNameTh)
                    .HasColumnName("FirstName_Th")
                    .HasMaxLength(100);

                entity.Property(e => e.FirstNameEn)
                    .HasColumnName("FirstName_En")
                    .HasMaxLength(100);

                entity.Property(e => e.LastNameTh)
                    .HasColumnName("LastName_Th")
                    .HasMaxLength(100);

                entity.Property(e => e.LastNameEn)
                    .HasColumnName("LastName_En")
                    .HasMaxLength(100);

                entity.Property(e => e.Email)
                    .HasColumnName("Email")
                    .HasMaxLength(255);

                entity.Property(e => e.Mobile)
                    .HasColumnName("Mobile")
                    .HasMaxLength(20);

                entity.Property(e => e.EmploymentDate)
                    .HasColumnName("Employment_Date")
                    .HasColumnType("date");

                entity.Property(e => e.TerminationDate)
                    .HasColumnName("Termination_Date")
                    .HasColumnType("date");

                entity.Property(e => e.EmployeeType)
                    .HasColumnName("Employee_Type")
                    .HasMaxLength(50);

                entity.Property(e => e.EmployeeStatus)
                    .HasColumnName("Employee_Status")
                    .HasMaxLength(10);

                entity.Property(e => e.SupervisorId)
                    .HasColumnName("Supervisor_Id")
                    .HasMaxLength(50);

                entity.Property(e => e.CompanyId)
                    .HasColumnName("Company_Id")
                    .HasMaxLength(50);

                entity.Property(e => e.BusinessUnitId)
                    .HasColumnName("BusinessUnit_Id")
                    .HasMaxLength(100);

                entity.Property(e => e.PositionId)
                    .HasColumnName("Position_Id")
                    .HasMaxLength(50);

                entity.Property(e => e.Salary)
                    .HasColumnName("Salary")
                    .HasMaxLength(255);

                entity.Property(e => e.IdCard)
                    .HasColumnName("IdCard")
                    .HasMaxLength(255);

                entity.Property(e => e.PassportNo)
                    .HasColumnName("Passport_No")
                    .HasMaxLength(100);
            });
            modelBuilder.Entity<EmployeeProfile>(entity =>
            {
                entity.ToTable("EMPLOYEE_PROFILE");

                entity.HasKey(e => e.Id);

                entity.Property(e => e.Id)
                    .HasColumnName("Id")
                    .IsRequired();

                entity.Property(e => e.EmployeeId)
                    .HasColumnName("Employee_Id")
                    .HasMaxLength(100)
                    .IsRequired();

                entity.Property(e => e.InternalPhone)
                    .HasColumnName("Internal_Phone")
                    .HasMaxLength(20);

                entity.Property(e => e.MilitaryStatus)
                    .HasColumnName("Military_Status")
                    .HasMaxLength(10);

                entity.Property(e => e.MailingAddrTh)
                    .HasColumnName("Mailing_Addr_Th")
                    .HasMaxLength(500);

                entity.Property(e => e.MailingAddrEn)
                    .HasColumnName("Mailing_Addr_En")
                    .HasMaxLength(500);

                entity.Property(e => e.MailingSubdistrict)
                    .HasColumnName("Mailing_Subdistrict")
                    .HasMaxLength(100);

                entity.Property(e => e.MailingDistrict)
                    .HasColumnName("Mailing_District")
                    .HasMaxLength(100);

                entity.Property(e => e.MailingProvince)
                    .HasColumnName("Mailing_Province")
                    .HasMaxLength(10);

                entity.Property(e => e.MailingCountry)
                    .HasColumnName("Mailing_Country")
                    .HasMaxLength(10);

                entity.Property(e => e.MailingPostCode)
                    .HasColumnName("Mailing_PostCode")
                    .HasMaxLength(10);

                entity.Property(e => e.MailingPhoneNo)
                    .HasColumnName("Mailing_Phone_No")
                    .HasMaxLength(20);

                entity.Property(e => e.RegisAddrTh)
                    .HasColumnName("Regis_Addr_Th")
                    .HasMaxLength(500);

                entity.Property(e => e.RegisAddrEn)
                    .HasColumnName("Regis_Addr_En")
                    .HasMaxLength(500);

                entity.Property(e => e.RegisSubdistrict)
                    .HasColumnName("Regis_Subdistrict")
                    .HasMaxLength(100);

                entity.Property(e => e.RegisDistrict)
                    .HasColumnName("Regis_District")
                    .HasMaxLength(100);

                entity.Property(e => e.RegisProvince)
                    .HasColumnName("Regis_Province")
                    .HasMaxLength(10);

                entity.Property(e => e.RegisCountry)
                    .HasColumnName("Regis_Country")
                    .HasMaxLength(10);

                entity.Property(e => e.RegisPostCode)
                    .HasColumnName("Regis_PostCode")
                    .HasMaxLength(10);

                entity.Property(e => e.RegisPhoneNo)
                    .HasColumnName("Regis_Phone_No")
                    .HasMaxLength(20);

                entity.Property(e => e.BloodGroup)
                    .HasColumnName("Blood_Group")
                    .HasMaxLength(10);

                entity.Property(e => e.Religion)
                    .HasColumnName("Religion")
                    .HasMaxLength(50);

                entity.Property(e => e.Race)
                    .HasColumnName("Race")
                    .HasMaxLength(50);

                entity.Property(e => e.Nationality)
                    .HasColumnName("Nationality")
                    .HasMaxLength(50);

                entity.Property(e => e.JobDetails)
                    .HasColumnName("Job_Details");

                entity.Property(e => e.NickName)
                    .HasColumnName("Nick_Name")
                    .HasMaxLength(100);
            });
            modelBuilder.Entity<ContractParty>(entity =>
            {
                entity.ToTable("Contract_Party");

                entity.HasKey(e => e.Id);

                entity.Property(e => e.Id)
                    .HasColumnName("Id")
                    .IsRequired();

                entity.Property(e => e.ContractPartyName)
                    .HasColumnName("Contract_Party_Name")
                    .HasMaxLength(100);

                entity.Property(e => e.RegType)
                    .HasColumnName("Reg_Type")
                    .HasMaxLength(50);

                entity.Property(e => e.RegIden)
                    .HasColumnName("Reg_Iden")
                    .HasMaxLength(50);

                entity.Property(e => e.RegDetail)
                    .HasColumnName("Reg_Detail")
                    .HasMaxLength(50);

                entity.Property(e => e.AddressNo)
                    .HasColumnName("Address_No")
                    .HasMaxLength(50);

                entity.Property(e => e.SubDistrict)
                    .HasColumnName("Sub_District")
                    .HasMaxLength(50);

                entity.Property(e => e.District)
                    .HasColumnName("District")
                    .HasMaxLength(50);

                entity.Property(e => e.Province)
                    .HasColumnName("Province")
                    .HasMaxLength(50);

                entity.Property(e => e.PostalCode)
                    .HasColumnName("Postal_Code")
                    .HasMaxLength(50);

                entity.Property(e => e.FlagActive)
                    .HasColumnName("Flag_Active")
                    .HasMaxLength(1);
            });

        }
    }

    public class EmployeeContract
    {
        public int Id { get; set; }
        public bool ContractFlag { get; set; }
        public string EmployeeId { get; set; }
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
        public string? Salary { get; set; }
        public string? IdCard { get; set; }
        public string? PassportNo { get; set; }
    }

    public class EmployeeProfile
    {
        public int Id { get; set; }
        public string EmployeeId { get; set; }
        public string? InternalPhone { get; set; }
        public string? MilitaryStatus { get; set; }
        public string? MailingAddrTh { get; set; }
        public string? MailingAddrEn { get; set; }
        public string? MailingSubdistrict { get; set; }
        public string? MailingDistrict { get; set; }
        public string? MailingProvince { get; set; }
        public string? MailingCountry { get; set; }
        public string? MailingPostCode { get; set; }
        public string? MailingPhoneNo { get; set; }
        public string? RegisAddrTh { get; set; }
        public string? RegisAddrEn { get; set; }
        public string? RegisSubdistrict { get; set; }
        public string? RegisDistrict { get; set; }
        public string? RegisProvince { get; set; }
        public string? RegisCountry { get; set; }
        public string? RegisPostCode { get; set; }
        public string? RegisPhoneNo { get; set; }
        public string? BloodGroup { get; set; }
        public string? Religion { get; set; }
        public string? Race { get; set; }
        public string? Nationality { get; set; }
        public string? JobDetails { get; set; }
        public string? NickName { get; set; }
    }
}