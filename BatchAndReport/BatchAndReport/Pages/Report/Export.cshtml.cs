using BatchAndReport.DAO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using OfficeOpenXml;
using PdfSharpCore.Pdf.IO;
using System.Diagnostics.Contracts;
using System.Threading.Tasks;
namespace BatchAndReport.Pages.Report
{
    public class ExportModel : PageModel
    {
        private readonly SmeDAO _smeDao;
        private readonly WordEContract_AllowanceService _AllowanceService;
        private readonly WordEContract_LoanPrinterService _wordEContract_LoanPrinterService;
        private readonly WordEContract_ContactToDoThingService _ContactToDoThingService;
        private readonly WordEContract_HireEmployee _HireEmployee;
        private readonly WordEContract_BorrowMoneyService _BorrowMoneyService;
        private readonly WordEContract_MaintenanceComputerService _maintenanceComputerService;
        private readonly WordEContract_LoanComputerService _LoanComputerService;
        private readonly WordEContract_BuyAgreeProgram _BuyAgreeProgram;
        private readonly WordEContract_BuyOrSellComputerService _BuyOrSellComputerService;
        private readonly WordEContract_BuyOrSellService _BuyOrSellService;
        private readonly WordEContract_DataSecretService _DataSecretService;
        private readonly WordEContract_MemorandumService _MemorandumService;
        private readonly WordEContract_PersernalProcessService _PersernalProcessService;
        private readonly WordEContract_SupportSMEsService _SupportSMEsService;
        private readonly WordEContract_JointOperationService _JointOperationService;
        private readonly WordEContract_ControlDataService _ControlDataService;
        private readonly WordEContract_DataPersonalService _DataPersonalService;
        private readonly WordEContract_ConsultantService _ConsultantService;
        private readonly WordEContract_Test_HeaderLOGOService _Test_HeaderLOGOService;
        private readonly IConfiguration _configuration;
        public ExportModel(SmeDAO smeDao, WordEContract_AllowanceService allowanceService
            , WordEContract_LoanPrinterService wordEContract_LoanPrinterService
            , WordEContract_ContactToDoThingService ContactToDoThingService
            , WordEContract_HireEmployee HireEmployee
            , WordEContract_BorrowMoneyService BorrowMoneyService
            , WordEContract_MaintenanceComputerService maintenanceComputerService
            , WordEContract_LoanComputerService LoanComputerService
            , WordEContract_BuyAgreeProgram BuyAgreeProgram
            , WordEContract_BuyOrSellComputerService BuyOrSellComputerService
            , WordEContract_BuyOrSellService BuyOrSellService
            , WordEContract_DataSecretService DataSecretService

            , WordEContract_MemorandumService MemorandumService
            , WordEContract_PersernalProcessService PersernalProcessService
            , WordEContract_SupportSMEsService SupportSMEsService
            , WordEContract_JointOperationService JointOperationService
            , WordEContract_ControlDataService ControlDataService
            , WordEContract_DataPersonalService DataPersonalService
            , WordEContract_ConsultantService ConsultantService
            , WordEContract_Test_HeaderLOGOService Test_HeaderLOGOService
            , IConfiguration configuration // <-- add this
            )
        {
            _smeDao = smeDao;
            _AllowanceService = allowanceService;
            this._wordEContract_LoanPrinterService = wordEContract_LoanPrinterService;
            _ContactToDoThingService = ContactToDoThingService;
            _HireEmployee = HireEmployee;
            _BorrowMoneyService = BorrowMoneyService;
            _maintenanceComputerService = maintenanceComputerService;
            this._LoanComputerService = LoanComputerService;
            this._BuyAgreeProgram = BuyAgreeProgram;
             _BuyOrSellComputerService = BuyOrSellComputerService;
            this._BuyOrSellService = BuyOrSellService;
            this._DataSecretService = DataSecretService;
            this._MemorandumService = MemorandumService;
            this._PersernalProcessService = PersernalProcessService;
            this._SupportSMEsService = SupportSMEsService;
            this._JointOperationService = JointOperationService;
            this._ControlDataService = ControlDataService;
            this._DataPersonalService = DataPersonalService;
            this._ConsultantService = ConsultantService;
            this._Test_HeaderLOGOService = Test_HeaderLOGOService;
            _configuration = configuration; // <-- initialize the configuration
        }
        public IActionResult OnGetPdf()
        {
            var wordDAO = new WordToPDFDAO(); // Create an instance of WordDAO
            var Resultpdf = wordDAO.OnGetPdfWithInterop(); // Call the method on the instance
            return Resultpdf; // Return an empty result since the PDF is handled in WordDAO
        }

        





        #region สสว. สัญญารับเงินอุดหนุน
        // This is your specific handler for the contract report
        public IActionResult OnGetWordContactAllowance()
        {

            var wordBytes = _AllowanceService.OnGetWordContact_Allowance();
            return File(wordBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "สัญญารับเงินอุดหนุน.docx");


        }
        // Helper for colored, bold, centered paragraph
        #endregion สสว. สัญญารับเงินอุดหนุน



        #region สสว. สัญญาเงินกู้ยืม โครงการพลิกฟื้นวิสาห 
        public IActionResult OnGetWordContactBorrowMoney()
        {

            var wordBytes = _BorrowMoneyService.OnGetWordContact_orrowMoney();
            return File(wordBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "สัญญาเงินกู้ยืมโครงการพลิกฟื้นวิสาห.docx");

        }
        #endregion สสว. สัญญาเงินกู้ยืม โครงการพลิกฟื้นวิสาห 

        #region  4.1.3.3. สัญญาจ้างลูกจ้าง EC
        public async Task<IActionResult> OnGetWordContact_EC(string ContractId="7")
        {
            var wordBytes = await _HireEmployee.OnGetWordContact_HireEmployee(ContractId);
            return File(wordBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "สัญญาจ้างลูกจ้าง.docx");

        }

        public async Task OnGetWordContact_EC_PDF(string ContractId = "8")
        {
            var wordBytes = await _HireEmployee.OnGetWordContact_HireEmployee_ToPDF(ContractId, "EC");
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "EC");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var filePath = Path.Combine(folderPath, "EC_" + ContractId + ".pdf");

            // Delete the file if it already exists
            if (System.IO.File.Exists(filePath))
            {
                System.IO.File.Delete(filePath);
            }
            await System.IO.File.WriteAllBytesAsync(filePath, wordBytes);

            // return File(wordBytes, "application/pdf", "สัญญาซื้อขาย.pdf");
        }
        public async Task<IActionResult> OnGetWordContact_EC_PDF_Preview(string ContractId = "8")
        {
            var wordBytes = await _HireEmployee.OnGetWordContact_HireEmployee_ToPDF(ContractId, "EC");
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "EC");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var fileName = $"EC_{ContractId}_Preview.pdf";
            var filePath = Path.Combine(folderPath, fileName);

            // Get password from appsettings.json
            string userPassword = _configuration["Password:PaswordPDF"];

            using (var inputStream = new MemoryStream(wordBytes))
            using (var outputStream = new MemoryStream())
            {
                var document = PdfSharpCore.Pdf.IO.PdfReader.Open(inputStream, PdfSharpCore.Pdf.IO.PdfDocumentOpenMode.Modify);

                var securitySettings = document.SecuritySettings;
                securitySettings.UserPassword = userPassword;
                securitySettings.OwnerPassword = userPassword;
                securitySettings.PermitPrint = true;
                securitySettings.PermitModifyDocument = false;
                securitySettings.PermitExtractContent = false;
                securitySettings.PermitAnnotations = false;

                document.Save(outputStream);

                // Optionally save to disk
                // await System.IO.File.WriteAllBytesAsync(filePath, outputStream.ToArray());

                return File(outputStream.ToArray(), "application/pdf", fileName);
            }
        }
        #endregion 4.1.3.3. สัญญาจ้างลูกจ้าง

        #region 4.1.1.2.15.สัญญาจ้างทำของ CWA
        public async Task<IActionResult> OnGetWordContact_CWA(string ContractId="1")
        {
            var wordBytes = await _ContactToDoThingService.OnGetWordContact_ToDoThing(ContractId);
            return File(wordBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "สัญญาจ้างทำของ.docx");
        }

        public async Task OnGetWordContact_CWA_PDF(string ContractId = "1")
        {
            var wordBytes = await _ContactToDoThingService.OnGetWordContact_ToDoThing_ToPDF(ContractId, "CWA");
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "CWA");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var filePath = Path.Combine(folderPath, "CWA_" + ContractId + ".pdf");

            // Delete the file if it already exists
            if (System.IO.File.Exists(filePath))
            {
                System.IO.File.Delete(filePath);
            }
            await System.IO.File.WriteAllBytesAsync(filePath, wordBytes);

            // return File(wordBytes, "application/pdf", "สัญญาซื้อขาย.pdf");
        }
        public async Task<IActionResult> OnGetWordContact_CWA_PDF_Preview(string ContractId = "1")
        {
            var wordBytes = await _ContactToDoThingService.OnGetWordContact_ToDoThing_ToPDF(ContractId, "CWA");
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "CWA");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var fileName = $"CWA_{ContractId}_Preview.pdf";
            var filePath = Path.Combine(folderPath, fileName);

            // Get password from appsettings.json
            string userPassword = _configuration["Password:PaswordPDF"];

            using (var inputStream = new MemoryStream(wordBytes))
            using (var outputStream = new MemoryStream())
            {
                var document = PdfSharpCore.Pdf.IO.PdfReader.Open(inputStream, PdfSharpCore.Pdf.IO.PdfDocumentOpenMode.Modify);

                var securitySettings = document.SecuritySettings;
                securitySettings.UserPassword = userPassword;
                securitySettings.OwnerPassword = userPassword;
                securitySettings.PermitPrint = true;
                securitySettings.PermitModifyDocument = false;
                securitySettings.PermitExtractContent = false;
                securitySettings.PermitAnnotations = false;

                document.Save(outputStream);

                // Optionally save to disk
                // await System.IO.File.WriteAllBytesAsync(filePath, outputStream.ToArray());

                return File(outputStream.ToArray(), "application/pdf", fileName);
            }
        }
        #endregion 4.1.1.2.15.สัญญาจ้างทำของ CWA

        #region 4.1.1.2.14.สัญญาจ้างผู้เชี่ยวชาญรายบุคคลหรือจ้างบริษัทที่ปรึกษา ร.317-60 CTR31760
        public async Task<IActionResult> OnGetWordContact_CTR31760(string ContractId ="1")
        {
            var wordBytes =await _ConsultantService.OnGetWordContact_ConsultantService(ContractId);
            return File(wordBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "สัญญาจ้างที่ปรึกษา.docx");
        }
        public async Task OnGetWordContact_CTR31760_PDF(string ContractId = "1")
        {
            var wordBytes = await _ConsultantService.OnGetWordContact_ConsultantService_ToPDF(ContractId, "CTR31760");
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "CTR31760");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var filePath = Path.Combine(folderPath, "CTR31760_" + ContractId + ".pdf");

            // Delete the file if it already exists
            if (System.IO.File.Exists(filePath))
            {
                System.IO.File.Delete(filePath);
            }
            await System.IO.File.WriteAllBytesAsync(filePath, wordBytes);

            // return File(wordBytes, "application/pdf", "สัญญาซื้อขาย.pdf");
        }
        public async Task<IActionResult> OnGetWordContact_CTR31760_PDF_Preview(string ContractId = "1")
        {
            var wordBytes = await _ConsultantService.OnGetWordContact_ConsultantService_ToPDF(ContractId, "CTR31760");
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "CTR31760");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var filePath = Path.Combine(folderPath, "CTR31760_" + ContractId + "_Preview.pdf");

            // Set your desired password here
            string userPassword = _configuration["Password:PaswordPDF"];

            // Load the PDF from the byte array
            using (var inputStream = new MemoryStream(wordBytes))
            using (var outputStream = new MemoryStream())
            {
                var document = PdfReader.Open(inputStream, PdfDocumentOpenMode.Modify);

                // Set up security settings
                var securitySettings = document.SecuritySettings;
                securitySettings.UserPassword = userPassword;
                securitySettings.OwnerPassword = userPassword;
                securitySettings.PermitPrint = true;
                securitySettings.PermitModifyDocument = false;
                securitySettings.PermitExtractContent = false;
                securitySettings.PermitAnnotations = false;

                document.Save(outputStream);

                // Optionally save to disk
                // await System.IO.File.WriteAllBytesAsync(filePath, outputStream.ToArray());

                // Return the password-protected PDF to the user
                return File(outputStream.ToArray(), "application/pdf", "CTR31760_" + ContractId + "_Preview.pdf");
            }
        }
        #endregion 4.1.1.2.14.สัญญาจ้างที่ปรึกษา CTR31760

        #region 4.1.1.2.13.สัญญาเช่าเครื่องถ่ายเอกสาร ร.314-60 PML31460

        public async Task<IActionResult> OnGetWordContact_PML31460(string ContractId ="1")
        {
            var wordBytes = await _wordEContract_LoanPrinterService.OnGetWordContact_LoanPrinter(ContractId);
            return File(wordBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "สัญญาเช่าเครื่องถ่ายเอกสาร ร.314-60.docx");
        }
        public async Task OnGetWordContact_PML31460_PDF(string ContractId = "1")
        {
            var wordBytes = await _wordEContract_LoanPrinterService.OnGetWordContact_LoanPrinter_ToPDF(ContractId, "PML31460");
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "PML31460");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var filePath = Path.Combine(folderPath, "PML31460_" + ContractId + ".pdf");

            // Delete the file if it already exists
            if (System.IO.File.Exists(filePath))
            {
                System.IO.File.Delete(filePath);
            }
            await System.IO.File.WriteAllBytesAsync(filePath, wordBytes);

            // return File(wordBytes, "application/pdf", "สัญญาซื้อขาย.pdf");
        }
        public async Task<IActionResult> OnGetWordContact_PML31460_PDF_Preview(string ContractId = "1")
        {
            var wordBytes = await _wordEContract_LoanPrinterService.OnGetWordContact_LoanPrinter_ToPDF(ContractId, "PML31460");
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "PML31460");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var fileName = $"PML31460_{ContractId}_Preview.pdf";
            var filePath = Path.Combine(folderPath, fileName);

            // Get password from appsettings.json
            string userPassword = _configuration["Password:PaswordPDF"];

            using (var inputStream = new MemoryStream(wordBytes))
            using (var outputStream = new MemoryStream())
            {
                var document = PdfReader.Open(inputStream, PdfDocumentOpenMode.Modify);

                var securitySettings = document.SecuritySettings;
                securitySettings.UserPassword = userPassword;
                securitySettings.OwnerPassword = userPassword;
                securitySettings.PermitPrint = true;
                securitySettings.PermitModifyDocument = false;
                securitySettings.PermitExtractContent = false;
                securitySettings.PermitAnnotations = false;

                document.Save(outputStream);

                // Optionally save to disk
                // await System.IO.File.WriteAllBytesAsync(filePath, outputStream.ToArray());

                return File(outputStream.ToArray(), "application/pdf", fileName);
            }
        }
        #endregion 4.1.1.2.13.สัญญาเช่าเครื่องถ่ายเอกสาร ร.314-60 PML31460

        #region 4.1.1.2.12.สัญญาจ้างบริการบำรุงรักษาและซ่อมแซมแก้ไขคอมพิวเตอร์ร.310-60 SMC31060
        public async Task<IActionResult> OnGetWordContact_SMC31060(string ContractId ="1")
        {
            var wordBytes = await _maintenanceComputerService.OnGetWordContact_MaintenanceComputer(ContractId);
            return File(wordBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "สัญญาจ้างบริการบำรุงรักษาและซ่อมแซมแก้ไขคอมพิวเตอร์ร.310-60.docx");
        }
        public async Task OnGetWordContact_SMC31060_PDF(string ContractId = "1")
        {
            var wordBytes = await _maintenanceComputerService.OnGetWordContact_MaintenanceComputer_ToPDF(ContractId,"SMC31060");
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "SMC31060");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var filePath = Path.Combine(folderPath, "SMC31060_" + ContractId + ".pdf");

            // Delete the file if it already exists
            if (System.IO.File.Exists(filePath))
            {
                System.IO.File.Delete(filePath);
            }
            await System.IO.File.WriteAllBytesAsync(filePath, wordBytes);

            // return File(wordBytes, "application/pdf", "สัญญาซื้อขาย.pdf");
        }

        public async Task<IActionResult> OnGetWordContact_SMC31060_PDF_Preview(string ContractId = "1")
        {
            var wordBytes = await _maintenanceComputerService.OnGetWordContact_MaintenanceComputer_ToPDF(ContractId, "SMC31060");
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "SMC31060");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var fileName = $"SMC31060_{ContractId}_Preview.pdf";
            var filePath = Path.Combine(folderPath, fileName);

            // Get password from appsettings.json
            string userPassword = _configuration["Password:PaswordPDF"];

            using (var inputStream = new MemoryStream(wordBytes))
            using (var outputStream = new MemoryStream())
            {
                var document = PdfSharpCore.Pdf.IO.PdfReader.Open(inputStream, PdfSharpCore.Pdf.IO.PdfDocumentOpenMode.Modify);

                var securitySettings = document.SecuritySettings;
                securitySettings.UserPassword = userPassword;
                securitySettings.OwnerPassword = userPassword;
                securitySettings.PermitPrint = true;
                securitySettings.PermitModifyDocument = false;
                securitySettings.PermitExtractContent = false;
                securitySettings.PermitAnnotations = false;

                document.Save(outputStream);

                // Optionally save to disk
                // await System.IO.File.WriteAllBytesAsync(filePath, outputStream.ToArray());

                return File(outputStream.ToArray(), "application/pdf", fileName);
            }
        }
        #endregion 4.1.1.2.12.สัญญาจ้างบริการบำรุงรักษาและซ่อมแซมแก้ไขคอมพิวเตอร์ร.310-60 SMC31060

        #region 4.1.1.2.11.สัญญาเช่าคอมพิวเตอร์ ร.309-60 CLA30960
        public async Task<IActionResult> OnGetWordContact_CLA30960(string ContractId ="1")
        {
            var wordBytes =await _LoanComputerService.OnGetWordContact_LoanComputer(ContractId);
            return File(wordBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "สัญญาเช่าคอมพิวเตอร์ ร.309-60.docx");
        }
        public async Task OnGetWordContact_CLA30960_PDF(string ContractId = "1")
        {
            var wordBytes = await _LoanComputerService.OnGetWordContact_LoanComputer_ToPDF(ContractId, "CLA30960");
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "CLA30960");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var filePath = Path.Combine(folderPath, "CLA30960_" + ContractId + ".pdf");

            // Delete the file if it already exists
            if (System.IO.File.Exists(filePath))
            {
                System.IO.File.Delete(filePath);
            }
            await System.IO.File.WriteAllBytesAsync(filePath, wordBytes);

            // return File(wordBytes, "application/pdf", "สัญญาซื้อขาย.pdf");
        }

        public async Task<IActionResult> OnGetWordContact_CLA30960_PDF_Preview(string ContractId = "1")
        {
            var wordBytes = await _LoanComputerService.OnGetWordContact_LoanComputer_ToPDF(ContractId, "CLA30960");
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "CLA30960");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var fileName = $"CLA30960_{ContractId}_Preview.pdf";
            var filePath = Path.Combine(folderPath, fileName);

            // Get password from appsettings.json
            string userPassword = _configuration["Password:PaswordPDF"];

            using (var inputStream = new MemoryStream(wordBytes))
            using (var outputStream = new MemoryStream())
            {
                var document = PdfSharpCore.Pdf.IO.PdfReader.Open(inputStream, PdfSharpCore.Pdf.IO.PdfDocumentOpenMode.Modify);

                var securitySettings = document.SecuritySettings;
                securitySettings.UserPassword = userPassword;
                securitySettings.OwnerPassword = userPassword;
                securitySettings.PermitPrint = true;
                securitySettings.PermitModifyDocument = false;
                securitySettings.PermitExtractContent = false;
                securitySettings.PermitAnnotations = false;

                document.Save(outputStream);

                // Optionally save to disk
                // await System.IO.File.WriteAllBytesAsync(filePath, outputStream.ToArray());

                return File(outputStream.ToArray(), "application/pdf", fileName);
            }
        }
        #endregion 4.1.1.2.11.สัญญาเช่าคอมพิวเตอร์ ร.309-60 CLA30960

        #region 4.1.1.2.10.สัญญาซื้อขายและอนุญาตให้ใช้สิทธิในโปรแกรมคอมพิวเตอร์ ร.308-60 SLA30860
        public async Task<IActionResult> OnGetWordContact_SLA30860(string ContractId ="1")
        {
            var wordBytes = await _BuyAgreeProgram.OnGetWordContact_BuyAgreeProgram(ContractId);
            return File(wordBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "สัญญาซื้อขายและอนุญาตให้ใช้สิทธิในโปรแกรมคอมพิวเตอร์.docx");
        }
        public async Task OnGetWordContact_SLA30860_PDF(string ContractId = "1")
        {
            var wordBytes = await _BuyAgreeProgram.OnGetWordContact_BuyAgreeProgram_ToPDF(ContractId, "SLA30860");
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "SLA30860");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var filePath = Path.Combine(folderPath, "SLA30860_" + ContractId + ".pdf");

            // Delete the file if it already exists
            if (System.IO.File.Exists(filePath))
            {
                System.IO.File.Delete(filePath);
            }
            await System.IO.File.WriteAllBytesAsync(filePath, wordBytes);

            // return File(wordBytes, "application/pdf", "สัญญาซื้อขาย.pdf");
        }

        public async Task<IActionResult> OnGetWordContact_SLA30860_PDF_Preview(string ContractId = "1")
        {
            var wordBytes = await _BuyAgreeProgram.OnGetWordContact_BuyAgreeProgram_ToPDF(ContractId, "SLA30860");
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "SLA30860");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var fileName = $"SLA30860_{ContractId}_Preview.pdf";
            var filePath = Path.Combine(folderPath, fileName);

            // Get password from appsettings.json
            string userPassword = _configuration["Password:PaswordPDF"];

            using (var inputStream = new MemoryStream(wordBytes))
            using (var outputStream = new MemoryStream())
            {
                var document = PdfSharpCore.Pdf.IO.PdfReader.Open(inputStream, PdfSharpCore.Pdf.IO.PdfDocumentOpenMode.Modify);

                var securitySettings = document.SecuritySettings;
                securitySettings.UserPassword = userPassword;
                securitySettings.OwnerPassword = userPassword;
                securitySettings.PermitPrint = true;
                securitySettings.PermitModifyDocument = false;
                securitySettings.PermitExtractContent = false;
                securitySettings.PermitAnnotations = false;

                document.Save(outputStream);

                // Optionally save to disk
                // await System.IO.File.WriteAllBytesAsync(filePath, outputStream.ToArray());

                return File(outputStream.ToArray(), "application/pdf", fileName);
            }
        }
        #endregion 4.1.1.2.10.สัญญาซื้อขายและอนุญาตให้ใช้สิทธิในโปรแกรมคอมพิวเตอร์ ร.308-60 SLA30860

        #region 4.1.1.2.9.สัญญาซื้อขายคอมพิวเตอร์ CPA
        public async Task<IActionResult> OnGetWordContact_CPA(string ContractId ="13")
        {
            var wordBytes = await _BuyOrSellComputerService.OnGetWordContact_BuyOrSellComputerService(ContractId);
            return File(wordBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "สัญญาซื้อขายคอมพิวเตอร์.docx");
        }
        public async Task OnGetWordContact_CPA_PDF(string ContractId = "14")
        {
            var wordBytes = await _BuyOrSellComputerService.OnGetWordContact_BuyOrSellComputerService_ToPDF(ContractId);

            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "CPA");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var filePath = Path.Combine(folderPath, "CPA_" + ContractId + ".pdf");

            // Delete the file if it already exists
            if (System.IO.File.Exists(filePath))
            {
                System.IO.File.Delete(filePath);
            }
            await System.IO.File.WriteAllBytesAsync(filePath, wordBytes);

            //return File(wordBytes, "application/pdf", "CPA_" + ContractId + ".pdf");
        }

        public async Task<IActionResult> OnGetWordContact_CPA_PDF_Preview(string ContractId = "14")
        {
            var wordBytes = await _BuyOrSellComputerService.OnGetWordContact_BuyOrSellComputerService_ToPDF(ContractId);

            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "CPA");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var fileName = $"CPA_{ContractId}_Preview.pdf";
            var filePath = Path.Combine(folderPath, fileName);

            // Get password from appsettings.json
            string userPassword = _configuration["Password:PaswordPDF"];

            using (var inputStream = new MemoryStream(wordBytes))
            using (var outputStream = new MemoryStream())
            {
                var document = PdfSharpCore.Pdf.IO.PdfReader.Open(inputStream, PdfSharpCore.Pdf.IO.PdfDocumentOpenMode.Modify);

                var securitySettings = document.SecuritySettings;
                securitySettings.UserPassword = userPassword;
                securitySettings.OwnerPassword = userPassword;
                securitySettings.PermitPrint = true;
                securitySettings.PermitModifyDocument = false;
                securitySettings.PermitExtractContent = false;
                securitySettings.PermitAnnotations = false;

                document.Save(outputStream);

                // Optionally save to disk
                // await System.IO.File.WriteAllBytesAsync(filePath, outputStream.ToArray());

                return File(outputStream.ToArray(), "application/pdf", fileName);
            }
        }
        #endregion 4.1.1.2.9.สัญญาซื้อขายคอมพิวเตอร์ CPA

        #region 4.1.1.2.8.สัญญาซื้อขาย ร.305-60 SPA30560

        public async Task<IActionResult> OnGetWordContact_SPA30560(string ContractId = "1")
        {
            var wordBytes = await _BuyOrSellService.OnGetWordContact_BuyOrSellService(ContractId);
            return File(wordBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "สัญญาซื้อขาย.docx");
        }

        public async Task OnGetWordContact_SPA30560_PDF(string ContractId = "4")
        {
            var wordBytes = await _BuyOrSellService.OnGetWordContact_BuyOrSellService_ToPDF(ContractId, "SPA30560");
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "SPA30560");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var filePath = Path.Combine(folderPath, "SPA30560_" + ContractId + ".pdf");

            // Delete the file if it already exists
            if (System.IO.File.Exists(filePath))
            {
                System.IO.File.Delete(filePath);
            }
            await System.IO.File.WriteAllBytesAsync(filePath, wordBytes);

            // return File(wordBytes, "application/pdf", "สัญญาซื้อขาย.pdf");
        }
        public async Task<IActionResult> OnGetWordContact_SPA30560_PDF_Preview(string ContractId = "4")
        {
            var wordBytes = await _BuyOrSellService.OnGetWordContact_BuyOrSellService_ToPDF(ContractId, "SPA30560");
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "SPA30560");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var fileName = $"SPA30560_{ContractId}_Preview.pdf";
            var filePath = Path.Combine(folderPath, fileName);

            // Get password from appsettings.json
            string userPassword = _configuration["Password:PaswordPDF"];

            using (var inputStream = new MemoryStream(wordBytes))
            using (var outputStream = new MemoryStream())
            {
                var document = PdfSharpCore.Pdf.IO.PdfReader.Open(inputStream, PdfSharpCore.Pdf.IO.PdfDocumentOpenMode.Modify);

                var securitySettings = document.SecuritySettings;
                securitySettings.UserPassword = userPassword;
                securitySettings.OwnerPassword = userPassword;
                securitySettings.PermitPrint = true;
                securitySettings.PermitModifyDocument = false;
                securitySettings.PermitExtractContent = false;
                securitySettings.PermitAnnotations = false;

                document.Save(outputStream);

                // Optionally save to disk
                // await System.IO.File.WriteAllBytesAsync(filePath, outputStream.ToArray());

                return File(outputStream.ToArray(), "application/pdf", fileName);
            }
        }
        #endregion 4.1.1.2.8.สัญญาซื้อขาย ร.305-60 SPA30560

        #region 4.1.1.2.7.สัญญาการรักษาข้อมูลที่เป็นความลับ NDA
        public async Task<IActionResult> OnGetWordContact_NDA(string ContractId ="1")
        {
            var wordBytes = await _DataSecretService.OnGetWordContact_DataSecretService(ContractId);
            return File(wordBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "สัญญาการรักษาข้อมูลที่เป็นความลับ.docx");
        }
        public async Task OnGetWordContact_NDA_PDF(string ContractId = "1")
        {
            var wordBytes = await _DataSecretService.OnGetWordContact_DataSecretService_ToPDF(ContractId,"NDA");
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "NDA");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var filePath = Path.Combine(folderPath, "NDA_" + ContractId + ".pdf");

            // Delete the file if it already exists
            if (System.IO.File.Exists(filePath))
            {
                System.IO.File.Delete(filePath);
            }
            await System.IO.File.WriteAllBytesAsync(filePath, wordBytes);

            // return File(wordBytes, "application/pdf", "บันทึกข้อตกลงการแบ่งปันข้อมูลส่วนบุคคล.pdf");
        }
        public async Task<IActionResult> OnGetWordContact_NDA_PDF_Preview(string ContractId = "1")
        {
            var wordBytes = await _DataSecretService.OnGetWordContact_DataSecretService_ToPDF(ContractId, "NDA");
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "NDA");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var fileName = $"NDA_{ContractId}_Preview.pdf";
            var filePath = Path.Combine(folderPath, fileName);

            // Get password from appsettings.json
            string userPassword = _configuration["Password:PaswordPDF"];

            using (var inputStream = new MemoryStream(wordBytes))
            using (var outputStream = new MemoryStream())
            {
                var document = PdfSharpCore.Pdf.IO.PdfReader.Open(inputStream, PdfSharpCore.Pdf.IO.PdfDocumentOpenMode.Modify);

                var securitySettings = document.SecuritySettings;
                securitySettings.UserPassword = userPassword;
                securitySettings.OwnerPassword = userPassword;
                securitySettings.PermitPrint = true;
                securitySettings.PermitModifyDocument = false;
                securitySettings.PermitExtractContent = false;
                securitySettings.PermitAnnotations = false;

                document.Save(outputStream);

                // Optionally save to disk
                // await System.IO.File.WriteAllBytesAsync(filePath, outputStream.ToArray());

                return File(outputStream.ToArray(), "application/pdf", fileName);
            }
        }

        #endregion 4.1.1.2.7.สัญญาการรักษาข้อมูลที่เป็นความลับ NDA

        #region 4.1.1.2.6.บันทึกข้อตกลงการแบ่งปันข้อมูลส่วนบุคคล PDSA
        public async Task<IActionResult> OnGetWordContact_PDSA(string ContractId ="3")
        {
            var wordBytes = await _DataPersonalService.OnGetWordContact_DataPersonalService(ContractId);
            return File(wordBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "บันทึกข้อตกลงการแบ่งปันข้อมูลส่วนบุคคล.docx");
        }
        public async Task OnGetWordContact_PDSA_PDF(string ContractId = "3")
        {
            var wordBytes = await _DataPersonalService.OnGetWordContact_DataPersonalService_ToPDF(ContractId,"PDSA");
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "PDSA");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var filePath = Path.Combine(folderPath, "PDSA_" + ContractId + ".pdf");

            // Delete the file if it already exists
            if (System.IO.File.Exists(filePath))
            {
                System.IO.File.Delete(filePath);
            }
            await System.IO.File.WriteAllBytesAsync(filePath, wordBytes);

           // return File(wordBytes, "application/pdf", "บันทึกข้อตกลงการแบ่งปันข้อมูลส่วนบุคคล.pdf");
        }

        public async Task<IActionResult> OnGetWordContact_PDSA_PDF_Preview(string ContractId = "3")
        {
            var wordBytes = await _DataPersonalService.OnGetWordContact_DataPersonalService_ToPDF(ContractId, "PDSA");
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "PDSA");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var fileName = $"PDSA_{ContractId}_Preview.pdf";
            var filePath = Path.Combine(folderPath, fileName);

            // Get password from appsettings.json
            string userPassword = _configuration["Password:PaswordPDF"];

            using (var inputStream = new MemoryStream(wordBytes))
            using (var outputStream = new MemoryStream())
            {
                var document = PdfSharpCore.Pdf.IO.PdfReader.Open(inputStream, PdfSharpCore.Pdf.IO.PdfDocumentOpenMode.Modify);

                var securitySettings = document.SecuritySettings;
                securitySettings.UserPassword = userPassword;
                securitySettings.OwnerPassword = userPassword;
                securitySettings.PermitPrint = true;
                securitySettings.PermitModifyDocument = false;
                securitySettings.PermitExtractContent = false;
                securitySettings.PermitAnnotations = false;

                document.Save(outputStream);

                // Optionally save to disk
                // await System.IO.File.WriteAllBytesAsync(filePath, outputStream.ToArray());

                return File(outputStream.ToArray(), "application/pdf", fileName);
            }
        }
        # endregion 4.1.1.2.6.บันทึกข้อตกลงการแบ่งปันข้อมูลส่วนบุคคล PDSA

        #region 4.1.1.2.5.บันทึกข้อตกลงการเป็นผู้ควบคุมข้อมูลส่วนบุคคลร่วมตัวอย่างหน้าจอ JDCA

        public async Task<IActionResult> OnGetWordContact_JDCA(string ContractId="1")
        {
            var wordBytes = await _ControlDataService.OnGetWordContact_ControlDataService(ContractId);
            return File(wordBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "บันทึกข้อตกลงการเป็นผู้ควบคุมข้อมูลส่วนบุคคลร่วม.docx");

        }
        public async Task OnGetWordContact_JDCA_PDF(string ContractId = "1")
        {
            var wordBytes = await _ControlDataService.OnGetWordContact_ControlDataServiceHtmlToPdf(ContractId,"JDCA");
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "JDCA");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var filePath = Path.Combine(folderPath, "JDCA_" + ContractId + ".pdf");

            // Delete the file if it already exists
            if (System.IO.File.Exists(filePath))
            {
                System.IO.File.Delete(filePath);
            }
            await System.IO.File.WriteAllBytesAsync(filePath, wordBytes);
            // return File(wordBytes, "application/pdf", "บันทึกข้อตกลงการเป็นผู้ควบคุมข้อมูลส่วนบุคคลร่วม.pdf");
        }
        public async Task<IActionResult> OnGetWordContact_JDCA_PDF_Preview(string ContractId = "1")
        {
            var wordBytes = await _ControlDataService.OnGetWordContact_ControlDataServiceHtmlToPdf(ContractId, "JDCA");
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "JDCA");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var fileName = $"JDCA_{ContractId}_Preview.pdf";
            var filePath = Path.Combine(folderPath, fileName);

            // Get password from appsettings.json
            string userPassword = _configuration["Password:PaswordPDF"];

            using (var inputStream = new MemoryStream(wordBytes))
            using (var outputStream = new MemoryStream())
            {
                var document = PdfSharpCore.Pdf.IO.PdfReader.Open(inputStream, PdfSharpCore.Pdf.IO.PdfDocumentOpenMode.Modify);

                var securitySettings = document.SecuritySettings;
                securitySettings.UserPassword = userPassword;
                securitySettings.OwnerPassword = userPassword;
                securitySettings.PermitPrint = true;
                securitySettings.PermitModifyDocument = false;
                securitySettings.PermitExtractContent = false;
                securitySettings.PermitAnnotations = false;

                document.Save(outputStream);

                // Optionally save to disk
                // await System.IO.File.WriteAllBytesAsync(filePath, outputStream.ToArray());

                return File(outputStream.ToArray(), "application/pdf", fileName);
            }
        }
        #endregion 4.1.1.2.5.บันทึกข้อตกลงการเป็นผู้ควบคุมข้อมูลส่วนบุคคลร่วมตัวอย่างหน้าจอ JDCA


        #region 4.1.1.2.4.บันทึกข้อตกลงการประมวลผลข้อมูลส่วนบุคคล PDPA
        public async Task<IActionResult> OnGetWordContact_PDPA(string ContractId = "1")
        {
            var wordBytes = await _PersernalProcessService.OnGetWordContact_PersernalProcessService(ContractId);
            return File(wordBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "บันทึกข้อตกลงการประมวลผลข้อมูลส่วนบุคคล.docx");
        }
        public async Task OnGetWordContact_PDPA_PDF(string ContractId = "4")
        {
            var wordBytes = await _PersernalProcessService.OnGetWordContact_PersernalProcessService_HtmlToPDF(ContractId, "PDPA");
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "PDPA");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var filePath = Path.Combine(folderPath, "PDPA_" + ContractId + ".pdf");

            // Delete the file if it already exists
            if (System.IO.File.Exists(filePath))
            {
                System.IO.File.Delete(filePath);
            }
            await System.IO.File.WriteAllBytesAsync(filePath, wordBytes);
            // return File(wordBytes, "application/pdf", "บันทึกข้อตกลงการประมวลผลข้อมูลส่วนบุคคล.pdf");
        }

        public async Task<IActionResult> OnGetWordContact_PDPA_PDF_Preview(string ContractId = "4")
        {
            var wordBytes = await _PersernalProcessService.OnGetWordContact_PersernalProcessService_HtmlToPDF(ContractId, "PDPA");
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "PDPA");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var fileName = $"PDPA_{ContractId}_Preview.pdf";
            var filePath = Path.Combine(folderPath, fileName);

            // Get password from appsettings.json
            string userPassword = _configuration["Password:PaswordPDF"];

            using (var inputStream = new MemoryStream(wordBytes))
            using (var outputStream = new MemoryStream())
            {
                var document = PdfSharpCore.Pdf.IO.PdfReader.Open(inputStream, PdfSharpCore.Pdf.IO.PdfDocumentOpenMode.Modify);

                var securitySettings = document.SecuritySettings;
                securitySettings.UserPassword = userPassword;
                securitySettings.OwnerPassword = userPassword;
                securitySettings.PermitPrint = true;
                securitySettings.PermitModifyDocument = false;
                securitySettings.PermitExtractContent = false;
                securitySettings.PermitAnnotations = false;

                document.Save(outputStream);

                // Optionally save to disk
                // await System.IO.File.WriteAllBytesAsync(filePath, outputStream.ToArray());

                return File(outputStream.ToArray(), "application/pdf", fileName);
            }
        }
        #endregion 4.1.1.2.4.บันทึกข้อตกลงการประมวลผลข้อมูลส่วนบุคคล PDPA

        #region 4.1.1.2.3.บันทึกข้อตกลงความร่วมมือ MOU
        public async Task<IActionResult> OnGetWordContact_MOU(string ContractId = "5")
        {
            var wordBytes = await _MemorandumService.OnGetWordContact_MemorandumService(ContractId);
            return File(wordBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "บันทึกข้อตกลงความร่วมมือ.docx");
        }
        public async Task OnGetWordContact_MOU_PDF(string ContractId = "7")
        {
            var wordBytes = await _MemorandumService.OnGetWordContact_MemorandumService_HtmlToPDF(ContractId,"MOU");
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "MOU");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var filePath = Path.Combine(folderPath, "MOU_" + ContractId + ".pdf");

            // Delete the file if it already exists
            if (System.IO.File.Exists(filePath))
            {
                System.IO.File.Delete(filePath);
            }
            await System.IO.File.WriteAllBytesAsync(filePath, wordBytes);
            //   return File(wordBytes, "application/pdf", "MOU_" + ContractId + ".pdf");
        }

        public async Task<IActionResult> OnGetWordContact_MOU_PDF_Preview(string ContractId = "7")
        {
            var wordBytes = await _MemorandumService.OnGetWordContact_MemorandumService_HtmlToPDF(ContractId, "MOU");
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "MOU");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var fileName = $"MOU_{ContractId}_Preview.pdf";
            var filePath = Path.Combine(folderPath, fileName);

            // Get password from appsettings.json
            string userPassword = _configuration["Password:PaswordPDF"];

            using (var inputStream = new MemoryStream(wordBytes))
            using (var outputStream = new MemoryStream())
            {
                var document = PdfSharpCore.Pdf.IO.PdfReader.Open(inputStream, PdfSharpCore.Pdf.IO.PdfDocumentOpenMode.Modify);

                var securitySettings = document.SecuritySettings;
                securitySettings.UserPassword = userPassword;
                securitySettings.OwnerPassword = userPassword;
                securitySettings.PermitPrint = true;
                securitySettings.PermitModifyDocument = false;
                securitySettings.PermitExtractContent = false;
                securitySettings.PermitAnnotations = false;

                document.Save(outputStream);

                // Optionally save to disk
                // await System.IO.File.WriteAllBytesAsync(filePath, outputStream.ToArray());

                return File(outputStream.ToArray(), "application/pdf", fileName);
            }
        }
        #endregion  4.1.1.2.3.บันทึกข้อตกลงความร่วมมือ MOU

        #region 4.1.1.2.2.สัญญารับเงินอุดหนุน GA
        public async Task<IActionResult> OnGetWordContact_GA(string ContractId = "1")
        {
            var wordBytes = await _SupportSMEsService.OnGetWordContact_SupportSMEsService(ContractId);
            return File(wordBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "บันทึกข้อตกลงความร่วมมือในการสนับสนุน SMEs.docx");
        }
        public async Task OnGetWordContact_GA_PDF(string ContractId = "1")
        {
            var pdfBytes = await _SupportSMEsService.OnGetWordContact_SupportSMEsService_HtmlToPDF(ContractId,"GA");

            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "GA");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var filePath = Path.Combine(folderPath, "GA_" + ContractId + ".pdf");

            // Delete the file if it already exists
            if (System.IO.File.Exists(filePath))
            {
                System.IO.File.Delete(filePath);
            }
            await System.IO.File.WriteAllBytesAsync(filePath, pdfBytes);
           // return File(pdfBytes, "application/pdf", "GA_"+ContractId+".pdf");
        }

        public async Task<IActionResult> OnGetWordContact_GA_PDF_Preview(string ContractId = "1")
        {
            var pdfBytes = await _SupportSMEsService.OnGetWordContact_SupportSMEsService_HtmlToPDF(ContractId, "GA");
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "GA");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var fileName = $"GA_{ContractId}_Preview.pdf";
            var filePath = Path.Combine(folderPath, fileName);

            // Get password from appsettings.json
            string userPassword = _configuration["Password:PaswordPDF"];

            using (var inputStream = new MemoryStream(pdfBytes))
            using (var outputStream = new MemoryStream())
            {
                var document = PdfSharpCore.Pdf.IO.PdfReader.Open(inputStream, PdfSharpCore.Pdf.IO.PdfDocumentOpenMode.Modify);

                var securitySettings = document.SecuritySettings;
                securitySettings.UserPassword = userPassword;
                securitySettings.OwnerPassword = userPassword;
                securitySettings.PermitPrint = true;
                securitySettings.PermitModifyDocument = false;
                securitySettings.PermitExtractContent = false;
                securitySettings.PermitAnnotations = false;

                document.Save(outputStream);

                // Optionally save to disk
                // await System.IO.File.WriteAllBytesAsync(filePath, outputStream.ToArray());

                return File(outputStream.ToArray(), "application/pdf", fileName);
            }
        }
        #endregion 4.1.1.2.2.สัญญารับเงินอุดหนุน GA

        #region 4.1.1.2.1.สัญญาร่วมดำเนินการ JOA
        public async Task<IActionResult> OnGetWordContact_JOA(string ContractId="32")
        {
            var wordBytes = await _JointOperationService.OnGetWordContact_JointOperationService(ContractId);
            return File(wordBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "สัญญาร่วมดำเนินการ.docx");
        }

        public async Task OnGetWordContact_JOA_PDF(string ContractId = "70")
        {
            var wordBytes = await _JointOperationService.OnGetWordContact_JointOperationServiceHtmlToPDF(ContractId);

            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "JOA");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var filePath = Path.Combine(folderPath, "JOA_" + ContractId + ".pdf");

            // Delete the file if it already exists
            if (System.IO.File.Exists(filePath))
            {
                System.IO.File.Delete(filePath);
            }
            await System.IO.File.WriteAllBytesAsync(filePath, wordBytes);

          //  return File(wordBytes, "application/pdf", "JOA_" + ContractId + ".pdf");
        }

        public async Task<IActionResult> OnGetWordContact_JOA_PDF_Preview(string ContractId = "70")
        {
            var wordBytes = await _JointOperationService.OnGetWordContact_JointOperationServiceHtmlToPDF(ContractId);
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "JOA");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var fileName = $"JOA_{ContractId}_Preview.pdf";
            var filePath = Path.Combine(folderPath, fileName);

            // Get password from appsettings.json
            string userPassword = _configuration["Password:PaswordPDF"];

            using (var inputStream = new MemoryStream(wordBytes))
            using (var outputStream = new MemoryStream())
            {
                var document = PdfSharpCore.Pdf.IO.PdfReader.Open(inputStream, PdfSharpCore.Pdf.IO.PdfDocumentOpenMode.Modify);

                var securitySettings = document.SecuritySettings;
                securitySettings.UserPassword = userPassword;
                securitySettings.OwnerPassword = userPassword;
                securitySettings.PermitPrint = true;
                securitySettings.PermitModifyDocument = false;
                securitySettings.PermitExtractContent = false;
                securitySettings.PermitAnnotations = false;

                document.Save(outputStream);

                // Optionally save to disk
                // await System.IO.File.WriteAllBytesAsync(filePath, outputStream.ToArray());

                return File(outputStream.ToArray(), "application/pdf", fileName);
            }
        }
        #endregion 4.1.1.2.1.สัญญาร่วมดำเนินการ JOA



        #region Test Header logo
        public async Task<IActionResult> OnGetWordContact_TestLogo(string PDPAid = "1")
        {
            var wordBytes = await _Test_HeaderLOGOService.OnGetWordContact_PersernalProcessService(PDPAid);
            return File(wordBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "บันทึกข้อตกลงการประมวลผลข้อมูลส่วนบุคคล.docx");
        }

        #endregion Test Header logo


        #region Word to PDF using Interop
        public async Task<IActionResult> OnGetWordtoPDF(string ContractId = "1")
        {
            var pdfBytes = await _SupportSMEsService.OnGetWordContact_SupportSMEsService_HtmlToPDF(ContractId,"68");
            return File(pdfBytes, "application/pdf", "บันทึกข้อตกลงการแบ่งปันข้อมูลส่วนบุคคล.pdf");
        }

        #endregion Word to PDF using Interop
    }
}
