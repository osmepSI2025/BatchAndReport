using BatchAndReport.DAO;
using DinkToPdf.Contracts;
using ICSharpCode.SharpZipLib.Zip;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.Data.SqlClient;
using PuppeteerSharp;
using Spire.Doc;
using System.IO.Compression;
using System.Runtime.InteropServices;
using Document = Spire.Doc.Document;
namespace BatchAndReport.Pages.Report
{
    public class ExportModel : PageModel
    {
        private readonly SmeDAO _smeDao;
        private readonly EContractDAO _eContractDao;
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
        private readonly WordEContract_MIWService _MIWService;
        private readonly WordEContract_MemorandumInWritingService _MemorandumInWritingService;
        private readonly IConfiguration _configuration;
        private readonly IConverter _pdfConverter; // เพิ่ม DI สำหรับ PDF Converter
        private readonly WordEContract_AMJOAService _AMJOAService;
        public ExportModel(SmeDAO smeDao, EContractDAO eContractDao, WordEContract_AllowanceService allowanceService
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
            , WordEContract_MemorandumInWritingService MemorandumInWritingService
            , IConfiguration configuration // <-- add this
            , WordEContract_MIWService MIWService
              , IConverter pdfConverter
            , WordEContract_AMJOAService AMJOAService
            )
        {
            _smeDao = smeDao;
            _eContractDao = eContractDao;
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
            this._MemorandumInWritingService = MemorandumInWritingService;
            _configuration = configuration; // <-- initialize the configuration
            this._MIWService = MIWService;
            _pdfConverter = pdfConverter;
            _AMJOAService = AMJOAService;
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
        public async Task<IActionResult> OnGetWordContact_EC(string ContractId = "7")
        {
            var wordBytes = await _HireEmployee.OnGetWordContact_HireEmployee(ContractId);
            return File(wordBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "สัญญาจ้างลูกจ้าง.docx");

        }

        public async Task OnGetWordContact_EC_PDF(string ContractId = "55")
        {
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
            var filePathView = Path.Combine(folderPath, "EC_" + ContractId + "_1.pdf");

            // Delete the file if it already exists
            if (System.IO.File.Exists(filePathView))
            {
                System.IO.File.Delete(filePathView);
            }
            var htmlContent = await _HireEmployee.OnGetWordContact_HireEmployee_ToPDF(ContractId, "EC");
            await new BrowserFetcher().DownloadAsync();
            await using var browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = true });
            await using var page = await browser.NewPageAsync();

            await page.SetContentAsync(htmlContent);

            var pdfOptions = new PdfOptions
            {
                Format = PuppeteerSharp.Media.PaperFormat.A4,

                Landscape = false,
                MarginOptions = new PuppeteerSharp.Media.MarginOptions
                {
                    Top = "20mm",
                    Bottom = "20mm",
                    Left = "20mm",
                    Right = "20mm"
                },
                PrintBackground = true

            };

            var pdfBytes = await page.PdfDataAsync(pdfOptions);


          
            await System.IO.File.WriteAllBytesAsync(filePath, pdfBytes);
            await System.IO.File.WriteAllBytesAsync(filePathView, pdfBytes);

            // return File(wordBytes, "application/pdf", "สัญญาซื้อขาย.pdf");
        }
        public async Task<IActionResult> OnGetWordContact_EC_PDF_Preview(string ContractId = "55", string Name = "สมใจ ทดสอบ")
        {
            var htmlContent = await _HireEmployee.OnGetWordContact_HireEmployee_ToPDF(ContractId, "EC");
            await new BrowserFetcher().DownloadAsync();
            await using var browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = true });
            await using var page = await browser.NewPageAsync();

            await page.SetContentAsync(htmlContent);

            var pdfOptions = new PdfOptions
            {
                Format = PuppeteerSharp.Media.PaperFormat.A4,

                Landscape = false,
                MarginOptions = new PuppeteerSharp.Media.MarginOptions
                {
                    Top = "20mm",
                    Bottom = "20mm",
                    Left = "20mm",
                    Right = "20mm"
                },
                PrintBackground = true

            };

            var pdfBytes = await page.PdfDataAsync(pdfOptions);
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "EC");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var fileName = $"EC_{ContractId}_Preview.pdf";
            var filePath = Path.Combine(folderPath, fileName);

            // Get password from appsettings.json
            string? userPassword = await GetPdfPasswordAsync(Name);;

            using (var inputStream = new MemoryStream(pdfBytes))
            using (var outputStream = new MemoryStream())
            {
                var document = PdfSharpCore.Pdf.IO.PdfReader.Open(inputStream, PdfSharpCore.Pdf.IO.PdfDocumentOpenMode.Modify);

                foreach (var pdfPage in document.Pages)
                {
                    using (var gfx = PdfSharpCore.Drawing.XGraphics.FromPdfPage(pdfPage))
                    {
                        var font = new PdfSharpCore.Drawing.XFont("Tahoma", 25, PdfSharpCore.Drawing.XFontStyle.Bold);
                        var text = $"สัญญาอิเลคทรอนิกส์พิมพ์ออกโดย {Name}\nวันที่ {DateTime.Now:dd/MM/yyyy HH:mm}";
                        var lines = text.Split('\n');
                        double lineHeight = font.GetHeight();
                        double totalHeight = lineHeight * lines.Length;
                        double y = (pdfPage.Height - totalHeight) / 2;

                        foreach (var line in lines)
                        {
                            var size = gfx.MeasureString(line, font);
                            double x = (pdfPage.Width - size.Width) / 2;
                            var state = gfx.Save();
                            gfx.TranslateTransform(pdfPage.Width / 2, pdfPage.Height / 2);
                            gfx.RotateTransform(-30);
                            gfx.TranslateTransform(-pdfPage.Width / 2, -pdfPage.Height / 2);

                            var brush = new PdfSharpCore.Drawing.XSolidBrush(
                                PdfSharpCore.Drawing.XColor.FromArgb(80, 255, 0, 0)); // semi-transparent red

                            gfx.DrawString(line, font, brush, x, y);
                            gfx.Restore(state);

                            y += lineHeight;
                        }
                    }
                }

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

        public async Task<IActionResult> OnGetWordContact_EC_JPEG(string ContractId = "55")
        {
            // 1. Generate PDF from EC contract
            var htmlContent = await _HireEmployee.OnGetWordContact_HireEmployee_ToPDF(ContractId, "EC");
            await new BrowserFetcher().DownloadAsync();
            await using var browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = true });
            await using var page = await browser.NewPageAsync();

            await page.SetContentAsync(htmlContent);

            var pdfOptions = new PdfOptions
            {
                Format = PuppeteerSharp.Media.PaperFormat.A4,

                Landscape = false,
                MarginOptions = new PuppeteerSharp.Media.MarginOptions
                {
                    Top = "20mm",
                    Bottom = "20mm",
                    Left = "20mm",
                    Right = "20mm"
                },
                PrintBackground = true

            };

            var pdfBytes = await page.PdfDataAsync(pdfOptions);

            // 2. Prepare folder structure
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "EC", "EC_" + ContractId, "EC_" + ContractId + "_JPEG");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var pdfPath = Path.Combine(folderPath, $"EC_{ContractId}.pdf");

            // 3. Save PDF to disk
            if (System.IO.File.Exists(pdfPath))
            {
                System.IO.File.Delete(pdfPath);
            }
            await System.IO.File.WriteAllBytesAsync(pdfPath, pdfBytes);

            // 4. Convert each PDF page to JPEG
            using (var pdfStream = new MemoryStream(pdfBytes))
            using (var document = PdfiumViewer.PdfDocument.Load(pdfStream))
            {
                for (int i = 0; i < document.PageCount; i++)
                {
                    using (var image = document.Render(i, 300, 300, true))
                    {
                        var jpegPath = Path.Combine(folderPath, $"EC_{ContractId}_p{i + 1}.jpg");
                        image.Save(jpegPath, System.Drawing.Imaging.ImageFormat.Jpeg);
                    }
                }
            }
            // 5. Delete the PDF file after conversion
            System.IO.File.Delete(pdfPath);

            // 6. Create zip file of JPEGs
            var folderPathZip = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "EC", "EC_" + ContractId);
            var zipPath = Path.Combine(folderPathZip, $"EC_{ContractId}_JPEG.zip");
            if (System.IO.File.Exists(zipPath))
            {
                System.IO.File.Delete(zipPath);
            }
            var jpegFiles = Directory.GetFiles(folderPath, $"EC_{ContractId}_p*.jpg");
            using (var zip = System.IO.Compression.ZipFile.Open(zipPath, ZipArchiveMode.Create))
            {
                foreach (var file in jpegFiles)
                {
                    zip.CreateEntryFromFile(file, Path.GetFileName(file));
                }
            }
            // 7. Delete JPEG files after zipping
            foreach (var file in jpegFiles)
            {
                System.IO.File.Delete(file);
            }

            // 8. Return the zip file as download
            var zipBytes = await System.IO.File.ReadAllBytesAsync(zipPath);
            return File(zipBytes, "application/zip", $"EC_{ContractId}_JPEG.zip");
        }

        public async Task<IActionResult> OnGetWordContact_EC_JPEG_Preview(string ContractId = "55")
        {
            // 1. Generate PDF from EC contract
            var htmlContent = await _HireEmployee.OnGetWordContact_HireEmployee_ToPDF(ContractId, "EC");
            await new BrowserFetcher().DownloadAsync();
            await using var browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = true });
            await using var page = await browser.NewPageAsync();

            await page.SetContentAsync(htmlContent);

            var pdfOptions = new PdfOptions
            {
                Format = PuppeteerSharp.Media.PaperFormat.A4,

                Landscape = false,
                MarginOptions = new PuppeteerSharp.Media.MarginOptions
                {
                    Top = "20mm",
                    Bottom = "20mm",
                    Left = "20mm",
                    Right = "20mm"
                },
                PrintBackground = true

            };

            var pdfBytes = await page.PdfDataAsync(pdfOptions);

            // 2. Prepare folder structure
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "EC", "EC_" + ContractId, "EC_" + ContractId + "_JPEG");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var pdfPath = Path.Combine(folderPath, $"EC_{ContractId}.pdf");

            // 3. Save PDF to disk
            if (System.IO.File.Exists(pdfPath))
            {
                System.IO.File.Delete(pdfPath);
            }
            await System.IO.File.WriteAllBytesAsync(pdfPath, pdfBytes);

            // 4. Convert each PDF page to JPEG
            using (var pdfStream = new MemoryStream(pdfBytes))
            using (var document = PdfiumViewer.PdfDocument.Load(pdfStream))
            {
                for (int i = 0; i < document.PageCount; i++)
                {
                    using (var image = document.Render(i, 300, 300, true))
                    {
                        var jpegPath = Path.Combine(folderPath, $"EC_{ContractId}_p{i + 1}.jpg");
                        image.Save(jpegPath, System.Drawing.Imaging.ImageFormat.Jpeg);
                    }
                }
            }
            // 5. Delete the PDF file after conversion
            System.IO.File.Delete(pdfPath);

            // 6. Create password-protected zip file of JPEGs
            var folderPathZip = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "EC", "EC_" + ContractId);
            var zipPath = Path.Combine(folderPathZip, $"EC_{ContractId}_JPEG_Preview.zip");
            if (System.IO.File.Exists(zipPath))
            {
                System.IO.File.Delete(zipPath);
            }
            var jpegFiles = Directory.GetFiles(folderPath, $"EC_{ContractId}_p*.jpg");
            string password = _configuration["Password:PaswordPDF"];

            using (var fsOut = System.IO.File.Create(zipPath))
            using (var zipStream = new ICSharpCode.SharpZipLib.Zip.ZipOutputStream(fsOut))
            {
                zipStream.SetLevel(9); // 0-9, 9 = best compression
                zipStream.Password = password; // Set password

                foreach (var file in jpegFiles)
                {
                    var entry = new ICSharpCode.SharpZipLib.Zip.ZipEntry(Path.GetFileName(file))
                    {
                        DateTime = DateTime.Now
                    };
                    zipStream.PutNextEntry(entry);

                    byte[] buffer = System.IO.File.ReadAllBytes(file);
                    zipStream.Write(buffer, 0, buffer.Length);
                    zipStream.CloseEntry();
                }
                zipStream.IsStreamOwner = true;
                zipStream.Close();
            }
            // 7. Delete JPEG files after zipping
            foreach (var file in jpegFiles)
            {
                System.IO.File.Delete(file);
            }

            // 8. Return the password-protected zip file as download
            var zipBytes = await System.IO.File.ReadAllBytesAsync(zipPath);
            return File(zipBytes, "application/zip", $"EC_{ContractId}_JPEG_Preview.zip");
        }

        public async Task<IActionResult> OnGetWordContact_EC_Word(string ContractId = "55")
        {
            // 1. Get HTML content for EC contract
            var htmlContent = await _HireEmployee.OnGetWordContact_HireEmployee_ToPDF(ContractId, "EC");

            // 2. Create a Word document from HTML using Spire.Doc
            var document = new Spire.Doc.Document();
            document.LoadFromStream(new MemoryStream(System.Text.Encoding.UTF8.GetBytes(htmlContent)), Spire.Doc.FileFormat.Html);

            // 3. Save the Word document to a MemoryStream
            using var ms = new MemoryStream();
            document.SaveToStream(ms, Spire.Doc.FileFormat.Docx);
            var wordBytes = ms.ToArray();

            // 4. Prepare folder structure
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "EC", "EC_" + ContractId);
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var filePath = Path.Combine(folderPath, $"EC_{ContractId}.docx");

            // 5. Save Word file to disk
            if (System.IO.File.Exists(filePath))
            {
                System.IO.File.Delete(filePath);
            }
            await System.IO.File.WriteAllBytesAsync(filePath, wordBytes);

            // 6. Return the Word file as download
            var resultBytes = await System.IO.File.ReadAllBytesAsync(filePath);
            return File(resultBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", $"EC_{ContractId}.docx");
        }
        public async Task<IActionResult> OnGetWordContact_EC_Word_Preview(string ContractId = "55")
        {
            // 1. Get HTML content for EC contract
            var htmlContent = await _HireEmployee.OnGetWordContact_HireEmployee_ToPDF(ContractId, "EC");

            // 2. Create a Word document from HTML using Spire.Doc
            var document = new Spire.Doc.Document();
            document.LoadFromStream(new MemoryStream(System.Text.Encoding.UTF8.GetBytes(htmlContent)), Spire.Doc.FileFormat.Html);

            // 3. Get password from appsettings.json
            string? userPassword = _configuration["Password:PaswordPDF"];
            if (!string.IsNullOrEmpty(userPassword))
            {
                document.Encrypt(userPassword);
            }

            // 4. Prepare folder structure
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "EC", "EC_" + ContractId);
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var filePath = Path.Combine(folderPath, $"EC_{ContractId}_Preview.docx");

            // 5. Delete the file if it already exists
            if (System.IO.File.Exists(filePath))
            {
                System.IO.File.Delete(filePath);
            }

            // 6. Save the password-protected Word document
            document.SaveToFile(filePath, FileFormat.Docx);

            // 7. Return the Word file as download
            var resultBytes = await System.IO.File.ReadAllBytesAsync(filePath);
            return File(resultBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", $"EC_{ContractId}_Preview.docx");
        }
        #endregion 4.1.3.3. สัญญาจ้างลูกจ้าง

        #region 4.1.1.2.15.สัญญาจ้างทำของ CWA
        public async Task<IActionResult> OnGetWordContact_CWA(string ContractId = "1")
        {
            var wordBytes = await _ContactToDoThingService.OnGetWordContact_ToDoThing(ContractId);
            return File(wordBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "สัญญาจ้างทำของ.docx");
        }

        public async Task OnGetWordContact_CWA_PDF(string ContractId = "1")
        {
            var htmlContent = await _ContactToDoThingService.OnGetWordContact_ToDoThing_ToPDF(ContractId, "CWA");
            await new BrowserFetcher().DownloadAsync();
            await using var browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = true });
            await using var page = await browser.NewPageAsync();

            await page.SetContentAsync(htmlContent);

            var pdfOptions = new PdfOptions
            {
                Format = PuppeteerSharp.Media.PaperFormat.A4,

                Landscape = false,
                MarginOptions = new PuppeteerSharp.Media.MarginOptions
                {
                    Top = "20mm",
                    Bottom = "20mm",
                    Left = "20mm",
                    Right = "20mm"
                },
                PrintBackground = true

            };

            var pdfBytes = await page.PdfDataAsync(pdfOptions);
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
            await System.IO.File.WriteAllBytesAsync(filePath, pdfBytes);

            // return File(wordBytes, "application/pdf", "สัญญาซื้อขาย.pdf");
        }
        public async Task<IActionResult> OnGetWordContact_CWA_PDF_Preview(string ContractId = "1", string Name = "สมใจ ทดสอบ")
        {
            // 1. Generate PDF from CWA contract
            var htmlContent = await _ContactToDoThingService.OnGetWordContact_ToDoThing_ToPDF(ContractId, "CWA");
            await new BrowserFetcher().DownloadAsync();
            await using var browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = true });
            await using var page = await browser.NewPageAsync();

            await page.SetContentAsync(htmlContent);

            var pdfOptions = new PdfOptions
            {
                Format = PuppeteerSharp.Media.PaperFormat.A4,

                Landscape = false,
                MarginOptions = new PuppeteerSharp.Media.MarginOptions
                {
                    Top = "20mm",
                    Bottom = "20mm",
                    Left = "20mm",
                    Right = "20mm"
                },
                PrintBackground = true

            };

            var pdfBytes = await page.PdfDataAsync(pdfOptions);
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "CWA");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var fileName = $"CWA_{ContractId}_Preview.pdf";
            var filePath = Path.Combine(folderPath, fileName);

            // Get password from appsettings.json
            string? userPassword = await GetPdfPasswordAsync(Name);;

            using (var inputStream = new MemoryStream(pdfBytes))
            using (var outputStream = new MemoryStream())
            {
                var document = PdfSharpCore.Pdf.IO.PdfReader.Open(inputStream, PdfSharpCore.Pdf.IO.PdfDocumentOpenMode.Modify);

                foreach (var pdfPage in document.Pages)
                {
                    using (var gfx = PdfSharpCore.Drawing.XGraphics.FromPdfPage(pdfPage))
                    {
                        var font = new PdfSharpCore.Drawing.XFont("Tahoma", 25, PdfSharpCore.Drawing.XFontStyle.Bold);
                        var text = $"สัญญาอิเลคทรอนิกส์พิมพ์ออกโดย {Name}\nวันที่ {DateTime.Now:dd/MM/yyyy HH:mm}";
                        var lines = text.Split('\n');
                        double lineHeight = font.GetHeight();
                        double totalHeight = lineHeight * lines.Length;
                        double y = (pdfPage.Height - totalHeight) / 2;

                        foreach (var line in lines)
                        {
                            var size = gfx.MeasureString(line, font);
                            double x = (pdfPage.Width - size.Width) / 2;
                            var state = gfx.Save();
                            gfx.TranslateTransform(pdfPage.Width / 2, pdfPage.Height / 2);
                            gfx.RotateTransform(-30);
                            gfx.TranslateTransform(-pdfPage.Width / 2, -pdfPage.Height / 2);

                            var brush = new PdfSharpCore.Drawing.XSolidBrush(
                                PdfSharpCore.Drawing.XColor.FromArgb(80, 255, 0, 0)); // semi-transparent red

                            gfx.DrawString(line, font, brush, x, y);
                            gfx.Restore(state);

                            y += lineHeight;
                        }
                    }
                }

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

        public async Task<IActionResult> OnGetWordContact_CWA_JPEG(string ContractId = "1")
        {
            // 1. Generate PDF from CWA contract
            var htmlContent = await _ContactToDoThingService.OnGetWordContact_ToDoThing_ToPDF(ContractId, "CWA");
            await new BrowserFetcher().DownloadAsync();
            await using var browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = true });
            await using var page = await browser.NewPageAsync();

            await page.SetContentAsync(htmlContent);

            var pdfOptions = new PdfOptions
            {
                Format = PuppeteerSharp.Media.PaperFormat.A4,

                Landscape = false,
                MarginOptions = new PuppeteerSharp.Media.MarginOptions
                {
                    Top = "20mm",
                    Bottom = "20mm",
                    Left = "20mm",
                    Right = "20mm"
                },
                PrintBackground = true

            };

            var pdfBytes = await page.PdfDataAsync(pdfOptions);

            // 2. Prepare folder structure
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "CWA", "CWA_" + ContractId, "CWA_" + ContractId + "_JPEG");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var pdfPath = Path.Combine(folderPath, $"CWA_{ContractId}.pdf");

            // 3. Save PDF to disk
            if (System.IO.File.Exists(pdfPath))
            {
                System.IO.File.Delete(pdfPath);
            }
            await System.IO.File.WriteAllBytesAsync(pdfPath, pdfBytes);

            // 4. Convert each PDF page to JPEG
            using (var pdfStream = new MemoryStream(pdfBytes))
            using (var document = PdfiumViewer.PdfDocument.Load(pdfStream))
            {
                for (int i = 0; i < document.PageCount; i++)
                {
                    using (var image = document.Render(i, 300, 300, true))
                    {
                        var jpegPath = Path.Combine(folderPath, $"CWA_{ContractId}_p{i + 1}.jpg");
                        image.Save(jpegPath, System.Drawing.Imaging.ImageFormat.Jpeg);
                    }
                }
            }
            // 5. Delete the PDF file after conversion
            System.IO.File.Delete(pdfPath);

            // 6. Create zip file of JPEGs
            var folderPathZip = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "CWA", "CWA_" + ContractId);
            var zipPath = Path.Combine(folderPathZip, $"CWA_{ContractId}_JPEG.zip");
            if (System.IO.File.Exists(zipPath))
            {
                System.IO.File.Delete(zipPath);
            }
            var jpegFiles = Directory.GetFiles(folderPath, $"CWA_{ContractId}_p*.jpg");
            using (var zip = System.IO.Compression.ZipFile.Open(zipPath, ZipArchiveMode.Create))
            {
                foreach (var file in jpegFiles)
                {
                    zip.CreateEntryFromFile(file, Path.GetFileName(file));
                }
            }
            // 7. Delete JPEG files after zipping
            foreach (var file in jpegFiles)
            {
                System.IO.File.Delete(file);
            }

            // 8. Return the zip file as download
            var zipBytes = await System.IO.File.ReadAllBytesAsync(zipPath);
            return File(zipBytes, "application/zip", $"CWA_{ContractId}_JPEG.zip");
        }
        public async Task<IActionResult> OnGetWordContact_CWA_JPEG_Preview(string ContractId = "1")
        {
        
            // 1. Generate PDF from CWA contract
            var htmlContent = await _ContactToDoThingService.OnGetWordContact_ToDoThing_ToPDF(ContractId, "CWA");
            await new BrowserFetcher().DownloadAsync();
            await using var browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = true });
            await using var page = await browser.NewPageAsync();

            await page.SetContentAsync(htmlContent);

            var pdfOptions = new PdfOptions
            {
                Format = PuppeteerSharp.Media.PaperFormat.A4,

                Landscape = false,
                MarginOptions = new PuppeteerSharp.Media.MarginOptions
                {
                    Top = "20mm",
                    Bottom = "20mm",
                    Left = "20mm",
                    Right = "20mm"
                },
                PrintBackground = true

            };

            var pdfBytes = await page.PdfDataAsync(pdfOptions);
            // 2. Prepare folder structure
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "CWA", "CWA_" + ContractId, "CWA_" + ContractId + "_JPEG");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var pdfPath = Path.Combine(folderPath, $"CWA_{ContractId}.pdf");

            // 3. Save PDF to disk
            if (System.IO.File.Exists(pdfPath))
            {
                System.IO.File.Delete(pdfPath);
            }
            await System.IO.File.WriteAllBytesAsync(pdfPath, pdfBytes);

            // 4. Convert each PDF page to JPEG
            using (var pdfStream = new MemoryStream(pdfBytes))
            using (var document = PdfiumViewer.PdfDocument.Load(pdfStream))
            {
                for (int i = 0; i < document.PageCount; i++)
                {
                    using (var image = document.Render(i, 300, 300, true))
                    {
                        var jpegPath = Path.Combine(folderPath, $"CWA_{ContractId}_p{i + 1}.jpg");
                        image.Save(jpegPath, System.Drawing.Imaging.ImageFormat.Jpeg);
                    }
                }
            }
            // 5. Delete the PDF file after conversion
            System.IO.File.Delete(pdfPath);

            // 6. Create password-protected zip file of JPEGs
            var folderPathZip = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "CWA", "CWA_" + ContractId);
            var zipPath = Path.Combine(folderPathZip, $"CWA_{ContractId}_JPEG_Preview.zip");
            if (System.IO.File.Exists(zipPath))
            {
                System.IO.File.Delete(zipPath);
            }
            var jpegFiles = Directory.GetFiles(folderPath, $"CWA_{ContractId}_p*.jpg");
            string password = _configuration["Password:PaswordPDF"];

            using (var fsOut = System.IO.File.Create(zipPath))
            using (var zipStream = new ICSharpCode.SharpZipLib.Zip.ZipOutputStream(fsOut))
            {
                zipStream.SetLevel(9); // 0-9, 9 = best compression
                zipStream.Password = password; // Set password

                foreach (var file in jpegFiles)
                {
                    var entry = new ICSharpCode.SharpZipLib.Zip.ZipEntry(Path.GetFileName(file))
                    {
                        DateTime = DateTime.Now
                    };
                    zipStream.PutNextEntry(entry);

                    byte[] buffer = System.IO.File.ReadAllBytes(file);
                    zipStream.Write(buffer, 0, buffer.Length);
                    zipStream.CloseEntry();
                }
                zipStream.IsStreamOwner = true;
                zipStream.Close();
            }
            // 7. Delete JPEG files after zipping
            foreach (var file in jpegFiles)
            {
                System.IO.File.Delete(file);
            }

            // 8. Return the password-protected zip file as download
            var zipBytes = await System.IO.File.ReadAllBytesAsync(zipPath);
            return File(zipBytes, "application/zip", $"CWA_{ContractId}_JPEG_Preview.zip");
        }
        public async Task<IActionResult> OnGetWordContact_CWA_Word(string ContractId = "1")
        {

            // 1. Generate PDF from CWA contract
            var htmlContent = await _ContactToDoThingService.OnGetWordContact_ToDoThing_ToPDF(ContractId, "CWA");

            // 2. Create a Word document from HTML using Spire.Doc
            var document = new Spire.Doc.Document();
            document.LoadFromStream(new MemoryStream(System.Text.Encoding.UTF8.GetBytes(htmlContent)), Spire.Doc.FileFormat.Html);

            // 3. Save the Word document to a MemoryStream
            using var ms = new MemoryStream();
            document.SaveToStream(ms, Spire.Doc.FileFormat.Docx);
            var wordBytes = ms.ToArray();

            // 4. Prepare folder structure
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "CWA", "CWA_" + ContractId);
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var filePath = Path.Combine(folderPath, $"CWA_{ContractId}.docx");

            // 5. Save Word file to disk
            if (System.IO.File.Exists(filePath))
            {
                System.IO.File.Delete(filePath);
            }
            await System.IO.File.WriteAllBytesAsync(filePath, wordBytes);

            // 6. Return the Word file as download
            var resultBytes = await System.IO.File.ReadAllBytesAsync(filePath);
            return File(resultBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", $"CWA_{ContractId}.docx");
        }
        public async Task<IActionResult> OnGetWordContact_CWA_Word_Preview(string ContractId = "1")
        {
            // 1. Get HTML content for CWA contract
            var htmlContent = await _ContactToDoThingService.OnGetWordContact_ToDoThing_ToPDF(ContractId, "CWA");

            // 2. Create a Word document from HTML using Spire.Doc
            var document = new Spire.Doc.Document();
            document.LoadFromStream(new MemoryStream(System.Text.Encoding.UTF8.GetBytes(htmlContent)), Spire.Doc.FileFormat.Html);

            // 3. Get password from appsettings.json
            string? userPassword = _configuration["Password:PaswordPDF"];
            if (!string.IsNullOrEmpty(userPassword))
            {
                document.Encrypt(userPassword);
            }

            // 4. Prepare folder structure
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "CWA", "CWA_" + ContractId);
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var filePath = Path.Combine(folderPath, $"CWA_{ContractId}_Preview.docx");

            // 5. Delete the file if it already exists
            if (System.IO.File.Exists(filePath))
            {
                System.IO.File.Delete(filePath);
            }

            // 6. Save the password-protected Word document
            document.SaveToFile(filePath, FileFormat.Docx);

            // 7. Return the Word file as download
            var resultBytes = await System.IO.File.ReadAllBytesAsync(filePath);
            return File(resultBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", $"CWA_{ContractId}_Preview.docx");
        }

        #endregion 4.1.1.2.15.สัญญาจ้างทำของ CWA

        #region 4.1.1.2.14.สัญญาจ้างผู้เชี่ยวชาญรายบุคคลหรือจ้างบริษัทที่ปรึกษา ร.317-60 CTR31760
        public async Task<IActionResult> OnGetWordContact_CTR31760(string ContractId = "1")
        {
            var wordBytes = await _ConsultantService.OnGetWordContact_ConsultantService(ContractId);
            return File(wordBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "สัญญาจ้างที่ปรึกษา.docx");
        }
        public async Task OnGetWordContact_CTR31760_PDF(string ContractId = "1")
        {
            var htmlContent = await _ConsultantService.OnGetWordContact_ConsultantService_ToPDF(ContractId, "CTR31760");
            await new BrowserFetcher().DownloadAsync();
            await using var browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = true });
            await using var page = await browser.NewPageAsync();

            await page.SetContentAsync(htmlContent);

            var pdfOptions = new PdfOptions
            {
                Format = PuppeteerSharp.Media.PaperFormat.A4,

                Landscape = false,
                MarginOptions = new PuppeteerSharp.Media.MarginOptions
                {
                    Top = "20mm",
                    Bottom = "20mm",
                    Left = "20mm",
                    Right = "20mm"
                },
                PrintBackground = true

            };

            var pdfBytes = await page.PdfDataAsync(pdfOptions);



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
            await System.IO.File.WriteAllBytesAsync(filePath, pdfBytes);

            // return File(wordBytes, "application/pdf", "สัญญาซื้อขาย.pdf");
        }
        public async Task<IActionResult> OnGetWordContact_CTR31760_PDF_Preview(string ContractId = "1", string Name = "สมใจ ทดสอบ")
        {
            var htmlContent = await _ConsultantService.OnGetWordContact_ConsultantService_ToPDF(ContractId, "CTR31760");
            await new BrowserFetcher().DownloadAsync();
            await using var browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = true });
            await using var page = await browser.NewPageAsync();

            await page.SetContentAsync(htmlContent);

            var pdfOptions = new PdfOptions
            {
                Format = PuppeteerSharp.Media.PaperFormat.A4,

                Landscape = false,
                MarginOptions = new PuppeteerSharp.Media.MarginOptions
                {
                    Top = "20mm",
                    Bottom = "20mm",
                    Left = "20mm",
                    Right = "20mm"
                },
                PrintBackground = true

            };

            var pdfBytes = await page.PdfDataAsync(pdfOptions);

            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "CTR31760");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var filePath = Path.Combine(folderPath, "CTR31760_" + ContractId + "_Preview.pdf");
            var fileName = $"CTR31760_{ContractId}_Preview.pdf";
            // Set your desired password here
            string? userPassword = await GetPdfPasswordAsync(Name);;

            // Load the PDF from the byte array
            using (var inputStream = new MemoryStream(pdfBytes))
            using (var outputStream = new MemoryStream())
            {
                var document = PdfSharpCore.Pdf.IO.PdfReader.Open(inputStream, PdfSharpCore.Pdf.IO.PdfDocumentOpenMode.Modify);

                foreach (var pdfPage in document.Pages)
                {
                    using (var gfx = PdfSharpCore.Drawing.XGraphics.FromPdfPage(pdfPage))
                    {
                        var font = new PdfSharpCore.Drawing.XFont("Tahoma", 25, PdfSharpCore.Drawing.XFontStyle.Bold);
                        var text = $"สัญญาอิเลคทรอนิกส์พิมพ์ออกโดย {Name}\nวันที่ {DateTime.Now:dd/MM/yyyy HH:mm}";
                        var lines = text.Split('\n');
                        double lineHeight = font.GetHeight();
                        double totalHeight = lineHeight * lines.Length;
                        double y = (pdfPage.Height - totalHeight) / 2;

                        foreach (var line in lines)
                        {
                            var size = gfx.MeasureString(line, font);
                            double x = (pdfPage.Width - size.Width) / 2;
                            var state = gfx.Save();
                            gfx.TranslateTransform(pdfPage.Width / 2, pdfPage.Height / 2);
                            gfx.RotateTransform(-30);
                            gfx.TranslateTransform(-pdfPage.Width / 2, -pdfPage.Height / 2);

                            var brush = new PdfSharpCore.Drawing.XSolidBrush(
                                PdfSharpCore.Drawing.XColor.FromArgb(80, 255, 0, 0)); // semi-transparent red

                            gfx.DrawString(line, font, brush, x, y);
                            gfx.Restore(state);

                            y += lineHeight;
                        }
                    }
                }

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
        public async Task<IActionResult> OnGetWordContact_CTR31760_JPEG(string ContractId = "1")
        {
            // 1. Generate PDF from CTR31760 contract
            var htmlContent = await _ConsultantService.OnGetWordContact_ConsultantService_ToPDF(ContractId, "CTR31760");
            await new BrowserFetcher().DownloadAsync();
            await using var browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = true });
            await using var page = await browser.NewPageAsync();

            await page.SetContentAsync(htmlContent);

            var pdfOptions = new PdfOptions
            {
                Format = PuppeteerSharp.Media.PaperFormat.A4,

                Landscape = false,
                MarginOptions = new PuppeteerSharp.Media.MarginOptions
                {
                    Top = "20mm",
                    Bottom = "20mm",
                    Left = "20mm",
                    Right = "20mm"
                },
                PrintBackground = true

            };

            var pdfBytes = await page.PdfDataAsync(pdfOptions);

            // 2. Prepare folder structure
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "CTR31760", "CTR31760_" + ContractId, "CTR31760_" + ContractId + "_JPEG");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var pdfPath = Path.Combine(folderPath, $"CTR31760_{ContractId}.pdf");

            // 3. Save PDF to disk
            if (System.IO.File.Exists(pdfPath))
            {
                System.IO.File.Delete(pdfPath);
            }
            await System.IO.File.WriteAllBytesAsync(pdfPath, pdfBytes);

            // 4. Convert each PDF page to JPEG
            using (var pdfStream = new MemoryStream(pdfBytes))
            using (var document = PdfiumViewer.PdfDocument.Load(pdfStream))
            {
                for (int i = 0; i < document.PageCount; i++)
                {
                    using (var image = document.Render(i, 300, 300, true))
                    {
                        var jpegPath = Path.Combine(folderPath, $"CTR31760_{ContractId}_p{i + 1}.jpg");
                        image.Save(jpegPath, System.Drawing.Imaging.ImageFormat.Jpeg);
                    }
                }
            }
            // 5. Delete the PDF file after conversion
            System.IO.File.Delete(pdfPath);

            // 6. Create zip file of JPEGs
            var folderPathZip = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "CTR31760", "CTR31760_" + ContractId);
            var zipPath = Path.Combine(folderPathZip, $"CTR31760_{ContractId}_JPEG.zip");
            if (System.IO.File.Exists(zipPath))
            {
                System.IO.File.Delete(zipPath);
            }
            var jpegFiles = Directory.GetFiles(folderPath, $"CTR31760_{ContractId}_p*.jpg");
            using (var zip = System.IO.Compression.ZipFile.Open(zipPath, ZipArchiveMode.Create))
            {
                foreach (var file in jpegFiles)
                {
                    zip.CreateEntryFromFile(file, Path.GetFileName(file));
                }
            }
            // 7. Delete JPEG files after zipping
            foreach (var file in jpegFiles)
            {
                System.IO.File.Delete(file);
            }

            // 8. Return the zip file as download
            var zipBytes = await System.IO.File.ReadAllBytesAsync(zipPath);
            return File(zipBytes, "application/zip", $"CTR31760_{ContractId}_JPEG.zip");
        }
        public async Task<IActionResult> OnGetWordContact_CTR31760_JPEG_Preview(string ContractId = "1")
        {
            // 1. Generate PDF from CTR31760 contract
            var htmlContent = await _ConsultantService.OnGetWordContact_ConsultantService_ToPDF(ContractId, "CTR31760");
            await new BrowserFetcher().DownloadAsync();
            await using var browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = true });
            await using var page = await browser.NewPageAsync();

            await page.SetContentAsync(htmlContent);

            var pdfOptions = new PdfOptions
            {
                Format = PuppeteerSharp.Media.PaperFormat.A4,

                Landscape = false,
                MarginOptions = new PuppeteerSharp.Media.MarginOptions
                {
                    Top = "20mm",
                    Bottom = "20mm",
                    Left = "20mm",
                    Right = "20mm"
                },
                PrintBackground = true

            };

            var pdfBytes = await page.PdfDataAsync(pdfOptions);


            // 2. Prepare folder structure
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "CTR31760", "CTR31760_" + ContractId, "CTR31760_" + ContractId + "_JPEG");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var pdfPath = Path.Combine(folderPath, $"CTR31760_{ContractId}.pdf");

            // 3. Save PDF to disk
            if (System.IO.File.Exists(pdfPath))
            {
                System.IO.File.Delete(pdfPath);
            }
            await System.IO.File.WriteAllBytesAsync(pdfPath, pdfBytes);

            // 4. Convert each PDF page to JPEG
            using (var pdfStream = new MemoryStream(pdfBytes))
            using (var document = PdfiumViewer.PdfDocument.Load(pdfStream))
            {
                for (int i = 0; i < document.PageCount; i++)
                {
                    using (var image = document.Render(i, 300, 300, true))
                    {
                        var jpegPath = Path.Combine(folderPath, $"CTR31760_{ContractId}_p{i + 1}.jpg");
                        image.Save(jpegPath, System.Drawing.Imaging.ImageFormat.Jpeg);
                    }
                }
            }
            // 5. Delete the PDF file after conversion
            System.IO.File.Delete(pdfPath);

            // 6. Create password-protected zip file of JPEGs
            var folderPathZip = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "CTR31760", "CTR31760_" + ContractId);
            var zipPath = Path.Combine(folderPathZip, $"CTR31760_{ContractId}_JPEG_Preview.zip");
            if (System.IO.File.Exists(zipPath))
            {
                System.IO.File.Delete(zipPath);
            }
            var jpegFiles = Directory.GetFiles(folderPath, $"CTR31760_{ContractId}_p*.jpg");
            string password = _configuration["Password:PaswordPDF"];

            using (var fsOut = System.IO.File.Create(zipPath))
            using (var zipStream = new ICSharpCode.SharpZipLib.Zip.ZipOutputStream(fsOut))
            {
                zipStream.SetLevel(9); // 0-9, 9 = best compression
                zipStream.Password = password; // Set password

                foreach (var file in jpegFiles)
                {
                    var entry = new ICSharpCode.SharpZipLib.Zip.ZipEntry(Path.GetFileName(file))
                    {
                        DateTime = DateTime.Now
                    };
                    zipStream.PutNextEntry(entry);

                    byte[] buffer = System.IO.File.ReadAllBytes(file);
                    zipStream.Write(buffer, 0, buffer.Length);
                    zipStream.CloseEntry();
                }
                zipStream.IsStreamOwner = true;
                zipStream.Close();
            }
            // 7. Delete JPEG files after zipping
            foreach (var file in jpegFiles)
            {
                System.IO.File.Delete(file);
            }

            // 8. Return the password-protected zip file as download
            var zipBytes = await System.IO.File.ReadAllBytesAsync(zipPath);
            return File(zipBytes, "application/zip", $"CTR31760_{ContractId}_JPEG_Preview.zip");
        }
        public async Task<IActionResult> OnGetWordContact_CTR31760_Word(string ContractId = "1")
        {
            // 1. Get HTML content for CTR31760 contract
            var htmlContent = await _ConsultantService.OnGetWordContact_ConsultantService_ToPDF(ContractId, "CTR31760");

            // 2. Create a Word document from HTML using Spire.Doc
            var document = new Spire.Doc.Document();
            document.LoadFromStream(new MemoryStream(System.Text.Encoding.UTF8.GetBytes(htmlContent)), Spire.Doc.FileFormat.Html);

            // 3. Save the Word document to a MemoryStream
            using var ms = new MemoryStream();
            document.SaveToStream(ms, Spire.Doc.FileFormat.Docx);
            var wordBytes = ms.ToArray();

            // 4. Prepare folder structure
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "CTR31760", "CTR31760_" + ContractId);
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var filePath = Path.Combine(folderPath, $"CTR31760_{ContractId}.docx");

            // 5. Save Word file to disk
            if (System.IO.File.Exists(filePath))
            {
                System.IO.File.Delete(filePath);
            }
            await System.IO.File.WriteAllBytesAsync(filePath, wordBytes);

            // 6. Return the Word file as download
            var resultBytes = await System.IO.File.ReadAllBytesAsync(filePath);
            return File(resultBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", $"CTR31760_{ContractId}.docx");
        }
        public async Task<IActionResult> OnGetWordContact_CTR31760_Word_Preview(string ContractId = "1")
        {
            // 1. Get the HTML content for CTR31760 contract
            var htmlContent = await _ConsultantService.OnGetWordContact_ConsultantService_ToPDF(ContractId, "CTR31760");

            // 2. Create a Word document from HTML using Spire.Doc
            var document = new Spire.Doc.Document();
            document.LoadFromStream(new MemoryStream(System.Text.Encoding.UTF8.GetBytes(htmlContent)), Spire.Doc.FileFormat.Html);

            // 3. Get password from appsettings.json
            string? userPassword = _configuration["Password:PaswordPDF"];
            if (!string.IsNullOrEmpty(userPassword))
            {
                document.Encrypt(userPassword);
            }

            // 4. Prepare folder structure
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "CTR31760", "CTR31760_" + ContractId);
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var filePath = Path.Combine(folderPath, $"CTR31760_{ContractId}_Preview.docx");

            // 5. Delete the file if it already exists
            if (System.IO.File.Exists(filePath))
            {
                System.IO.File.Delete(filePath);
            }

            // 6. Save the password-protected Word document
            document.SaveToFile(filePath, FileFormat.Docx);

            // 7. Return the Word file as download
            var resultBytes = await System.IO.File.ReadAllBytesAsync(filePath);
            return File(resultBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", $"CTR31760_{ContractId}_Preview.docx");
        }
        #endregion 4.1.1.2.14.สัญญาจ้างที่ปรึกษา CTR31760

        #region 4.1.1.2.13.สัญญาเช่าเครื่องถ่ายเอกสาร ร.314-60 PML31460

        public async Task<IActionResult> OnGetWordContact_PML31460(string ContractId = "1")
        {
            var wordBytes = await _wordEContract_LoanPrinterService.OnGetWordContact_LoanPrinter(ContractId);
            return File(wordBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "สัญญาเช่าเครื่องถ่ายเอกสาร ร.314-60.docx");
        }
        public async Task OnGetWordContact_PML31460_PDF(string ContractId = "1")
        {
            var htmlContent = await _wordEContract_LoanPrinterService.OnGetWordContact_LoanPrinter_ToPDF(ContractId, "PML31460");
            await new BrowserFetcher().DownloadAsync();
            await using var browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = true });
            await using var page = await browser.NewPageAsync();

            await page.SetContentAsync(htmlContent);

            var pdfOptions = new PdfOptions
            {
                Format = PuppeteerSharp.Media.PaperFormat.A4,

                Landscape = false,
                MarginOptions = new PuppeteerSharp.Media.MarginOptions
                {
                    Top = "20mm",
                    Bottom = "20mm",
                    Left = "20mm",
                    Right = "20mm"
                },
                PrintBackground = true

            };

            var pdfBytes = await page.PdfDataAsync(pdfOptions);

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
            await System.IO.File.WriteAllBytesAsync(filePath, pdfBytes);

            // return File(wordBytes, "application/pdf", "สัญญาซื้อขาย.pdf");
        }
        public async Task<IActionResult> OnGetWordContact_PML31460_PDF_Preview(string ContractId = "1", string Name = "สมใจ ทดสอบ")
        {
            var htmlContent = await _wordEContract_LoanPrinterService.OnGetWordContact_LoanPrinter_ToPDF(ContractId, "PML31460");
            await new BrowserFetcher().DownloadAsync();
            await using var browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = true });
            await using var page = await browser.NewPageAsync();

            await page.SetContentAsync(htmlContent);

            var pdfOptions = new PdfOptions
            {
                Format = PuppeteerSharp.Media.PaperFormat.A4,

                Landscape = false,
                MarginOptions = new PuppeteerSharp.Media.MarginOptions
                {
                    Top = "20mm",
                    Bottom = "20mm",
                    Left = "20mm",
                    Right = "20mm"
                },
                PrintBackground = true

            };

            var pdfBytes = await page.PdfDataAsync(pdfOptions);

            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "PML31460");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var fileName = $"PML31460_{ContractId}_Preview.pdf";
            var filePath = Path.Combine(folderPath, fileName);

            // Get password from appsettings.json
            string? userPassword = await GetPdfPasswordAsync(Name);;

            using (var inputStream = new MemoryStream(pdfBytes))
            using (var outputStream = new MemoryStream())
            {
                var document = PdfSharpCore.Pdf.IO.PdfReader.Open(inputStream, PdfSharpCore.Pdf.IO.PdfDocumentOpenMode.Modify);

                foreach (var pdfPage in document.Pages)
                {
                    using (var gfx = PdfSharpCore.Drawing.XGraphics.FromPdfPage(pdfPage))
                    {
                        var font = new PdfSharpCore.Drawing.XFont("Tahoma", 25, PdfSharpCore.Drawing.XFontStyle.Bold);
                        var text = $"สัญญาอิเลคทรอนิกส์พิมพ์ออกโดย {Name}\nวันที่ {DateTime.Now:dd/MM/yyyy HH:mm}";
                        var lines = text.Split('\n');
                        double lineHeight = font.GetHeight();
                        double totalHeight = lineHeight * lines.Length;
                        double y = (pdfPage.Height - totalHeight) / 2;

                        foreach (var line in lines)
                        {
                            var size = gfx.MeasureString(line, font);
                            double x = (pdfPage.Width - size.Width) / 2;
                            var state = gfx.Save();
                            gfx.TranslateTransform(pdfPage.Width / 2, pdfPage.Height / 2);
                            gfx.RotateTransform(-30);
                            gfx.TranslateTransform(-pdfPage.Width / 2, -pdfPage.Height / 2);

                            var brush = new PdfSharpCore.Drawing.XSolidBrush(
                                PdfSharpCore.Drawing.XColor.FromArgb(80, 255, 0, 0)); // semi-transparent red

                            gfx.DrawString(line, font, brush, x, y);
                            gfx.Restore(state);

                            y += lineHeight;
                        }
                    }
                }

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
        public async Task<IActionResult> OnGetWordContact_PML31460_JPEG(string ContractId = "1")
        {
            // 1. Generate PDF from PML31460 contract
            var htmlContent = await _wordEContract_LoanPrinterService.OnGetWordContact_LoanPrinter_ToPDF(ContractId, "PML31460");
            await new BrowserFetcher().DownloadAsync();
            await using var browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = true });
            await using var page = await browser.NewPageAsync();

            await page.SetContentAsync(htmlContent);

            var pdfOptions = new PdfOptions
            {
                Format = PuppeteerSharp.Media.PaperFormat.A4,

                Landscape = false,
                MarginOptions = new PuppeteerSharp.Media.MarginOptions
                {
                    Top = "20mm",
                    Bottom = "20mm",
                    Left = "20mm",
                    Right = "20mm"
                },
                PrintBackground = true

            };

            var pdfBytes = await page.PdfDataAsync(pdfOptions);
            // 2. Prepare folder structure
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "PML31460", "PML31460_" + ContractId, "PML31460_" + ContractId + "_JPEG");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var pdfPath = Path.Combine(folderPath, $"PML31460_{ContractId}.pdf");

            // 3. Save PDF to disk
            if (System.IO.File.Exists(pdfPath))
            {
                System.IO.File.Delete(pdfPath);
            }
            await System.IO.File.WriteAllBytesAsync(pdfPath, pdfBytes);

            // 4. Convert each PDF page to JPEG
            using (var pdfStream = new MemoryStream(pdfBytes))
            using (var document = PdfiumViewer.PdfDocument.Load(pdfStream))
            {
                for (int i = 0; i < document.PageCount; i++)
                {
                    using (var image = document.Render(i, 300, 300, true))
                    {
                        var jpegPath = Path.Combine(folderPath, $"PML31460_{ContractId}_p{i + 1}.jpg");
                        image.Save(jpegPath, System.Drawing.Imaging.ImageFormat.Jpeg);
                    }
                }
            }
            // 5. Delete the PDF file after conversion
            System.IO.File.Delete(pdfPath);

            // 6. Create zip file of JPEGs
            var folderPathZip = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "PML31460", "PML31460_" + ContractId);
            var zipPath = Path.Combine(folderPathZip, $"PML31460_{ContractId}_JPEG.zip");
            if (System.IO.File.Exists(zipPath))
            {
                System.IO.File.Delete(zipPath);
            }
            var jpegFiles = Directory.GetFiles(folderPath, $"PML31460_{ContractId}_p*.jpg");
            using (var zip = System.IO.Compression.ZipFile.Open(zipPath, ZipArchiveMode.Create))
            {
                foreach (var file in jpegFiles)
                {
                    zip.CreateEntryFromFile(file, Path.GetFileName(file));
                }
            }
            // 7. Delete JPEG files after zipping
            foreach (var file in jpegFiles)
            {
                System.IO.File.Delete(file);
            }

            // 8. Return the zip file as download
            var zipBytes = await System.IO.File.ReadAllBytesAsync(zipPath);
            return File(zipBytes, "application/zip", $"PML31460_{ContractId}_JPEG.zip");
        }
        public async Task<IActionResult> OnGetWordContact_PML31460_JPEG_Preview(string ContractId = "1")
        {
            // 1. Generate PDF from PML31460 contract
            var htmlContent = await _wordEContract_LoanPrinterService.OnGetWordContact_LoanPrinter_ToPDF(ContractId, "PML31460");
            await new BrowserFetcher().DownloadAsync();
            await using var browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = true });
            await using var page = await browser.NewPageAsync();

            await page.SetContentAsync(htmlContent);

            var pdfOptions = new PdfOptions
            {
                Format = PuppeteerSharp.Media.PaperFormat.A4,

                Landscape = false,
                MarginOptions = new PuppeteerSharp.Media.MarginOptions
                {
                    Top = "20mm",
                    Bottom = "20mm",
                    Left = "20mm",
                    Right = "20mm"
                },
                PrintBackground = true

            };

            var pdfBytes = await page.PdfDataAsync(pdfOptions);
            // 2. Prepare folder structure
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "PML31460", "PML31460_" + ContractId, "PML31460_" + ContractId + "_JPEG");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var pdfPath = Path.Combine(folderPath, $"PML31460_{ContractId}.pdf");

            // 3. Save PDF to disk
            if (System.IO.File.Exists(pdfPath))
            {
                System.IO.File.Delete(pdfPath);
            }
            await System.IO.File.WriteAllBytesAsync(pdfPath, pdfBytes);

            // 4. Convert each PDF page to JPEG
            using (var pdfStream = new MemoryStream(pdfBytes))
            using (var document = PdfiumViewer.PdfDocument.Load(pdfStream))
            {
                for (int i = 0; i < document.PageCount; i++)
                {
                    using (var image = document.Render(i, 300, 300, true))
                    {
                        var jpegPath = Path.Combine(folderPath, $"PML31460_{ContractId}_p{i + 1}.jpg");
                        image.Save(jpegPath, System.Drawing.Imaging.ImageFormat.Jpeg);
                    }
                }
            }
            // 5. Delete the PDF file after conversion
            System.IO.File.Delete(pdfPath);

            // 6. Create password-protected zip file of JPEGs
            var folderPathZip = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "PML31460", "PML31460_" + ContractId);
            var zipPath = Path.Combine(folderPathZip, $"PML31460_{ContractId}_JPEG_Preview.zip");
            if (System.IO.File.Exists(zipPath))
            {
                System.IO.File.Delete(zipPath);
            }
            var jpegFiles = Directory.GetFiles(folderPath, $"PML31460_{ContractId}_p*.jpg");
            string password = _configuration["Password:PaswordPDF"];

            using (var fsOut = System.IO.File.Create(zipPath))
            using (var zipStream = new ICSharpCode.SharpZipLib.Zip.ZipOutputStream(fsOut))
            {
                zipStream.SetLevel(9); // 0-9, 9 = best compression
                zipStream.Password = password; // Set password

                foreach (var file in jpegFiles)
                {
                    var entry = new ICSharpCode.SharpZipLib.Zip.ZipEntry(Path.GetFileName(file))
                    {
                        DateTime = DateTime.Now
                    };
                    zipStream.PutNextEntry(entry);

                    byte[] buffer = System.IO.File.ReadAllBytes(file);
                    zipStream.Write(buffer, 0, buffer.Length);
                    zipStream.CloseEntry();
                }
                zipStream.IsStreamOwner = true;
                zipStream.Close();
            }
            // 7. Delete JPEG files after zipping
            foreach (var file in jpegFiles)
            {
                System.IO.File.Delete(file);
            }

            // 8. Return the password-protected zip file as download
            var zipBytes = await System.IO.File.ReadAllBytesAsync(zipPath);
            return File(zipBytes, "application/zip", $"PML31460_{ContractId}_JPEG_Preview.zip");
        }
        public async Task<IActionResult> OnGetWordContact_PML31460_Word(string ContractId = "1")
        {
            // 1. Get HTML content for PML31460 contract
            var htmlContent = await _wordEContract_LoanPrinterService.OnGetWordContact_LoanPrinter_ToPDF(ContractId, "PML31460");

            // 2. Create a Word document from HTML using Spire.Doc
            var document = new Spire.Doc.Document();
            document.LoadFromStream(new MemoryStream(System.Text.Encoding.UTF8.GetBytes(htmlContent)), Spire.Doc.FileFormat.Html);

            // 3. Save the Word document to a MemoryStream
            using var ms = new MemoryStream();
            document.SaveToStream(ms, Spire.Doc.FileFormat.Docx);
            var wordBytes = ms.ToArray();

            // 4. Prepare folder structure
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "PML31460", "PML31460_" + ContractId);
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var filePath = Path.Combine(folderPath, $"PML31460_{ContractId}.docx");

            // 5. Save Word file to disk
            if (System.IO.File.Exists(filePath))
            {
                System.IO.File.Delete(filePath);
            }
            await System.IO.File.WriteAllBytesAsync(filePath, wordBytes);

            // 6. Return the Word file as download
            var resultBytes = await System.IO.File.ReadAllBytesAsync(filePath);
            return File(resultBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", $"PML31460_{ContractId}.docx");
        }
        public async Task<IActionResult> OnGetWordContact_PML31460_Word_Preview(string ContractId = "1")
        {
            // 1. Get HTML content for PML31460 contract
            var htmlContent = await _wordEContract_LoanPrinterService.OnGetWordContact_LoanPrinter_ToPDF(ContractId, "PML31460");

            // 2. Create a Word document from HTML using Spire.Doc
            var document = new Spire.Doc.Document();
            document.LoadFromStream(new MemoryStream(System.Text.Encoding.UTF8.GetBytes(htmlContent)), Spire.Doc.FileFormat.Html);

            // 3. Get password from appsettings.json
            string? userPassword = _configuration["Password:PaswordPDF"];
            if (!string.IsNullOrEmpty(userPassword))
            {
                document.Encrypt(userPassword);
            }

            // 4. Prepare folder structure
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "PML31460", "PML31460_" + ContractId);
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var filePath = Path.Combine(folderPath, $"PML31460_{ContractId}_Preview.docx");

            // 5. Delete the file if it already exists
            if (System.IO.File.Exists(filePath))
            {
                System.IO.File.Delete(filePath);
            }

            // 6. Save the password-protected Word document
            document.SaveToFile(filePath, FileFormat.Docx);

            // 7. Return the Word file as download
            var resultBytes = await System.IO.File.ReadAllBytesAsync(filePath);
            return File(resultBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", $"PML31460_{ContractId}_Preview.docx");
        }

        #endregion 4.1.1.2.13.สัญญาเช่าเครื่องถ่ายเอกสาร ร.314-60 PML31460

        #region 4.1.1.2.12.สัญญาจ้างบริการบำรุงรักษาและซ่อมแซมแก้ไขคอมพิวเตอร์ร.310-60 SMC31060
     
        public async Task OnGetWordContact_SMC31060_PDF(string ContractId = "1")
        {
            var wordBytes = await _maintenanceComputerService.OnGetWordContact_MaintenanceComputer_ToPDF(ContractId, "SMC31060");
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

        public async Task<IActionResult> OnGetWordContact_SMC31060_PDF_Preview(string ContractId = "1", string Name = "สมใจ ทดสอบ")
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
            string? userPassword = await GetPdfPasswordAsync(Name);;

            using (var inputStream = new MemoryStream(wordBytes))
            using (var outputStream = new MemoryStream())
            {
                var document = PdfSharpCore.Pdf.IO.PdfReader.Open(inputStream, PdfSharpCore.Pdf.IO.PdfDocumentOpenMode.Modify);

                // Add watermark to each page
                foreach (var page in document.Pages)
                {
                    using (var gfx = PdfSharpCore.Drawing.XGraphics.FromPdfPage(page))
                    {
                        var font = new PdfSharpCore.Drawing.XFont("Tahoma", 25, PdfSharpCore.Drawing.XFontStyle.Bold);
                        var text = $"สัญญาอิเลคทรอนิกส์พิมพ์ออกโดย {Name}\nวันที่ {DateTime.Now:dd/MM/yyyy HH:mm}";
                        var lines = text.Split('\n');

                        // Measure the height of one line
                        double lineHeight = font.GetHeight();

                        // Calculate total height for centering
                        double totalHeight = lineHeight * lines.Length;
                        double y = (page.Height - totalHeight) / 2;

                        // Center horizontally
                        foreach (var line in lines)
                        {
                            var size = gfx.MeasureString(line, font);
                            double x = (page.Width - size.Width) / 2;

                            // Draw the watermark diagonally with transparency
                            var state = gfx.Save();
                            gfx.TranslateTransform(page.Width / 2, page.Height / 2);
                            gfx.RotateTransform(-30);
                            gfx.TranslateTransform(-page.Width / 2, -page.Height / 2);

                            var brush = new PdfSharpCore.Drawing.XSolidBrush(
                                PdfSharpCore.Drawing.XColor.FromArgb(80, 255, 0, 0)); // semi-transparent red

                            gfx.DrawString(line, font, brush, x, y);
                            gfx.Restore(state);

                            y += lineHeight;
                        }
                    }
                }

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
        public async Task<IActionResult> OnGetWordContact_SMC31060_JPEG(string ContractId = "1")
        {
            // 1. Generate PDF from SMC31060 contract
            var pdfBytes = await _maintenanceComputerService.OnGetWordContact_MaintenanceComputer_ToPDF(ContractId, "SMC31060");

            // 2. Prepare folder structure
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "SMC31060", "SMC31060_" + ContractId, "SMC31060_" + ContractId + "_JPEG");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var pdfPath = Path.Combine(folderPath, $"SMC31060_{ContractId}.pdf");

            // 3. Save PDF to disk
            if (System.IO.File.Exists(pdfPath))
            {
                System.IO.File.Delete(pdfPath);
            }
            await System.IO.File.WriteAllBytesAsync(pdfPath, pdfBytes);

            // 4. Convert each PDF page to JPEG
            using (var pdfStream = new MemoryStream(pdfBytes))
            using (var document = PdfiumViewer.PdfDocument.Load(pdfStream))
            {
                for (int i = 0; i < document.PageCount; i++)
                {
                    using (var image = document.Render(i, 300, 300, true))
                    {
                        var jpegPath = Path.Combine(folderPath, $"SMC31060_{ContractId}_p{i + 1}.jpg");
                        image.Save(jpegPath, System.Drawing.Imaging.ImageFormat.Jpeg);
                    }
                }
            }
            // 5. Delete the PDF file after conversion
            System.IO.File.Delete(pdfPath);

            // 6. Create zip file of JPEGs
            var folderPathZip = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "SMC31060", "SMC31060_" + ContractId);
            var zipPath = Path.Combine(folderPathZip, $"SMC31060_{ContractId}_JPEG.zip");
            if (System.IO.File.Exists(zipPath))
            {
                System.IO.File.Delete(zipPath);
            }
            var jpegFiles = Directory.GetFiles(folderPath, $"SMC31060_{ContractId}_p*.jpg");
            using (var zip = System.IO.Compression.ZipFile.Open(zipPath, ZipArchiveMode.Create))
            {
                foreach (var file in jpegFiles)
                {
                    zip.CreateEntryFromFile(file, Path.GetFileName(file));
                }
            }
            // 7. Delete JPEG files after zipping
            foreach (var file in jpegFiles)
            {
                System.IO.File.Delete(file);
            }

            // 8. Return the zip file as download
            var zipBytes = await System.IO.File.ReadAllBytesAsync(zipPath);
            return File(zipBytes, "application/zip", $"SMC31060_{ContractId}_JPEG.zip");
        }
        public async Task<IActionResult> OnGetWordContact_SMC31060_JPEG_Preview(string ContractId = "1")
        {
            // 1. Generate PDF from SMC31060 contract
            var pdfBytes = await _maintenanceComputerService.OnGetWordContact_MaintenanceComputer_ToPDF(ContractId, "SMC31060");

            // 2. Prepare folder structure
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "SMC31060", "SMC31060_" + ContractId, "SMC31060_" + ContractId + "_JPEG");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var pdfPath = Path.Combine(folderPath, $"SMC31060_{ContractId}.pdf");

            // 3. Save PDF to disk
            if (System.IO.File.Exists(pdfPath))
            {
                System.IO.File.Delete(pdfPath);
            }
            await System.IO.File.WriteAllBytesAsync(pdfPath, pdfBytes);

            // 4. Convert each PDF page to JPEG
            using (var pdfStream = new MemoryStream(pdfBytes))
            using (var document = PdfiumViewer.PdfDocument.Load(pdfStream))
            {
                for (int i = 0; i < document.PageCount; i++)
                {
                    using (var image = document.Render(i, 300, 300, true))
                    {
                        var jpegPath = Path.Combine(folderPath, $"SMC31060_{ContractId}_p{i + 1}.jpg");
                        image.Save(jpegPath, System.Drawing.Imaging.ImageFormat.Jpeg);
                    }
                }
            }
            // 5. Delete the PDF file after conversion
            System.IO.File.Delete(pdfPath);

            // 6. Create password-protected zip file of JPEGs
            var folderPathZip = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "SMC31060", "SMC31060_" + ContractId);
            var zipPath = Path.Combine(folderPathZip, $"SMC31060_{ContractId}_JPEG_Preview.zip");
            if (System.IO.File.Exists(zipPath))
            {
                System.IO.File.Delete(zipPath);
            }
            var jpegFiles = Directory.GetFiles(folderPath, $"SMC31060_{ContractId}_p*.jpg");
            string password = _configuration["Password:PaswordPDF"];

            using (var fsOut = System.IO.File.Create(zipPath))
            using (var zipStream = new ICSharpCode.SharpZipLib.Zip.ZipOutputStream(fsOut))
            {
                zipStream.SetLevel(9); // 0-9, 9 = best compression
                zipStream.Password = password; // Set password

                foreach (var file in jpegFiles)
                {
                    var entry = new ICSharpCode.SharpZipLib.Zip.ZipEntry(Path.GetFileName(file))
                    {
                        DateTime = DateTime.Now
                    };
                    zipStream.PutNextEntry(entry);

                    byte[] buffer = System.IO.File.ReadAllBytes(file);
                    zipStream.Write(buffer, 0, buffer.Length);
                    zipStream.CloseEntry();
                }
                zipStream.IsStreamOwner = true;
                zipStream.Close();
            }
            // 7. Delete JPEG files after zipping
            foreach (var file in jpegFiles)
            {
                System.IO.File.Delete(file);
            }

            // 8. Return the password-protected zip file as download
            var zipBytes = await System.IO.File.ReadAllBytesAsync(zipPath);
            return File(zipBytes, "application/zip", $"SMC31060_{ContractId}_JPEG_Preview.zip");
        }
        public async Task<IActionResult> OnGetWordContact_SMC31060_Word(string ContractId = "1")
        {
            // 1. Get the Word document for SMC31060 contract
            var wordBytes = await _maintenanceComputerService.OnGetWordContact_MaintenanceComputer(ContractId);

            // 2. Prepare folder structure
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "SMC31060", "SMC31060_" + ContractId);
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var filePath = Path.Combine(folderPath, $"SMC31060_{ContractId}.docx");

            // 3. Save Word file to disk
            if (System.IO.File.Exists(filePath))
            {
                System.IO.File.Delete(filePath);
            }
            await System.IO.File.WriteAllBytesAsync(filePath, wordBytes);

            // 4. Return the Word file as download
            var resultBytes = await System.IO.File.ReadAllBytesAsync(filePath);
            return File(resultBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", $"SMC31060_{ContractId}.docx");
        }
        public async Task<IActionResult> OnGetWordContact_SMC31060_Word_Preview(string ContractId = "1")
        {
            // 1. Get the Word document for SMC31060 contract
            var wordBytes = await _maintenanceComputerService.OnGetWordContact_MaintenanceComputer(ContractId);

            // 2. Prepare folder structure
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "SMC31060", "SMC31060_" + ContractId);
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var filePath = Path.Combine(folderPath, $"SMC31060_{ContractId}_Preview.docx");

            // 3. Delete the file if it already exists
            if (System.IO.File.Exists(filePath))
            {
                System.IO.File.Delete(filePath);
            }

            // 4. Get password from appsettings.json
            string? userPassword = _configuration["Password:PaswordPDF"];

            // 5. Load the Word file from memory and apply password protection
            using (var ms = new MemoryStream(wordBytes))
            {
                Document doc = new Document();
                doc.LoadFromStream(ms, FileFormat.Docx);

                // Apply password protection
                doc.Encrypt(userPassword);

                // Save the password-protected file
                doc.SaveToFile(filePath, FileFormat.Docx);
            }

            // 6. Return the password-protected Word file as download
            var resultBytes = await System.IO.File.ReadAllBytesAsync(filePath);
            return File(resultBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", $"SMC31060_{ContractId}_Preview.docx");
        }

        #endregion 4.1.1.2.12.สัญญาจ้างบริการบำรุงรักษาและซ่อมแซมแก้ไขคอมพิวเตอร์ร.310-60 SMC31060

        #region 4.1.1.2.11.สัญญาเช่าคอมพิวเตอร์ ร.309-60 CLA30960
       
        public async Task OnGetWordContact_CLA30960_PDF(string ContractId = "1")
        {
            var htmlContent = await _LoanComputerService.OnGetWordContact_LoanComputer_ToPDF(ContractId, "CLA30960");
            await new BrowserFetcher().DownloadAsync();
            await using var browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = true });
            await using var page = await browser.NewPageAsync();

            await page.SetContentAsync(htmlContent);

            var pdfOptions = new PdfOptions
            {
                Format = PuppeteerSharp.Media.PaperFormat.A4,

                Landscape = false,
                MarginOptions = new PuppeteerSharp.Media.MarginOptions
                {
                    Top = "20mm",
                    Bottom = "20mm",
                    Left = "20mm",
                    Right = "20mm"
                },
                PrintBackground = true

            };

            var pdfBytes = await page.PdfDataAsync(pdfOptions);
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
            await System.IO.File.WriteAllBytesAsync(filePath, pdfBytes);

            // return File(wordBytes, "application/pdf", "สัญญาซื้อขาย.pdf");
        }
        public async Task<IActionResult> OnGetWordContact_CLA30960_PDF_Preview(string ContractId = "1", string Name = "สมใจ ทดสอบ")
        {
            var htmlContent = await _LoanComputerService.OnGetWordContact_LoanComputer_ToPDF(ContractId, "CLA30960");
            await new BrowserFetcher().DownloadAsync();
            await using var browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = true });
            await using var page = await browser.NewPageAsync();

            await page.SetContentAsync(htmlContent);

            var pdfOptions = new PdfOptions
            {
                Format = PuppeteerSharp.Media.PaperFormat.A4,

                Landscape = false,
                MarginOptions = new PuppeteerSharp.Media.MarginOptions
                {
                    Top = "20mm",
                    Bottom = "20mm",
                    Left = "20mm",
                    Right = "20mm"
                },
                PrintBackground = true

            };

            var pdfBytes = await page.PdfDataAsync(pdfOptions);
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "CLA30960");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var fileName = $"CLA30960_{ContractId}_Preview.pdf";
            var filePath = Path.Combine(folderPath, fileName);

            // Get password from appsettings.json
            string? userPassword = await GetPdfPasswordAsync(Name);

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
        public async Task<IActionResult> OnGetWordContact_CLA30960_JPEG(string ContractId = "1")
        {
            // 1. Generate PDF from CLA30960 contract
            var htmlContent = await _LoanComputerService.OnGetWordContact_LoanComputer_ToPDF(ContractId, "CLA30960");
            await new BrowserFetcher().DownloadAsync();
            await using var browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = true });
            await using var page = await browser.NewPageAsync();

            await page.SetContentAsync(htmlContent);

            var pdfOptions = new PdfOptions
            {
                Format = PuppeteerSharp.Media.PaperFormat.A4,

                Landscape = false,
                MarginOptions = new PuppeteerSharp.Media.MarginOptions
                {
                    Top = "20mm",
                    Bottom = "20mm",
                    Left = "20mm",
                    Right = "20mm"
                },
                PrintBackground = true

            };

            var pdfBytes = await page.PdfDataAsync(pdfOptions);

            // 2. Prepare folder structure
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "CLA30960", "CLA30960_" + ContractId, "CLA30960_" + ContractId + "_JPEG");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var pdfPath = Path.Combine(folderPath, $"CLA30960_{ContractId}.pdf");

            // 3. Save PDF to disk
            if (System.IO.File.Exists(pdfPath))
            {
                System.IO.File.Delete(pdfPath);
            }
            await System.IO.File.WriteAllBytesAsync(pdfPath, pdfBytes);

            // 4. Convert each PDF page to JPEG
            using (var pdfStream = new MemoryStream(pdfBytes))
            using (var document = PdfiumViewer.PdfDocument.Load(pdfStream))
            {
                for (int i = 0; i < document.PageCount; i++)
                {
                    using (var image = document.Render(i, 300, 300, true))
                    {
                        var jpegPath = Path.Combine(folderPath, $"CLA30960_{ContractId}_p{i + 1}.jpg");
                        image.Save(jpegPath, System.Drawing.Imaging.ImageFormat.Jpeg);
                    }
                }
            }
            // 5. Delete the PDF file after conversion
            System.IO.File.Delete(pdfPath);

            // 6. Create zip file of JPEGs
            var folderPathZip = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "CLA30960", "CLA30960_" + ContractId);
            var zipPath = Path.Combine(folderPathZip, $"CLA30960_{ContractId}_JPEG.zip");
            if (System.IO.File.Exists(zipPath))
            {
                System.IO.File.Delete(zipPath);
            }
            var jpegFiles = Directory.GetFiles(folderPath, $"CLA30960_{ContractId}_p*.jpg");
            using (var zip = System.IO.Compression.ZipFile.Open(zipPath, ZipArchiveMode.Create))
            {
                foreach (var file in jpegFiles)
                {
                    zip.CreateEntryFromFile(file, Path.GetFileName(file));
                }
            }
            // 7. Delete JPEG files after zipping
            foreach (var file in jpegFiles)
            {
                System.IO.File.Delete(file);
            }

            // 8. Return the zip file as download
            var zipBytes = await System.IO.File.ReadAllBytesAsync(zipPath);
            return File(zipBytes, "application/zip", $"CLA30960_{ContractId}_JPEG.zip");
        }

        public async Task<IActionResult> OnGetWordContact_CLA30960_JPEG_Preview(string ContractId = "1")
        {
            // 1. Generate PDF from CLA30960 contract
            var htmlContent = await _LoanComputerService.OnGetWordContact_LoanComputer_ToPDF(ContractId, "CLA30960");
            await new BrowserFetcher().DownloadAsync();
            await using var browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = true });
            await using var page = await browser.NewPageAsync();

            await page.SetContentAsync(htmlContent);

            var pdfOptions = new PdfOptions
            {
                Format = PuppeteerSharp.Media.PaperFormat.A4,

                Landscape = false,
                MarginOptions = new PuppeteerSharp.Media.MarginOptions
                {
                    Top = "20mm",
                    Bottom = "20mm",
                    Left = "20mm",
                    Right = "20mm"
                },
                PrintBackground = true

            };

            var pdfBytes = await page.PdfDataAsync(pdfOptions);

            // 2. Prepare folder structure
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "CLA30960", "CLA30960_" + ContractId, "CLA30960_" + ContractId + "_JPEG");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var pdfPath = Path.Combine(folderPath, $"CLA30960_{ContractId}.pdf");

            // 3. Save PDF to disk
            if (System.IO.File.Exists(pdfPath))
            {
                System.IO.File.Delete(pdfPath);
            }
            await System.IO.File.WriteAllBytesAsync(pdfPath, pdfBytes);

            // 4. Convert each PDF page to JPEG
            using (var pdfStream = new MemoryStream(pdfBytes))
            using (var document = PdfiumViewer.PdfDocument.Load(pdfStream))
            {
                for (int i = 0; i < document.PageCount; i++)
                {
                    using (var image = document.Render(i, 300, 300, true))
                    {
                        var jpegPath = Path.Combine(folderPath, $"CLA30960_{ContractId}_p{i + 1}.jpg");
                        image.Save(jpegPath, System.Drawing.Imaging.ImageFormat.Jpeg);
                    }
                }
            }
            // 5. Delete the PDF file after conversion
            System.IO.File.Delete(pdfPath);

            // 6. Create password-protected zip file of JPEGs
            var folderPathZip = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "CLA30960", "CLA30960_" + ContractId);
            var zipPath = Path.Combine(folderPathZip, $"CLA30960_{ContractId}_JPEG_Preview.zip");
            if (System.IO.File.Exists(zipPath))
            {
                System.IO.File.Delete(zipPath);
            }
            var jpegFiles = Directory.GetFiles(folderPath, $"CLA30960_{ContractId}_p*.jpg");
            string password = _configuration["Password:PaswordPDF"];

            using (var fsOut = System.IO.File.Create(zipPath))
            using (var zipStream = new ZipOutputStream(fsOut))
            {
                zipStream.SetLevel(9); // 0-9, 9 = best compression
                zipStream.Password = password; // Set password

                foreach (var file in jpegFiles)
                {
                    var entry = new ZipEntry(Path.GetFileName(file))
                    {
                        DateTime = DateTime.Now
                    };
                    zipStream.PutNextEntry(entry);

                    byte[] buffer = System.IO.File.ReadAllBytes(file);
                    zipStream.Write(buffer, 0, buffer.Length);
                    zipStream.CloseEntry();
                }
                zipStream.IsStreamOwner = true;
                zipStream.Close();
            }
            // 7. Delete JPEG files after zipping
            foreach (var file in jpegFiles)
            {
                System.IO.File.Delete(file);
            }

            // 8. Return the password-protected zip file as download
            var zipBytes = await System.IO.File.ReadAllBytesAsync(zipPath);
            return File(zipBytes, "application/zip", $"CLA30960_{ContractId}_JPEG_Preview.zip");
        }

        public async Task<IActionResult> OnGetWordContact_CLA30960_Word(string ContractId = "1")
        {
            // 1. Get HTML content for CLA30960 contract
            var htmlContent = await _LoanComputerService.OnGetWordContact_LoanComputer_ToPDF(ContractId, "CLA30960");

            // 2. Create a Word document from HTML using Spire.Doc
            var document = new Spire.Doc.Document();
            document.LoadFromStream(new MemoryStream(System.Text.Encoding.UTF8.GetBytes(htmlContent)), Spire.Doc.FileFormat.Html);

            // 3. Save the Word document to a MemoryStream
            using var ms = new MemoryStream();
            document.SaveToStream(ms, Spire.Doc.FileFormat.Docx);
            var wordBytes = ms.ToArray();

            // 4. Prepare folder structure
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "CLA30960", "CLA30960_" + ContractId);
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var filePath = Path.Combine(folderPath, $"CLA30960_{ContractId}.docx");

            // 5. Save Word file to disk
            if (System.IO.File.Exists(filePath))
            {
                System.IO.File.Delete(filePath);
            }
            await System.IO.File.WriteAllBytesAsync(filePath, wordBytes);

            // 6. Return the Word file as download
            var resultBytes = await System.IO.File.ReadAllBytesAsync(filePath);
            return File(resultBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", $"CLA30960_{ContractId}.docx");
        }

        public async Task<IActionResult> OnGetWordContact_CLA30960_Word_Preview(string ContractId = "1")
        {
            // 1. Get HTML content for CLA30960 contract
            var htmlContent = await _LoanComputerService.OnGetWordContact_LoanComputer_ToPDF(ContractId, "CLA30960");

            // 2. Create a Word document from HTML using Spire.Doc
            var document = new Spire.Doc.Document();
            document.LoadFromStream(new MemoryStream(System.Text.Encoding.UTF8.GetBytes(htmlContent)), Spire.Doc.FileFormat.Html);

            // 3. Get password from appsettings.json
            string? userPassword = _configuration["Password:PaswordPDF"];
            if (!string.IsNullOrEmpty(userPassword))
            {
                document.Encrypt(userPassword);
            }

            // 4. Prepare folder structure
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "CLA30960", "CLA30960_" + ContractId);
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var filePath = Path.Combine(folderPath, $"CLA30960_{ContractId}_Preview.docx");

            // 5. Delete the file if it already exists
            if (System.IO.File.Exists(filePath))
            {
                System.IO.File.Delete(filePath);
            }

            // 6. Save the password-protected Word document
            document.SaveToFile(filePath, FileFormat.Docx);

            // 7. Return the Word file as download
            var resultBytes = await System.IO.File.ReadAllBytesAsync(filePath);
            return File(resultBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", $"CLA30960_{ContractId}_Preview.docx");
        }      
        
        #endregion 4.1.1.2.11.สัญญาเช่าคอมพิวเตอร์ ร.309-60 CLA30960

        #region 4.1.1.2.10.สัญญาซื้อขายและอนุญาตให้ใช้สิทธิในโปรแกรมคอมพิวเตอร์ ร.308-60 SLA30860
    
        public async Task OnGetWordContact_SLA30860_PDF(string ContractId = "1")
        {
            var htmlContent = await _BuyAgreeProgram.OnGetWordContact_BuyAgreeProgram_ToPDF(ContractId, "SLA30860");
            await new BrowserFetcher().DownloadAsync();
            await using var browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = true });
            await using var page = await browser.NewPageAsync();

            await page.SetContentAsync(htmlContent);

            var pdfOptions = new PdfOptions
            {
                Format = PuppeteerSharp.Media.PaperFormat.A4,

                Landscape = false,
                MarginOptions = new PuppeteerSharp.Media.MarginOptions
                {
                    Top = "20mm",
                    Bottom = "20mm",
                    Left = "20mm",
                    Right = "20mm"
                },
                PrintBackground = true

            };

            var pdfBytes = await page.PdfDataAsync(pdfOptions);

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
            await System.IO.File.WriteAllBytesAsync(filePath, pdfBytes);

            // return File(wordBytes, "application/pdf", "สัญญาซื้อขาย.pdf");
        }

        public async Task<IActionResult> OnGetWordContact_SLA30860_PDF_Preview(string ContractId = "1", string Name = "สมใจ ทดสอบ")
        {
            var htmlContent = await _BuyAgreeProgram.OnGetWordContact_BuyAgreeProgram_ToPDF(ContractId, "SLA30860");
            await new BrowserFetcher().DownloadAsync();
            await using var browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = true });
            await using var page = await browser.NewPageAsync();

            await page.SetContentAsync(htmlContent);

            var pdfOptions = new PdfOptions
            {
                Format = PuppeteerSharp.Media.PaperFormat.A4,

                Landscape = false,
                MarginOptions = new PuppeteerSharp.Media.MarginOptions
                {
                    Top = "20mm",
                    Bottom = "20mm",
                    Left = "20mm",
                    Right = "20mm"
                },
                PrintBackground = true

            };

            var pdfBytes = await page.PdfDataAsync(pdfOptions);

            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "SLA30860");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var fileName = $"SLA30860_{ContractId}_Preview.pdf";
            var filePath = Path.Combine(folderPath, fileName);

            // Get password from appsettings.json
            string? userPassword = await GetPdfPasswordAsync(Name);;

            using (var inputStream = new MemoryStream(pdfBytes))
            using (var outputStream = new MemoryStream())
            {
                var document = PdfSharpCore.Pdf.IO.PdfReader.Open(inputStream, PdfSharpCore.Pdf.IO.PdfDocumentOpenMode.Modify);

                foreach (var pdfPage in document.Pages)
                {
                    using (var gfx = PdfSharpCore.Drawing.XGraphics.FromPdfPage(pdfPage))
                    {
                        var font = new PdfSharpCore.Drawing.XFont("Tahoma", 25, PdfSharpCore.Drawing.XFontStyle.Bold);
                        var text = $"สัญญาอิเลคทรอนิกส์พิมพ์ออกโดย {Name}\nวันที่ {DateTime.Now:dd/MM/yyyy HH:mm}";
                        var lines = text.Split('\n');
                        double lineHeight = font.GetHeight();
                        double totalHeight = lineHeight * lines.Length;
                        double y = (pdfPage.Height - totalHeight) / 2;

                        foreach (var line in lines)
                        {
                            var size = gfx.MeasureString(line, font);
                            double x = (pdfPage.Width - size.Width) / 2;
                            var state = gfx.Save();
                            gfx.TranslateTransform(pdfPage.Width / 2, pdfPage.Height / 2);
                            gfx.RotateTransform(-30);
                            gfx.TranslateTransform(-pdfPage.Width / 2, -pdfPage.Height / 2);

                            var brush = new PdfSharpCore.Drawing.XSolidBrush(
                                PdfSharpCore.Drawing.XColor.FromArgb(80, 255, 0, 0)); // semi-transparent red

                            gfx.DrawString(line, font, brush, x, y);
                            gfx.Restore(state);

                            y += lineHeight;
                        }
                    }
                }

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

        public async Task<IActionResult> OnGetWordContact_SLA30860_JPEG(string ContractId = "1")
        {
            // 1. Generate PDF from SLA30860 contract
            var htmlContent = await _BuyAgreeProgram.OnGetWordContact_BuyAgreeProgram_ToPDF(ContractId, "SLA30860");
            await new BrowserFetcher().DownloadAsync();
            await using var browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = true });
            await using var page = await browser.NewPageAsync();

            await page.SetContentAsync(htmlContent);

            var pdfOptions = new PdfOptions
            {
                Format = PuppeteerSharp.Media.PaperFormat.A4,

                Landscape = false,
                MarginOptions = new PuppeteerSharp.Media.MarginOptions
                {
                    Top = "20mm",
                    Bottom = "20mm",
                    Left = "20mm",
                    Right = "20mm"
                },
                PrintBackground = true

            };

            var pdfBytes = await page.PdfDataAsync(pdfOptions);

            // 2. Prepare folder structure
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "SLA30860", "SLA30860_" + ContractId, "SLA30860_" + ContractId + "_JPEG");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var pdfPath = Path.Combine(folderPath, $"SLA30860_{ContractId}.pdf");

            // 3. Save PDF to disk
            if (System.IO.File.Exists(pdfPath))
            {
                System.IO.File.Delete(pdfPath);
            }
            await System.IO.File.WriteAllBytesAsync(pdfPath, pdfBytes);

            // 4. Convert each PDF page to JPEG
            using (var pdfStream = new MemoryStream(pdfBytes))
            using (var document = PdfiumViewer.PdfDocument.Load(pdfStream))
            {
                for (int i = 0; i < document.PageCount; i++)
                {
                    using (var image = document.Render(i, 300, 300, true))
                    {
                        var jpegPath = Path.Combine(folderPath, $"SLA30860_{ContractId}_p{i + 1}.jpg");
                        image.Save(jpegPath, System.Drawing.Imaging.ImageFormat.Jpeg);
                    }
                }
            }
            // 5. Delete the PDF file after conversion
            System.IO.File.Delete(pdfPath);

            // 6. Create zip file of JPEGs
            var folderPathZip = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "SLA30860", "SLA30860_" + ContractId);
            var zipPath = Path.Combine(folderPathZip, $"SLA30860_{ContractId}_JPEG.zip");
            if (System.IO.File.Exists(zipPath))
            {
                System.IO.File.Delete(zipPath);
            }
            var jpegFiles = Directory.GetFiles(folderPath, $"SLA30860_{ContractId}_p*.jpg");
            using (var zip = System.IO.Compression.ZipFile.Open(zipPath, ZipArchiveMode.Create))
            {
                foreach (var file in jpegFiles)
                {
                    zip.CreateEntryFromFile(file, Path.GetFileName(file));
                }
            }
            // 7. Delete JPEG files after zipping
            foreach (var file in jpegFiles)
            {
                System.IO.File.Delete(file);
            }

            // 8. Return the zip file as download
            var zipBytes = await System.IO.File.ReadAllBytesAsync(zipPath);
            return File(zipBytes, "application/zip", $"SLA30860_{ContractId}_JPEG.zip");
        }
        public async Task<IActionResult> OnGetWordContact_SLA30860_JPEG_Preview(string ContractId = "1")
        {
            // 1. Generate PDF from SLA30860 contract
            var htmlContent = await _BuyAgreeProgram.OnGetWordContact_BuyAgreeProgram_ToPDF(ContractId, "SLA30860");
            await new BrowserFetcher().DownloadAsync();
            await using var browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = true });
            await using var page = await browser.NewPageAsync();

            await page.SetContentAsync(htmlContent);

            var pdfOptions = new PdfOptions
            {
                Format = PuppeteerSharp.Media.PaperFormat.A4,

                Landscape = false,
                MarginOptions = new PuppeteerSharp.Media.MarginOptions
                {
                    Top = "20mm",
                    Bottom = "20mm",
                    Left = "20mm",
                    Right = "20mm"
                },
                PrintBackground = true

            };

            var pdfBytes = await page.PdfDataAsync(pdfOptions);

            // 2. Prepare folder structure
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "SLA30860", "SLA30860_" + ContractId, "SLA30860_" + ContractId + "_JPEG");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var pdfPath = Path.Combine(folderPath, $"SLA30860_{ContractId}.pdf");

            // 3. Save PDF to disk
            if (System.IO.File.Exists(pdfPath))
            {
                System.IO.File.Delete(pdfPath);
            }
            await System.IO.File.WriteAllBytesAsync(pdfPath, pdfBytes);

            // 4. Convert each PDF page to JPEG
            using (var pdfStream = new MemoryStream(pdfBytes))
            using (var document = PdfiumViewer.PdfDocument.Load(pdfStream))
            {
                for (int i = 0; i < document.PageCount; i++)
                {
                    using (var image = document.Render(i, 300, 300, true))
                    {
                        var jpegPath = Path.Combine(folderPath, $"SLA30860_{ContractId}_p{i + 1}.jpg");
                        image.Save(jpegPath, System.Drawing.Imaging.ImageFormat.Jpeg);
                    }
                }
            }
            // 5. Delete the PDF file after conversion
            System.IO.File.Delete(pdfPath);

            // 6. Create password-protected zip file of JPEGs
            var folderPathZip = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "SLA30860", "SLA30860_" + ContractId);
            var zipPath = Path.Combine(folderPathZip, $"SLA30860_{ContractId}_JPEG_Preview.zip");
            if (System.IO.File.Exists(zipPath))
            {
                System.IO.File.Delete(zipPath);
            }
            var jpegFiles = Directory.GetFiles(folderPath, $"SLA30860_{ContractId}_p*.jpg");
            string password = _configuration["Password:PaswordPDF"];

            using (var fsOut = System.IO.File.Create(zipPath))
            using (var zipStream = new ZipOutputStream(fsOut))
            {
                zipStream.SetLevel(9); // 0-9, 9 = best compression
                zipStream.Password = password; // Set password

                foreach (var file in jpegFiles)
                {
                    var entry = new ZipEntry(Path.GetFileName(file))
                    {
                        DateTime = DateTime.Now
                    };
                    zipStream.PutNextEntry(entry);

                    byte[] buffer = System.IO.File.ReadAllBytes(file);
                    zipStream.Write(buffer, 0, buffer.Length);
                    zipStream.CloseEntry();
                }
                zipStream.IsStreamOwner = true;
                zipStream.Close();
            }
            // 7. Delete JPEG files after zipping
            foreach (var file in jpegFiles)
            {
                System.IO.File.Delete(file);
            }

            // 8. Return the password-protected zip file as download
            var zipBytes = await System.IO.File.ReadAllBytesAsync(zipPath);
            return File(zipBytes, "application/zip", $"SLA30860_{ContractId}_JPEG_Preview.zip");
        }
        public async Task<IActionResult> OnGetWordContact_SLA30860_Word(string ContractId = "1")
        {
            // 1. Get HTML content for SLA30860 contract
            var htmlContent = await _BuyAgreeProgram.OnGetWordContact_BuyAgreeProgram_ToPDF(ContractId, "SLA30860");

            // 2. Create a Word document from HTML using Spire.Doc
            var document = new Spire.Doc.Document();
            document.LoadFromStream(new MemoryStream(System.Text.Encoding.UTF8.GetBytes(htmlContent)), Spire.Doc.FileFormat.Html);

            // 3. Save the Word document to a MemoryStream
            using var ms = new MemoryStream();
            document.SaveToStream(ms, Spire.Doc.FileFormat.Docx);
            var wordBytes = ms.ToArray();

            // 4. Prepare folder structure
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "SLA30860", "SLA30860_" + ContractId);
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var filePath = Path.Combine(folderPath, $"SLA30860_{ContractId}.docx");

            // 5. Save Word file to disk
            if (System.IO.File.Exists(filePath))
            {
                System.IO.File.Delete(filePath);
            }
            await System.IO.File.WriteAllBytesAsync(filePath, wordBytes);

            // 6. Return the Word file as download
            var resultBytes = await System.IO.File.ReadAllBytesAsync(filePath);
            return File(resultBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", $"SLA30860_{ContractId}.docx");
        }
        public async Task<IActionResult> OnGetWordContact_SLA30860_Word_Preview(string ContractId = "1")
        {
            // 1. Get HTML content for SLA30860 contract
            var htmlContent = await _BuyAgreeProgram.OnGetWordContact_BuyAgreeProgram_ToPDF(ContractId, "SLA30860");

            // 2. Create a Word document from HTML using Spire.Doc
            var document = new Spire.Doc.Document();
            document.LoadFromStream(new MemoryStream(System.Text.Encoding.UTF8.GetBytes(htmlContent)), Spire.Doc.FileFormat.Html);

            // 3. Get password from appsettings.json
            string? userPassword = _configuration["Password:PaswordPDF"];
            if (!string.IsNullOrEmpty(userPassword))
            {
                document.Encrypt(userPassword);
            }

            // 4. Prepare folder structure
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "SLA30860", "SLA30860_" + ContractId);
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var filePath = Path.Combine(folderPath, $"SLA30860_{ContractId}_Preview.docx");

            // 5. Delete the file if it already exists
            if (System.IO.File.Exists(filePath))
            {
                System.IO.File.Delete(filePath);
            }

            // 6. Save the password-protected Word document
            document.SaveToFile(filePath, FileFormat.Docx);

            // 7. Return the Word file as download
            var resultBytes = await System.IO.File.ReadAllBytesAsync(filePath);
            return File(resultBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", $"SLA30860_{ContractId}_Preview.docx");
        }

        #endregion 4.1.1.2.10.สัญญาซื้อขายและอนุญาตให้ใช้สิทธิในโปรแกรมคอมพิวเตอร์ ร.308-60 SLA30860

        #region 4.1.1.2.9.สัญญาซื้อขายคอมพิวเตอร์ CPA
     
        public async Task OnGetWordContact_CPA_PDF(string ContractId = "14")
        {
            var htmlContent = await _BuyOrSellComputerService.OnGetWordContact_BuyOrSellComputerService_ToPDF(ContractId);
            await new BrowserFetcher().DownloadAsync();
            await using var browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = true });
            await using var page = await browser.NewPageAsync();

            await page.SetContentAsync(htmlContent);

            var pdfOptions = new PdfOptions
            {
                Format = PuppeteerSharp.Media.PaperFormat.A4,

                Landscape = false,
                MarginOptions = new PuppeteerSharp.Media.MarginOptions
                {
                    Top = "20mm",
                    Bottom = "20mm",
                    Left = "20mm",
                    Right = "20mm"
                },
                PrintBackground = true

            };

            var pdfBytes = await page.PdfDataAsync(pdfOptions);

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
            await System.IO.File.WriteAllBytesAsync(filePath, pdfBytes);

            //return File(wordBytes, "application/pdf", "CPA_" + ContractId + ".pdf");
        }

        public async Task<IActionResult> OnGetWordContact_CPA_PDF_Preview(string ContractId = "14", string Name = "สมใจ ทดสอบ")
        {
            var htmlContent = await _BuyOrSellComputerService.OnGetWordContact_BuyOrSellComputerService_ToPDF(ContractId);
            await new BrowserFetcher().DownloadAsync();
            await using var browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = true });
            await using var page = await browser.NewPageAsync();

            await page.SetContentAsync(htmlContent);

            var pdfOptions = new PdfOptions
            {
                Format = PuppeteerSharp.Media.PaperFormat.A4,

                Landscape = false,
                MarginOptions = new PuppeteerSharp.Media.MarginOptions
                {
                    Top = "20mm",
                    Bottom = "20mm",
                    Left = "20mm",
                    Right = "20mm"
                },
                PrintBackground = true

            };

            var pdfBytes = await page.PdfDataAsync(pdfOptions);

            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "CPA");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var fileName = $"CPA_{ContractId}_Preview.pdf";
            var filePath = Path.Combine(folderPath, fileName);

            // Get password from appsettings.json
            string? userPassword = await GetPdfPasswordAsync(Name);;

            using (var inputStream = new MemoryStream(pdfBytes))
            using (var outputStream = new MemoryStream())
            {
                var document = PdfSharpCore.Pdf.IO.PdfReader.Open(inputStream, PdfSharpCore.Pdf.IO.PdfDocumentOpenMode.Modify);

                foreach (var pdfPage in document.Pages)
                {
                    using (var gfx = PdfSharpCore.Drawing.XGraphics.FromPdfPage(pdfPage))
                    {
                        var font = new PdfSharpCore.Drawing.XFont("Tahoma", 25, PdfSharpCore.Drawing.XFontStyle.Bold);
                        var text = $"สัญญาอิเลคทรอนิกส์พิมพ์ออกโดย {Name}\nวันที่ {DateTime.Now:dd/MM/yyyy HH:mm}";
                        var lines = text.Split('\n');
                        double lineHeight = font.GetHeight();
                        double totalHeight = lineHeight * lines.Length;
                        double y = (pdfPage.Height - totalHeight) / 2;

                        foreach (var line in lines)
                        {
                            var size = gfx.MeasureString(line, font);
                            double x = (pdfPage.Width - size.Width) / 2;
                            var state = gfx.Save();
                            gfx.TranslateTransform(pdfPage.Width / 2, pdfPage.Height / 2);
                            gfx.RotateTransform(-30);
                            gfx.TranslateTransform(-pdfPage.Width / 2, -pdfPage.Height / 2);

                            var brush = new PdfSharpCore.Drawing.XSolidBrush(
                                PdfSharpCore.Drawing.XColor.FromArgb(80, 255, 0, 0)); // semi-transparent red

                            gfx.DrawString(line, font, brush, x, y);
                            gfx.Restore(state);

                            y += lineHeight;
                        }
                    }
                }

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

        public async Task<IActionResult> OnGetWordContact_CPA_JPEG(string ContractId = "14")
        {
            // 1. Generate PDF from CPA contract
            var htmlContent = await _BuyOrSellComputerService.OnGetWordContact_BuyOrSellComputerService_ToPDF(ContractId);
            await new BrowserFetcher().DownloadAsync();
            await using var browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = true });
            await using var page = await browser.NewPageAsync();

            await page.SetContentAsync(htmlContent);

            var pdfOptions = new PdfOptions
            {
                Format = PuppeteerSharp.Media.PaperFormat.A4,

                Landscape = false,
                MarginOptions = new PuppeteerSharp.Media.MarginOptions
                {
                    Top = "20mm",
                    Bottom = "20mm",
                    Left = "20mm",
                    Right = "20mm"
                },
                PrintBackground = true

            };

            var pdfBytes = await page.PdfDataAsync(pdfOptions);
            // 2. Prepare folder structure
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "CPA", "CPA_" + ContractId, "CPA_" + ContractId + "_JPEG");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var pdfPath = Path.Combine(folderPath, $"CPA_{ContractId}.pdf");

            // 3. Save PDF to disk
            if (System.IO.File.Exists(pdfPath))
            {
                System.IO.File.Delete(pdfPath);
            }
            await System.IO.File.WriteAllBytesAsync(pdfPath, pdfBytes);

            // 4. Convert each PDF page to JPEG
            using (var pdfStream = new MemoryStream(pdfBytes))
            using (var document = PdfiumViewer.PdfDocument.Load(pdfStream))
            {
                for (int i = 0; i < document.PageCount; i++)
                {
                    using (var image = document.Render(i, 300, 300, true))
                    {
                        var jpegPath = Path.Combine(folderPath, $"CPA_{ContractId}_p{i + 1}.jpg");
                        image.Save(jpegPath, System.Drawing.Imaging.ImageFormat.Jpeg);
                    }
                }
            }
            // 5. Delete the PDF file after conversion
            System.IO.File.Delete(pdfPath);

            // 6. Create zip file of JPEGs
            var folderPathZip = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "CPA", "CPA_" + ContractId);
            var zipPath = Path.Combine(folderPathZip, $"CPA_{ContractId}_JPEG.zip");
            if (System.IO.File.Exists(zipPath))
            {
                System.IO.File.Delete(zipPath);
            }
            var jpegFiles = Directory.GetFiles(folderPath, $"CPA_{ContractId}_p*.jpg");
            using (var zip = System.IO.Compression.ZipFile.Open(zipPath, ZipArchiveMode.Create))
            {
                foreach (var file in jpegFiles)
                {
                    zip.CreateEntryFromFile(file, Path.GetFileName(file));
                }
            }
            // 7. Delete JPEG files after zipping
            foreach (var file in jpegFiles)
            {
                System.IO.File.Delete(file);
            }

            // 8. Return the zip file as download
            var zipBytes = await System.IO.File.ReadAllBytesAsync(zipPath);
            return File(zipBytes, "application/zip", $"CPA_{ContractId}_JPEG.zip");
        }
        public async Task<IActionResult> OnGetWordContact_CPA_JPEG_Preview(string ContractId = "14")
        {
            // 1. Generate PDF from CPA contract
            var htmlContent = await _BuyOrSellComputerService.OnGetWordContact_BuyOrSellComputerService_ToPDF(ContractId);
            await new BrowserFetcher().DownloadAsync();
            await using var browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = true });
            await using var page = await browser.NewPageAsync();

            await page.SetContentAsync(htmlContent);

            var pdfOptions = new PdfOptions
            {
                Format = PuppeteerSharp.Media.PaperFormat.A4,

                Landscape = false,
                MarginOptions = new PuppeteerSharp.Media.MarginOptions
                {
                    Top = "20mm",
                    Bottom = "20mm",
                    Left = "20mm",
                    Right = "20mm"
                },
                PrintBackground = true

            };

            var pdfBytes = await page.PdfDataAsync(pdfOptions);

            // 2. Prepare folder structure
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "CPA", "CPA_" + ContractId, "CPA_" + ContractId + "_JPEG");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var pdfPath = Path.Combine(folderPath, $"CPA_{ContractId}.pdf");

            // 3. Save PDF to disk
            if (System.IO.File.Exists(pdfPath))
            {
                System.IO.File.Delete(pdfPath);
            }
            await System.IO.File.WriteAllBytesAsync(pdfPath, pdfBytes);

            // 4. Convert each PDF page to JPEG
            using (var pdfStream = new MemoryStream(pdfBytes))
            using (var document = PdfiumViewer.PdfDocument.Load(pdfStream))
            {
                for (int i = 0; i < document.PageCount; i++)
                {
                    using (var image = document.Render(i, 300, 300, true))
                    {
                        var jpegPath = Path.Combine(folderPath, $"CPA_{ContractId}_p{i + 1}.jpg");
                        image.Save(jpegPath, System.Drawing.Imaging.ImageFormat.Jpeg);
                    }
                }
            }
            // 5. Delete the PDF file after conversion
            System.IO.File.Delete(pdfPath);

            // 6. Create password-protected zip file of JPEGs
            var folderPathZip = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "CPA", "CPA_" + ContractId);
            var zipPath = Path.Combine(folderPathZip, $"CPA_{ContractId}_JPEG_Preview.zip");
            if (System.IO.File.Exists(zipPath))
            {
                System.IO.File.Delete(zipPath);
            }
            var jpegFiles = Directory.GetFiles(folderPath, $"CPA_{ContractId}_p*.jpg");
            string password = _configuration["Password:PaswordPDF"];

            using (var fsOut = System.IO.File.Create(zipPath))
            using (var zipStream = new ZipOutputStream(fsOut))
            {
                zipStream.SetLevel(9); // 0-9, 9 = best compression
                zipStream.Password = password; // Set password

                foreach (var file in jpegFiles)
                {
                    var entry = new ZipEntry(Path.GetFileName(file))
                    {
                        DateTime = DateTime.Now
                    };
                    zipStream.PutNextEntry(entry);

                    byte[] buffer = System.IO.File.ReadAllBytes(file);
                    zipStream.Write(buffer, 0, buffer.Length);
                    zipStream.CloseEntry();
                }
                zipStream.IsStreamOwner = true;
                zipStream.Close();
            }
            // 7. Delete JPEG files after zipping
            foreach (var file in jpegFiles)
            {
                System.IO.File.Delete(file);
            }

            // 8. Return the password-protected zip file as download
            var zipBytes = await System.IO.File.ReadAllBytesAsync(zipPath);
            return File(zipBytes, "application/zip", $"CPA_{ContractId}_JPEG_Preview.zip");
        }
        public async Task<IActionResult> OnGetWordContact_CPA_Word(string ContractId = "14")
        {
            // 1. Get HTML content for CPA contract
            var htmlContent = await _BuyOrSellComputerService.OnGetWordContact_BuyOrSellComputerService_ToPDF(ContractId);

            // 2. Create a Word document from HTML using Spire.Doc
            var document = new Spire.Doc.Document();
            document.LoadFromStream(new MemoryStream(System.Text.Encoding.UTF8.GetBytes(htmlContent)), Spire.Doc.FileFormat.Html);

            // 3. Save the Word document to a MemoryStream
            using var ms = new MemoryStream();
            document.SaveToStream(ms, Spire.Doc.FileFormat.Docx);
            var wordBytes = ms.ToArray();

            // 4. Prepare folder structure
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "CPA", "CPA_" + ContractId);
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var filePath = Path.Combine(folderPath, $"CPA_{ContractId}.docx");

            // 5. Save Word file to disk
            if (System.IO.File.Exists(filePath))
            {
                System.IO.File.Delete(filePath);
            }
            await System.IO.File.WriteAllBytesAsync(filePath, wordBytes);

            // 6. Return the Word file as download
            var resultBytes = await System.IO.File.ReadAllBytesAsync(filePath);
            return File(resultBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", $"CPA_{ContractId}.docx");
        }
        public async Task<IActionResult> OnGetWordContact_CPA_Word_Preview(string ContractId = "14")
        {
            // 1. Get HTML content for CPA contract
            var htmlContent = await _BuyOrSellComputerService.OnGetWordContact_BuyOrSellComputerService_ToPDF(ContractId);

            // 2. Create a Word document from HTML using Spire.Doc
            var document = new Spire.Doc.Document();
            document.LoadFromStream(new MemoryStream(System.Text.Encoding.UTF8.GetBytes(htmlContent)), Spire.Doc.FileFormat.Html);

            // 3. Get password from appsettings.json
            string? userPassword = _configuration["Password:PaswordPDF"];

            // 4. Apply password protection if password is set
            if (!string.IsNullOrEmpty(userPassword))
            {
                document.Encrypt(userPassword);
            }

            // 5. Save the Word document to a MemoryStream
            using var ms = new MemoryStream();
            document.SaveToStream(ms, Spire.Doc.FileFormat.Docx);
            var wordBytes = ms.ToArray();

            // 6. Prepare folder structure
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "CPA", "CPA_" + ContractId);
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var filePath = Path.Combine(folderPath, $"CPA_{ContractId}_Preview.docx");

            // 7. Save Word file to disk
            if (System.IO.File.Exists(filePath))
            {
                System.IO.File.Delete(filePath);
            }
            await System.IO.File.WriteAllBytesAsync(filePath, wordBytes);

            // 8. Return the Word file as download
            var resultBytes = await System.IO.File.ReadAllBytesAsync(filePath);
            return File(resultBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", $"CPA_{ContractId}_Preview.docx");
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
            var htmlContent = await _BuyOrSellService.OnGetWordContact_BuyOrSellService_ToPDF(ContractId, "SPA30560");
            await new BrowserFetcher().DownloadAsync();
            await using var browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = true });
            await using var page = await browser.NewPageAsync();

            await page.SetContentAsync(htmlContent);

            var pdfOptions = new PdfOptions
            {
                Format = PuppeteerSharp.Media.PaperFormat.A4,

                Landscape = false,
                MarginOptions = new PuppeteerSharp.Media.MarginOptions
                {
                    Top = "20mm",
                    Bottom = "20mm",
                    Left = "20mm",
                    Right = "20mm"
                },
                PrintBackground = true

            };

            var pdfBytes = await page.PdfDataAsync(pdfOptions);


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
            await System.IO.File.WriteAllBytesAsync(filePath, pdfBytes);

            // return File(wordBytes, "application/pdf", "สัญญาซื้อขาย.pdf");
        }
        public async Task<IActionResult> OnGetWordContact_SPA30560_PDF_Preview(string ContractId = "4", string Name = "สมใจ ทดสอบ")
        {
            var htmlContent = await _BuyOrSellService.OnGetWordContact_BuyOrSellService_ToPDF(ContractId, "SPA30560");
            await new BrowserFetcher().DownloadAsync();
            await using var browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = true });
            await using var page = await browser.NewPageAsync();

            await page.SetContentAsync(htmlContent);

            var pdfOptions = new PdfOptions
            {
                Format = PuppeteerSharp.Media.PaperFormat.A4,

                Landscape = false,
                MarginOptions = new PuppeteerSharp.Media.MarginOptions
                {
                    Top = "20mm",
                    Bottom = "20mm",
                    Left = "20mm",
                    Right = "20mm"
                },
                PrintBackground = true

            };

            var pdfBytes = await page.PdfDataAsync(pdfOptions);

            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "SPA30560");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var fileName = $"SPA30560_{ContractId}_Preview.pdf";
            var filePath = Path.Combine(folderPath, fileName);

            // Get password from appsettings.json
            string? userPassword = await GetPdfPasswordAsync(Name);;

            using (var inputStream = new MemoryStream(pdfBytes))
            using (var outputStream = new MemoryStream())
            {
                var document = PdfSharpCore.Pdf.IO.PdfReader.Open(inputStream, PdfSharpCore.Pdf.IO.PdfDocumentOpenMode.Modify);

                foreach (var pdfPage in document.Pages)
                {
                    using (var gfx = PdfSharpCore.Drawing.XGraphics.FromPdfPage(pdfPage))
                    {
                        var font = new PdfSharpCore.Drawing.XFont("Tahoma", 25, PdfSharpCore.Drawing.XFontStyle.Bold);
                        var text = $"สัญญาอิเลคทรอนิกส์พิมพ์ออกโดย {Name}\nวันที่ {DateTime.Now:dd/MM/yyyy HH:mm}";
                        var lines = text.Split('\n');
                        double lineHeight = font.GetHeight();
                        double totalHeight = lineHeight * lines.Length;
                        double y = (pdfPage.Height - totalHeight) / 2;

                        foreach (var line in lines)
                        {
                            var size = gfx.MeasureString(line, font);
                            double x = (pdfPage.Width - size.Width) / 2;
                            var state = gfx.Save();
                            gfx.TranslateTransform(pdfPage.Width / 2, pdfPage.Height / 2);
                            gfx.RotateTransform(-30);
                            gfx.TranslateTransform(-pdfPage.Width / 2, -pdfPage.Height / 2);

                            var brush = new PdfSharpCore.Drawing.XSolidBrush(
                                PdfSharpCore.Drawing.XColor.FromArgb(80, 255, 0, 0)); // semi-transparent red

                            gfx.DrawString(line, font, brush, x, y);
                            gfx.Restore(state);

                            y += lineHeight;
                        }
                    }
                }

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
        public async Task<IActionResult> OnGetWordContact_SPA30560_JPEG(string ContractId = "4")
        {
            // 1. Generate PDF from SPA30560 contract
            var htmlContent = await _BuyOrSellService.OnGetWordContact_BuyOrSellService_ToPDF(ContractId, "SPA30560");
            await new BrowserFetcher().DownloadAsync();
            await using var browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = true });
            await using var page = await browser.NewPageAsync();

            await page.SetContentAsync(htmlContent);

            var pdfOptions = new PdfOptions
            {
                Format = PuppeteerSharp.Media.PaperFormat.A4,

                Landscape = false,
                MarginOptions = new PuppeteerSharp.Media.MarginOptions
                {
                    Top = "20mm",
                    Bottom = "20mm",
                    Left = "20mm",
                    Right = "20mm"
                },
                PrintBackground = true

            };

            var pdfBytes = await page.PdfDataAsync(pdfOptions);

            // 2. Prepare folder structure
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "SPA30560", "SPA30560_" + ContractId, "SPA30560_" + ContractId + "_JPEG");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var pdfPath = Path.Combine(folderPath, $"SPA30560_{ContractId}.pdf");

            // 3. Save PDF to disk
            if (System.IO.File.Exists(pdfPath))
            {
                System.IO.File.Delete(pdfPath);
            }
            await System.IO.File.WriteAllBytesAsync(pdfPath, pdfBytes);

            // 4. Convert each PDF page to JPEG
            using (var pdfStream = new MemoryStream(pdfBytes))
            using (var document = PdfiumViewer.PdfDocument.Load(pdfStream))
            {
                for (int i = 0; i < document.PageCount; i++)
                {
                    using (var image = document.Render(i, 300, 300, true))
                    {
                        var jpegPath = Path.Combine(folderPath, $"SPA30560_{ContractId}_p{i + 1}.jpg");
                        image.Save(jpegPath, System.Drawing.Imaging.ImageFormat.Jpeg);
                    }
                }
            }
            // 5. Delete the PDF file after conversion
            System.IO.File.Delete(pdfPath);

            // 6. Create zip file of JPEGs
            var folderPathZip = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "SPA30560", "SPA30560_" + ContractId);
            var zipPath = Path.Combine(folderPathZip, $"SPA30560_{ContractId}_JPEG.zip");
            if (System.IO.File.Exists(zipPath))
            {
                System.IO.File.Delete(zipPath);
            }
            var jpegFiles = Directory.GetFiles(folderPath, $"SPA30560_{ContractId}_p*.jpg");
            using (var zip = System.IO.Compression.ZipFile.Open(zipPath, ZipArchiveMode.Create))
            {
                foreach (var file in jpegFiles)
                {
                    zip.CreateEntryFromFile(file, Path.GetFileName(file));
                }
            }
            // 7. Delete JPEG files after zipping
            foreach (var file in jpegFiles)
            {
                System.IO.File.Delete(file);
            }

            // 8. Return the zip file as download
            var zipBytes = await System.IO.File.ReadAllBytesAsync(zipPath);
            return File(zipBytes, "application/zip", $"SPA30560_{ContractId}_JPEG.zip");
        }

        public async Task<IActionResult> OnGetWordContact_SPA30560_JPEG_Preview(string ContractId = "4")
        {
            // 1. Generate PDF from SPA30560 contract
            var htmlContent = await _BuyOrSellService.OnGetWordContact_BuyOrSellService_ToPDF(ContractId, "SPA30560");
            await new BrowserFetcher().DownloadAsync();
            await using var browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = true });
            await using var page = await browser.NewPageAsync();

            await page.SetContentAsync(htmlContent);

            var pdfOptions = new PdfOptions
            {
                Format = PuppeteerSharp.Media.PaperFormat.A4,

                Landscape = false,
                MarginOptions = new PuppeteerSharp.Media.MarginOptions
                {
                    Top = "20mm",
                    Bottom = "20mm",
                    Left = "20mm",
                    Right = "20mm"
                },
                PrintBackground = true

            };

            var pdfBytes = await page.PdfDataAsync(pdfOptions);

            // 2. Prepare folder structure
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "SPA30560", "SPA30560_" + ContractId, "SPA30560_" + ContractId + "_JPEG");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var pdfPath = Path.Combine(folderPath, $"SPA30560_{ContractId}.pdf");

            // 3. Save PDF to disk
            if (System.IO.File.Exists(pdfPath))
            {
                System.IO.File.Delete(pdfPath);
            }
            await System.IO.File.WriteAllBytesAsync(pdfPath, pdfBytes);

            // 4. Convert each PDF page to JPEG
            using (var pdfStream = new MemoryStream(pdfBytes))
            using (var document = PdfiumViewer.PdfDocument.Load(pdfStream))
            {
                for (int i = 0; i < document.PageCount; i++)
                {
                    using (var image = document.Render(i, 300, 300, true))
                    {
                        var jpegPath = Path.Combine(folderPath, $"SPA30560_{ContractId}_p{i + 1}.jpg");
                        image.Save(jpegPath, System.Drawing.Imaging.ImageFormat.Jpeg);
                    }
                }
            }
            // 5. Delete the PDF file after conversion
            System.IO.File.Delete(pdfPath);

            // 6. Create password-protected zip file of JPEGs
            var folderPathZip = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "SPA30560", "SPA30560_" + ContractId);
            var zipPath = Path.Combine(folderPathZip, $"SPA30560_{ContractId}_JPEG_Preview.zip");
            if (System.IO.File.Exists(zipPath))
            {
                System.IO.File.Delete(zipPath);
            }
            var jpegFiles = Directory.GetFiles(folderPath, $"SPA30560_{ContractId}_p*.jpg");
            string password = _configuration["Password:PaswordPDF"];

            using (var fsOut = System.IO.File.Create(zipPath))
            using (var zipStream = new ZipOutputStream(fsOut))
            {
                zipStream.SetLevel(9); // 0-9, 9 = best compression
                zipStream.Password = password; // Set password

                foreach (var file in jpegFiles)
                {
                    var entry = new ZipEntry(Path.GetFileName(file))
                    {
                        DateTime = DateTime.Now
                    };
                    zipStream.PutNextEntry(entry);

                    byte[] buffer = System.IO.File.ReadAllBytes(file);
                    zipStream.Write(buffer, 0, buffer.Length);
                    zipStream.CloseEntry();
                }
                zipStream.IsStreamOwner = true;
                zipStream.Close();
            }
            // 7. Delete JPEG files after zipping
            foreach (var file in jpegFiles)
            {
                System.IO.File.Delete(file);
            }

            // 8. Return the password-protected zip file as download
            var zipBytes = await System.IO.File.ReadAllBytesAsync(zipPath);
            return File(zipBytes, "application/zip", $"SPA30560_{ContractId}_JPEG_Preview.zip");
        }

        public async Task<IActionResult> OnGetWordContact_SPA30560_Word(string ContractId = "4")
        {
            // 1. Get HTML content for SPA30560 contract
            var htmlContent = await _BuyOrSellService.OnGetWordContact_BuyOrSellService_ToPDF(ContractId, "SPA30560");

            // 2. Create a Word document from HTML using Spire.Doc
            var document = new Spire.Doc.Document();
            document.LoadFromStream(new MemoryStream(System.Text.Encoding.UTF8.GetBytes(htmlContent)), Spire.Doc.FileFormat.Html);

            // 3. Save the Word document to a MemoryStream
            using var ms = new MemoryStream();
            document.SaveToStream(ms, Spire.Doc.FileFormat.Docx);
            var wordBytes = ms.ToArray();

            // 4. Prepare folder structure
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "SPA30560", "SPA30560_" + ContractId);
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var filePath = Path.Combine(folderPath, $"SPA30560_{ContractId}.docx");

            // 5. Save Word file to disk
            if (System.IO.File.Exists(filePath))
            {
                System.IO.File.Delete(filePath);
            }
            await System.IO.File.WriteAllBytesAsync(filePath, wordBytes);

            // 6. Return the Word file as download
            var resultBytes = await System.IO.File.ReadAllBytesAsync(filePath);
            return File(resultBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", $"SPA30560_{ContractId}.docx");
        }
        public async Task<IActionResult> OnGetWordContact_SPA30560_Word_Preview(string ContractId = "4")
        {
            // 1. Get the Word document for SPA30560 contract
            var htmlContent = await _BuyOrSellService.OnGetWordContact_BuyOrSellService_ToPDF(ContractId, "SPA30560");

            // 2. Create a Word document from HTML using Spire.Doc
            var document = new Spire.Doc.Document();
            document.LoadFromStream(new MemoryStream(System.Text.Encoding.UTF8.GetBytes(htmlContent)), Spire.Doc.FileFormat.Html);

            // 3. Save the Word document to a MemoryStream
            using var ms = new MemoryStream();
            document.SaveToStream(ms, Spire.Doc.FileFormat.Docx);
            var wordBytes = ms.ToArray();

            // 4. Prepare folder structure
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "SPA30560", "SPA30560_" + ContractId);
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var filePath = Path.Combine(folderPath, $"SPA30560_{ContractId}.docx");

            // 5. Save Word file to disk
            if (System.IO.File.Exists(filePath))
            {
                System.IO.File.Delete(filePath);
            }
            await System.IO.File.WriteAllBytesAsync(filePath, wordBytes);

            // 6. Return the Word file as download
            var resultBytes = await System.IO.File.ReadAllBytesAsync(filePath);
            return File(resultBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", $"SPA30560_{ContractId}.docx");
        }

        #endregion 4.1.1.2.8.สัญญาซื้อขาย ร.305-60 SPA30560

        #region 4.1.1.2.7.สัญญาการรักษาข้อมูลที่เป็นความลับ NDA
   
        public async Task OnGetWordContact_NDA_PDF(string ContractId = "4")
        {
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
            var filePathView = Path.Combine(folderPath, "NDA_" + ContractId + "_1.pdf");

            // Delete the file if it already exists
            if (System.IO.File.Exists(filePathView))
            {
                System.IO.File.Delete(filePathView);
            }
            var htmlContent = await _DataSecretService.OnGetWordContact_DataSecretService_ToPDF(ContractId, "NDA");
            await new BrowserFetcher().DownloadAsync();
            await using var browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = true });
            await using var page = await browser.NewPageAsync();

            await page.SetContentAsync(htmlContent);

            var pdfOptions = new PdfOptions
            {
                Format = PuppeteerSharp.Media.PaperFormat.A4,

                Landscape = false,
                MarginOptions = new PuppeteerSharp.Media.MarginOptions
                {
                    Top = "20mm",
                    Bottom = "20mm",
                    Left = "20mm",
                    Right = "20mm"
                },
                PrintBackground = true

            };

            var pdfBytes = await page.PdfDataAsync(pdfOptions);


          
            await System.IO.File.WriteAllBytesAsync(filePath, pdfBytes);
            await System.IO.File.WriteAllBytesAsync(filePathView, pdfBytes);
            
            // return File(wordBytes, "application/pdf", "บันทึกข้อตกลงการแบ่งปันข้อมูลส่วนบุคคล.pdf");
        }
        public async Task<IActionResult> OnGetWordContact_NDA_PDF_Preview(string ContractId = "1", string Name = "สมใจ ทดสอบ")
        {
            var htmlContent = await _DataSecretService.OnGetWordContact_DataSecretService_ToPDF(ContractId, "NDA");
            await new BrowserFetcher().DownloadAsync();
            await using var browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = true });
            await using var page = await browser.NewPageAsync();

            await page.SetContentAsync(htmlContent);

            var pdfOptions = new PdfOptions
            {
                Format = PuppeteerSharp.Media.PaperFormat.A4,

                Landscape = false,
                MarginOptions = new PuppeteerSharp.Media.MarginOptions
                {
                    Top = "20mm",
                    Bottom = "20mm",
                    Left = "20mm",
                    Right = "20mm"
                },
                PrintBackground = true

            };

            var pdfBytes = await page.PdfDataAsync(pdfOptions);

            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "NDA");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var fileName = $"NDA_{ContractId}_Preview.pdf";
            var filePath = Path.Combine(folderPath, fileName);

            // Get password from appsettings.json
            string? userPassword = await GetPdfPasswordAsync(Name);;

            using (var inputStream = new MemoryStream(pdfBytes))
            using (var outputStream = new MemoryStream())
            {
                var document = PdfSharpCore.Pdf.IO.PdfReader.Open(inputStream, PdfSharpCore.Pdf.IO.PdfDocumentOpenMode.Modify);

                foreach (var pdfPage in document.Pages)
                {
                    using (var gfx = PdfSharpCore.Drawing.XGraphics.FromPdfPage(pdfPage))
                    {
                        var font = new PdfSharpCore.Drawing.XFont("Tahoma", 25, PdfSharpCore.Drawing.XFontStyle.Bold);
                        var text = $"สัญญาอิเลคทรอนิกส์พิมพ์ออกโดย {Name}\nวันที่ {DateTime.Now:dd/MM/yyyy HH:mm}";
                        var lines = text.Split('\n');
                        double lineHeight = font.GetHeight();
                        double totalHeight = lineHeight * lines.Length;
                        double y = (pdfPage.Height - totalHeight) / 2;

                        foreach (var line in lines)
                        {
                            var size = gfx.MeasureString(line, font);
                            double x = (pdfPage.Width - size.Width) / 2;
                            var state = gfx.Save();
                            gfx.TranslateTransform(pdfPage.Width / 2, pdfPage.Height / 2);
                            gfx.RotateTransform(-30);
                            gfx.TranslateTransform(-pdfPage.Width / 2, -pdfPage.Height / 2);

                            var brush = new PdfSharpCore.Drawing.XSolidBrush(
                                PdfSharpCore.Drawing.XColor.FromArgb(80, 255, 0, 0)); // semi-transparent red

                            gfx.DrawString(line, font, brush, x, y);
                            gfx.Restore(state);

                            y += lineHeight;
                        }
                    }
                }

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

        public async Task<IActionResult> OnGetWordContact_NDA_PDF_JPEG(string ContractId = "1")
        {
            // 1. Generate PDF from NDA contract
            var htmlContent = await _DataSecretService.OnGetWordContact_DataSecretService_ToPDF(ContractId, "NDA");
            await new BrowserFetcher().DownloadAsync();
            await using var browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = true });
            await using var page = await browser.NewPageAsync();

            await page.SetContentAsync(htmlContent);

            var pdfOptions = new PdfOptions
            {
                Format = PuppeteerSharp.Media.PaperFormat.A4,

                Landscape = false,
                MarginOptions = new PuppeteerSharp.Media.MarginOptions
                {
                    Top = "20mm",
                    Bottom = "20mm",
                    Left = "20mm",
                    Right = "20mm"
                },
                PrintBackground = true

            };

            var pdfBytes = await page.PdfDataAsync(pdfOptions);


            // 2. Prepare folder structure
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "NDA", "NDA_" + ContractId, "NDA_" + ContractId + "_JPEG");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var pdfPath = Path.Combine(folderPath, $"NDA_{ContractId}.pdf");

            // 3. Save PDF to disk
            if (System.IO.File.Exists(pdfPath))
            {
                System.IO.File.Delete(pdfPath);
            }
            await System.IO.File.WriteAllBytesAsync(pdfPath, pdfBytes);

            // 4. Convert each PDF page to JPEG
            using (var pdfStream = new MemoryStream(pdfBytes))
            using (var document = PdfiumViewer.PdfDocument.Load(pdfStream))
            {
                for (int i = 0; i < document.PageCount; i++)
                {
                    using (var image = document.Render(i, 300, 300, true))
                    {
                        var jpegPath = Path.Combine(folderPath, $"NDA_{ContractId}_p{i + 1}.jpg");
                        image.Save(jpegPath, System.Drawing.Imaging.ImageFormat.Jpeg);
                    }
                }
            }
            // 5. Delete the PDF file after conversion
            System.IO.File.Delete(pdfPath);

            // 6. Create zip file of JPEGs
            var folderPathZip = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "NDA", "NDA_" + ContractId);
            var zipPath = Path.Combine(folderPathZip, $"NDA_{ContractId}_JPEG.zip");
            if (System.IO.File.Exists(zipPath))
            {
                System.IO.File.Delete(zipPath);
            }
            var jpegFiles = Directory.GetFiles(folderPath, $"NDA_{ContractId}_p*.jpg");
            using (var zip = System.IO.Compression.ZipFile.Open(zipPath, ZipArchiveMode.Create))
            {
                foreach (var file in jpegFiles)
                {
                    zip.CreateEntryFromFile(file, Path.GetFileName(file));
                }
            }
            // 7. Delete JPEG files after zipping
            foreach (var file in jpegFiles)
            {
                System.IO.File.Delete(file);
            }

            // 8. Return the zip file as download
            var zipBytes = await System.IO.File.ReadAllBytesAsync(zipPath);
            return File(zipBytes, "application/zip", $"NDA_{ContractId}_JPEG.zip");
        }
        public async Task<IActionResult> OnGetWordContact_NDA_PDF_JPEG_Preview(string ContractId = "1")
        {
            // 1. Generate PDF from NDA contract
            var htmlContent = await _DataSecretService.OnGetWordContact_DataSecretService_ToPDF(ContractId, "NDA");
            await new BrowserFetcher().DownloadAsync();
            await using var browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = true });
            await using var page = await browser.NewPageAsync();

            await page.SetContentAsync(htmlContent);

            var pdfOptions = new PdfOptions
            {
                Format = PuppeteerSharp.Media.PaperFormat.A4,

                Landscape = false,
                MarginOptions = new PuppeteerSharp.Media.MarginOptions
                {
                    Top = "20mm",
                    Bottom = "20mm",
                    Left = "20mm",
                    Right = "20mm"
                },
                PrintBackground = true

            };

            var pdfBytes = await page.PdfDataAsync(pdfOptions);

            // 2. Prepare folder structure
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "NDA", "NDA_" + ContractId, "NDA_" + ContractId + "_JPEG");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var pdfPath = Path.Combine(folderPath, $"NDA_{ContractId}.pdf");

            // 3. Save PDF to disk
            if (System.IO.File.Exists(pdfPath))
            {
                System.IO.File.Delete(pdfPath);
            }
            await System.IO.File.WriteAllBytesAsync(pdfPath, pdfBytes);

            // 4. Convert each PDF page to JPEG
            using (var pdfStream = new MemoryStream(pdfBytes))
            using (var document = PdfiumViewer.PdfDocument.Load(pdfStream))
            {
                for (int i = 0; i < document.PageCount; i++)
                {
                    using (var image = document.Render(i, 300, 300, true))
                    {
                        var jpegPath = Path.Combine(folderPath, $"NDA_{ContractId}_p{i + 1}.jpg");
                        image.Save(jpegPath, System.Drawing.Imaging.ImageFormat.Jpeg);
                    }
                }
            }
            // 5. Delete the PDF file after conversion
            System.IO.File.Delete(pdfPath);

            // 6. Create password-protected zip file of JPEGs
            var folderPathZip = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "NDA", "NDA_" + ContractId);
            var zipPath = Path.Combine(folderPathZip, $"NDA_{ContractId}_JPEG_Preview.zip");
            if (System.IO.File.Exists(zipPath))
            {
                System.IO.File.Delete(zipPath);
            }
            var jpegFiles = Directory.GetFiles(folderPath, $"NDA_{ContractId}_p*.jpg");
            string password = _configuration["Password:PaswordPDF"];

            using (var fsOut = System.IO.File.Create(zipPath))
            using (var zipStream = new ZipOutputStream(fsOut))
            {
                zipStream.SetLevel(9); // 0-9, 9 = best compression
                zipStream.Password = password; // Set password

                foreach (var file in jpegFiles)
                {
                    var entry = new ZipEntry(Path.GetFileName(file))
                    {
                        DateTime = DateTime.Now
                    };
                    zipStream.PutNextEntry(entry);

                    byte[] buffer = System.IO.File.ReadAllBytes(file);
                    zipStream.Write(buffer, 0, buffer.Length);
                    zipStream.CloseEntry();
                }
                zipStream.IsStreamOwner = true;
                zipStream.Close();
            }
            // 7. Delete JPEG files after zipping
            foreach (var file in jpegFiles)
            {
                System.IO.File.Delete(file);
            }

            // 8. Return the password-protected zip file as download
            var zipBytes = await System.IO.File.ReadAllBytesAsync(zipPath);
            return File(zipBytes, "application/zip", $"NDA_{ContractId}_JPEG_Preview.zip");
        }

        public async Task<IActionResult> OnGetWordContact_NDA_Word(string ContractId = "1")
        {

            // 1. Get HTML content from the service (for NDA Word export)
            var htmlContent = await _DataSecretService.OnGetWordContact_DataSecretService_ToPDF(ContractId, "NDA");

            // 2. Create a Word document from HTML using Spire.Doc
            var document = new Spire.Doc.Document();
            document.LoadFromStream(new MemoryStream(System.Text.Encoding.UTF8.GetBytes(htmlContent)), Spire.Doc.FileFormat.Html);

            // 3. Save the Word document to a MemoryStream
            using var ms = new MemoryStream();
            document.SaveToStream(ms, Spire.Doc.FileFormat.Docx);
            var wordBytes = ms.ToArray();

            // 4. Prepare folder structure
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "NDA", "NDA_" + ContractId);
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var filePath = Path.Combine(folderPath, $"NDA_{ContractId}.docx");

            // 5. Save Word file to disk
            if (System.IO.File.Exists(filePath))
            {
                System.IO.File.Delete(filePath);
            }
            await System.IO.File.WriteAllBytesAsync(filePath, wordBytes);

            // 6. Return the Word file as download
            var resultBytes = await System.IO.File.ReadAllBytesAsync(filePath);
            return File(resultBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", $"NDA_{ContractId}.docx");
        }

        public async Task<IActionResult> OnGetWordContact_NDA_Word_Preview(string ContractId = "1")
        {
            // 1. Get HTML content from the service
            var htmlContent = await _DataSecretService.OnGetWordContact_DataSecretService_ToPDF(ContractId, "NDA");

            // 2. Create a Word document from HTML using Spire.Doc
            var document = new Spire.Doc.Document();
            document.LoadFromStream(new MemoryStream(System.Text.Encoding.UTF8.GetBytes(htmlContent)), Spire.Doc.FileFormat.Html);

            // 3. Save the Word document to a MemoryStream
            using var ms = new MemoryStream();
            document.SaveToStream(ms, Spire.Doc.FileFormat.Docx);
            var wordBytes = ms.ToArray();

            // 4. Prepare folder structure
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "NDA", "NDA_" + ContractId);
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var filePath = Path.Combine(folderPath, $"NDA_{ContractId}_Preview.docx");

            // 5. Save Word file to disk
            if (System.IO.File.Exists(filePath))
            {
                System.IO.File.Delete(filePath);
            }
            await System.IO.File.WriteAllBytesAsync(filePath, wordBytes);

            // 6. Return the Word file as download
            var resultBytes = await System.IO.File.ReadAllBytesAsync(filePath);
            return File(resultBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", $"NDA_{ContractId}_Preview.docx");
        }
        #endregion 4.1.1.2.7.สัญญาการรักษาข้อมูลที่เป็นความลับ NDA

        #region 4.1.1.2.6.บันทึกข้อตกลงการแบ่งปันข้อมูลส่วนบุคคล PDSA

        public async Task OnGetWordContact_PDSA_PDF(string ContractId = "1")
        {
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
            var filePathView = Path.Combine(folderPath, "PDSA_" + ContractId + "_1.pdf");

            // Delete the file if it already exists
            if (System.IO.File.Exists(filePathView))
            {
                System.IO.File.Delete(filePathView);
            }
            var htmlContent = await _DataPersonalService.OnGetWordContact_DataPersonalService_ToPDF(ContractId, "PDSA");
            await new BrowserFetcher().DownloadAsync();
            await using var browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = true });
            await using var page = await browser.NewPageAsync();

            await page.SetContentAsync(htmlContent);

            var pdfOptions = new PdfOptions
            {
                Format = PuppeteerSharp.Media.PaperFormat.A4,

                Landscape = false,
                MarginOptions = new PuppeteerSharp.Media.MarginOptions
                {
                    Top = "20mm",
                    Bottom = "20mm",
                    Left = "20mm",
                    Right = "20mm"
                },
                PrintBackground = true

            };

            var pdfBytes = await page.PdfDataAsync(pdfOptions);

           
            await System.IO.File.WriteAllBytesAsync(filePath, pdfBytes);
            await System.IO.File.WriteAllBytesAsync(filePathView, pdfBytes);
            
            // return File(wordBytes, "application/pdf", "บันทึกข้อตกลงการแบ่งปันข้อมูลส่วนบุคคล.pdf");
        }

        public async Task<IActionResult> OnGetWordContact_PDSA_PDF_Preview(string ContractId = "1", string Name = "สมใจ ทดสอบ")
        {
            var htmlContent = await _DataPersonalService.OnGetWordContact_DataPersonalService_ToPDF(ContractId, "PDSA");
            await new BrowserFetcher().DownloadAsync();
            await using var browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = true });
            await using var page = await browser.NewPageAsync();

            await page.SetContentAsync(htmlContent);

            var pdfOptions = new PdfOptions
            {
                Format = PuppeteerSharp.Media.PaperFormat.A4,

                Landscape = false,
                MarginOptions = new PuppeteerSharp.Media.MarginOptions
                {
                    Top = "20mm",
                    Bottom = "20mm",
                    Left = "20mm",
                    Right = "20mm"
                },
                PrintBackground = true

            };

            var pdfBytes = await page.PdfDataAsync(pdfOptions);

            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "PDSA");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var fileName = $"PDSA_{ContractId}_Preview.pdf";
            var filePath = Path.Combine(folderPath, fileName);

            // Get password from appsettings.json
            string? userPassword = await GetPdfPasswordAsync(Name);;

            using (var inputStream = new MemoryStream(pdfBytes))
            using (var outputStream = new MemoryStream())
            {
                var document = PdfSharpCore.Pdf.IO.PdfReader.Open(inputStream, PdfSharpCore.Pdf.IO.PdfDocumentOpenMode.Modify);

                foreach (var pdfPage in document.Pages)
                {
                    using (var gfx = PdfSharpCore.Drawing.XGraphics.FromPdfPage(pdfPage))
                    {
                        var font = new PdfSharpCore.Drawing.XFont("Tahoma", 25, PdfSharpCore.Drawing.XFontStyle.Bold);
                        var text = $"สัญญาอิเลคทรอนิกส์พิมพ์ออกโดย {Name}\nวันที่ {DateTime.Now:dd/MM/yyyy HH:mm}";
                        var lines = text.Split('\n');
                        double lineHeight = font.GetHeight();
                        double totalHeight = lineHeight * lines.Length;
                        double y = (pdfPage.Height - totalHeight) / 2;

                        foreach (var line in lines)
                        {
                            var size = gfx.MeasureString(line, font);
                            double x = (pdfPage.Width - size.Width) / 2;
                            var state = gfx.Save();
                            gfx.TranslateTransform(pdfPage.Width / 2, pdfPage.Height / 2);
                            gfx.RotateTransform(-30);
                            gfx.TranslateTransform(-pdfPage.Width / 2, -pdfPage.Height / 2);

                            var brush = new PdfSharpCore.Drawing.XSolidBrush(
                                PdfSharpCore.Drawing.XColor.FromArgb(80, 255, 0, 0)); // semi-transparent red

                            gfx.DrawString(line, font, brush, x, y);
                            gfx.Restore(state);

                            y += lineHeight;
                        }
                    }
                }

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
        public async Task<IActionResult> OnGetWordContact_PDSA_JPEG(string ContractId = "3")
        {
            // 1. Generate PDF from PDSA contract
            var htmlContent = await _DataPersonalService.OnGetWordContact_DataPersonalService_ToPDF(ContractId, "PDSA");
            await new BrowserFetcher().DownloadAsync();
            await using var browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = true });
            await using var page = await browser.NewPageAsync();

            await page.SetContentAsync(htmlContent);

            var pdfOptions = new PdfOptions
            {
                Format = PuppeteerSharp.Media.PaperFormat.A4,

                Landscape = false,
                MarginOptions = new PuppeteerSharp.Media.MarginOptions
                {
                    Top = "20mm",
                    Bottom = "20mm",
                    Left = "20mm",
                    Right = "20mm"
                },
                PrintBackground = true

            };

            var pdfBytes = await page.PdfDataAsync(pdfOptions);

            // 2. Prepare folder structure
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "PDSA", "PDSA_" + ContractId, "PDSA_" + ContractId + "_JPEG");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var pdfPath = Path.Combine(folderPath, $"PDSA_{ContractId}.pdf");

            // 3. Save PDF to disk
            if (System.IO.File.Exists(pdfPath))
            {
                System.IO.File.Delete(pdfPath);
            }
            await System.IO.File.WriteAllBytesAsync(pdfPath, pdfBytes);

            // 4. Convert each PDF page to JPEG
            using (var pdfStream = new MemoryStream(pdfBytes))
            using (var document = PdfiumViewer.PdfDocument.Load(pdfStream))
            {
                for (int i = 0; i < document.PageCount; i++)
                {
                    using (var image = document.Render(i, 300, 300, true))
                    {
                        var jpegPath = Path.Combine(folderPath, $"PDSA_{ContractId}_p{i + 1}.jpg");
                        image.Save(jpegPath, System.Drawing.Imaging.ImageFormat.Jpeg);
                    }
                }
            }
            // 5. Delete the PDF file after conversion
            System.IO.File.Delete(pdfPath);

            // 6. Create zip file of JPEGs
            var folderPathZip = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "PDSA", "PDSA_" + ContractId);
            var zipPath = Path.Combine(folderPathZip, $"PDSA_{ContractId}_JPEG.zip");
            if (System.IO.File.Exists(zipPath))
            {
                System.IO.File.Delete(zipPath);
            }
            var jpegFiles = Directory.GetFiles(folderPath, $"PDSA_{ContractId}_p*.jpg");
            using (var zip = System.IO.Compression.ZipFile.Open(zipPath, ZipArchiveMode.Create))
            {
                foreach (var file in jpegFiles)
                {
                    zip.CreateEntryFromFile(file, Path.GetFileName(file));
                }
            }
            // 7. Delete JPEG files after zipping
            foreach (var file in jpegFiles)
            {
                System.IO.File.Delete(file);
            }

            // 8. Return the zip file as download
            var zipBytes = await System.IO.File.ReadAllBytesAsync(zipPath);
            return File(zipBytes, "application/zip", $"PDSA_{ContractId}_JPEG.zip");
        }
        public async Task<IActionResult> OnGetWordContact_PDSA_JPEG_Preview(string ContractId = "3")
        {
            // 1. Generate PDF from PDSA contract
            var htmlContent = await _DataPersonalService.OnGetWordContact_DataPersonalService_ToPDF(ContractId, "PDSA");
            await new BrowserFetcher().DownloadAsync();
            await using var browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = true });
            await using var page = await browser.NewPageAsync();

            await page.SetContentAsync(htmlContent);

            var pdfOptions = new PdfOptions
            {
                Format = PuppeteerSharp.Media.PaperFormat.A4,

                Landscape = false,
                MarginOptions = new PuppeteerSharp.Media.MarginOptions
                {
                    Top = "20mm",
                    Bottom = "20mm",
                    Left = "20mm",
                    Right = "20mm"
                },
                PrintBackground = true

            };

            var pdfBytes = await page.PdfDataAsync(pdfOptions);

            // 2. Prepare folder structure
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "PDSA", "PDSA_" + ContractId, "PDSA_" + ContractId + "_JPEG");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var pdfPath = Path.Combine(folderPath, $"PDSA_{ContractId}.pdf");

            // 3. Save PDF to disk
            if (System.IO.File.Exists(pdfPath))
            {
                System.IO.File.Delete(pdfPath);
            }
            await System.IO.File.WriteAllBytesAsync(pdfPath, pdfBytes);

            // 4. Convert each PDF page to JPEG
            using (var pdfStream = new MemoryStream(pdfBytes))
            using (var document = PdfiumViewer.PdfDocument.Load(pdfStream))
            {
                for (int i = 0; i < document.PageCount; i++)
                {
                    using (var image = document.Render(i, 300, 300, true))
                    {
                        var jpegPath = Path.Combine(folderPath, $"PDSA_{ContractId}_p{i + 1}.jpg");
                        image.Save(jpegPath, System.Drawing.Imaging.ImageFormat.Jpeg);
                    }
                }
            }
            // 5. Delete the PDF file after conversion
            System.IO.File.Delete(pdfPath);

            // 6. Create password-protected zip file of JPEGs
            var folderPathZip = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "PDSA", "PDSA_" + ContractId);
            var zipPath = Path.Combine(folderPathZip, $"PDSA_{ContractId}_JPEG_Preview.zip");
            if (System.IO.File.Exists(zipPath))
            {
                System.IO.File.Delete(zipPath);
            }
            var jpegFiles = Directory.GetFiles(folderPath, $"PDSA_{ContractId}_p*.jpg");
            string password = _configuration["Password:PaswordPDF"];

            using (var fsOut = System.IO.File.Create(zipPath))
            using (var zipStream = new ZipOutputStream(fsOut))
            {
                zipStream.SetLevel(9); // 0-9, 9 = best compression
                zipStream.Password = password; // Set password

                foreach (var file in jpegFiles)
                {
                    var entry = new ZipEntry(Path.GetFileName(file))
                    {
                        DateTime = DateTime.Now
                    };
                    zipStream.PutNextEntry(entry);

                    byte[] buffer = System.IO.File.ReadAllBytes(file);
                    zipStream.Write(buffer, 0, buffer.Length);
                    zipStream.CloseEntry();
                }
                zipStream.IsStreamOwner = true;
                zipStream.Close();
            }
            // 7. Delete JPEG files after zipping
            foreach (var file in jpegFiles)
            {
                System.IO.File.Delete(file);
            }

            // 8. Return the password-protected zip file as download
            var zipBytes = await System.IO.File.ReadAllBytesAsync(zipPath);
            return File(zipBytes, "application/zip", $"PDSA_{ContractId}_JPEG_Preview.zip");
        }
        // Update for OnGetWordContact_PDSA_Word
        public async Task<IActionResult> OnGetWordContact_PDSA_Word(string ContractId = "3")
        {
            // 1. Get HTML content from the service (for PDSA Word export)
            var htmlContent = await _DataPersonalService.OnGetWordContact_DataPersonalService_ToPDF(ContractId, "PDSA");

            // 2. Create a Word document from HTML using Spire.Doc
            var document = new Spire.Doc.Document();
            document.LoadFromStream(new MemoryStream(System.Text.Encoding.UTF8.GetBytes(htmlContent)), Spire.Doc.FileFormat.Html);

            // 3. Save the Word document to a MemoryStream
            using var ms = new MemoryStream();
            document.SaveToStream(ms, Spire.Doc.FileFormat.Docx);
            var wordBytes = ms.ToArray();

            // 4. Prepare folder structure
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "PDSA", $"PDSA_{ContractId}");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var filePath = Path.Combine(folderPath, $"PDSA_{ContractId}.docx");

            // 5. Save Word file to disk
            if (System.IO.File.Exists(filePath))
            {
                System.IO.File.Delete(filePath);
            }
            await System.IO.File.WriteAllBytesAsync(filePath, wordBytes);

            // 6. Return the Word file as download
            return File(wordBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", $"PDSA_{ContractId}.docx");
        }
        public async Task<IActionResult> OnGetWordContact_PDSA_Word_Preview(string ContractId = "3")
        {
            // 1. Get HTML content for PDSA Word export
            var htmlContent = await _DataPersonalService.OnGetWordContact_DataPersonalService_ToPDF(ContractId, "PDSA");

            // 2. Create Word document from HTML
            using var htmlStream = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(htmlContent));
            var document = new Spire.Doc.Document();
            document.LoadFromStream(htmlStream, Spire.Doc.FileFormat.Html);

            // 3. Save Word document to memory
            using var wordStream = new MemoryStream();
            document.SaveToStream(wordStream, Spire.Doc.FileFormat.Docx);
            var wordBytes = wordStream.ToArray();

            // 4. Prepare output folder
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "PDSA", $"PDSA_{ContractId}");
            Directory.CreateDirectory(folderPath);
            var filePath = Path.Combine(folderPath, $"PDSA_{ContractId}.docx");

            // 5. Remove existing file if present
            if (System.IO.File.Exists(filePath))
            {
                System.IO.File.Delete(filePath);
            }

            // 6. Get password from configuration
            string? userPassword = _configuration["Password:PaswordPDF"];
            if (string.IsNullOrWhiteSpace(userPassword))
            {
                return BadRequest("Password for document protection is not configured.");
            }

            // 7. Apply password protection and save to disk
            using (var protectedStream = new MemoryStream(wordBytes))
            {
                var protectedDoc = new Spire.Doc.Document();
                protectedDoc.LoadFromStream(protectedStream, Spire.Doc.FileFormat.Docx);
                protectedDoc.Encrypt(userPassword);
                protectedDoc.SaveToFile(filePath, Spire.Doc.FileFormat.Docx);
            }

            // 8. Return the password-protected Word file for download
            var resultBytes = await System.IO.File.ReadAllBytesAsync(filePath);
            return File(resultBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", $"PDSA_{ContractId}_Preview.docx");
        }

        #endregion 4.1.1.2.6.บันทึกข้อตกลงการแบ่งปันข้อมูลส่วนบุคคล PDSA

        #region 4.1.1.2.5.บันทึกข้อตกลงการเป็นผู้ควบคุมข้อมูลส่วนบุคคลร่วมตัวอย่างหน้าจอ JDCA

      
        public async Task OnGetWordContact_JDCA_PDF(string ContractId = "2")
        {
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
            var filePathView = Path.Combine(folderPath, "JDCA_" + ContractId + "_1.pdf");

            // Delete the file if it already exists
            if (System.IO.File.Exists(filePathView))
            {
                System.IO.File.Delete(filePathView);
            }
            var htmlContent = await _ControlDataService.OnGetWordContact_ControlDataServiceHtmlToPdf(ContractId, "JDCA");
            await new BrowserFetcher().DownloadAsync();
            await using var browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = true });
            await using var page = await browser.NewPageAsync();

            await page.SetContentAsync(htmlContent);

            var pdfOptions = new PdfOptions
            {
                Format = PuppeteerSharp.Media.PaperFormat.A4,

                Landscape = false,
                MarginOptions = new PuppeteerSharp.Media.MarginOptions
                {
                    Top = "20mm",
                    Bottom = "20mm",
                    Left = "20mm",
                    Right = "20mm"
                },
                PrintBackground = true

            };

            var pdfBytes = await page.PdfDataAsync(pdfOptions);

           
            await System.IO.File.WriteAllBytesAsync(filePath, pdfBytes);
            await System.IO.File.WriteAllBytesAsync(filePathView, pdfBytes);
            // return File(wordBytes, "application/pdf", "บันทึกข้อตกลงการเป็นผู้ควบคุมข้อมูลส่วนบุคคลร่วม.pdf");
        }
        public async Task<IActionResult> OnGetWordContact_JDCA_PDF_Preview(string ContractId = "2", string Name = "สมใจ ทดสอบ")
        {
            var htmlContent = await _ControlDataService.OnGetWordContact_ControlDataServiceHtmlToPdf(ContractId, "JDCA");
            await new BrowserFetcher().DownloadAsync();
            await using var browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = true });
            await using var page = await browser.NewPageAsync();

            await page.SetContentAsync(htmlContent);

            var pdfOptions = new PdfOptions
            {
                Format = PuppeteerSharp.Media.PaperFormat.A4,

                Landscape = false,
                MarginOptions = new PuppeteerSharp.Media.MarginOptions
                {
                    Top = "20mm",
                    Bottom = "20mm",
                    Left = "20mm",
                    Right = "20mm"
                },
                PrintBackground = true

            };

            var pdfBytes = await page.PdfDataAsync(pdfOptions);


            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "JDCA");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var fileName = $"JDCA_{ContractId}_Preview.pdf";
            var filePath = Path.Combine(folderPath, fileName);

            // Get password from appsettings.json
            string? userPassword = await GetPdfPasswordAsync(Name);;

            using (var inputStream = new MemoryStream(pdfBytes))
            using (var outputStream = new MemoryStream())
            {
                var document = PdfSharpCore.Pdf.IO.PdfReader.Open(inputStream, PdfSharpCore.Pdf.IO.PdfDocumentOpenMode.Modify);

                foreach (var pdfPage in document.Pages)
                {
                    using (var gfx = PdfSharpCore.Drawing.XGraphics.FromPdfPage(pdfPage))
                    {
                        var font = new PdfSharpCore.Drawing.XFont("Tahoma", 25, PdfSharpCore.Drawing.XFontStyle.Bold);
                        var text = $"สัญญาอิเลคทรอนิกส์พิมพ์ออกโดย {Name}\nวันที่ {DateTime.Now:dd/MM/yyyy HH:mm}";
                        var lines = text.Split('\n');
                        double lineHeight = font.GetHeight();
                        double totalHeight = lineHeight * lines.Length;
                        double y = (pdfPage.Height - totalHeight) / 2;

                        foreach (var line in lines)
                        {
                            var size = gfx.MeasureString(line, font);
                            double x = (pdfPage.Width - size.Width) / 2;
                            var state = gfx.Save();
                            gfx.TranslateTransform(pdfPage.Width / 2, pdfPage.Height / 2);
                            gfx.RotateTransform(-30);
                            gfx.TranslateTransform(-pdfPage.Width / 2, -pdfPage.Height / 2);

                            var brush = new PdfSharpCore.Drawing.XSolidBrush(
                                PdfSharpCore.Drawing.XColor.FromArgb(80, 255, 0, 0)); // semi-transparent red

                            gfx.DrawString(line, font, brush, x, y);
                            gfx.Restore(state);

                            y += lineHeight;
                        }
                    }
                }

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

        public async Task<IActionResult> OnGetWordContact_JDCA_JPEG(string ContractId = "5")
        {
            // 1. Generate PDF from JDCA contract
            var htmlContent = await _ControlDataService.OnGetWordContact_ControlDataServiceHtmlToPdf(ContractId, "JDCA");
            await new BrowserFetcher().DownloadAsync();
            await using var browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = true });
            await using var page = await browser.NewPageAsync();

            await page.SetContentAsync(htmlContent);

            var pdfOptions = new PdfOptions
            {
                Format = PuppeteerSharp.Media.PaperFormat.A4,

                Landscape = false,
                MarginOptions = new PuppeteerSharp.Media.MarginOptions
                {
                    Top = "20mm",
                    Bottom = "20mm",
                    Left = "20mm",
                    Right = "20mm"
                },
                PrintBackground = true

            };

            var pdfBytes = await page.PdfDataAsync(pdfOptions);

            // 2. Prepare folder structure
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "JDCA", "JDCA_" + ContractId, "JDCA_" + ContractId + "_JPEG");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var pdfPath = Path.Combine(folderPath, $"JDCA_{ContractId}.pdf");

            // 3. Save PDF to disk
            if (System.IO.File.Exists(pdfPath))
            {
                System.IO.File.Delete(pdfPath);
            }
            await System.IO.File.WriteAllBytesAsync(pdfPath, pdfBytes);

            // 4. Convert each PDF page to JPEG
            using (var pdfStream = new MemoryStream(pdfBytes))
            using (var document = PdfiumViewer.PdfDocument.Load(pdfStream))
            {
                for (int i = 0; i < document.PageCount; i++)
                {
                    using (var image = document.Render(i, 300, 300, true))
                    {
                        var jpegPath = Path.Combine(folderPath, $"JDCA_{ContractId}_p{i + 1}.jpg");
                        image.Save(jpegPath, System.Drawing.Imaging.ImageFormat.Jpeg);
                    }
                }
            }
            // 5. Delete the PDF file after conversion
            System.IO.File.Delete(pdfPath);

            // 6. Create zip file of JPEGs
            var folderPathZip = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "JDCA", "JDCA_" + ContractId);
            var zipPath = Path.Combine(folderPathZip, $"JDCA_{ContractId}_JPEG.zip");
            if (System.IO.File.Exists(zipPath))
            {
                System.IO.File.Delete(zipPath);
            }
            var jpegFiles = Directory.GetFiles(folderPath, $"JDCA_{ContractId}_p*.jpg");
            using (var zip = System.IO.Compression.ZipFile.Open(zipPath, ZipArchiveMode.Create))
            {
                foreach (var file in jpegFiles)
                {
                    zip.CreateEntryFromFile(file, Path.GetFileName(file));
                }
            }
            // 7. Delete JPEG files after zipping
            foreach (var file in jpegFiles)
            {
                System.IO.File.Delete(file);
            }

            // 8. Return the zip file as download
            var zipBytes = await System.IO.File.ReadAllBytesAsync(zipPath);
            return File(zipBytes, "application/zip", $"JDCA_{ContractId}_JPEG.zip");
        }
        public async Task<IActionResult> OnGetWordContact_JDCA_JPEG_Preview(string ContractId = "5")
        {
            // 1. Generate PDF from JDCA contract
            var htmlContent = await _ControlDataService.OnGetWordContact_ControlDataServiceHtmlToPdf(ContractId, "JDCA");
            await new BrowserFetcher().DownloadAsync();
            await using var browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = true });
            await using var page = await browser.NewPageAsync();

            await page.SetContentAsync(htmlContent);

            var pdfOptions = new PdfOptions
            {
                Format = PuppeteerSharp.Media.PaperFormat.A4,

                Landscape = false,
                MarginOptions = new PuppeteerSharp.Media.MarginOptions
                {
                    Top = "20mm",
                    Bottom = "20mm",
                    Left = "20mm",
                    Right = "20mm"
                },
                PrintBackground = true

            };

            var pdfBytes = await page.PdfDataAsync(pdfOptions);

            // 2. Prepare folder structure
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "JDCA", "JDCA_" + ContractId, "JDCA_" + ContractId + "_JPEG");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var pdfPath = Path.Combine(folderPath, $"JDCA_{ContractId}.pdf");

            // 3. Save PDF to disk
            if (System.IO.File.Exists(pdfPath))
            {
                System.IO.File.Delete(pdfPath);
            }
            await System.IO.File.WriteAllBytesAsync(pdfPath, pdfBytes);

            // 4. Convert each PDF page to JPEG
            using (var pdfStream = new MemoryStream(pdfBytes))
            using (var document = PdfiumViewer.PdfDocument.Load(pdfStream))
            {
                for (int i = 0; i < document.PageCount; i++)
                {
                    using (var image = document.Render(i, 300, 300, true))
                    {
                        var jpegPath = Path.Combine(folderPath, $"JDCA_{ContractId}_p{i + 1}.jpg");
                        image.Save(jpegPath, System.Drawing.Imaging.ImageFormat.Jpeg);
                    }
                }
            }
            // 5. Delete the PDF file after conversion
            System.IO.File.Delete(pdfPath);

            // 6. Create password-protected zip file of JPEGs
            var folderPathZip = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "JDCA", "JDCA_" + ContractId);
            var zipPath = Path.Combine(folderPathZip, $"JDCA_{ContractId}_JPEG_Preview.zip");
            if (System.IO.File.Exists(zipPath))
            {
                System.IO.File.Delete(zipPath);
            }
            var jpegFiles = Directory.GetFiles(folderPath, $"JDCA_{ContractId}_p*.jpg");
            string password = _configuration["Password:PaswordPDF"];

            using (var fsOut = System.IO.File.Create(zipPath))
            using (var zipStream = new ZipOutputStream(fsOut))
            {
                zipStream.SetLevel(9); // 0-9, 9 = best compression
                zipStream.Password = password; // Set password

                foreach (var file in jpegFiles)
                {
                    var entry = new ZipEntry(Path.GetFileName(file))
                    {
                        DateTime = DateTime.Now
                    };
                    zipStream.PutNextEntry(entry);

                    byte[] buffer = System.IO.File.ReadAllBytes(file);
                    zipStream.Write(buffer, 0, buffer.Length);
                    zipStream.CloseEntry();
                }
                zipStream.IsStreamOwner = true;
                zipStream.Close();
            }
            // 7. Delete JPEG files after zipping
            foreach (var file in jpegFiles)
            {
                System.IO.File.Delete(file);
            }

            // 8. Return the password-protected zip file as download
            var zipBytes = await System.IO.File.ReadAllBytesAsync(zipPath);
            return File(zipBytes, "application/zip", $"JDCA_{ContractId}_JPEG_Preview.zip");
        }
        public async Task<IActionResult> OnGetWordContact_JDCA_Word(string ContractId = "5")
        {
            // 1. Get HTML content from the service
            var htmlContent = await _ControlDataService.OnGetWordContact_ControlDataServiceHtmlToPdf(ContractId, "JDCA");

            // 2. Create a Word document from HTML using Spire.Doc
            var document = new Spire.Doc.Document();
            document.LoadFromStream(new MemoryStream(System.Text.Encoding.UTF8.GetBytes(htmlContent)), Spire.Doc.FileFormat.Html);

            // 3. Save the Word document to a MemoryStream
            using var ms = new MemoryStream();
            document.SaveToStream(ms, Spire.Doc.FileFormat.Docx);
            var wordBytes = ms.ToArray();

            // 4. Prepare folder structure
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "JDCA", $"JDCA_{ContractId}");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var filePath = Path.Combine(folderPath, $"JDCA_{ContractId}.docx");

            // 5. Save Word file to disk
            if (System.IO.File.Exists(filePath))
            {
                System.IO.File.Delete(filePath);
            }
            await System.IO.File.WriteAllBytesAsync(filePath, wordBytes);

            // 6. Return the Word file as download
            return File(wordBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", $"JDCA_{ContractId}.docx");
        }
        public async Task<IActionResult> OnGetWordContact_JDCA_Word_Preview(string ContractId = "5")
        {
            // 1. Get HTML content from the service
            var htmlContent = await _ControlDataService.OnGetWordContact_ControlDataServiceHtmlToPdf(ContractId, "JDCA");

            // 2. Create a Word document from HTML using Spire.Doc
            var document = new Spire.Doc.Document();
            document.LoadFromStream(new MemoryStream(System.Text.Encoding.UTF8.GetBytes(htmlContent)), Spire.Doc.FileFormat.Html);

            // 3. Save the Word document to a MemoryStream
            using var ms = new MemoryStream();
            document.SaveToStream(ms, Spire.Doc.FileFormat.Docx);
            var wordBytes = ms.ToArray();

            // 4. Prepare folder structure
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "JDCA", $"JDCA_{ContractId}");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var filePath = Path.Combine(folderPath, $"JDCA_{ContractId}_Preview.docx");

            // 5. Delete the file if it already exists
            if (System.IO.File.Exists(filePath))
            {
                System.IO.File.Delete(filePath);
            }

            // 6. Get password from appsettings.json
            string? userPassword = _configuration["Password:PaswordPDF"];

            // 7. Load the Word file from memory and apply password protection
            using (var msProtect = new MemoryStream(wordBytes))
            {
                var doc = new Spire.Doc.Document();
                doc.LoadFromStream(msProtect, Spire.Doc.FileFormat.Docx);

                // Apply password protection
                doc.Encrypt(userPassword);

                // Save the password-protected file
                doc.SaveToFile(filePath, Spire.Doc.FileFormat.Docx);
            }

            // 8. Return the password-protected Word file as download
            var resultBytes = await System.IO.File.ReadAllBytesAsync(filePath);
            return File(resultBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", $"JDCA_{ContractId}_Preview.docx");
        }

        #endregion 4.1.1.2.5.บันทึกข้อตกลงการเป็นผู้ควบคุมข้อมูลส่วนบุคคลร่วมตัวอย่างหน้าจอ JDCA


        #region 4.1.1.2.4.บันทึกข้อตกลงการประมวลผลข้อมูลส่วนบุคคล PDPA
 
        public async Task OnGetWordContact_PDPA_PDF(string ContractId = "2")
        {
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
            var filePathView = Path.Combine(folderPath, "PDPA_" + ContractId + "_1.pdf");

            // Delete the file if it already exists
            if (System.IO.File.Exists(filePathView))
            {
                System.IO.File.Delete(filePathView);
            }
            var htmlContent = await _PersernalProcessService.OnGetWordContact_PersernalProcessService_HtmlToPDF(ContractId, "PDPA");
            await new BrowserFetcher().DownloadAsync();
            await using var browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = true });
            await using var page = await browser.NewPageAsync();

            await page.SetContentAsync(htmlContent);

            var pdfOptions = new PdfOptions
            {
                Format = PuppeteerSharp.Media.PaperFormat.A4,

                Landscape = false,
                MarginOptions = new PuppeteerSharp.Media.MarginOptions
                {
                    Top = "20mm",
                    Bottom = "20mm",
                    Left = "20mm",
                    Right = "20mm"
                },
                PrintBackground = true

            };

            var pdfBytes = await page.PdfDataAsync(pdfOptions);


            
              await System.IO.File.WriteAllBytesAsync(filePath, pdfBytes);
            await System.IO.File.WriteAllBytesAsync(filePathView, pdfBytes);
            // return File(wordBytes, "application/pdf", "บันทึกข้อตกลงการประมวลผลข้อมูลส่วนบุคคล.pdf");
        }

        public async Task<IActionResult> OnGetWordContact_PDPA_PDF_Preview(string ContractId = "1", string Name = "สมใจ ทดสอบ")
        {
            var htmlContent = await _PersernalProcessService.OnGetWordContact_PersernalProcessService_HtmlToPDF(ContractId, "PDPA");

            await new BrowserFetcher().DownloadAsync();
            await using var browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = true });
            await using var page = await browser.NewPageAsync();

            await page.SetContentAsync(htmlContent);

            var pdfOptions = new PdfOptions
            {
                Format = PuppeteerSharp.Media.PaperFormat.A4,

                Landscape = false,
                MarginOptions = new PuppeteerSharp.Media.MarginOptions
                {
                    Top = "20mm",
                    Bottom = "20mm",
                    Left = "20mm",
                    Right = "20mm"
                },
                PrintBackground = true

            };

            var pdfBytes = await page.PdfDataAsync(pdfOptions);

            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "PDPA");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var fileName = $"PDPA_{ContractId}_Preview.pdf";
            var filePath = Path.Combine(folderPath, fileName);

            // Get password from appsettings.json
            string? userPassword = await GetPdfPasswordAsync(Name);;

            using (var inputStream = new MemoryStream(pdfBytes))
            using (var outputStream = new MemoryStream())
            {
                var document = PdfSharpCore.Pdf.IO.PdfReader.Open(inputStream, PdfSharpCore.Pdf.IO.PdfDocumentOpenMode.Modify);

                foreach (var pdfPage in document.Pages)
                {
                    using (var gfx = PdfSharpCore.Drawing.XGraphics.FromPdfPage(pdfPage))
                    {
                        var font = new PdfSharpCore.Drawing.XFont("Tahoma", 25, PdfSharpCore.Drawing.XFontStyle.Bold);
                        var text = $"สัญญาอิเลคทรอนิกส์พิมพ์ออกโดย {Name}\nวันที่ {DateTime.Now:dd/MM/yyyy HH:mm}";
                        var lines = text.Split('\n');
                        double lineHeight = font.GetHeight();
                        double totalHeight = lineHeight * lines.Length;
                        double y = (pdfPage.Height - totalHeight) / 2;

                        foreach (var line in lines)
                        {
                            var size = gfx.MeasureString(line, font);
                            double x = (pdfPage.Width - size.Width) / 2;
                            var state = gfx.Save();
                            gfx.TranslateTransform(pdfPage.Width / 2, pdfPage.Height / 2);
                            gfx.RotateTransform(-30);
                            gfx.TranslateTransform(-pdfPage.Width / 2, -pdfPage.Height / 2);

                            var brush = new PdfSharpCore.Drawing.XSolidBrush(
                                PdfSharpCore.Drawing.XColor.FromArgb(80, 255, 0, 0)); // semi-transparent red

                            gfx.DrawString(line, font, brush, x, y);
                            gfx.Restore(state);

                            y += lineHeight;
                        }
                    }
                }

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
        public async Task<IActionResult> OnGetWordContact_PDPA_JPEG(string ContractId = "1")
        {
            // 1. Generate PDF from PDPA contract
            var htmlContent = await _PersernalProcessService.OnGetWordContact_PersernalProcessService_HtmlToPDF(ContractId, "PDPA");
            await new BrowserFetcher().DownloadAsync();
            await using var browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = true });
            await using var page = await browser.NewPageAsync();

            await page.SetContentAsync(htmlContent);

            var pdfOptions = new PdfOptions
            {
                Format = PuppeteerSharp.Media.PaperFormat.A4,

                Landscape = false,
                MarginOptions = new PuppeteerSharp.Media.MarginOptions
                {
                    Top = "20mm",
                    Bottom = "20mm",
                    Left = "20mm",
                    Right = "20mm"
                },
                PrintBackground = true

            };

            var pdfBytes = await page.PdfDataAsync(pdfOptions);


            // 2. Prepare folder structure
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "PDPA", "PDPA_" + ContractId, "PDPA_" + ContractId + "_JPEG");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var pdfPath = Path.Combine(folderPath, $"PDPA_{ContractId}.pdf");

            // 3. Save PDF to disk
            if (System.IO.File.Exists(pdfPath))
            {
                System.IO.File.Delete(pdfPath);
            }
            await System.IO.File.WriteAllBytesAsync(pdfPath, pdfBytes);

            // 4. Convert each PDF page to JPEG
            using (var pdfStream = new MemoryStream(pdfBytes))
            using (var document = PdfiumViewer.PdfDocument.Load(pdfStream))
            {
                for (int i = 0; i < document.PageCount; i++)
                {
                    using (var image = document.Render(i, 300, 300, true))
                    {
                        var jpegPath = Path.Combine(folderPath, $"PDPA_{ContractId}_p{i + 1}.jpg");
                        image.Save(jpegPath, System.Drawing.Imaging.ImageFormat.Jpeg);
                    }
                }
            }
            // 5. Delete the PDF file after conversion
            System.IO.File.Delete(pdfPath);

            // 6. Create zip file of JPEGs
            var folderPathZip = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "PDPA", "PDPA_" + ContractId);
            var zipPath = Path.Combine(folderPathZip, $"PDPA_{ContractId}_JPEG.zip");
            if (System.IO.File.Exists(zipPath))
            {
                System.IO.File.Delete(zipPath);
            }
            var jpegFiles = Directory.GetFiles(folderPath, $"PDPA_{ContractId}_p*.jpg");
            using (var zip = System.IO.Compression.ZipFile.Open(zipPath, ZipArchiveMode.Create))
            {
                foreach (var file in jpegFiles)
                {
                    zip.CreateEntryFromFile(file, Path.GetFileName(file));
                }
            }
            // 7. Delete JPEG files after zipping
            foreach (var file in jpegFiles)
            {
                System.IO.File.Delete(file);
            }

            // 8. Return the zip file as download
            var zipBytes = await System.IO.File.ReadAllBytesAsync(zipPath);
            return File(zipBytes, "application/zip", $"PDPA_{ContractId}_JPEG.zip");
        }
        public async Task<IActionResult> OnGetWordContact_PDPA_JPEG_Preview(string ContractId = "1")
        {
            // 1. Generate PDF from PDPA contract
            var htmlContent = await _PersernalProcessService.OnGetWordContact_PersernalProcessService_HtmlToPDF(ContractId, "PDPA");
            await new BrowserFetcher().DownloadAsync();
            await using var browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = true });
            await using var page = await browser.NewPageAsync();

            await page.SetContentAsync(htmlContent);

            var pdfOptions = new PdfOptions
            {
                Format = PuppeteerSharp.Media.PaperFormat.A4,

                Landscape = false,
                MarginOptions = new PuppeteerSharp.Media.MarginOptions
                {
                    Top = "20mm",
                    Bottom = "20mm",
                    Left = "20mm",
                    Right = "20mm"
                },
                PrintBackground = true

            };

            var pdfBytes = await page.PdfDataAsync(pdfOptions);

            // 2. Prepare folder structure
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "PDPA", "PDPA_" + ContractId, "PDPA_" + ContractId + "_JPEG");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var pdfPath = Path.Combine(folderPath, $"PDPA_{ContractId}.pdf");

            // 3. Save PDF to disk
            if (System.IO.File.Exists(pdfPath))
            {
                System.IO.File.Delete(pdfPath);
            }
            await System.IO.File.WriteAllBytesAsync(pdfPath, pdfBytes);

            // 4. Convert each PDF page to JPEG
            using (var pdfStream = new MemoryStream(pdfBytes))
            using (var document = PdfiumViewer.PdfDocument.Load(pdfStream))
            {
                for (int i = 0; i < document.PageCount; i++)
                {
                    using (var image = document.Render(i, 300, 300, true))
                    {
                        var jpegPath = Path.Combine(folderPath, $"PDPA_{ContractId}_p{i + 1}.jpg");
                        image.Save(jpegPath, System.Drawing.Imaging.ImageFormat.Jpeg);
                    }
                }
            }
            // 5. Delete the PDF file after conversion
            System.IO.File.Delete(pdfPath);

            // 6. Create password-protected zip file of JPEGs
            var folderPathZip = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "PDPA", "PDPA_" + ContractId);
            var zipPath = Path.Combine(folderPathZip, $"PDPA_{ContractId}_JPEG_Preview.zip");
            if (System.IO.File.Exists(zipPath))
            {
                System.IO.File.Delete(zipPath);
            }
            var jpegFiles = Directory.GetFiles(folderPath, $"PDPA_{ContractId}_p*.jpg");
            string password = _configuration["Password:PaswordPDF"];

            using (var fsOut = System.IO.File.Create(zipPath))
            using (var zipStream = new ZipOutputStream(fsOut))
            {
                zipStream.SetLevel(9); // 0-9, 9 = best compression
                zipStream.Password = password; // Set password

                foreach (var file in jpegFiles)
                {
                    var entry = new ZipEntry(Path.GetFileName(file))
                    {
                        DateTime = DateTime.Now
                    };
                    zipStream.PutNextEntry(entry);

                    byte[] buffer = System.IO.File.ReadAllBytes(file);
                    zipStream.Write(buffer, 0, buffer.Length);
                    zipStream.CloseEntry();
                }
                zipStream.IsStreamOwner = true;
                zipStream.Close();
            }
            // 7. Delete JPEG files after zipping
            foreach (var file in jpegFiles)
            {
                System.IO.File.Delete(file);
            }

            // 8. Return the password-protected zip file as download
            var zipBytes = await System.IO.File.ReadAllBytesAsync(zipPath);
            return File(zipBytes, "application/zip", $"PDPA_{ContractId}_JPEG_Preview.zip");
        }
        public async Task<IActionResult> OnGetWordContact_PDPA_Word(string ContractId = "1")
        {
            // 1. Get HTML content from the service
            var htmlContent = await _PersernalProcessService.OnGetWordContact_PersernalProcessService_HtmlToPDF(ContractId, "PDPA");


            // 2. Create a Word document from HTML using Spire.Doc
            var document = new Spire.Doc.Document();
            document.LoadFromStream(new MemoryStream(System.Text.Encoding.UTF8.GetBytes(htmlContent)), Spire.Doc.FileFormat.Html);

            // 3. Save the Word document to a MemoryStream
            using var ms = new MemoryStream();
            document.SaveToStream(ms, Spire.Doc.FileFormat.Docx);
            var wordBytes = ms.ToArray();

            // 4. Prepare folder structure
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "PDPA", $"PDPA_{ContractId}");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var filePath = Path.Combine(folderPath, $"PDPA_{ContractId}.docx");

            // 5. Save Word file to disk
            if (System.IO.File.Exists(filePath))
            {
                System.IO.File.Delete(filePath);
            }
            await System.IO.File.WriteAllBytesAsync(filePath, wordBytes);

            // 6. Return the Word file as download
            return File(wordBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", $"PDPA_{ContractId}.docx");
        }
        public async Task<IActionResult> OnGetWordContact_PDPA_Word_Preview(string ContractId = "1")
        {
            // 1. Get HTML content from the service
            var htmlContent = await _PersernalProcessService.OnGetWordContact_PersernalProcessService_HtmlToPDF(ContractId, "PDPA");

            // 2. Create a Word document from HTML using Spire.Doc
            var document = new Spire.Doc.Document();
            document.LoadFromStream(new MemoryStream(System.Text.Encoding.UTF8.GetBytes(htmlContent)), Spire.Doc.FileFormat.Html);

            // 3. Save the Word document to a MemoryStream
            using var ms = new MemoryStream();
            document.SaveToStream(ms, Spire.Doc.FileFormat.Docx);
            var wordBytes = ms.ToArray();

            // 4. Prepare folder structure
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "PDPA", $"PDPA_{ContractId}");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var filePath = Path.Combine(folderPath, $"PDPA_{ContractId}_Preview.docx");

            // 5. Delete the file if it already exists
            if (System.IO.File.Exists(filePath))
            {
                System.IO.File.Delete(filePath);
            }

            // 6. Get password from appsettings.json
            string? userPassword = _configuration["Password:PaswordPDF"];

            // 7. Load the Word file from memory and apply password protection
            using (var msProtect = new MemoryStream(wordBytes))
            {
                var doc = new Spire.Doc.Document();
                doc.LoadFromStream(msProtect, Spire.Doc.FileFormat.Docx);

                // Apply password protection
                doc.Encrypt(userPassword);

                // Save the password-protected file
                doc.SaveToFile(filePath, Spire.Doc.FileFormat.Docx);
            }

            // 8. Return the password-protected Word file as download
            var resultBytes = await System.IO.File.ReadAllBytesAsync(filePath);
            return File(resultBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", $"PDPA_{ContractId}_Preview.docx");
        }



        #endregion 4.1.1.2.4.บันทึกข้อตกลงการประมวลผลข้อมูลส่วนบุคคล PDPA

        #region 4.1.1.2.3.บันทึกข้อตกลงความร่วมมือ MOU
 
        public async Task OnGetWordContact_MOU_PDF(string ContractId = "7")
        {
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

            var filePathView = Path.Combine(folderPath, "MOU_" + ContractId + "_1.pdf");

            // Delete the file if it already exists
            if (System.IO.File.Exists(filePathView))
            {
                System.IO.File.Delete(filePathView);
            }

            var htmlContent = await _MemorandumService.OnGetWordContact_MemorandumService_HtmlToPDF(ContractId, "MOU");
            await new BrowserFetcher().DownloadAsync();
            await using var browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = true });
            await using var page = await browser.NewPageAsync();

            await page.SetContentAsync(htmlContent);

            var pdfOptions = new PdfOptions
            {
                Format = PuppeteerSharp.Media.PaperFormat.A4,

                Landscape = false,
                MarginOptions = new PuppeteerSharp.Media.MarginOptions
                {
                    Top = "20mm",
                    Bottom = "20mm",
                    Left = "20mm",
                    Right = "20mm"
                },
                PrintBackground = true

            };

            var pdfBytes = await page.PdfDataAsync(pdfOptions);

           
            await System.IO.File.WriteAllBytesAsync(filePath, pdfBytes);
            await System.IO.File.WriteAllBytesAsync(filePathView, pdfBytes);
            //   return File(wordBytes, "application/pdf", "MOU_" + ContractId + ".pdf");
        }

        public async Task<IActionResult> OnGetWordContact_MOU_PDF_Preview(string ContractId = "2", string Name = "สมใจ ทดสอบ")
        {
            var htmlContent = await _MemorandumService.OnGetWordContact_MemorandumService_HtmlToPDF(ContractId, "MOU");
            await new BrowserFetcher().DownloadAsync();
            await using var browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = true });
            await using var page = await browser.NewPageAsync();

            await page.SetContentAsync(htmlContent);

            var pdfOptions = new PdfOptions
            {
                Format = PuppeteerSharp.Media.PaperFormat.A4,

                Landscape = false,
                MarginOptions = new PuppeteerSharp.Media.MarginOptions
                {
                    Top = "20mm",
                    Bottom = "20mm",
                    Left = "20mm",
                    Right = "20mm"
                },
                PrintBackground = true

            };

            var pdfBytes = await page.PdfDataAsync(pdfOptions);


            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "MOU");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var fileName = $"MOU_{ContractId}_Preview.pdf";
            var filePath = Path.Combine(folderPath, fileName);

            // Get password from appsettings.json
            string? userPassword = await GetPdfPasswordAsync(Name);;

            using (var inputStream = new MemoryStream(pdfBytes))
            using (var outputStream = new MemoryStream())
            {
                var document = PdfSharpCore.Pdf.IO.PdfReader.Open(inputStream, PdfSharpCore.Pdf.IO.PdfDocumentOpenMode.Modify);

                foreach (var pdfPage in document.Pages)
                {
                    using (var gfx = PdfSharpCore.Drawing.XGraphics.FromPdfPage(pdfPage))
                    {
                        var font = new PdfSharpCore.Drawing.XFont("Tahoma", 25, PdfSharpCore.Drawing.XFontStyle.Bold);
                        var text = $"สัญญาอิเลคทรอนิกส์พิมพ์ออกโดย {Name}\nวันที่ {DateTime.Now:dd/MM/yyyy HH:mm}";
                        var lines = text.Split('\n');
                        double lineHeight = font.GetHeight();
                        double totalHeight = lineHeight * lines.Length;
                        double y = (pdfPage.Height - totalHeight) / 2;

                        foreach (var line in lines)
                        {
                            var size = gfx.MeasureString(line, font);
                            double x = (pdfPage.Width - size.Width) / 2;
                            var state = gfx.Save();
                            gfx.TranslateTransform(pdfPage.Width / 2, pdfPage.Height / 2);
                            gfx.RotateTransform(-30);
                            gfx.TranslateTransform(-pdfPage.Width / 2, -pdfPage.Height / 2);

                            var brush = new PdfSharpCore.Drawing.XSolidBrush(
                                PdfSharpCore.Drawing.XColor.FromArgb(80, 255, 0, 0)); // semi-transparent red

                            gfx.DrawString(line, font, brush, x, y);
                            gfx.Restore(state);

                            y += lineHeight;
                        }
                    }
                }

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
        public async Task<IActionResult> OnGetWordContact_MOU_JPEG(string ContractId = "19")
        {
            // 1. Generate PDF from MOU contract
            var htmlContent = await _MemorandumService.OnGetWordContact_MemorandumService_HtmlToPDF(ContractId, "MOU");
            await new BrowserFetcher().DownloadAsync();
            await using var browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = true });
            await using var page = await browser.NewPageAsync();

            await page.SetContentAsync(htmlContent);

            var pdfOptions = new PdfOptions
            {
                Format = PuppeteerSharp.Media.PaperFormat.A4,

                Landscape = false,
                MarginOptions = new PuppeteerSharp.Media.MarginOptions
                {
                    Top = "20mm",
                    Bottom = "20mm",
                    Left = "20mm",
                    Right = "20mm"
                },
                PrintBackground = true

            };

            var pdfBytes = await page.PdfDataAsync(pdfOptions);


            // 2. Prepare folder structure
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "MOU", "MOU_" + ContractId, "MOU_" + ContractId + "_JPEG");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var pdfPath = Path.Combine(folderPath, $"MOU_{ContractId}.pdf");

            // 3. Save PDF to disk
            if (System.IO.File.Exists(pdfPath))
            {
                System.IO.File.Delete(pdfPath);
            }
            await System.IO.File.WriteAllBytesAsync(pdfPath, pdfBytes);

            // 4. Convert each PDF page to JPEG
            using (var pdfStream = new MemoryStream(pdfBytes))
            using (var document = PdfiumViewer.PdfDocument.Load(pdfStream))
            {
                for (int i = 0; i < document.PageCount; i++)
                {
                    using (var image = document.Render(i, 300, 300, true))
                    {
                        var jpegPath = Path.Combine(folderPath, $"MOU_{ContractId}_p{i + 1}.jpg");
                        image.Save(jpegPath, System.Drawing.Imaging.ImageFormat.Jpeg);
                    }
                }
            }
            // 5. Delete the PDF file after conversion
            System.IO.File.Delete(pdfPath);

            // 6. Create zip file of JPEGs
            var folderPathZip = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "MOU", "MOU_" + ContractId);
            var zipPath = Path.Combine(folderPathZip, $"MOU_{ContractId}_JPEG.zip");
            if (System.IO.File.Exists(zipPath))
            {
                System.IO.File.Delete(zipPath);
            }
            var jpegFiles = Directory.GetFiles(folderPath, $"MOU_{ContractId}_p*.jpg");
            using (var zip = System.IO.Compression.ZipFile.Open(zipPath, ZipArchiveMode.Create))
            {
                foreach (var file in jpegFiles)
                {
                    zip.CreateEntryFromFile(file, Path.GetFileName(file));
                }
            }
            // 7. Delete JPEG files after zipping
            foreach (var file in jpegFiles)
            {
                System.IO.File.Delete(file);
            }

            // 8. Return the zip file as download
            var zipBytes = await System.IO.File.ReadAllBytesAsync(zipPath);
            return File(zipBytes, "application/zip", $"MOU_{ContractId}_JPEG.zip");
        }

        public async Task<IActionResult> OnGetWordContact_MOU_JPEG_Preview(string ContractId = "1")
        {
            // 1. Generate PDF from MOU contract
            var htmlContent = await _MemorandumService.OnGetWordContact_MemorandumService_HtmlToPDF(ContractId, "MOU");
            await new BrowserFetcher().DownloadAsync();
            await using var browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = true });
            await using var page = await browser.NewPageAsync();

            await page.SetContentAsync(htmlContent);

            var pdfOptions = new PdfOptions
            {
                Format = PuppeteerSharp.Media.PaperFormat.A4,

                Landscape = false,
                MarginOptions = new PuppeteerSharp.Media.MarginOptions
                {
                    Top = "20mm",
                    Bottom = "20mm",
                    Left = "20mm",
                    Right = "20mm"
                },
                PrintBackground = true

            };

            var pdfBytes = await page.PdfDataAsync(pdfOptions);


            // 2. Prepare folder structure
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "MOU", "MOU_" + ContractId, "MOU_" + ContractId + "_JPEG");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var pdfPath = Path.Combine(folderPath, $"MOU_{ContractId}.pdf");

            // 3. Save PDF to disk
            if (System.IO.File.Exists(pdfPath))
            {
                System.IO.File.Delete(pdfPath);
            }
            await System.IO.File.WriteAllBytesAsync(pdfPath, pdfBytes);

            // 4. Convert each PDF page to JPEG
            using (var pdfStream = new MemoryStream(pdfBytes))
            using (var document = PdfiumViewer.PdfDocument.Load(pdfStream))
            {
                for (int i = 0; i < document.PageCount; i++)
                {
                    using (var image = document.Render(i, 300, 300, true))
                    {
                        var jpegPath = Path.Combine(folderPath, $"MOU_{ContractId}_p{i + 1}.jpg");
                        image.Save(jpegPath, System.Drawing.Imaging.ImageFormat.Jpeg);
                    }
                }
            }
            // 5. Delete the PDF file after conversion
            System.IO.File.Delete(pdfPath);

            // 6. Create password-protected zip file of JPEGs
            var folderPathZip = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "MOU", "MOU_" + ContractId);
            var zipPath = Path.Combine(folderPathZip, $"MOU_{ContractId}_JPEG_Preview.zip");
            if (System.IO.File.Exists(zipPath))
            {
                System.IO.File.Delete(zipPath);
            }
            var jpegFiles = Directory.GetFiles(folderPath, $"MOU_{ContractId}_p*.jpg");
            string password = _configuration["Password:PaswordPDF"];

            using (var fsOut = System.IO.File.Create(zipPath))
            using (var zipStream = new ZipOutputStream(fsOut))
            {
                zipStream.SetLevel(9); // 0-9, 9 = best compression
                zipStream.Password = password; // Set password

                foreach (var file in jpegFiles)
                {
                    var entry = new ZipEntry(Path.GetFileName(file))
                    {
                        DateTime = DateTime.Now
                    };
                    zipStream.PutNextEntry(entry);

                    byte[] buffer = System.IO.File.ReadAllBytes(file);
                    zipStream.Write(buffer, 0, buffer.Length);
                    zipStream.CloseEntry();
                }
                zipStream.IsStreamOwner = true;
                zipStream.Close();
            }
            // 7. Delete JPEG files after zipping
            foreach (var file in jpegFiles)
            {
                System.IO.File.Delete(file);
            }

            // 8. Return the password-protected zip file as download
            var zipBytes = await System.IO.File.ReadAllBytesAsync(zipPath);
            return File(zipBytes, "application/zip", $"MOU_{ContractId}_JPEG_Preview.zip");
        }
        //  var pdfBytes = await _MemorandumService.OnGetWordContact_MemorandumService_HtmlToPDF(ContractId, "MOU");

        public async Task<IActionResult> OnGetWordContact_MOU_Word(string ContractId = "1")
        {
            // 1. Get HTML content for MOU contract
            var htmlContent = await _MemorandumService.OnGetWordContact_MemorandumService_HtmlToPDF(ContractId, "MOU");

            // 2. Create a Word document from HTML using Spire.Doc
            var document = new Spire.Doc.Document();
            document.LoadFromStream(new MemoryStream(System.Text.Encoding.UTF8.GetBytes(htmlContent)), Spire.Doc.FileFormat.Html);

            // 3. Save the Word document to a MemoryStream
            using var ms = new MemoryStream();
            document.SaveToStream(ms, Spire.Doc.FileFormat.Docx);
            var wordBytes = ms.ToArray();

            // 4. Prepare folder structure
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "MOU", "MOU_" + ContractId);
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var filePath = Path.Combine(folderPath, $"MOU_{ContractId}.docx");

            // 5. Save Word file to disk
            if (System.IO.File.Exists(filePath))
            {
                System.IO.File.Delete(filePath);
            }
            await System.IO.File.WriteAllBytesAsync(filePath, wordBytes);

            // 6. Return the Word file as download
            var resultBytes = await System.IO.File.ReadAllBytesAsync(filePath);
            return File(resultBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", $"MOU_{ContractId}.docx");
        }
        public async Task<IActionResult> OnGetWordContact_MOU_Word_Preview(string ContractId = "1")
        {
            // 1. Get the Word document for MOU contract
            var htmlContent = await _MemorandumService.OnGetWordContact_MemorandumService_HtmlToPDF(ContractId, "MOU");

            // 2. Create a Word document from HTML using Spire.Doc
            var document = new Spire.Doc.Document();
            document.LoadFromStream(new MemoryStream(System.Text.Encoding.UTF8.GetBytes(htmlContent)), Spire.Doc.FileFormat.Html);

            // 3. Get password from appsettings.json
            string? userPassword = _configuration["Password:PaswordPDF"];
            if (!string.IsNullOrEmpty(userPassword))
            {
                document.Encrypt(userPassword);
            }

            // 4. Prepare folder structure
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "MOU", "MOU_" + ContractId);
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var filePath = Path.Combine(folderPath, $"MOU_{ContractId}_Preview.docx");

            // 5. Delete the file if it already exists
            if (System.IO.File.Exists(filePath))
            {
                System.IO.File.Delete(filePath);
            }

            // 6. Save the password-protected Word document
            document.SaveToFile(filePath, FileFormat.Docx);

            // 7. Return the Word file as download
            var resultBytes = await System.IO.File.ReadAllBytesAsync(filePath);
            return File(resultBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", $"MOU_{ContractId}_Preview.docx");
        }
        #endregion  4.1.1.2.3.บันทึกข้อตกลงความร่วมมือ MOU

        #region 4.1.1.2.xxxx.บันทึกข้อตกลงความเข้าใจ MOA
        public async Task OnGetWordContact_MOA_PDF(string ContractId = "4")
        {

            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "MOA");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var filePath = Path.Combine(folderPath, "MOA_" + ContractId + ".pdf");

            // Delete the file if it already exists
            if (System.IO.File.Exists(filePath))
            {
                System.IO.File.Delete(filePath);
            }
            var filePathView = Path.Combine(folderPath, "MOA_" + ContractId + "_1.pdf");

            // Delete the file if it already exists
            if (System.IO.File.Exists(filePathView))
            {
                System.IO.File.Delete(filePathView);
            }

            var htmlContent = await _MemorandumInWritingService.OnGetWordContact_MemorandumInWritingService_HtmlToPDF(ContractId, "MOA");
            await new BrowserFetcher().DownloadAsync();
            await using var browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = true });
            await using var page = await browser.NewPageAsync();

            await page.SetContentAsync(htmlContent);

            var pdfOptions = new PdfOptions
            {
                Format = PuppeteerSharp.Media.PaperFormat.A4,

                Landscape = false,
                MarginOptions = new PuppeteerSharp.Media.MarginOptions
                {
                    Top = "20mm",
                    Bottom = "20mm",
                    Left = "20mm",
                    Right = "20mm"
                },
                PrintBackground = true

            };

            var pdfBytes = await page.PdfDataAsync(pdfOptions);

            await System.IO.File.WriteAllBytesAsync(filePath, pdfBytes);
            //   return File(wordBytes, "application/pdf", "MOU_" + ContractId + ".pdf");
        }
        public async Task<IActionResult> OnGetWordContact_MOA_PDF_Preview(string ContractId = "8", string Name = "สมใจ ทดสอบ")
        {
            var htmlContent = await _MemorandumInWritingService.OnGetWordContact_MemorandumInWritingService_HtmlToPDF(ContractId, "MOA");
            await new BrowserFetcher().DownloadAsync();
            await using var browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = true });
            await using var page = await browser.NewPageAsync();

            await page.SetContentAsync(htmlContent);

            var pdfOptions = new PdfOptions
            {
                Format = PuppeteerSharp.Media.PaperFormat.A4,

                Landscape = false,
                MarginOptions = new PuppeteerSharp.Media.MarginOptions
                {
                    Top = "20mm",
                    Bottom = "20mm",
                    Left = "20mm",
                    Right = "20mm"
                },
                PrintBackground = true

            };

            var pdfBytes = await page.PdfDataAsync(pdfOptions);

            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "MOA");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var fileName = $"MOA_{ContractId}_Preview.pdf";
            var filePath = Path.Combine(folderPath, fileName);

            // Get password from appsettings.json
            string? userPassword = await GetPdfPasswordAsync(Name);;

            using (var inputStream = new MemoryStream(pdfBytes))
            using (var outputStream = new MemoryStream())
            {
                var document = PdfSharpCore.Pdf.IO.PdfReader.Open(inputStream, PdfSharpCore.Pdf.IO.PdfDocumentOpenMode.Modify);

                foreach (var pdfPage in document.Pages)
                {
                    using (var gfx = PdfSharpCore.Drawing.XGraphics.FromPdfPage(pdfPage))
                    {
                        var font = new PdfSharpCore.Drawing.XFont("Tahoma", 25, PdfSharpCore.Drawing.XFontStyle.Bold);
                        var text = $"สัญญาอิเลคทรอนิกส์พิมพ์ออกโดย {Name}\nวันที่ {DateTime.Now:dd/MM/yyyy HH:mm}";
                        var lines = text.Split('\n');
                        double lineHeight = font.GetHeight();
                        double totalHeight = lineHeight * lines.Length;
                        double y = (pdfPage.Height - totalHeight) / 2;

                        foreach (var line in lines)
                        {
                            var size = gfx.MeasureString(line, font);
                            double x = (pdfPage.Width - size.Width) / 2;
                            var state = gfx.Save();
                            gfx.TranslateTransform(pdfPage.Width / 2, pdfPage.Height / 2);
                            gfx.RotateTransform(-30);
                            gfx.TranslateTransform(-pdfPage.Width / 2, -pdfPage.Height / 2);

                            var brush = new PdfSharpCore.Drawing.XSolidBrush(
                                PdfSharpCore.Drawing.XColor.FromArgb(80, 255, 0, 0)); // semi-transparent red

                            gfx.DrawString(line, font, brush, x, y);
                            gfx.Restore(state);

                            y += lineHeight;
                        }
                    }
                }

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
        public async Task<IActionResult> OnGetWordContact_MOA_JPEG(string ContractId = "8")
        {
            // 1. Generate PDF from MOU contract
            var htmlContent = await _MemorandumInWritingService.OnGetWordContact_MemorandumInWritingService_HtmlToPDF(ContractId, "MOA");
            await new BrowserFetcher().DownloadAsync();
            await using var browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = true });
            await using var page = await browser.NewPageAsync();

            await page.SetContentAsync(htmlContent);

            var pdfOptions = new PdfOptions
            {
                Format = PuppeteerSharp.Media.PaperFormat.A4,

                Landscape = false,
                MarginOptions = new PuppeteerSharp.Media.MarginOptions
                {
                    Top = "20mm",
                    Bottom = "20mm",
                    Left = "20mm",
                    Right = "20mm"
                },
                PrintBackground = true

            };

            var pdfBytes = await page.PdfDataAsync(pdfOptions);

            // 2. Prepare folder structure
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "MOA", "MOA_" + ContractId, "MOA_" + ContractId + "_JPEG");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var pdfPath = Path.Combine(folderPath, $"MOA_{ContractId}.pdf");

            // 3. Save PDF to disk
            if (System.IO.File.Exists(pdfPath))
            {
                System.IO.File.Delete(pdfPath);
            }
            await System.IO.File.WriteAllBytesAsync(pdfPath, pdfBytes);

            // 4. Convert each PDF page to JPEG
            using (var pdfStream = new MemoryStream(pdfBytes))
            using (var document = PdfiumViewer.PdfDocument.Load(pdfStream))
            {
                for (int i = 0; i < document.PageCount; i++)
                {
                    using (var image = document.Render(i, 300, 300, true))
                    {
                        var jpegPath = Path.Combine(folderPath, $"MOA_{ContractId}_p{i + 1}.jpg");
                        image.Save(jpegPath, System.Drawing.Imaging.ImageFormat.Jpeg);
                    }
                }
            }
            // 5. Delete the PDF file after conversion
            System.IO.File.Delete(pdfPath);

            // 6. Create zip file of JPEGs
            var folderPathZip = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "MOA", "MOA_" + ContractId);
            var zipPath = Path.Combine(folderPathZip, $"MOA_{ContractId}_JPEG.zip");
            if (System.IO.File.Exists(zipPath))
            {
                System.IO.File.Delete(zipPath);
            }
            var jpegFiles = Directory.GetFiles(folderPath, $"MOA_{ContractId}_p*.jpg");
            using (var zip = System.IO.Compression.ZipFile.Open(zipPath, ZipArchiveMode.Create))
            {
                foreach (var file in jpegFiles)
                {
                    zip.CreateEntryFromFile(file, Path.GetFileName(file));
                }
            }
            // 7. Delete JPEG files after zipping
            foreach (var file in jpegFiles)
            {
                System.IO.File.Delete(file);
            }

            // 8. Return the zip file as download
            var zipBytes = await System.IO.File.ReadAllBytesAsync(zipPath);
            return File(zipBytes, "application/zip", $"MOA_{ContractId}_JPEG.zip");
        }

        public async Task<IActionResult> OnGetWordContact_MOA_JPEG_Preview(string ContractId = "8")
        {
            // 1. Generate PDF from MOU contract
            var htmlContent = await _MemorandumInWritingService.OnGetWordContact_MemorandumInWritingService_HtmlToPDF(ContractId, "MOA");
            await new BrowserFetcher().DownloadAsync();
            await using var browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = true });
            await using var page = await browser.NewPageAsync();

            await page.SetContentAsync(htmlContent);

            var pdfOptions = new PdfOptions
            {
                Format = PuppeteerSharp.Media.PaperFormat.A4,

                Landscape = false,
                MarginOptions = new PuppeteerSharp.Media.MarginOptions
                {
                    Top = "20mm",
                    Bottom = "20mm",
                    Left = "20mm",
                    Right = "20mm"
                },
                PrintBackground = true

            };

            var pdfBytes = await page.PdfDataAsync(pdfOptions);

            // 2. Prepare folder structure
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "MOA", "MOA_" + ContractId, "MOA_" + ContractId + "_JPEG");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var pdfPath = Path.Combine(folderPath, $"MOA_{ContractId}.pdf");

            // 3. Save PDF to disk
            if (System.IO.File.Exists(pdfPath))
            {
                System.IO.File.Delete(pdfPath);
            }
            await System.IO.File.WriteAllBytesAsync(pdfPath, pdfBytes);

            // 4. Convert each PDF page to JPEG
            using (var pdfStream = new MemoryStream(pdfBytes))
            using (var document = PdfiumViewer.PdfDocument.Load(pdfStream))
            {
                for (int i = 0; i < document.PageCount; i++)
                {
                    using (var image = document.Render(i, 300, 300, true))
                    {
                        var jpegPath = Path.Combine(folderPath, $"MOA_{ContractId}_p{i + 1}.jpg");
                        image.Save(jpegPath, System.Drawing.Imaging.ImageFormat.Jpeg);
                    }
                }
            }
            // 5. Delete the PDF file after conversion
            System.IO.File.Delete(pdfPath);

            // 6. Create password-protected zip file of JPEGs
            var folderPathZip = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "MOA", "MOA_" + ContractId);
            var zipPath = Path.Combine(folderPathZip, $"MOA_{ContractId}_JPEG_Preview.zip");
            if (System.IO.File.Exists(zipPath))
            {
                System.IO.File.Delete(zipPath);
            }
            var jpegFiles = Directory.GetFiles(folderPath, $"MOA_{ContractId}_p*.jpg");
            string password = _configuration["Password:PaswordPDF"];

            using (var fsOut = System.IO.File.Create(zipPath))
            using (var zipStream = new ZipOutputStream(fsOut))
            {
                zipStream.SetLevel(9); // 0-9, 9 = best compression
                zipStream.Password = password; // Set password

                foreach (var file in jpegFiles)
                {
                    var entry = new ZipEntry(Path.GetFileName(file))
                    {
                        DateTime = DateTime.Now
                    };
                    zipStream.PutNextEntry(entry);

                    byte[] buffer = System.IO.File.ReadAllBytes(file);
                    zipStream.Write(buffer, 0, buffer.Length);
                    zipStream.CloseEntry();
                }
                zipStream.IsStreamOwner = true;
                zipStream.Close();
            }
            // 7. Delete JPEG files after zipping
            foreach (var file in jpegFiles)
            {
                System.IO.File.Delete(file);
            }

            // 8. Return the password-protected zip file as download
            var zipBytes = await System.IO.File.ReadAllBytesAsync(zipPath);
            return File(zipBytes, "application/zip", $"MOA_{ContractId}_JPEG_Preview.zip");
        }
        public async Task<IActionResult> OnGetWordContact_MOA_Word(string ContractId = "8")
        {
            // 1. Get HTML content from the service
            var htmlContent = await _MemorandumInWritingService.OnGetWordContact_MemorandumInWritingService_HtmlToPDF(ContractId, "MOA");

            // 2. Create a Word document from HTML using Spire.Doc
            var document = new Spire.Doc.Document();
            document.LoadFromStream(new MemoryStream(System.Text.Encoding.UTF8.GetBytes(htmlContent)), Spire.Doc.FileFormat.Html);

            // 3. Save the Word document to a MemoryStream
            using var ms = new MemoryStream();
            document.SaveToStream(ms, Spire.Doc.FileFormat.Docx);
            var wordBytes = ms.ToArray();

            // 4. Prepare folder structure
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "MOA", $"MOA_{ContractId}");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var filePath = Path.Combine(folderPath, $"MOA_{ContractId}.docx");

            // 5. Save Word file to disk
            if (System.IO.File.Exists(filePath))
            {
                System.IO.File.Delete(filePath);
            }
            await System.IO.File.WriteAllBytesAsync(filePath, wordBytes);

            // 6. Return the Word file as download
            return File(wordBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", $"MOA_{ContractId}.docx");
        }
        public async Task<IActionResult> OnGetWordContact_MOA_Word_Preview(string ContractId = "8")
        {
            // 1. Get HTML content from the service
            var htmlContent = await _MemorandumInWritingService.OnGetWordContact_MemorandumInWritingService_HtmlToPDF(ContractId, "MOA");

            // 2. Create a Word document from HTML using Spire.Doc
            var document = new Spire.Doc.Document();
            document.LoadFromStream(new MemoryStream(System.Text.Encoding.UTF8.GetBytes(htmlContent)), Spire.Doc.FileFormat.Html);

            // 3. Save the Word document to a MemoryStream
            using var ms = new MemoryStream();
            document.SaveToStream(ms, Spire.Doc.FileFormat.Docx);
            var wordBytes = ms.ToArray();

            // 4. Prepare folder structure
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "MOA", $"MOA_{ContractId}");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var filePath = Path.Combine(folderPath, $"MOA_{ContractId}_Preview.docx");

            // 5. Delete the file if it already exists
            if (System.IO.File.Exists(filePath))
            {
                System.IO.File.Delete(filePath);
            }

            // 6. Get password from appsettings.json
            string? userPassword = _configuration["Password:PaswordPDF"];

            // 7. Load the Word file from memory and apply password protection
            using (var msProtect = new MemoryStream(wordBytes))
            {
                var doc = new Spire.Doc.Document();
                doc.LoadFromStream(msProtect, Spire.Doc.FileFormat.Docx);

                // Apply password protection
                doc.Encrypt(userPassword);

                // Save the password-protected file
                doc.SaveToFile(filePath, Spire.Doc.FileFormat.Docx);
            }

            // 8. Return the password-protected Word file as download
            var resultBytes = await System.IO.File.ReadAllBytesAsync(filePath);
            return File(resultBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", $"MOA_{ContractId}_Preview.docx");
        }     
        #endregion  4.1.1.2.3.บันทึกข้อตกลงความเข้าใจ MOA




        #region 4.1.1.2.2.สัญญารับเงินอุดหนุน GA
        public async Task<IActionResult> OnGetWordContact_GA(string ContractId = "1")
        {
            var wordBytes = await _SupportSMEsService.OnGetWordContact_SupportSMEsService(ContractId);
            return File(wordBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "บันทึกข้อตกลงความร่วมมือในการสนับสนุน SMEs.docx");
        }
        public async Task OnGetWordContact_GA_PDF(string ContractId = "1")
        {
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
            var filePathView = Path.Combine(folderPath, "GA_" + ContractId + "_1.pdf");

            // Delete the file if it already exists
            if (System.IO.File.Exists(filePathView))
            {
                System.IO.File.Delete(filePathView);
            }


            var htmlContent = await _SupportSMEsService.OnGetWordContact_SupportSMEsService_HtmlToPDF(ContractId, "GA");
            // 2. Convert HTML to PDF using DinkToPdf
            await new BrowserFetcher().DownloadAsync();
            await using var browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = true });
            await using var page = await browser.NewPageAsync();

            await page.SetContentAsync(htmlContent);

            var pdfOptions = new PdfOptions
            {
                Format = PuppeteerSharp.Media.PaperFormat.A4,

                Landscape = false,
                MarginOptions = new PuppeteerSharp.Media.MarginOptions
                {
                    Top = "20mm",
                    Bottom = "20mm",
                    Left = "20mm",
                    Right = "20mm"
                },
                PrintBackground = true

            };

            var pdfBytes = await page.PdfDataAsync(pdfOptions);

        
            await System.IO.File.WriteAllBytesAsync(filePath, pdfBytes);
            await System.IO.File.WriteAllBytesAsync(filePathView, pdfBytes);
            
            // return File(pdfBytes, "application/pdf", "GA_"+ContractId+".pdf");
        }

        public async Task<IActionResult> OnGetWordContact_GA_PDF_Preview(string ContractId = "1", string Name = "สมใจ ทดสอบ")
        {
            var htmlContent = await _SupportSMEsService.OnGetWordContact_SupportSMEsService_HtmlToPDF(ContractId, "GA");
            await new BrowserFetcher().DownloadAsync();
            await using var browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = true });
            await using var page = await browser.NewPageAsync();

            await page.SetContentAsync(htmlContent);

            var pdfOptions = new PdfOptions
            {
                Format = PuppeteerSharp.Media.PaperFormat.A4,

                Landscape = false,
                MarginOptions = new PuppeteerSharp.Media.MarginOptions
                {
                    Top = "20mm",
                    Bottom = "20mm",
                    Left = "20mm",
                    Right = "20mm"
                },
                PrintBackground = true

            };

            var pdfBytes = await page.PdfDataAsync(pdfOptions);


            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "GA");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var fileName = $"GA_{ContractId}_Preview.pdf";
            var filePath = Path.Combine(folderPath, fileName);

            string? userPassword = await GetPdfPasswordAsync(Name);;

            using (var inputStream = new MemoryStream(pdfBytes))
            using (var outputStream = new MemoryStream())
            {
                var document = PdfSharpCore.Pdf.IO.PdfReader.Open(inputStream, PdfSharpCore.Pdf.IO.PdfDocumentOpenMode.Modify);

                foreach (var pdfPage in document.Pages)
                {
                    using (var gfx = PdfSharpCore.Drawing.XGraphics.FromPdfPage(pdfPage))
                    {
                        var font = new PdfSharpCore.Drawing.XFont("Tahoma", 25, PdfSharpCore.Drawing.XFontStyle.Bold);
                        var text = $"สัญญาอิเลคทรอนิกส์พิมพ์ออกโดย {Name}\nวันที่ {DateTime.Now:dd/MM/yyyy HH:mm}";
                        var lines = text.Split('\n');
                        double lineHeight = font.GetHeight();
                        double totalHeight = lineHeight * lines.Length;
                        double y = (pdfPage.Height - totalHeight) / 2;

                        foreach (var line in lines)
                        {
                            var size = gfx.MeasureString(line, font);
                            double x = (pdfPage.Width - size.Width) / 2;
                            var state = gfx.Save();
                            gfx.TranslateTransform(pdfPage.Width / 2, pdfPage.Height / 2);
                            gfx.RotateTransform(-30);
                            gfx.TranslateTransform(-pdfPage.Width / 2, -pdfPage.Height / 2);

                            var brush = new PdfSharpCore.Drawing.XSolidBrush(
                                PdfSharpCore.Drawing.XColor.FromArgb(80, 255, 0, 0)); // semi-transparent red

                            gfx.DrawString(line, font, brush, x, y);
                            gfx.Restore(state);

                            y += lineHeight;
                        }
                    }
                }

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
        public async Task<IActionResult> OnGetWordContact_GA_JPEG(string ContractId = "1")
        {
            // 1. Generate PDF from GA contract (HTML to PDF)
            var htmlContent = await _SupportSMEsService.OnGetWordContact_SupportSMEsService_HtmlToPDF(ContractId, "GA");
            await new BrowserFetcher().DownloadAsync();
            await using var browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = true });
            await using var page = await browser.NewPageAsync();

            await page.SetContentAsync(htmlContent);

            var pdfOptions = new PdfOptions
            {
                Format = PuppeteerSharp.Media.PaperFormat.A4,

                Landscape = false,
                MarginOptions = new PuppeteerSharp.Media.MarginOptions
                {
                    Top = "20mm",
                    Bottom = "20mm",
                    Left = "20mm",
                    Right = "20mm"
                },
                PrintBackground = true

            };

            var pdfBytes = await page.PdfDataAsync(pdfOptions);


            // 2. Prepare folder structure
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "GA", $"GA_{ContractId}", $"GA_{ContractId}_JPEG");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var pdfPath = Path.Combine(folderPath, $"GA_{ContractId}.pdf");

            // 3. Save PDF to disk
            if (System.IO.File.Exists(pdfPath))
            {
                System.IO.File.Delete(pdfPath);
            }
            await System.IO.File.WriteAllBytesAsync(pdfPath, pdfBytes);

            // 4. Convert each PDF page to JPEG
            using (var pdfStream = new MemoryStream(pdfBytes))
            using (var document = PdfiumViewer.PdfDocument.Load(pdfStream))
            {
                for (int i = 0; i < document.PageCount; i++)
                {
                    using (var image = document.Render(i, 300, 300, true))
                    {
                        var jpegPath = Path.Combine(folderPath, $"GA_{ContractId}_p{i + 1}.jpg");
                        image.Save(jpegPath, System.Drawing.Imaging.ImageFormat.Jpeg);
                    }
                }
            }
            // 5. Delete the PDF file after conversion
            System.IO.File.Delete(pdfPath);

            // 6. Create zip file of JPEGs
            var folderPathZip = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "GA", $"GA_{ContractId}");
            var zipPath = Path.Combine(folderPathZip, $"GA_{ContractId}_JPEG.zip");
            if (System.IO.File.Exists(zipPath))
            {
                System.IO.File.Delete(zipPath);
            }
            var jpegFiles = Directory.GetFiles(folderPath, $"GA_{ContractId}_p*.jpg");
            using (var zip = System.IO.Compression.ZipFile.Open(zipPath, ZipArchiveMode.Create))
            {
                foreach (var file in jpegFiles)
                {
                    zip.CreateEntryFromFile(file, Path.GetFileName(file));
                }
            }
            // 7. Delete JPEG files after zipping
            foreach (var file in jpegFiles)
            {
                System.IO.File.Delete(file);
            }

            // 8. Return the zip file as download
            var zipBytes = await System.IO.File.ReadAllBytesAsync(zipPath);
            return File(zipBytes, "application/zip", $"GA_{ContractId}_JPEG.zip");
        }
        public async Task<IActionResult> OnGetWordContact_GA_JPEG_Preview(string ContractId = "1")
        {
            // 1. Generate PDF from GA contract (HTML to PDF)
            var htmlContent = await _SupportSMEsService.OnGetWordContact_SupportSMEsService_HtmlToPDF(ContractId, "GA");
            await new BrowserFetcher().DownloadAsync();
            await using var browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = true });
            await using var page = await browser.NewPageAsync();

            await page.SetContentAsync(htmlContent);

            var pdfOptions = new PdfOptions
            {
                Format = PuppeteerSharp.Media.PaperFormat.A4,

                Landscape = false,
                MarginOptions = new PuppeteerSharp.Media.MarginOptions
                {
                    Top = "20mm",
                    Bottom = "20mm",
                    Left = "20mm",
                    Right = "20mm"
                },
                PrintBackground = true

            };

            var pdfBytes = await page.PdfDataAsync(pdfOptions);


            // 2. Prepare folder structure
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "GA", $"GA_{ContractId}", $"GA_{ContractId}_JPEG");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var pdfPath = Path.Combine(folderPath, $"GA_{ContractId}.pdf");

            // 3. Save PDF to disk
            if (System.IO.File.Exists(pdfPath))
            {
                System.IO.File.Delete(pdfPath);
            }
            await System.IO.File.WriteAllBytesAsync(pdfPath, pdfBytes);

            // 4. Convert each PDF page to JPEG
            using (var pdfStream = new MemoryStream(pdfBytes))
            using (var document = PdfiumViewer.PdfDocument.Load(pdfStream))
            {
                for (int i = 0; i < document.PageCount; i++)
                {
                    using (var image = document.Render(i, 300, 300, true))
                    {
                        var jpegPath = Path.Combine(folderPath, $"GA_{ContractId}_p{i + 1}.jpg");
                        image.Save(jpegPath, System.Drawing.Imaging.ImageFormat.Jpeg);
                    }
                }
            }
            // 5. Delete the PDF file after conversion
            System.IO.File.Delete(pdfPath);

            // 6. Create password-protected zip file of JPEGs
            var folderPathZip = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "GA", $"GA_{ContractId}");
            var zipPath = Path.Combine(folderPathZip, $"GA_{ContractId}_JPEG_Preview.zip");
            if (System.IO.File.Exists(zipPath))
            {
                System.IO.File.Delete(zipPath);
            }
            var jpegFiles = Directory.GetFiles(folderPath, $"GA_{ContractId}_p*.jpg");
            string password = _configuration["Password:PaswordPDF"];

            using (var fsOut = System.IO.File.Create(zipPath))
            using (var zipStream = new ICSharpCode.SharpZipLib.Zip.ZipOutputStream(fsOut))
            {
                zipStream.SetLevel(9); // 0-9, 9 = best compression
                zipStream.Password = password; // Set password

                foreach (var file in jpegFiles)
                {
                    var entry = new ICSharpCode.SharpZipLib.Zip.ZipEntry(Path.GetFileName(file))
                    {
                        DateTime = DateTime.Now
                    };
                    zipStream.PutNextEntry(entry);

                    byte[] buffer = System.IO.File.ReadAllBytes(file);
                    zipStream.Write(buffer, 0, buffer.Length);
                    zipStream.CloseEntry();
                }
                zipStream.IsStreamOwner = true;
                zipStream.Close();
            }
            // 7. Delete JPEG files after zipping
            foreach (var file in jpegFiles)
            {
                System.IO.File.Delete(file);
            }

            // 8. Return the password-protected zip file as download
            var zipBytes = await System.IO.File.ReadAllBytesAsync(zipPath);
            return File(zipBytes, "application/zip", $"GA_{ContractId}_JPEG_Preview.zip");
        }
        public async Task<IActionResult> OnGetWordContact_GA_Word(string ContractId = "1")
        {
            // 1. Get HTML content from the service
            var htmlContent = await _SupportSMEsService.OnGetWordContact_SupportSMEsService_HtmlToPDF(ContractId, "GA");

            // 2. Create a Word document from HTML using Spire.Doc
            var document = new Spire.Doc.Document();
            document.LoadFromStream(new MemoryStream(System.Text.Encoding.UTF8.GetBytes(htmlContent)), Spire.Doc.FileFormat.Html);

            // 3. Save the Word document to a MemoryStream
            using var ms = new MemoryStream();
            document.SaveToStream(ms, Spire.Doc.FileFormat.Docx);
            var wordBytes = ms.ToArray();

            // 4. Prepare folder structure
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "GA", $"GA_{ContractId}");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var filePath = Path.Combine(folderPath, $"GA_{ContractId}.docx");

            // 5. Save Word file to disk
            if (System.IO.File.Exists(filePath))
            {
                System.IO.File.Delete(filePath);
            }
            await System.IO.File.WriteAllBytesAsync(filePath, wordBytes);

            // 6. Return the Word file as download
            return File(wordBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", $"GA_{ContractId}.docx");
        }
        public async Task<IActionResult> OnGetWordContact_GA_Word_Preview(string ContractId = "1")
        {
            // 1. Get HTML content from the service
            var htmlContent = await _SupportSMEsService.OnGetWordContact_SupportSMEsService_HtmlToPDF(ContractId, "GA");

            // 2. Create a Word document from HTML using Spire.Doc
            var document = new Spire.Doc.Document();
            document.LoadFromStream(new MemoryStream(System.Text.Encoding.UTF8.GetBytes(htmlContent)), Spire.Doc.FileFormat.Html);

            // 3. Save the Word document to a MemoryStream
            using var ms = new MemoryStream();
            document.SaveToStream(ms, Spire.Doc.FileFormat.Docx);
            var wordBytes = ms.ToArray();

            // 4. Prepare folder structure
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "GA", $"GA_{ContractId}");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var filePath = Path.Combine(folderPath, $"GA_{ContractId}_Preview.docx");

            // 5. Delete the file if it already exists
            if (System.IO.File.Exists(filePath))
            {
                System.IO.File.Delete(filePath);
            }

            // 6. Get password from appsettings.json
            string? userPassword = _configuration["Password:PaswordPDF"];

            // 7. Load the Word file from memory and apply password protection
            using (var msProtect = new MemoryStream(wordBytes))
            {
                var doc = new Spire.Doc.Document();
                doc.LoadFromStream(msProtect, Spire.Doc.FileFormat.Docx);

                // Apply password protection
                doc.Encrypt(userPassword);

                // Save the password-protected file
                doc.SaveToFile(filePath, Spire.Doc.FileFormat.Docx);
            }

            // 8. Return the password-protected Word file as download
            var resultBytes = await System.IO.File.ReadAllBytesAsync(filePath);
            return File(resultBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", $"GA_{ContractId}_Preview.docx");
        }
        #endregion 4.1.1.2.2.สัญญารับเงินอุดหนุน GA

        #region 4.1.1.2.1.สัญญาร่วมดำเนินการ JOA
        // Add this P/Invoke at the top of your file or in a static class
        [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        static extern bool MoveFileEx(string lpExistingFileName, string? lpNewFileName, MoveFileFlags dwFlags);

        [Flags]
        enum MoveFileFlags
        {
            MOVEFILE_REPLACE_EXISTING = 0x1,
            MOVEFILE_COPY_ALLOWED = 0x2,
            MOVEFILE_DELAY_UNTIL_REBOOT = 0x4,
            MOVEFILE_WRITE_THROUGH = 0x8
        }
        public async Task OnGetWordContact_JOA_PDF(string ContractId = "95", string Name = "สมใจ ทดสอบ")
        {
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "JOA");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var filePath = Path.Combine(folderPath, $"JOA_{ContractId}.pdf");
            var filePathView = Path.Combine(folderPath, $"JOA_{ContractId}_1.pdf");
            // Delete old file if exists
            if (System.IO.File.Exists(filePath))
            {
                try
                {
                    System.IO.File.Delete(filePath);
                }
                catch (IOException)
                {
                    // If file is locked, schedule for deletion on next reboot (Windows only)
                    MoveFileEx(filePath, null, MoveFileFlags.MOVEFILE_DELAY_UNTIL_REBOOT);
                }
            }
            if (System.IO.File.Exists(filePathView))
            {
                try
                {
                    System.IO.File.Delete(filePathView);
                }
                catch (IOException)
                {
                    // If file is locked, schedule for deletion on next reboot (Windows only)
                    MoveFileEx(filePath, null, MoveFileFlags.MOVEFILE_DELAY_UNTIL_REBOOT);
                }
            }

            // Generate new PDF file
            var htmlContent = await _JointOperationService.OnGetWordContact_JointOperationServiceHtmlToPDF(ContractId);
            await new BrowserFetcher().DownloadAsync();
            await using var browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = true });
            await using var page = await browser.NewPageAsync();
            await page.SetContentAsync(htmlContent);

            var pdfOptions = new PdfOptions
            {
                Format = PuppeteerSharp.Media.PaperFormat.A4,
                Landscape = false,
                MarginOptions = new PuppeteerSharp.Media.MarginOptions
                {
                    Top = "20mm",
                    Bottom = "20mm",
                    Left = "20mm",
                    Right = "20mm"
                },
                PrintBackground = true
            };

            var pdfBytes = await page.PdfDataAsync(pdfOptions);
            await System.IO.File.WriteAllBytesAsync(filePath, pdfBytes);
            await System.IO.File.WriteAllBytesAsync(filePathView, pdfBytes);
        }
        public async Task<string?> GetPdfPasswordAsync(string? empId, CancellationToken ct = default)
        {
            if (string.IsNullOrWhiteSpace(empId))
                return null;

            // HrDAO จะคืนรหัสจาก DB หรือ fallback เอง
            return await _eContractDao.GetPdfPasswordByEmpIdAsync(empId, ct);
        }

        public async Task<IActionResult> OnGetWordContact_JOA_PDF_Preview(string ContractId = "3", string Name = "สมใจ ทดสอบ")
        {
            // 1. Get HTML content
            var htmlContent = await _JointOperationService.OnGetWordContact_JointOperationServiceHtmlToPDF(ContractId);

            await new BrowserFetcher().DownloadAsync();
            await using var browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = true });
            await using var page = await browser.NewPageAsync();

            await page.SetContentAsync(htmlContent);

            var pdfOptions = new PdfOptions
            {
                Format = PuppeteerSharp.Media.PaperFormat.A4,

                Landscape = false,
                MarginOptions = new PuppeteerSharp.Media.MarginOptions
                {
                    Top = "20mm",
                    Bottom = "20mm",
                    Left = "20mm",
                    Right = "20mm"
                },
                PrintBackground = true

            };

            var pdfBytes = await page.PdfDataAsync(pdfOptions);


            // 3. Prepare folder structure
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "JOA", "JOA_" + ContractId);
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var fileName = $"JOA_{ContractId}_Preview.pdf";
            var filePath = Path.Combine(folderPath, fileName);

            // 4. Get password from appsettings.json
            //string? userPassword = await GetPdfPasswordAsync(Name);;
            string? userPassword = await GetPdfPasswordAsync(Name);

            // 5. Add watermark and password protection
            using (var inputStream = new MemoryStream(pdfBytes))
            using (var outputStream = new MemoryStream())
            {
                var document = PdfSharpCore.Pdf.IO.PdfReader.Open(inputStream, PdfSharpCore.Pdf.IO.PdfDocumentOpenMode.Modify);

                foreach (var pdfPage in document.Pages)
                {
                    using (var gfx = PdfSharpCore.Drawing.XGraphics.FromPdfPage(pdfPage))
                    {
                        var font = new PdfSharpCore.Drawing.XFont("Tahoma", 25, PdfSharpCore.Drawing.XFontStyle.Bold);
                        var text = $"สัญญาอิเลคทรอนิกส์พิมพ์ออกโดย {Name}\nวันที่ {DateTime.Now:dd/MM/yyyy HH:mm}";
                        var lines = text.Split('\n');
                        double lineHeight = font.GetHeight();
                        double totalHeight = lineHeight * lines.Length;
                        double y = (pdfPage.Height - totalHeight) / 2;

                        foreach (var line in lines)
                        {
                            var size = gfx.MeasureString(line, font);
                            double x = (pdfPage.Width - size.Width) / 2;
                            var state = gfx.Save();
                            gfx.TranslateTransform(pdfPage.Width / 2, pdfPage.Height / 2);
                            gfx.RotateTransform(-30);
                            gfx.TranslateTransform(-pdfPage.Width / 2, -pdfPage.Height / 2);

                            var brush = new PdfSharpCore.Drawing.XSolidBrush(
                                PdfSharpCore.Drawing.XColor.FromArgb(80, 255, 0, 0)); // semi-transparent red

                            gfx.DrawString(line, font, brush, x, y);
                            gfx.Restore(state);

                            y += lineHeight;
                        }
                    }
                }

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
        public async Task<IActionResult> OnGetWordContact_JOA_JPEG(string ContractId = "1")
        {
            // 1. Generate PDF from JOA contract (HTML to PDF)
            var htmlContent = await _JointOperationService.OnGetWordContact_JointOperationServiceHtmlToPDF(ContractId);
            await new BrowserFetcher().DownloadAsync();
            await using var browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = true });
            await using var page = await browser.NewPageAsync();

            await page.SetContentAsync(htmlContent);

            var pdfOptions = new PdfOptions
            {
                Format = PuppeteerSharp.Media.PaperFormat.A4,

                Landscape = false,
                MarginOptions = new PuppeteerSharp.Media.MarginOptions
                {
                    Top = "20mm",
                    Bottom = "20mm",
                    Left = "20mm",
                    Right = "20mm"
                },
                PrintBackground = true

            };

            var pdfBytes = await page.PdfDataAsync(pdfOptions);

            // 2. Prepare folder structure
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "JOA", $"JOA_{ContractId}", $"JOA_{ContractId}_JPEG");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var pdfPath = Path.Combine(folderPath, $"JOA_{ContractId}.pdf");

            // 3. Save PDF to disk
            if (System.IO.File.Exists(pdfPath))
            {
                System.IO.File.Delete(pdfPath);
            }
            await System.IO.File.WriteAllBytesAsync(pdfPath, pdfBytes);

            // 4. Convert each PDF page to JPEG
            try
            {
                using (var pdfStream = new MemoryStream(pdfBytes))
                using (var document = PdfiumViewer.PdfDocument.Load(pdfStream))
                {
                    for (int i = 0; i < document.PageCount; i++)
                    {
                        using (var image = document.Render(i, 300, 300, true))
                        {
                            var jpegPath = Path.Combine(folderPath, $"JOA_{ContractId}_p{i + 1}.jpg");
                            image.Save(jpegPath, System.Drawing.Imaging.ImageFormat.Jpeg);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // Clean up if conversion fails
                if (System.IO.File.Exists(pdfPath))
                {
                    System.IO.File.Delete(pdfPath);
                }
                return BadRequest($"PDF to JPEG conversion failed: {ex.Message}");
            }

            // 5. Delete the PDF file after conversion
            if (System.IO.File.Exists(pdfPath))
            {
                System.IO.File.Delete(pdfPath);
            }

            // 6. Create zip file of JPEGs
            var folderPathZip = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "JOA", $"JOA_{ContractId}");
            var zipPath = Path.Combine(folderPathZip, $"JOA_{ContractId}_JPEG.zip");
            if (System.IO.File.Exists(zipPath))
            {
                System.IO.File.Delete(zipPath);
            }
            var jpegFiles = Directory.GetFiles(folderPath, $"JOA_{ContractId}_p*.jpg");
            try
            {
                using (var zip = System.IO.Compression.ZipFile.Open(zipPath, ZipArchiveMode.Create))
                {
                    foreach (var file in jpegFiles)
                    {
                        zip.CreateEntryFromFile(file, Path.GetFileName(file));
                    }
                }
            }
            catch (Exception ex)
            {
                return BadRequest($"JPEG to ZIP failed: {ex.Message}");
            }

            // 7. Delete JPEG files after zipping
            foreach (var file in jpegFiles)
            {
                System.IO.File.Delete(file);
            }

            // 8. Return the zip file as download
            var zipBytes = await System.IO.File.ReadAllBytesAsync(zipPath);
            return File(zipBytes, "application/zip", $"JOA_{ContractId}_JPEG.zip");
        }
        public async Task<IActionResult> OnGetWordContact_JOA_JPEG_Preview(string ContractId = "1")
        {
            // 1. Generate PDF from JOA contract (HTML to PDF)
            var htmlContent = await _JointOperationService.OnGetWordContact_JointOperationServiceHtmlToPDF(ContractId);
            await new BrowserFetcher().DownloadAsync();
            await using var browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = true });
            await using var page = await browser.NewPageAsync();

            await page.SetContentAsync(htmlContent);

            var pdfOptions = new PdfOptions
            {
                Format = PuppeteerSharp.Media.PaperFormat.A4,

                Landscape = false,
                MarginOptions = new PuppeteerSharp.Media.MarginOptions
                {
                    Top = "20mm",
                    Bottom = "20mm",
                    Left = "20mm",
                    Right = "20mm"
                },
                PrintBackground = true

            };

            var pdfBytes = await page.PdfDataAsync(pdfOptions);

            // 2. Prepare folder structure
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "JOA", $"JOA_{ContractId}", $"JOA_{ContractId}_JPEG");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var pdfPath = Path.Combine(folderPath, $"JOA_{ContractId}.pdf");

            // 3. Save PDF to disk
            if (System.IO.File.Exists(pdfPath))
            {
                System.IO.File.Delete(pdfPath);
            }
            await System.IO.File.WriteAllBytesAsync(pdfPath, pdfBytes);

            // 4. Convert each PDF page to JPEG
            try
            {
                using (var pdfStream = new MemoryStream(pdfBytes))
                using (var document = PdfiumViewer.PdfDocument.Load(pdfStream))
                {
                    for (int i = 0; i < document.PageCount; i++)
                    {
                        using (var image = document.Render(i, 300, 300, true))
                        {
                            var jpegPath = Path.Combine(folderPath, $"JOA_{ContractId}_p{i + 1}.jpg");
                            image.Save(jpegPath, System.Drawing.Imaging.ImageFormat.Jpeg);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                if (System.IO.File.Exists(pdfPath))
                {
                    System.IO.File.Delete(pdfPath);
                }
                return BadRequest($"PDF to JPEG conversion failed: {ex.Message}");
            }

            // 5. Delete the PDF file after conversion
            if (System.IO.File.Exists(pdfPath))
            {
                System.IO.File.Delete(pdfPath);
            }

            // 6. Create password-protected zip file of JPEGs
            var folderPathZip = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "JOA", $"JOA_{ContractId}");
            var zipPath = Path.Combine(folderPathZip, $"JOA_{ContractId}_JPEG_Preview.zip");
            if (System.IO.File.Exists(zipPath))
            {
                System.IO.File.Delete(zipPath);
            }
            var jpegFiles = Directory.GetFiles(folderPath, $"JOA_{ContractId}_p*.jpg");
            string password = _configuration["Password:PaswordPDF"];

            try
            {
                using (var fsOut = System.IO.File.Create(zipPath))
                using (var zipStream = new ICSharpCode.SharpZipLib.Zip.ZipOutputStream(fsOut))
                {
                    zipStream.SetLevel(9); // 0-9, 9 = best compression
                    zipStream.Password = password; // Set password

                    foreach (var file in jpegFiles)
                    {
                        var entry = new ICSharpCode.SharpZipLib.Zip.ZipEntry(Path.GetFileName(file))
                        {
                            DateTime = DateTime.Now
                        };
                        zipStream.PutNextEntry(entry);

                        byte[] buffer = System.IO.File.ReadAllBytes(file);
                        zipStream.Write(buffer, 0, buffer.Length);
                        zipStream.CloseEntry();
                    }
                    zipStream.IsStreamOwner = true;
                    zipStream.Close();
                }
            }
            catch (Exception ex)
            {
                return BadRequest($"JPEG to ZIP failed: {ex.Message}");
            }

            // 7. Delete JPEG files after zipping
            foreach (var file in jpegFiles)
            {
                System.IO.File.Delete(file);
            }

            // 8. Return the password-protected zip file as download
            var zipBytes = await System.IO.File.ReadAllBytesAsync(zipPath);
            return File(zipBytes, "application/zip", $"JOA_{ContractId}_JPEG_Preview.zip");
        }
        public async Task<IActionResult> OnGetWordContact_JOA_Word(string ContractId = "1")
        {
            // 1. Get HTML content from the service
            var htmlContent = await _JointOperationService.OnGetWordContact_JointOperationServiceHtmlToPDF(ContractId);

            // 2. Create a Word document from HTML using Spire.Doc
            var document = new Spire.Doc.Document();
            document.LoadFromStream(new MemoryStream(System.Text.Encoding.UTF8.GetBytes(htmlContent)), Spire.Doc.FileFormat.Html);

            // 3. Save the Word document to a MemoryStream
            using var ms = new MemoryStream();
            document.SaveToStream(ms, Spire.Doc.FileFormat.Docx);
            var wordBytes = ms.ToArray();

            // 4. Prepare folder structure
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "JOA", $"JOA_{ContractId}");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var filePath = Path.Combine(folderPath, $"JOA_{ContractId}.docx");

            // 5. Save Word file to disk
            if (System.IO.File.Exists(filePath))
            {
                System.IO.File.Delete(filePath);
            }
            await System.IO.File.WriteAllBytesAsync(filePath, wordBytes);

            // 6. Return the Word file as download
            return File(wordBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", $"JOA_{ContractId}.docx");
        }
        public async Task<IActionResult> OnGetWordContact_JOA_Word_Preview(string ContractId = "1")
        {
            // 1. Get HTML content from the service
            var htmlContent = await _JointOperationService.OnGetWordContact_JointOperationServiceHtmlToPDF(ContractId);

            // 2. Create a Word document from HTML using Spire.Doc
            var document = new Spire.Doc.Document();
            document.LoadFromStream(new MemoryStream(System.Text.Encoding.UTF8.GetBytes(htmlContent)), Spire.Doc.FileFormat.Html);

            // 3. Save the Word document to a MemoryStream
            using var ms = new MemoryStream();
            document.SaveToStream(ms, Spire.Doc.FileFormat.Docx);
            var wordBytes = ms.ToArray();

            // 4. Prepare folder structure
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "JOA", $"JOA_{ContractId}");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var filePath = Path.Combine(folderPath, $"JOA_{ContractId}_Preview.docx");

            // 5. Delete the file if it already exists
            if (System.IO.File.Exists(filePath))
            {
                System.IO.File.Delete(filePath);
            }

            // 6. Get password from appsettings.json
            string? userPassword = _configuration["Password:PaswordPDF"];

            // 7. Load the Word file from memory and apply password protection
            using (var msProtect = new MemoryStream(wordBytes))
            {
                var doc = new Spire.Doc.Document();
                doc.LoadFromStream(msProtect, Spire.Doc.FileFormat.Docx);

                // Apply password protection
                doc.Encrypt(userPassword);

                // Save the password-protected file
                doc.SaveToFile(filePath, Spire.Doc.FileFormat.Docx);
            }

            // 8. Return the password-protected Word file as download
            var resultBytes = await System.IO.File.ReadAllBytesAsync(filePath);
            return File(resultBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", $"JOA_{ContractId}_Preview.docx");
        }        
        #endregion 4.1.1.2.1.สัญญาร่วมดำเนินการ JOA


        #region 4.1.1.2.16 แบบฟอร์มบันทึกข้อตกลงเป็นหนังสือ MIW

        public async Task OnGetWordContact_MIW_PDF(string ContractId = "3")
        {

            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "MIW");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var filePath = Path.Combine(folderPath, "MIW_" + ContractId + ".pdf");
            var filePathView = Path.Combine(folderPath, "MIW_" + ContractId + "_1.pdf");

            if (System.IO.File.Exists(filePath))
            {
                System.IO.File.Delete(filePath);
            }
            if (System.IO.File.Exists(filePathView))
            {
                System.IO.File.Delete(filePathView);
            }
            // 1. Get HTML content from the service

            var htmlContent = await _MIWService.OnGetWordContact_MIWServiceHtmlToPDF(ContractId);

            // 2. Convert HTML to PDF using DinkToPdf
            await new BrowserFetcher().DownloadAsync();
            await using var browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = true });
            await using var page = await browser.NewPageAsync();

            await page.SetContentAsync(htmlContent);

            var pdfOptions = new PdfOptions
            {
                Format = PuppeteerSharp.Media.PaperFormat.A4,

                Landscape = false,
                MarginOptions = new PuppeteerSharp.Media.MarginOptions
                {
                    Top = "20mm",
                    Bottom = "20mm",
                    Left = "20mm",
                    Right = "20mm"
                },
                PrintBackground = true

            };

            var pdfBytes = await page.PdfDataAsync(pdfOptions);

            
            await System.IO.File.WriteAllBytesAsync(filePath, pdfBytes);
            await System.IO.File.WriteAllBytesAsync(filePathView, pdfBytes);
            // Optionally, return the file as a download:
            // return File(pdfBytes, "application/pdf", "MIW_" + ContractId + ".pdf");
        }
        public async Task<IActionResult> OnGetWordContact_MIW_PDF_Preview(string ContractId = "1", string Name = "สมใจ ทดสอบ")
        {
            // 1. Get HTML content from the service
            var htmlContent = await _MIWService.OnGetWordContact_MIWServiceHtmlToPDF(ContractId);

            // 2. Convert HTML to PDF using DinkToPdf
            await new BrowserFetcher().DownloadAsync();
            await using var browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = true });
            await using var page = await browser.NewPageAsync();

            await page.SetContentAsync(htmlContent);

            var pdfOptions = new PdfOptions
            {
                Format = PuppeteerSharp.Media.PaperFormat.A4,

                Landscape = false,
                MarginOptions = new PuppeteerSharp.Media.MarginOptions
                {
                    Top = "20mm",
                    Bottom = "20mm",
                    Left = "20mm",
                    Right = "20mm"
                },
                PrintBackground = true

            };

            var pdfBytes = await page.PdfDataAsync(pdfOptions);

            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "MIW", "MIW_" + ContractId);
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var fileName = $"MIW_{ContractId}_Preview.pdf";
            var filePath = Path.Combine(folderPath, fileName);

            // Get password from appsettings.json
            string? userPassword = await GetPdfPasswordAsync(Name);;

            using (var inputStream = new MemoryStream(pdfBytes))
            using (var outputStream = new MemoryStream())
            {
                var document = PdfSharpCore.Pdf.IO.PdfReader.Open(inputStream, PdfSharpCore.Pdf.IO.PdfDocumentOpenMode.Modify);

                foreach (var pdfPage in document.Pages)
                {
                    using (var gfx = PdfSharpCore.Drawing.XGraphics.FromPdfPage(pdfPage))
                    {
                        var font = new PdfSharpCore.Drawing.XFont("Tahoma", 25, PdfSharpCore.Drawing.XFontStyle.Bold);
                        var text = $"สัญญาอิเลคทรอนิกส์พิมพ์ออกโดย {Name}\nวันที่ {DateTime.Now:dd/MM/yyyy HH:mm}";
                        var lines = text.Split('\n');
                        double lineHeight = font.GetHeight();
                        double totalHeight = lineHeight * lines.Length;
                        double y = (pdfPage.Height - totalHeight) / 2;

                        foreach (var line in lines)
                        {
                            var size = gfx.MeasureString(line, font);
                            double x = (pdfPage.Width - size.Width) / 2;
                            var state = gfx.Save();
                            gfx.TranslateTransform(pdfPage.Width / 2, pdfPage.Height / 2);
                            gfx.RotateTransform(-30);
                            gfx.TranslateTransform(-pdfPage.Width / 2, -pdfPage.Height / 2);

                            var brush = new PdfSharpCore.Drawing.XSolidBrush(
                                PdfSharpCore.Drawing.XColor.FromArgb(80, 255, 0, 0)); // semi-transparent red

                            gfx.DrawString(line, font, brush, x, y);
                            gfx.Restore(state);

                            y += lineHeight;
                        }
                    }
                }

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
        public async Task<IActionResult> OnGetWordContact_MIW_JPEG(string ContractId = "12")
        {
            // 1. Get HTML content and convert to PDF bytes
            var htmlContent = await _MIWService.OnGetWordContact_MIWServiceHtmlToPDF(ContractId);
            await new BrowserFetcher().DownloadAsync();
            await using var browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = true });
            await using var page = await browser.NewPageAsync();

            await page.SetContentAsync(htmlContent);

            var pdfOptions = new PdfOptions
            {
                Format = PuppeteerSharp.Media.PaperFormat.A4,

                Landscape = false,
                MarginOptions = new PuppeteerSharp.Media.MarginOptions
                {
                    Top = "20mm",
                    Bottom = "20mm",
                    Left = "20mm",
                    Right = "20mm"
                },
                PrintBackground = true

            };

            var pdfBytes = await page.PdfDataAsync(pdfOptions);

            // 2. Prepare folder structure
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "MIW", "MIW_" + ContractId, "MIW_" + ContractId + "_JPEG");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var pdfPath = Path.Combine(folderPath, $"MIW_{ContractId}.pdf");

            // 3. Save PDF to disk
            if (System.IO.File.Exists(pdfPath))
            {
                System.IO.File.Delete(pdfPath);
            }
            await System.IO.File.WriteAllBytesAsync(pdfPath, pdfBytes);

            // 4. Convert each PDF page to JPEG
            using (var pdfStream = new MemoryStream(pdfBytes))
            using (var document = PdfiumViewer.PdfDocument.Load(pdfStream))
            {
                for (int i = 0; i < document.PageCount; i++)
                {
                    using (var image = document.Render(i, 300, 300, true))
                    {
                        var jpegPath = Path.Combine(folderPath, $"MIW_{ContractId}_p{i + 1}.jpg");
                        image.Save(jpegPath, System.Drawing.Imaging.ImageFormat.Jpeg);
                    }
                }
            }
            // 5. Delete the PDF file after conversion
            System.IO.File.Delete(pdfPath);

            // 6. Create zip file of JPEGs
            var folderPathZip = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "MIW", "MIW_" + ContractId);
            var zipPath = Path.Combine(folderPathZip, $"MIW_{ContractId}_JPEG.zip");
            if (System.IO.File.Exists(zipPath))
            {
                System.IO.File.Delete(zipPath);
            }
            var jpegFiles = Directory.GetFiles(folderPath, $"MIW_{ContractId}_p*.jpg");
            using (var zip = System.IO.Compression.ZipFile.Open(zipPath, ZipArchiveMode.Create))
            {
                foreach (var file in jpegFiles)
                {
                    zip.CreateEntryFromFile(file, Path.GetFileName(file));
                }
            }
            // 7. Delete JPEG files after zipping
            foreach (var file in jpegFiles)
            {
                System.IO.File.Delete(file);
            }

            // 8. Return the zip file as download
            var zipBytes = await System.IO.File.ReadAllBytesAsync(zipPath);
            return File(zipBytes, "application/zip", $"MIW_{ContractId}_JPEG.zip");
        }
        public async Task<IActionResult> OnGetWordContact_MIW_JPEG_Preview(string ContractId = "12")
        {
            // 1. Get HTML content and convert to PDF bytes
            var htmlContent = await _MIWService.OnGetWordContact_MIWServiceHtmlToPDF(ContractId);
            await new BrowserFetcher().DownloadAsync();
            await using var browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = true });
            await using var page = await browser.NewPageAsync();

            await page.SetContentAsync(htmlContent);

            var pdfOptions = new PdfOptions
            {
                Format = PuppeteerSharp.Media.PaperFormat.A4,

                Landscape = false,
                MarginOptions = new PuppeteerSharp.Media.MarginOptions
                {
                    Top = "20mm",
                    Bottom = "20mm",
                    Left = "20mm",
                    Right = "20mm"
                },
                PrintBackground = true

            };

            var pdfBytes = await page.PdfDataAsync(pdfOptions);

            // 2. Prepare folder structure
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "MIW", "MIW_" + ContractId, "MIW_" + ContractId + "_JPEG");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var pdfPath = Path.Combine(folderPath, $"MIW_{ContractId}.pdf");

            // 3. Save PDF to disk
            if (System.IO.File.Exists(pdfPath))
            {
                System.IO.File.Delete(pdfPath);
            }
            await System.IO.File.WriteAllBytesAsync(pdfPath, pdfBytes);

            // 4. Convert each PDF page to JPEG
            using (var pdfStream = new MemoryStream(pdfBytes))
            using (var document = PdfiumViewer.PdfDocument.Load(pdfStream))
            {
                for (int i = 0; i < document.PageCount; i++)
                {
                    using (var image = document.Render(i, 300, 300, true))
                    {
                        var jpegPath = Path.Combine(folderPath, $"MIW_{ContractId}_p{i + 1}.jpg");
                        image.Save(jpegPath, System.Drawing.Imaging.ImageFormat.Jpeg);
                    }
                }
            }
            // 5. Delete the PDF file after conversion
            System.IO.File.Delete(pdfPath);

            // 6. Create password-protected zip file of JPEGs
            var folderPathZip = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "MIW", "MIW_" + ContractId);
            var zipPath = Path.Combine(folderPathZip, $"MIW_{ContractId}_JPEG_Preview.zip");
            if (System.IO.File.Exists(zipPath))
            {
                System.IO.File.Delete(zipPath);
            }
            var jpegFiles = Directory.GetFiles(folderPath, $"MIW_{ContractId}_p*.jpg");
            string password = _configuration["Password:PaswordPDF"];

            using (var fsOut = System.IO.File.Create(zipPath))
            using (var zipStream = new ICSharpCode.SharpZipLib.Zip.ZipOutputStream(fsOut))
            {
                zipStream.SetLevel(9); // 0-9, 9 = best compression
                zipStream.Password = password; // Set password

                foreach (var file in jpegFiles)
                {
                    var entry = new ICSharpCode.SharpZipLib.Zip.ZipEntry(Path.GetFileName(file))
                    {
                        DateTime = DateTime.Now
                    };
                    zipStream.PutNextEntry(entry);

                    byte[] buffer = System.IO.File.ReadAllBytes(file);
                    zipStream.Write(buffer, 0, buffer.Length);
                    zipStream.CloseEntry();
                }
                zipStream.IsStreamOwner = true;
                zipStream.Close();
            }
            // 7. Delete JPEG files after zipping
            foreach (var file in jpegFiles)
            {
                System.IO.File.Delete(file);
            }

            // 8. Return the password-protected zip file as download
            var zipBytes = await System.IO.File.ReadAllBytesAsync(zipPath);
            return File(zipBytes, "application/zip", $"MIW_{ContractId}_JPEG_Preview.zip");
        }
        public async Task<IActionResult> OnGetWordContact_MIW_Word(string ContractId = "12")
        {
            // 1. Get HTML content from the service
            var htmlContent = await _MIWService.OnGetWordContact_MIWServiceHtmlToPDF(ContractId);

            // 2. Create a Word document from HTML using Spire.Doc
            var document = new Spire.Doc.Document();
            document.LoadFromStream(new MemoryStream(System.Text.Encoding.UTF8.GetBytes(htmlContent)), Spire.Doc.FileFormat.Html);

            // 3. Save the Word document to a MemoryStream
            using var ms = new MemoryStream();
            document.SaveToStream(ms, Spire.Doc.FileFormat.Docx);
            var wordBytes = ms.ToArray();

            // 4. Prepare folder structure
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "MIW", "MIW_" + ContractId);
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var filePath = Path.Combine(folderPath, $"MIW_{ContractId}.docx");

            // 5. Save Word file to disk
            if (System.IO.File.Exists(filePath))
            {
                System.IO.File.Delete(filePath);
            }
            await System.IO.File.WriteAllBytesAsync(filePath, wordBytes);

            // 6. Return the Word file as download
            return File(wordBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", $"MIW_{ContractId}.docx");
        }
        public async Task<IActionResult> OnGetWordContact_MIW_Word_Preview(string ContractId = "12")
        {
            // 1. Get HTML content from the service
            var htmlContent = await _MIWService.OnGetWordContact_MIWServiceHtmlToPDF(ContractId);

            // 2. Create a Word document from HTML using Spire.Doc
            var document = new Spire.Doc.Document();
            document.LoadFromStream(new MemoryStream(System.Text.Encoding.UTF8.GetBytes(htmlContent)), Spire.Doc.FileFormat.Html);

            // 3. Save the Word document to a MemoryStream
            using var ms = new MemoryStream();
            document.SaveToStream(ms, Spire.Doc.FileFormat.Docx);
            var wordBytes = ms.ToArray();

            // 4. Prepare folder structure
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "MIW", "MIW_" + ContractId);
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var filePath = Path.Combine(folderPath, $"MIW_{ContractId}_Preview.docx");

            // 5. Delete the file if it already exists
            if (System.IO.File.Exists(filePath))
            {
                System.IO.File.Delete(filePath);
            }

            // 6. Get password from appsettings.json
            string? userPassword = _configuration["Password:PaswordPDF"];

            // 7. Load the Word file from memory and apply password protection
            using (var msProtect = new MemoryStream(wordBytes))
            {
                Document doc = new Document();
                doc.LoadFromStream(msProtect, FileFormat.Docx);

                // Apply password protection
                doc.Encrypt(userPassword);

                // Save the password-protected file
                doc.SaveToFile(filePath, FileFormat.Docx);
            }

            // 8. Return the password-protected Word file as download
            var resultBytes = await System.IO.File.ReadAllBytesAsync(filePath);
            return File(resultBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", $"MIW_{ContractId}_Preview.docx");
        }
        #endregion  4.1.1.2.16 แบบฟอร์มบันทึกข้อตกลงเป็นหนังสือ MIW


        #region  4.1.6 เอกสารแนบท้ายบันทึกข้อตกลงความร่วมมือและสัญญาร่วมดำเนินการ AMJOA

        public async Task OnGetWordContact_AMJOA_PDF(string ContractId = "2")
        {
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "AMJOA");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var filePath = Path.Combine(folderPath, "AMJOA_" + ContractId + ".pdf");

            if (System.IO.File.Exists(filePath))
            {
                System.IO.File.Delete(filePath);
            }
            var filePathView = Path.Combine(folderPath, "AMJOA_" + ContractId + "_1.pdf");

            if (System.IO.File.Exists(filePathView))
            {
                System.IO.File.Delete(filePathView);
            }
            // 1. Get HTML content from the service
            var htmlContent = await _AMJOAService.OnGetWordContact_AMJOAServiceHtmlToPDF(ContractId);

            // 2. Convert HTML to PDF using DinkToPdf
            await new BrowserFetcher().DownloadAsync();
            await using var browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = true });
            await using var page = await browser.NewPageAsync();

            await page.SetContentAsync(htmlContent);

            var pdfOptions = new PdfOptions
            {
                Format = PuppeteerSharp.Media.PaperFormat.A4,

                Landscape = false,
                MarginOptions = new PuppeteerSharp.Media.MarginOptions
                {
                    Top = "20mm",
                    Bottom = "20mm",
                    Left = "20mm",
                    Right = "20mm"
                },
                PrintBackground = true

            };

            var pdfBytes = await page.PdfDataAsync(pdfOptions);

         
            await System.IO.File.WriteAllBytesAsync(filePath, pdfBytes);
            await System.IO.File.WriteAllBytesAsync(filePathView, pdfBytes);

            // Optionally, return the file as a download:
            // return File(pdfBytes, "application/pdf", "AMJOA_" + ContractId + ".pdf");
        }
        public async Task<IActionResult> OnGetWordContact_AMJOA_PDF_Preview(string ContractId = "2", string Name = "สมใจ ทดสอบ")
        {
            // 1. Get HTML content from the service
            var htmlContent = await _AMJOAService.OnGetWordContact_AMJOAServiceHtmlToPDF(ContractId);

            // 2. Convert HTML to PDF using DinkToPdf
            await new BrowserFetcher().DownloadAsync();
            await using var browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = true });
            await using var page = await browser.NewPageAsync();

            await page.SetContentAsync(htmlContent);

            var pdfOptions = new PdfOptions
            {
                Format = PuppeteerSharp.Media.PaperFormat.A4,

                Landscape = false,
                MarginOptions = new PuppeteerSharp.Media.MarginOptions
                {
                    Top = "20mm",
                    Bottom = "20mm",
                    Left = "20mm",
                    Right = "20mm"
                },
                PrintBackground = true

            };

            var pdfBytes = await page.PdfDataAsync(pdfOptions);

            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "AMJOA", "AMJOA_" + ContractId);
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var fileName = $"AMJOA_{ContractId}_Preview.pdf";
            var filePath = Path.Combine(folderPath, fileName);

            // Get password from appsettings.json
            string? userPassword = await GetPdfPasswordAsync(Name);;

            using (var inputStream = new MemoryStream(pdfBytes))
            using (var outputStream = new MemoryStream())
            {
                var document = PdfSharpCore.Pdf.IO.PdfReader.Open(inputStream, PdfSharpCore.Pdf.IO.PdfDocumentOpenMode.Modify);

                foreach (var pdfPage in document.Pages)
                {
                    using (var gfx = PdfSharpCore.Drawing.XGraphics.FromPdfPage(pdfPage))
                    {
                        var font = new PdfSharpCore.Drawing.XFont("Tahoma", 25, PdfSharpCore.Drawing.XFontStyle.Bold);
                        var text = $"สัญญาอิเลคทรอนิกส์พิมพ์ออกโดย {Name}\nวันที่ {DateTime.Now:dd/MM/yyyy HH:mm}";
                        var lines = text.Split('\n');
                        double lineHeight = font.GetHeight();
                        double totalHeight = lineHeight * lines.Length;
                        double y = (pdfPage.Height - totalHeight) / 2;

                        foreach (var line in lines)
                        {
                            var size = gfx.MeasureString(line, font);
                            double x = (pdfPage.Width - size.Width) / 2;
                            var state = gfx.Save();
                            gfx.TranslateTransform(pdfPage.Width / 2, pdfPage.Height / 2);
                            gfx.RotateTransform(-30);
                            gfx.TranslateTransform(-pdfPage.Width / 2, -pdfPage.Height / 2);

                            var brush = new PdfSharpCore.Drawing.XSolidBrush(
                                PdfSharpCore.Drawing.XColor.FromArgb(80, 255, 0, 0)); // semi-transparent red

                            gfx.DrawString(line, font, brush, x, y);
                            gfx.Restore(state);

                            y += lineHeight;
                        }
                    }
                }

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
        public async Task<IActionResult> OnGetWordContact_AMJOA_JPEG(string ContractId = "12")
        {
            // 1. Get HTML content and convert to PDF bytes
            var htmlContent = await _AMJOAService.OnGetWordContact_AMJOAServiceHtmlToPDF(ContractId);
            await new BrowserFetcher().DownloadAsync();
            await using var browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = true });
            await using var page = await browser.NewPageAsync();

            await page.SetContentAsync(htmlContent);

            var pdfOptions = new PdfOptions
            {
                Format = PuppeteerSharp.Media.PaperFormat.A4,

                Landscape = false,
                MarginOptions = new PuppeteerSharp.Media.MarginOptions
                {
                    Top = "20mm",
                    Bottom = "20mm",
                    Left = "20mm",
                    Right = "20mm"
                },
                PrintBackground = true

            };

            var pdfBytes = await page.PdfDataAsync(pdfOptions);

            // 2. Prepare folder structure
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "AMJOA", "AMJOA_" + ContractId, "AMJOA_" + ContractId + "_JPEG");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var pdfPath = Path.Combine(folderPath, $"AMJOA_{ContractId}.pdf");

            // 3. Save PDF to disk
            if (System.IO.File.Exists(pdfPath))
            {
                System.IO.File.Delete(pdfPath);
            }
            await System.IO.File.WriteAllBytesAsync(pdfPath, pdfBytes);

            // 4. Convert each PDF page to JPEG
            using (var pdfStream = new MemoryStream(pdfBytes))
            using (var document = PdfiumViewer.PdfDocument.Load(pdfStream))
            {
                for (int i = 0; i < document.PageCount; i++)
                {
                    using (var image = document.Render(i, 300, 300, true))
                    {
                        var jpegPath = Path.Combine(folderPath, $"AMJOA_{ContractId}_p{i + 1}.jpg");
                        image.Save(jpegPath, System.Drawing.Imaging.ImageFormat.Jpeg);
                    }
                }
            }
            // 5. Delete the PDF file after conversion
            System.IO.File.Delete(pdfPath);

            // 6. Create zip file of JPEGs
            var folderPathZip = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "AMJOA", "AMJOA_" + ContractId);
            var zipPath = Path.Combine(folderPathZip, $"AMJOA_{ContractId}_JPEG.zip");
            if (System.IO.File.Exists(zipPath))
            {
                System.IO.File.Delete(zipPath);
            }
            var jpegFiles = Directory.GetFiles(folderPath, $"AMJOA_{ContractId}_p*.jpg");
            using (var zip = System.IO.Compression.ZipFile.Open(zipPath, ZipArchiveMode.Create))
            {
                foreach (var file in jpegFiles)
                {
                    zip.CreateEntryFromFile(file, Path.GetFileName(file));
                }
            }
            // 7. Delete JPEG files after zipping
            foreach (var file in jpegFiles)
            {
                System.IO.File.Delete(file);
            }

            // 8. Return the zip file as download
            var zipBytes = await System.IO.File.ReadAllBytesAsync(zipPath);
            return File(zipBytes, "application/zip", $"AMJOA_{ContractId}_JPEG.zip");
        }
        public async Task<IActionResult> OnGetWordContact_AMJOA_JPEG_Preview(string ContractId = "12")
        {
            // 1. Get HTML content and convert to PDF bytes
            var htmlContent = await _AMJOAService.OnGetWordContact_AMJOAServiceHtmlToPDF(ContractId);
            await new BrowserFetcher().DownloadAsync();
            await using var browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = true });
            await using var page = await browser.NewPageAsync();

            await page.SetContentAsync(htmlContent);

            var pdfOptions = new PdfOptions
            {
                Format = PuppeteerSharp.Media.PaperFormat.A4,

                Landscape = false,
                MarginOptions = new PuppeteerSharp.Media.MarginOptions
                {
                    Top = "20mm",
                    Bottom = "20mm",
                    Left = "20mm",
                    Right = "20mm"
                },
                PrintBackground = true

            };

            var pdfBytes = await page.PdfDataAsync(pdfOptions);

            // 2. Prepare folder structure
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "AMJOA", "AMJOA_" + ContractId, "AMJOA_" + ContractId + "_JPEG");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var pdfPath = Path.Combine(folderPath, $"AMJOA_{ContractId}.pdf");

            // 3. Save PDF to disk
            if (System.IO.File.Exists(pdfPath))
            {
                System.IO.File.Delete(pdfPath);
            }
            await System.IO.File.WriteAllBytesAsync(pdfPath, pdfBytes);

            // 4. Convert each PDF page to JPEG
            using (var pdfStream = new MemoryStream(pdfBytes))
            using (var document = PdfiumViewer.PdfDocument.Load(pdfStream))
            {
                for (int i = 0; i < document.PageCount; i++)
                {
                    using (var image = document.Render(i, 300, 300, true))
                    {
                        var jpegPath = Path.Combine(folderPath, $"AMJOA_{ContractId}_p{i + 1}.jpg");
                        image.Save(jpegPath, System.Drawing.Imaging.ImageFormat.Jpeg);
                    }
                }
            }
            // 5. Delete the PDF file after conversion
            System.IO.File.Delete(pdfPath);

            // 6. Create password-protected zip file of JPEGs
            var folderPathZip = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "AMJOA", "AMJOA_" + ContractId);
            var zipPath = Path.Combine(folderPathZip, $"AMJOA_{ContractId}_JPEG_Preview.zip");
            if (System.IO.File.Exists(zipPath))
            {
                System.IO.File.Delete(zipPath);
            }
            var jpegFiles = Directory.GetFiles(folderPath, $"AMJOA_{ContractId}_p*.jpg");
            string password = _configuration["Password:PaswordPDF"];

            using (var fsOut = System.IO.File.Create(zipPath))
            using (var zipStream = new ICSharpCode.SharpZipLib.Zip.ZipOutputStream(fsOut))
            {
                zipStream.SetLevel(9); // 0-9, 9 = best compression
                zipStream.Password = password; // Set password

                foreach (var file in jpegFiles)
                {
                    var entry = new ICSharpCode.SharpZipLib.Zip.ZipEntry(Path.GetFileName(file))
                    {
                        DateTime = DateTime.Now
                    };
                    zipStream.PutNextEntry(entry);

                    byte[] buffer = System.IO.File.ReadAllBytes(file);
                    zipStream.Write(buffer, 0, buffer.Length);
                    zipStream.CloseEntry();
                }
                zipStream.IsStreamOwner = true;
                zipStream.Close();
            }
            // 7. Delete JPEG files after zipping
            foreach (var file in jpegFiles)
            {
                System.IO.File.Delete(file);
            }

            // 8. Return the password-protected zip file as download
            var zipBytes = await System.IO.File.ReadAllBytesAsync(zipPath);
            return File(zipBytes, "application/zip", $"AMJOA_{ContractId}_JPEG_Preview.zip");
        }
        public async Task<IActionResult> OnGetWordContact_AMJOA_Word(string ContractId = "12")
        {
            // 1. Get HTML content from the service
            var htmlContent = await _AMJOAService.OnGetWordContact_AMJOAServiceHtmlToPDF(ContractId);

            // 2. Create a Word document from HTML using Spire.Doc
            var document = new Spire.Doc.Document();
            document.LoadFromStream(new MemoryStream(System.Text.Encoding.UTF8.GetBytes(htmlContent)), Spire.Doc.FileFormat.Html);

            // 3. Save the Word document to a MemoryStream
            using var ms = new MemoryStream();
            document.SaveToStream(ms, Spire.Doc.FileFormat.Docx);
            var wordBytes = ms.ToArray();

            // 4. Prepare folder structure
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "AMJOA", "AMJOA_" + ContractId);
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var filePath = Path.Combine(folderPath, $"AMJOA_{ContractId}.docx");

            // 5. Save Word file to disk
            if (System.IO.File.Exists(filePath))
            {
                System.IO.File.Delete(filePath);
            }
            await System.IO.File.WriteAllBytesAsync(filePath, wordBytes);

            // 6. Return the Word file as download
            return File(wordBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", $"AMJOA_{ContractId}.docx");
        }
        public async Task<IActionResult> OnGetWordContact_AMJOA_Word_Preview(string ContractId = "70")
        {
            // 1. Get HTML content from the service
            var htmlContent = await _AMJOAService.OnGetWordContact_AMJOAServiceHtmlToPDF(ContractId);

            // 2. Create a Word document from HTML using Spire.Doc
            var document = new Spire.Doc.Document();
            document.LoadFromStream(new MemoryStream(System.Text.Encoding.UTF8.GetBytes(htmlContent)), Spire.Doc.FileFormat.Html);

            // 3. Save the Word document to a MemoryStream
            using var ms = new MemoryStream();
            document.SaveToStream(ms, Spire.Doc.FileFormat.Docx);
            var wordBytes = ms.ToArray();

            // 4. Prepare folder structure
            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", "AMJOA", "AMJOA_" + ContractId);
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            var filePath = Path.Combine(folderPath, $"AMJOA_{ContractId}_Preview.docx");

            // 5. Delete the file if it already exists
            if (System.IO.File.Exists(filePath))
            {
                System.IO.File.Delete(filePath);
            }

            // 6. Get password from appsettings.json
            string? userPassword = _configuration["Password:PaswordPDF"];

            // 7. Load the Word file from memory and apply password protection
            using (var msProtect = new MemoryStream(wordBytes))
            {
                Document doc = new Document();
                doc.LoadFromStream(msProtect, FileFormat.Docx);

                // Apply password protection
                doc.Encrypt(userPassword);

                // Save the password-protected file
                doc.SaveToFile(filePath, FileFormat.Docx);
            }

            // 8. Return the password-protected Word file as download
            var resultBytes = await System.IO.File.ReadAllBytesAsync(filePath);
            return File(resultBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", $"AMJOA_{ContractId}_Preview.docx");
        }
        #endregion  4.1.6 เอกสารแนบท้ายบันทึกข้อตกลงความร่วมมือและสัญญาร่วมดำเนินการ AMJOA

        // Helper: Generate a simple hash and salt for Word protection (not secure, demo only)
        private static void GenerateWordPasswordHash(string password, out string hash, out string salt)
        {
            // This is a simple, non-secure example. For real security, use a proper OpenXML password hash implementation.
            // Here, we use a fixed salt and a SHA1 hash for demonstration.
            var saltBytes = new byte[] { 0x00, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08, 0x09, 0x0A, 0x0B, 0x0C, 0x0D, 0x0E, 0x0F };
            using (var sha1 = System.Security.Cryptography.SHA1.Create())
            {
                var passwordBytes = System.Text.Encoding.Unicode.GetBytes(password);
                var combined = new byte[saltBytes.Length + passwordBytes.Length];
                Buffer.BlockCopy(saltBytes, 0, combined, 0, saltBytes.Length);
                Buffer.BlockCopy(passwordBytes, 0, combined, saltBytes.Length, passwordBytes.Length);
                var hashBytes = sha1.ComputeHash(combined);
                hash = Convert.ToBase64String(hashBytes);
                salt = Convert.ToBase64String(saltBytes);
            }
        }

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
            var pdfBytes = await _SupportSMEsService.OnGetWordContact_SupportSMEsService_HtmlToPDF(ContractId, "68");
            return File(pdfBytes, "application/pdf", "บันทึกข้อตกลงการแบ่งปันข้อมูลส่วนบุคคล.pdf");
        }

        #endregion Word to PDF using Interop
    }
}
