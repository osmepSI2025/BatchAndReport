using BatchAndReport.DAO;
using BatchAndReport.Models;
using BatchAndReport.Repository;
using BatchAndReport.Services;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using PuppeteerSharp;
using System.Diagnostics.Contracts;
using System.Text;

using System.Text.Json;

namespace BatchAndReport.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class DgaEsignController : ControllerBase
    {
        private readonly EContractDAO _eContractDao;
        private readonly IApiInformationRepository _repositoryApi;
        private readonly ICallAPIService _serviceApi;
        private readonly IWordEContractService _serviceWord;
        private readonly DgaSignDAO _dgaSignDao;
        private readonly WordEContract_JointOperationService _JointOperationService;
        private readonly WordEContract_MemorandumService _MemorandumService;
        private readonly WordEContract_PersernalProcessService _PersernalProcessService;
        private readonly WordEContract_DataPersonalService _DataPersonalService;
        private readonly WordEContract_ControlDataService _ControlDataService;
        private readonly WordEContract_DataSecretService _DataSecretService;
        private readonly WordEContract_SupportSMEsService _SupportSMEsService;
        private readonly WordEContract_AMJOAService _AMJOAService;
        private readonly WordEContract_HireEmployee _HireEmployee;
        private readonly WordEContract_MIWService _MIWService;
        private readonly WordEContract_MemorandumInWritingService _MemorandumInWritingService;
        private readonly ILogger<DgaEsignController> _logger;
        private readonly IWebHostEnvironment _env; // added
        public DgaEsignController(
            EContractDAO eContractDao,
            IApiInformationRepository repositoryApi,
            ICallAPIService serviceApi,
            IWordEContractService serviceWord,
            DgaSignDAO dgaSignDao,
             WordEContract_JointOperationService jointOperationService,
             WordEContract_MemorandumService memorandumService,
                WordEContract_PersernalProcessService persernalProcessService,
                WordEContract_DataPersonalService dataPersonalService,
                WordEContract_ControlDataService controlDataService,
                WordEContract_DataSecretService dataSecretService,
                 WordEContract_SupportSMEsService supportSMEsService,
              WordEContract_AMJOAService aMJOAService,
                WordEContract_HireEmployee hireEmployee,
                WordEContract_MIWService mIWService
            ,
                WordEContract_MemorandumInWritingService memorandumInWritingService,
                ILogger<DgaEsignController> logger,
                  IWebHostEnvironment env // added parameter
            )
        {
            _eContractDao = eContractDao;
            _repositoryApi = repositoryApi;
            _serviceApi = serviceApi;
            _serviceWord = serviceWord;
            _dgaSignDao = dgaSignDao;
            _JointOperationService = jointOperationService;
            _MemorandumService = memorandumService;
            _PersernalProcessService = persernalProcessService;
            _DataPersonalService = dataPersonalService;
            _ControlDataService = controlDataService;
            _DataSecretService = dataSecretService;
            _SupportSMEsService = supportSMEsService;
            _AMJOAService = aMJOAService;
            _HireEmployee = hireEmployee;
            _MIWService = mIWService;
            _MemorandumInWritingService = memorandumInWritingService;
            _logger = logger;
            _env = env; // assign
        }


        [HttpGet("GetDgaCert")]
        public async Task<IActionResult> GetDgaCert(string ContractType = "JOA", string ContractId = "95" ,string EmailSign = "si_noreply@sme.go.th")
        {
            _logger.LogInformation("Start GetDgaCert - ContractType={ContractType}, ContractId={ContractId}, EmailSign={EmailSign}", ContractType, ContractId, EmailSign);
            try
            {
                #region Get Master Data
                var apiInfo = await _dgaSignDao.GetDgaEsignUrlAsync();
                var api = apiInfo.Find(x => x.ServiceCode == "Token");
                if (api == null)
                {
                    _logger.LogWarning("GetDgaCert - API information not found for GetToken");
                    return NotFound(new
                    {
                        message = "API information not found for GetToken"
                    });
                }
                var dgaConfig = await _dgaSignDao.GetDgaEsignConfigAsync();
                if (dgaConfig == null || dgaConfig.Count == 0)
                {
                    _logger.LogWarning("GetDgaCert - DGA configuration not found");
                    return NotFound(new
                    {
                        message = "DGA configuration not found"
                    });
                }

                var dgaTemplate = await _dgaSignDao.GetDgaEsignTemplateAsync();
                if (dgaTemplate == null)
                {
                    _logger.LogWarning("GetDgaCert - DGA template not found for ContractType={ContractType}", ContractType);
                    return NotFound(new
                    {
                        message = $"DGA template not found for ContractType: {ContractType}"
                    });
                }

                var selectedTemplate = dgaTemplate.Find(x => x.ContractType.ToUpper() == ContractType.ToUpper() && x.FlagActive == "Y");

                if (selectedTemplate == null)
                {
                    _logger.LogWarning("GetDgaCert - Active DGA template not found for ContractType={ContractType}", ContractType);
                    return NotFound(new
                    {
                        message = $"Active DGA template not found for ContractType: {ContractType}"
                    });
                }

                string ConsumerKey = dgaConfig.FirstOrDefault().ConsumerKey.Trim();
                string ConsumerSecret = dgaConfig.FirstOrDefault().ConsumerSecret.Trim();
                string Email = dgaConfig.FirstOrDefault().Email.Trim();

                #endregion Master Data


                #region Get PDF
                DgaRegisterDocModels dgaDoc = new DgaRegisterDocModels();

                // get pdf
                var htmlContent = await GetPdfByContractType(ContractType, ContractId);
                if(htmlContent == null || htmlContent.Length == 0)
                {
                    _logger.LogWarning("GetDgaCert - generated PDF is empty for ContractType={ContractType}, ContractId={ContractId}", ContractType, ContractId);
                    return NotFound(new
                    {
                        message = "Generated PDF is empty",
                        ContractType = ContractType,
                        ContractId = ContractId
                    });
                }


                #endregion
                #region GetToken
                var apiToken = apiInfo.Find(x => x.ServiceCode == "Token");
                if (apiToken == null)
                {
                    _logger.LogWarning("GetDgaCert - API information not found for Token");
                    return NotFound(new
                    {
                        message = "API information not found for Token"
                    });
                }
                // Call GetToken and extract the JSON string from the IActionResult
                var tokenResult = await GetToken(ConsumerKey, ConsumerSecret, EmailSign, apiToken.UrlDev.ToString()) as ObjectResult;
                string tokenJson = tokenResult?.Value?.GetType().GetProperty("apiResponse")?.GetValue(tokenResult.Value)?.ToString();

                // Deserialize and extract the token
                using var doc = JsonDocument.Parse(tokenJson);
                string token = doc.RootElement.GetProperty("Result").GetString();

                #endregion



                #region send Register PDF to DGA

                var apiRegis = apiInfo.Find(x => x.ServiceCode == "RegisterDoc");
                if (apiRegis == null)
                {
                    _logger.LogWarning("GetDgaCert - API information not found for RegisterDoc");
                    return NotFound(new
                    {
                        message = "API information not found for RegisterDoc"
                    });
                }

                DGADocumentModels docx = await RegisterPDF(ConsumerKey, token, apiRegis.UrlDev, selectedTemplate.TemplateID, htmlContent);
                #endregion  send PDF to DGA


                #region Check download pdf

                //8.API ลงลายมือชื่ออิเลกทรอนิกส์แบบองค์กร
                var apiCert = apiInfo.Find(x => x.ServiceCode == "CertifiedSign");
                if (apiCert == null)
                {
                    _logger.LogWarning("GetDgaCert - API information not found for CertifiedSign");
                    return NotFound(new
                    {
                        message = "API information not found for CertifiedSign"
                    });
                }
                string SignatureID = "";
                var signResult = docx != null && !string.IsNullOrEmpty(docx.DocumentID)
                    ? await GetCertifiedSign(token, docx.DocumentID, ConsumerKey, apiCert.UrlDev.ToString())
                    : null;

                if (signResult is ObjectResult objectResult)
                {
                    // Get the apiResponse property from the result
                    var apiResponse = objectResult.Value?.GetType().GetProperty("apiResponse")?.GetValue(objectResult.Value)?.ToString();
                    if (!string.IsNullOrEmpty(apiResponse))
                    {
                        using var docSign = JsonDocument.Parse(apiResponse);
                        if (docSign.RootElement.TryGetProperty("SignatureID", out var sigProp))
                        {
                            SignatureID = sigProp.GetString() ?? "";
                        }
                    }
                }

                var apiDownloadSignedPdf = apiInfo.Find(x => x.ServiceCode == "DownloadSignedPdf");
                if (apiDownloadSignedPdf == null)
                {
                    _logger.LogWarning("GetDgaCert - API information not found for DownloadSignedPdf");
                    return NotFound(new
                    {
                        message = "API information not found for RegisterDoc"
                    });
                }
                #region Sava Transaction
                string savePath = $"/Document/{ContractType.ToUpper()}/DGA/{ContractType.ToUpper()}_{ContractId}.pdf";
                DgaEsignDocumentModels dgaResult = new DgaEsignDocumentModels();
                dgaResult.WFTypeCode = ContractType;
                 dgaResult.ContractID = int.Parse(ContractId);
                dgaResult.DGA_DocumentID = docx.DocumentID;
                dgaResult.DGA_TemplateID = selectedTemplate.TemplateID;
                dgaResult.DGA_SignatureID = SignatureID; //ยังไม่มี
                dgaResult.DGA_DocumentDataFile = htmlContent;
                dgaResult.DGA_DocumentPathFile = savePath;
                dgaResult.SignBy = Email;
                dgaResult.CreateBy = Email;
                dgaResult.CreateDate = DateTime.Now;


                var saveTrans = await _dgaSignDao.InsertDgaEsignDocumentAsync(dgaResult);

                #endregion Sava Transaction
                var savefile = await DownloadSignedPdf(docx.DocumentID, ConsumerKey, token, apiDownloadSignedPdf.UrlDev, ContractType, ContractId);
                _logger.LogInformation("End GetDgaCert - success - DocumentID={DocumentID}", docx?.DocumentID);
                return Ok();

                #endregion Check download pdf



            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "GetDgaCert - error for ContractType={ContractType}, ContractId={ContractId}", ContractType, ContractId);
                return StatusCode(500, new
                {
                    message = "Internal Server Error",
                    error = ex.Message,
                    inner = ex.InnerException?.Message,
                    stack = ex.StackTrace
                });
            }
        }

        // 3 API ขอ Token

        [HttpGet("GetToken")]
        public async Task<IActionResult> GetToken(string ConsumerKey, string ConsumerSecret, string Email,string UrlToken)
        {
            _logger.LogInformation("Start GetToken - ConsumerKey={ConsumerKey}, Email={Email}", ConsumerKey, Email);
            try
            {
                using var httpClient = new HttpClient();
                // Set required headers
                httpClient.DefaultRequestHeaders.Add("Consumer-Key", ConsumerKey);
                httpClient.DefaultRequestHeaders.Add("Consumer-Secret", ConsumerSecret);

             //   var url = "https://trial.dga.or.th/ws/auth/validate?ConsumerSecret=" + ConsumerSecret + "&AgentID=" + Email + "";
                var url = UrlToken.Replace("[Secret]", ConsumerSecret).Replace("[IdCard]", Email);
                var response = await httpClient.GetAsync(url);
                var responseBody = await response.Content.ReadAsStringAsync();
                _logger.LogInformation("End GetToken - status={StatusCode}", response.StatusCode);
                return StatusCode((int)response.StatusCode, new
                {
                    message = "success",
                    apiResponse = responseBody
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "GetToken - error");
                return StatusCode(500, new
                {
                    message = "Internal Server Error",
                    error = ex.Message,
                    inner = ex.InnerException?.Message,
                    stack = ex.StackTrace
                });
            }
        }

        [HttpPost("RegisterPDF")]

        public async Task<DGADocumentModels> RegisterPDF(string ConsumerKey, string token, string urlDga, string templateID, byte[] pdfBytes = null)
        {
            _logger.LogInformation("Start RegisterPDF - urlDga={UrlDga}, templateID={TemplateID}, pdfBytesLength={Len}", urlDga, templateID, pdfBytes?.Length ?? 0);
            DGADocumentModels docx = new DGADocumentModels();
            try
            {

                using var httpClient = new HttpClient();

                // Set required headers
                httpClient.DefaultRequestHeaders.Add("Consumer-Key", ConsumerKey);
                httpClient.DefaultRequestHeaders.Add("Token", token);


                // Prepare request URL with parameters
                 
                string url = urlDga.Replace("[TemplateID]", templateID) + "&Timestamp=true";

                using var form = new MultipartFormDataContent();

                //4 API เพื่อลงทะเบียนเอกสาร (DocumentID) จากการ Upload PDF
                // Add PDF file
                var pdfContent = new ByteArrayContent(pdfBytes);
                pdfContent.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("application/pdf");
                form.Add(pdfContent, "Content", "JOA_95.pdf");


                // Add other fields
                form.Add(new StringContent("สัญญาฉบับนี้"), "Clause");
                form.Add(new StringContent("https://econtract.dga.or.th/xxxxx"), "Link");
                form.Add(new StringContent(""), "Page");
                form.Add(new StringContent("50"), "Left");
                form.Add(new StringContent("20"), "Bottom");

                // Send PUT request as required by DGA API for document registration
                var response = await httpClient.PutAsync(url, form);
                var responseBody = await response.Content.ReadAsStringAsync();
                docx = JsonSerializer.Deserialize<DGADocumentModels>(responseBody);
                _logger.LogInformation("End RegisterPDF - DocumentID={DocumentID}", docx?.DocumentID);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "RegisterPDF - error");
                return null;
            }

            return docx;
        }

        //7 API ลงลายมือชื่ออิเลกทรอนิกส์แบบบุคคล ด้วยรูปภาพลายเซ็น
        //[HttpGet("GetDocumentId")]
        //public async Task<IActionResult> GetDocumentId(string token, string docId, string ConsumerKey,string UrlDoc)
        //{
        //    _logger.LogInformation("Start GetDocumentId - docId={DocId}", docId);
        //    try
        //    {
        //        // Prepare the payload
        //        var payload = new
        //        {
        //            DocumentID = docId,
        //            Reason = "ทดสอบเหตุผล JOA",
        //            Signature = new
        //            {
        //                //   Page = "",
        //                Left = "150",
        //                Bottom = "150",
        //                Width = "150",
        //                Height = "60",
        //                Image = "" // Replace with actual Base64 string of the signature image
        //            }
        //        };

        //        var jsonPayload = JsonSerializer.Serialize(payload);

        //        using var httpClient = new HttpClient();

        //        // Set required headers
        //        httpClient.DefaultRequestHeaders.Add("Consumer-Key", ConsumerKey);
        //        httpClient.DefaultRequestHeaders.Add("Token", token);

        //        var content = new StringContent(jsonPayload, Encoding.UTF8, "application/json");
        //        var url = "https://trial.dga.or.th/api/edoc/signature/egov/v1/image/signed";

        //        var response = await httpClient.PostAsync(url, content);
        //        var responseBody = await response.Content.ReadAsStringAsync();

        //        _logger.LogInformation("End GetDocumentId - status={StatusCode}", response.StatusCode);
        //        return StatusCode((int)response.StatusCode, new
        //        {
        //            message = "success",
        //            apiResponse = responseBody
        //        });
        //    }
        //    catch (Exception ex)
        //    {
        //        _logger.LogError(ex, "GetDocumentId - error for docId={DocId}", docId);
        //        return StatusCode(500, new
        //        {
        //            message = "Internal Server Error",
        //            error = ex.Message,
        //            inner = ex.InnerException?.Message,
        //            stack = ex.StackTrace
        //        });
        //    }
        //}

        // 8.API ลงลายมือชื่ออิเลกทรอนิกส์แบบองค์กร
        [HttpGet("GetCertifiedSign")]
        public async Task<IActionResult> GetCertifiedSign(string token, string docId, string ConsumerKey,string UrlCert)
        {
            _logger.LogInformation("Start GetCertifiedSign - docId={DocId}", docId);
            // convert image to base64
            string base64Image = "";
            try
            {
                var webRoot = _env.WebRootPath ?? Path.Combine(Directory.GetCurrentDirectory(), "wwwroot");
                var imagePath = Path.Combine(webRoot, "images", "no-signature.png");
                if (System.IO.File.Exists(imagePath))
                {
                    var bytes = await System.IO.File.ReadAllBytesAsync(imagePath);
                    base64Image = Convert.ToBase64String(bytes);
                }
                else
                {
                    _logger.LogWarning("GetCertifiedSign - no-signature.png not found at {ImagePath}", imagePath);
                }
            }
            catch (Exception e)
            {
                _logger.LogWarning(e, "GetCertifiedSign - failed to convert image to base64, using empty string fallback");
            }

            try
            {
                // Prepare the payload
                var payload = new
                {
                    CertificateID = "",
                    DocumentID = docId,
                    Reason = "ทดสอบเหตุผล",
                    Agent = "สมใจ นายทดสอบ",
                    Position = "ชื่อตําแหน่ง",
                    Signature = new
                    {
                           Page = "",
                        Left = "100",
                        Bottom = "20",
                        Width = "100",
                        Height = "50",
                        Image = base64Image // Replace with actual Base64 string of the signature image
                        //Image = "" // Replace with actual Base64 string of the signature image
                    },
                    Content = new[]
        {
        new
        {
            Type = "TEXT",
            Value = "สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม สสว.",
            Size = 12,
            Page = "",
            Left = 200,
            Bottom = 20,
            Width = 200
        }
    }
                };

                var jsonPayload = JsonSerializer.Serialize(payload);

                using var httpClient = new HttpClient();

                // Set required headers
                httpClient.DefaultRequestHeaders.Add("Consumer-Key", ConsumerKey);
                httpClient.DefaultRequestHeaders.Add("Token", token);

                var content = new StringContent(jsonPayload, Encoding.UTF8, "application/json");

                var url = UrlCert;
                var response = await httpClient.PostAsync(url, content);
                var responseBody = await response.Content.ReadAsStringAsync();

                _logger.LogInformation("End GetCertifiedSign - status={StatusCode}", response.StatusCode);
                return StatusCode((int)response.StatusCode, new
                {
                    message = "success",
                    apiResponse = responseBody
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "GetCertifiedSign - error for docId={DocId}", docId);
                return StatusCode(500, new
                {
                    message = "Internal Server Error",
                    error = ex.Message,
                    inner = ex.InnerException?.Message,
                    stack = ex.StackTrace
                });
            }
        }


        //9 API Download PDF Signed 
        [HttpGet("DownloadSignedPdf")]
        public async Task<IActionResult> DownloadSignedPdf([FromQuery] string documentId, string ConsumerKey, string Token,string apiurl,string contype,string conId)
        {
            _logger.LogInformation("Start DownloadSignedPdf - DocumentId={DocumentId}, Contype={Contype}, ConId={ConId}", documentId, contype, conId);
            try
            {
                using var httpClient = new HttpClient();

                // Set required headers
                httpClient.DefaultRequestHeaders.Add("Consumer-Key", ConsumerKey);
                httpClient.DefaultRequestHeaders.Add("Token", Token);

                // Prepare request URL
               
                string url = $"{apiurl}".Replace("[DocumentID]", documentId);
                

                var response = await httpClient.GetAsync(url);

                if (!response.IsSuccessStatusCode)
                {
                    var errorBody = await response.Content.ReadAsStringAsync();
                    _logger.LogWarning("DownloadSignedPdf - failed status={StatusCode} for DocumentId={DocumentId}", response.StatusCode, documentId);
                    return StatusCode((int)response.StatusCode, new
                    {
                        message = "Failed to download PDF",
                        apiResponse = errorBody
                    });
                }

                var pdfBytes = await response.Content.ReadAsByteArrayAsync();

                // update pdfBytes to DGA_DocumentDataFile
                var updateResult = await _dgaSignDao.UpdateDgaEsignDocumentAsync(documentId, pdfBytes);
           
                string savePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", contype, "DGA", $"{contype}_{conId}.pdf");
                Directory.CreateDirectory(Path.GetDirectoryName(savePath)!);
                await System.IO.File.WriteAllBytesAsync(savePath, pdfBytes);

                _logger.LogInformation("End DownloadSignedPdf - saved to {SavePath}", savePath);
                // Return PDF file
                return File(pdfBytes, "application/pdf", $"Signed_{documentId}.pdf");

            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "DownloadSignedPdf - error for DocumentId={DocumentId}", documentId);
                return StatusCode(500, new
                {
                    message = "Internal Server Error",
                    error = ex.Message,
                    inner = ex.InnerException?.Message,
                    stack = ex.StackTrace
                });
            }
        }

        //get pdf by Contract Type
        [HttpGet("GetPdfByContractType")]
        public async Task<byte[]> GetPdfByContractType(string contractType, string contractId)
        {
            _logger.LogInformation("Start GetPdfByContractType - contractType={ContractType}, contractId={ContractId}", contractType, contractId);
            try
            {
                var htmlContent = "";
                switch (contractType.ToUpper())
                {
                    case "JOA":
                        htmlContent = await _JointOperationService.OnGetWordContact_JointOperationServiceHtmlToPDF(contractId);
                        break;
                    case "MOU": //4.1.1.2.3.บันทึกข้อตกลงความร่วมมือ MOU
                        htmlContent = await _MemorandumService.OnGetWordContact_MemorandumService_HtmlToPDF(contractId, "MOU");
                        break;
                    case "PDPA": //4.1.1.2.4.บันทึกข้อตกลงการประมวลผลข้อมูลส่วนบุคคล PDPA
                        htmlContent = await _PersernalProcessService.OnGetWordContact_PersernalProcessService_HtmlToPDF(contractId, "PDPA");
                        break;
                    case "PDSA": //4.1.1.2.6.บันทึกข้อตกลงการแบ่งปันข้อมูลส่วนบุคคล PDSA
                        htmlContent = await _DataPersonalService.OnGetWordContact_DataPersonalService_ToPDF(contractId, "PDSA");
                        break;
                    case "JDCA": // 4.1.1.2.5.บันทึกข้อตกลงการเป็นผู้ควบคุมข้อมูลส่วนบุคคลร่วมตัวอย่างหน้าจอ JDCA
                        htmlContent = await _ControlDataService.OnGetWordContact_ControlDataServiceHtmlToPdf(contractId, "JDCA");
                        break;
                    case "NDA": //4.1.1.2.7.สัญญาการรักษาข้อมูลที่เป็นความลับ NDA
                        htmlContent = await _DataSecretService.OnGetWordContact_DataSecretService_ToPDF(contractId, "NDA");
                        break;
                    case "GA": //4.1.1.2.2.สัญญารับเงินอุดหนุน GA BDS
                        htmlContent = await _SupportSMEsService.OnGetWordContact_SupportSMEsService_HtmlToPDF(contractId, "GA");
                        break;
                    case "AMJOA":
                        htmlContent = await _AMJOAService.OnGetWordContact_AMJOAServiceHtmlToPDF(contractId);
                        break;
                    case "MIW":
                        htmlContent = await _MIWService.OnGetWordContact_MIWServiceHtmlToPDF(contractId);
                        break;
                    case "MOA":
                        htmlContent = await _MemorandumInWritingService.OnGetWordContact_MemorandumInWritingService_HtmlToPDF(contractId, "MOA");
                        break;
                    case "EC":
                        htmlContent = await _HireEmployee.OnGetWordContact_HireEmployee_ToPDF(contractId, "EC");
                        break;
                    default:
                        throw new ArgumentException("Unsupported contract type");
                }



                if (string.IsNullOrWhiteSpace(htmlContent))
                {
                    _logger.LogWarning("GetPdfByContractType - generated HTML is empty for ContractType={ContractType}, ContractId={ContractId}", contractType, contractId);
                    return null;
                }

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
                _logger.LogInformation("End GetPdfByContractType - produced PDF length={Length}", pdfBytes?.Length ?? 0);
                return pdfBytes;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "GetPdfByContractType - error for ContractType={ContractType}, ContractId={ContractId}", contractType, contractId);
                // Log the exception as needed, e.g. using ILogger
                // _logger.LogError(ex, "Failed to generate PDF for contract {ContractId}", contractId);
                return null;
            }
        }


        [HttpGet("GetFilePDFCert")]
        public async Task<IActionResult> GetFilePDFCert([FromQuery] string ContractType = "JOA", string ContractId = "8")
        {
            _logger.LogInformation("Start GetFilePDFCert - ContractType={ContractType}, ContractId={ContractId}", ContractType, ContractId);
            try
            {
                // get DGA record(s) for contract
                var records = await _dgaSignDao.GetDgaEsignAsync(ContractType, int.Parse(ContractId));

                if (records == null || records.Count == 0)
                {
                    _logger.LogWarning("GetFilePDFCert - No DGA document records found for ContractType={ContractType}, ContractId={ContractId}", ContractType, ContractId);
                    return NotFound(new { message = "No DGA document records found" });
                }

                // select the most recent record by ID (safe without LINQ)
                var selected = (DgaEsignModels?)null;
                var maxId = int.MinValue;
                foreach (var r in records)
                {
                    if (r != null && r.ID > maxId)
                    {
                        maxId = r.ID;
                        selected = r;
                    }
                }

                if (selected == null)
                {
                    _logger.LogWarning("GetFilePDFCert - No DGA document record selected for ContractType={ContractType}, ContractId={ContractId}", ContractType, ContractId);
                    return NotFound(new { message = "No DGA document record selected" });
                }

                var pdfBytes = selected.DGA_DocumentDataFile;
                if (pdfBytes == null || pdfBytes.Length == 0)
                {
                    _logger.LogWarning("GetFilePDFCert - PDF binary not found for DocumentID={DocumentID}", selected.DGA_DocumentID);
                    return NotFound(new
                    {
                        message = "PDF binary not found in DGA_DocumentDataFile",
                        DocumentID = selected.DGA_DocumentID,
                        Path = selected.DGA_DocumentPathFile
                    });
                }

                // optional: persist to disk under wwwroot/Document/{ContractType}/DGA/{ContractType}_{ContractId}.pdf
                try
                {
                    var fileName = $"{ContractType.ToUpper()}_{ContractId}.pdf";
                    var saveDir = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", ContractType.ToUpper(), "DGA");
                    Directory.CreateDirectory(saveDir);
                    var savePath = Path.Combine(saveDir, fileName);
                    await System.IO.File.WriteAllBytesAsync(savePath, pdfBytes);
                    _logger.LogInformation("GetFilePDFCert - persisted PDF to {SavePath}", savePath);
                }
                catch
                {
                    // ignore disk write errors — still return the file to the caller
                    _logger.LogWarning("GetFilePDFCert - failed to persist PDF to disk (ignored)");
                }

                // return PDF as file download
                var downloadFileName = $"{selected.DGA_DocumentID ?? $"{ContractType}_{ContractId}"}.pdf";
                _logger.LogInformation("End GetFilePDFCert - returning file {FileName}", downloadFileName);
                return File(pdfBytes, "application/pdf", downloadFileName);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "GetFilePDFCert - error for ContractType={ContractType}, ContractId={ContractId}", ContractType, ContractId);
                return StatusCode(500, new
                {
                    message = "Internal Server Error",
                    error = ex.Message,
                    inner = ex.InnerException?.Message,
                    stack = ex.StackTrace
                });
            }
        }

    }
}