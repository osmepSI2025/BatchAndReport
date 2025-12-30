using BatchAndReport.DAO;
using BatchAndReport.Models;
using BatchAndReport.Repository;
using BatchAndReport.Services;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
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
        private readonly IHttpClientFactory _httpClientFactory;
        private readonly IBrowserPdfService _browserPdfService;
        private readonly ILogger<DgaEsignController> _logger;

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
            WordEContract_MIWService mIWService,
            WordEContract_MemorandumInWritingService memorandumInWritingService,
            IHttpClientFactory httpClientFactory,
            IBrowserPdfService browserPdfService,
            ILogger<DgaEsignController> logger)
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
            _httpClientFactory = httpClientFactory;
            _browserPdfService = browserPdfService;
            _logger = logger;
        }

        [HttpGet("GetDgaCert")]
        public async Task<IActionResult> GetDgaCert(string ContractType = "JOA", string ContractId = "8", string EmailSign = "si_noreply@sme.go.th")
        {
            try
            {
                #region Get Master Data
                var apiInfo = await _dgaSignDao.GetDgaEsignUrlAsync();
                var api = apiInfo?.Find(x => x.ServiceCode == "Token");
                if (api == null)
                {
                    _logger.LogWarning("GetDgaCert: Token API info not found");
                    return NotFound(new { message = "API information not found for GetToken" });
                }

                var dgaConfig = await _dgaSignDao.GetDgaEsignConfigAsync();
                if (dgaConfig == null || dgaConfig.Count == 0)
                {
                    _logger.LogWarning("GetDgaCert: DGA configuration missing");
                    return NotFound(new { message = "DGA configuration not found" });
                }

                var dgaTemplate = await _dgaSignDao.GetDgaEsignTemplateAsync();
                if (dgaTemplate == null)
                    return NotFound(new { message = $"DGA template not found for ContractType: {ContractType}" });

                var selectedTemplate = dgaTemplate.Find(x => string.Equals(x.ContractType, ContractType, System.StringComparison.OrdinalIgnoreCase) && x.FlagActive == "Y");
                if (selectedTemplate == null)
                    return NotFound(new { message = $"Active DGA template not found for ContractType: {ContractType}" });

                var cfg = dgaConfig.First();
                string ConsumerKey = cfg.ConsumerKey?.Trim() ?? string.Empty;
                string ConsumerSecret = cfg.ConsumerSecret?.Trim() ?? string.Empty;
                string Email = cfg.Email?.Trim() ?? string.Empty;
                #endregion

                #region Get PDF
                // get pdf bytes via HTML -> PDF
                var pdfBytes = await GetPdfByContractType(ContractType, ContractId);
                if (pdfBytes == null || pdfBytes.Length == 0)
                {
                    _logger.LogWarning("GetDgaCert: PDF generation returned empty for ContractType={ContractType}, ContractId={ContractId}", ContractType, ContractId);
                    return StatusCode(500, new { message = "Failed to generate PDF content" });
                }
                #endregion

                #region GetToken
                var tokenResult = await GetToken(ConsumerKey, ConsumerSecret, EmailSign) as ObjectResult;
                var tokenJson = tokenResult?.Value?.GetType().GetProperty("apiResponse")?.GetValue(tokenResult.Value)?.ToString();
                if (string.IsNullOrWhiteSpace(tokenJson))
                {
                    _logger.LogWarning("GetDgaCert: Token API returned empty response");
                    return StatusCode(500, new { message = "Failed to obtain token from DGA" });
                }

                using var doc = JsonDocument.Parse(tokenJson);
                if (!doc.RootElement.TryGetProperty("Result", out var tokenElement))
                {
                    _logger.LogWarning("GetDgaCert: Token JSON missing Result property. Raw: {raw}", tokenJson);
                    return StatusCode(500, new { message = "Invalid token payload from DGA" });
                }
                string token = tokenElement.GetString() ?? string.Empty;
                #endregion

                #region send Register PDF to DGA
                var apiRegis = apiInfo.Find(x => x.ServiceCode == "RegisterDoc");
                if (apiRegis == null) return NotFound(new { message = "API information not found for RegisterDoc" });

                var docx = await RegisterPDF(ConsumerKey, token, apiRegis.UrlDev, selectedTemplate.TemplateID, pdfBytes);
                if (docx == null || string.IsNullOrEmpty(docx.DocumentID))
                {
                    _logger.LogWarning("GetDgaCert: RegisterPDF failed or returned empty DocumentID");
                    return StatusCode(500, new { message = "Failed to register document with DGA" });
                }
                #endregion

                #region Check download pdf / sign
                string SignatureID = "";
                var signResult = await GetCertifiedSign(token, docx.DocumentID, ConsumerKey);
                if (signResult is ObjectResult certObj)
                {
                    var apiResponse = certObj.Value?.GetType()?.GetProperty("apiResponse")?.GetValue(certObj.Value)?.ToString();
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
                if (apiDownloadSignedPdf == null) return NotFound(new { message = "API information not found for DownloadSignedPdf" });

                var savePath = $"/Document/{ContractType.ToUpper()}/DGA/{ContractType.ToUpper()}_{ContractId}.pdf";
                DgaEsignDocumentModels dgaResult = new()
                {
                    WFTypeCode = ContractType,
                    ContractID = int.Parse(ContractId),
                    DGA_DocumentID = docx.DocumentID,
                    DGA_TemplateID = selectedTemplate.TemplateID,
                    DGA_SignatureID = SignatureID,
                    DGA_DocumentDataFile = pdfBytes,
                    DGA_DocumentPathFile = savePath,
                    SignBy = Email,
                    CreateBy = Email,
                    CreateDate = DateTime.Now
                };

                await _dgaSignDao.InsertDgaEsignDocumentAsync(dgaResult);

                // download signed PDF and save it
                await DownloadSignedPdf(docx.DocumentID, ConsumerKey, token, apiDownloadSignedPdf.UrlDev, ContractType, ContractId);
                #endregion

                return Ok(new { message = "Document registered and processing started", documentId = docx.DocumentID });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "GetDgaCert failed for ContractType={ContractType} ContractId={ContractId}", ContractType, ContractId);
                return StatusCode(500, new
                {
                    message = "Internal Server Error",
                    error = ex.Message,
                    inner = ex.InnerException?.Message,
                    stack = ex.StackTrace
                });
            }
        }

        [HttpGet("GetToken")]
        public async Task<IActionResult> GetToken(string ConsumerKey, string ConsumerSecret, string Email)
        {
            try
            {
                using var httpClient = _httpClientFactory.CreateClient();
                httpClient.DefaultRequestHeaders.Remove("Consumer-Key");
                httpClient.DefaultRequestHeaders.Remove("Consumer-Secret");
                httpClient.DefaultRequestHeaders.Add("Consumer-Key", ConsumerKey);
                httpClient.DefaultRequestHeaders.Add("Consumer-Secret", ConsumerSecret);

                var url = "https://trial.dga.or.th/ws/auth/validate?ConsumerSecret=" + ConsumerSecret + "&AgentID=" + Email;
                var response = await httpClient.GetAsync(url);
                var responseBody = await response.Content.ReadAsStringAsync();
                return StatusCode((int)response.StatusCode, new
                {
                    message = "success",
                    apiResponse = responseBody
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "GetToken failed");
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
        public async Task<DGADocumentModels?> RegisterPDF(string ConsumerKey, string token, string urlDga, string templateID, byte[]? pdfBytes = null)
        {
            if (pdfBytes == null || pdfBytes.Length == 0)
            {
                _logger.LogWarning("RegisterPDF called with empty pdfBytes");
                return null;
            }

            try
            {
                using var httpClient = _httpClientFactory.CreateClient();
                httpClient.DefaultRequestHeaders.Remove("Consumer-Key");
                httpClient.DefaultRequestHeaders.Remove("Token");
                httpClient.DefaultRequestHeaders.Add("Consumer-Key", ConsumerKey);
                httpClient.DefaultRequestHeaders.Add("Token", token);

                string url = urlDga.Replace("[TemplateID]", templateID) + "&Timestamp=true";

                using var form = new MultipartFormDataContent();
                var pdfContent = new ByteArrayContent(pdfBytes);
                pdfContent.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("application/pdf");
                form.Add(pdfContent, "Content", "document.pdf");

                form.Add(new StringContent("สัญญาฉบับนี้"), "Clause");
                form.Add(new StringContent("https://econtract.dga.or.th/xxxxx"), "Link");
                form.Add(new StringContent(""), "Page");
                form.Add(new StringContent("50"), "Left");
                form.Add(new StringContent("20"), "Bottom");

                var response = await httpClient.PutAsync(url, form);
                var responseBody = await response.Content.ReadAsStringAsync();
                var docx = JsonSerializer.Deserialize<DGADocumentModels>(responseBody);
                return docx;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "RegisterPDF failed");
                return null;
            }
        }

        [HttpGet("GetDocumentId")]
        public async Task<IActionResult> GetDocumentId(string token, string docId, string ConsumerKey)
        {
            try
            {
                var payload = new
                {
                    DocumentID = docId,
                    Reason = "ทดสอบเหตุผล JOA",
                    Signature = new
                    {
                        Left = "150",
                        Bottom = "150",
                        Width = "150",
                        Height = "60",
                        Image = ""
                    }
                };

                var jsonPayload = JsonSerializer.Serialize(payload);
                using var httpClient = _httpClientFactory.CreateClient();
                httpClient.DefaultRequestHeaders.Remove("Consumer-Key");
                httpClient.DefaultRequestHeaders.Remove("Token");
                httpClient.DefaultRequestHeaders.Add("Consumer-Key", ConsumerKey);
                httpClient.DefaultRequestHeaders.Add("Token", token);

                var content = new StringContent(jsonPayload, Encoding.UTF8, "application/json");
                var url = "https://trial.dga.or.th/api/edoc/signature/egov/v1/image/signed";
                var response = await httpClient.PostAsync(url, content);
                var responseBody = await response.Content.ReadAsStringAsync();

                return StatusCode((int)response.StatusCode, new
                {
                    message = "success",
                    apiResponse = responseBody
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "GetDocumentId failed");
                return StatusCode(500, new
                {
                    message = "Internal Server Error",
                    error = ex.Message,
                    inner = ex.InnerException?.Message,
                    stack = ex.StackTrace
                });
            }
        }

        [HttpGet("GetCertifiedSign")]
        public async Task<IActionResult> GetCertifiedSign(string token, string docId, string ConsumerKey)
        {
            try
            {
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
                        Image = ""
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
                using var httpClient = _httpClientFactory.CreateClient();
                httpClient.DefaultRequestHeaders.Remove("Consumer-Key");
                httpClient.DefaultRequestHeaders.Remove("Token");
                httpClient.DefaultRequestHeaders.Add("Consumer-Key", ConsumerKey);
                httpClient.DefaultRequestHeaders.Add("Token", token);

                var content = new StringContent(jsonPayload, Encoding.UTF8, "application/json");
                var url = "https://trial.dga.or.th/api/edoc/signature/egov/v1/organization/certified";
                var response = await httpClient.PostAsync(url, content);
                var responseBody = await response.Content.ReadAsStringAsync();

                return StatusCode((int)response.StatusCode, new
                {
                    message = "success",
                    apiResponse = responseBody
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "GetCertifiedSign failed");
                return StatusCode(500, new
                {
                    message = "Internal Server Error",
                    error = ex.Message,
                    inner = ex.InnerException?.Message,
                    stack = ex.StackTrace
                });
            }
        }

        [HttpGet("DownloadSignedPdf")]
        public async Task<IActionResult> DownloadSignedPdf([FromQuery] string documentId, string ConsumerKey, string Token, string apiurl, string contype, string conId)
        {
            try
            {
                using var httpClient = _httpClientFactory.CreateClient();
                httpClient.DefaultRequestHeaders.Remove("Consumer-Key");
                httpClient.DefaultRequestHeaders.Remove("Token");
                httpClient.DefaultRequestHeaders.Add("Consumer-Key", ConsumerKey);
                httpClient.DefaultRequestHeaders.Add("Token", Token);

                string url = $"{apiurl}".Replace("[DocumentID]", documentId);
                var response = await httpClient.GetAsync(url);

                if (!response.IsSuccessStatusCode)
                {
                    var errorBody = await response.Content.ReadAsStringAsync();
                    _logger.LogWarning("DownloadSignedPdf failed with status {Status}. Body: {Body}", response.StatusCode, errorBody);
                    return StatusCode((int)response.StatusCode, new
                    {
                        message = "Failed to download PDF",
                        apiResponse = errorBody
                    });
                }

                var pdfBytes = await response.Content.ReadAsByteArrayAsync();

                await _dgaSignDao.UpdateDgaEsignDocumentAsync(documentId, pdfBytes);

                string savePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", contype, "DGA", $"{contype}_{conId}.pdf");
                Directory.CreateDirectory(Path.GetDirectoryName(savePath)!);
                await System.IO.File.WriteAllBytesAsync(savePath, pdfBytes);

                return File(pdfBytes, "application/pdf", $"Signed_{documentId}.pdf");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "DownloadSignedPdf failed");
                return StatusCode(500, new
                {
                    message = "Internal Server Error",
                    error = ex.Message,
                    inner = ex.InnerException?.Message,
                    stack = ex.StackTrace
                });
            }
        }

        [HttpGet("GetPdfByContractType")]
        public async Task<byte[]?> GetPdfByContractType(string contractType, string contractId)
        {
            try
            {
                string htmlContent = contractType.ToUpper() switch
                {
                    "JOA" => await _JointOperationService.OnGetWordContact_JointOperationServiceHtmlToPDF(contractId),
                    "MOU" => await _MemorandumService.OnGetWordContact_MemorandumService_HtmlToPDF(contractId, "MOU"),
                    "PDPA" => await _PersernalProcessService.OnGetWordContact_PersernalProcessService_HtmlToPDF(contractId, "PDPA"),
                    "PDSA" => await _DataPersonalService.OnGetWordContact_DataPersonalService_ToPDF(contractId, "PDSA"),
                    "JDCA" => await _ControlDataService.OnGetWordContact_ControlDataServiceHtmlToPdf(contractId, "JDCA"),
                    "NDA" => await _DataSecretService.OnGetWordContact_DataSecretService_ToPDF(contractId, "NDA"),
                    "GA" => await _SupportSMEsService.OnGetWordContact_SupportSMEsService_HtmlToPDF(contractId, "GA"),
                    "AMJOA" => await _AMJOAService.OnGetWordContact_AMJOAServiceHtmlToPDF(contractId),
                    "MIW" => await _MIWService.OnGetWordContact_MIWServiceHtmlToPDF(contractId),
                    "MOA" => await _MemorandumInWritingService.OnGetWordContact_MemorandumInWritingService_HtmlToPDF(contractId, "MOA"),
                    "EC" => await _HireEmployee.OnGetWordContact_HireEmployee_ToPDF(contractId, "EC"),
                    _ => throw new ArgumentException("Unsupported contract type")
                };

                if (string.IsNullOrWhiteSpace(htmlContent))
                    return null;

                var pdfBytes = await _browserPdfService.PdfFromHtmlAsync(htmlContent);
                return pdfBytes;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "GetPdfByContractType failed for type {ContractType}, id {ContractId}", contractType, contractId);
                return null;
            }
        }

        [HttpGet("GetFilePDFCert")]
        public async Task<IActionResult> GetFilePDFCert([FromQuery] string ContractType = "JOA", string ContractId = "8")
        {
            try
            {
                var records = await _dgaSignDao.GetDgaEsignAsync(ContractType, int.Parse(ContractId));
                if (records == null || records.Count == 0)
                    return NotFound(new { message = "No DGA document records found" });

                DgaEsignModels? selected = null;
                var maxId = int.MinValue;
                foreach (var r in records)
                {
                    if (r != null && r.ID > maxId)
                    {
                        maxId = r.ID;
                        selected = r;
                    }
                }

                if (selected == null) return NotFound(new { message = "No DGA document record selected" });

                var pdfBytes = selected.DGA_DocumentDataFile;
                if (pdfBytes == null || pdfBytes.Length == 0)
                {
                    return NotFound(new
                    {
                        message = "PDF binary not found in DGA_DocumentDataFile",
                        DocumentID = selected.DGA_DocumentID,
                        Path = selected.DGA_DocumentPathFile
                    });
                }

                try
                {
                    var fileName = $"{ContractType.ToUpper()}_{ContractId}.pdf";
                    var saveDir = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Document", ContractType.ToUpper(), "DGA");
                    Directory.CreateDirectory(saveDir);
                    var savePath = Path.Combine(saveDir, fileName);
                    await System.IO.File.WriteAllBytesAsync(savePath, pdfBytes);
                }
                catch
                {
                    // ignore disk write errors
                }

                var downloadFileName = $"{selected.DGA_DocumentID ?? $"{ContractType}_{ContractId}"}.pdf";
                return File(pdfBytes, "application/pdf", downloadFileName);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "GetFilePDFCert failed for ContractType={ContractType} id={ContractId}", ContractType, ContractId);
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