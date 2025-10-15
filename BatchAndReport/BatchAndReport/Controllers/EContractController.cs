using BatchAndReport.DAO;
using BatchAndReport.Models;
using BatchAndReport.Repository;
using BatchAndReport.Services;
using iText.Kernel.Crypto;
using Microsoft.AspNetCore.Mvc;
using Org.BouncyCastle.Crypto.Engines;
using Org.BouncyCastle.Crypto.Modes;
using Org.BouncyCastle.Crypto.Paddings;
using Org.BouncyCastle.Crypto.Parameters;

using Org.BouncyCastle.Security;
using System;
using System.Text;

using System.Text.Json;

namespace BatchAndReport.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class EContractController : ControllerBase
    {
        private readonly EContractDAO _eContractDao;
        private readonly IApiInformationRepository _repositoryApi;
        private readonly ICallAPIService _serviceApi;
        private readonly IWordEContractService _serviceWord;
        private readonly IConfiguration _configuration; // Add this line

        public EContractController(
            EContractDAO eContractDao,
            IApiInformationRepository repositoryApi,
            ICallAPIService serviceApi,
            IConfiguration configuration,
            IWordEContractService serviceWord)
        {
            _eContractDao = eContractDao;
            _repositoryApi = repositoryApi;
            _serviceApi = serviceApi;
            _configuration = configuration;
            _serviceWord = serviceWord;
        }

        //[HttpGet("GetEmpContract")]
        //public async Task<IActionResult> GetEmpContract([FromQuery] string employmentDate)
        //{
        //    try
        //    {
        //        var apiList = await _repositoryApi.GetAllAsync(new MapiInformationModels { ServiceNameCode = "employee-contract" });
        //        var apiParam = apiList.Select(x => new MapiInformationModels
        //        {
        //            ServiceNameCode = x.ServiceNameCode,
        //            ApiKey = x.ApiKey,
        //            AuthorizationType = x.AuthorizationType,
        //            ContentType = x.ContentType,
        //            CreateDate = x.CreateDate,
        //            Id = x.Id,
        //            MethodType = x.MethodType,
        //            ServiceNameTh = x.ServiceNameTh,
        //            Urldevelopment = x.Urldevelopment,
        //            Urlproduction = x.Urlproduction,
        //            Username = x.Username,
        //            Password = x.Password,
        //            UpdateDate = x.UpdateDate
        //        }).FirstOrDefault();

        //        if (apiParam == null)
        //            return BadRequest("API info not found.");

        //        // Prepare filter model
        //        var filterModel = new
        //        {
        //            employmentDate = employmentDate
        //        };

        //        var result = await _serviceApi.GetDataByParamApiAsync(apiParam, employmentDate);

        //        var employees = JsonSerializer.Deserialize<ApiListEmployeeContractResponse>(result.ToString());

        //        if (employees?.Results == null || !employees.Results.Any())
        //            return BadRequest("No employee contract data found.");

        //        string keyString = "5EU6l0SddrsT5HI6MhIxFkwT8JHRUyxz";
        //        // Map to model
        //        var contractModels = employees.Results
        //            .Where(emp => emp != null) // Ensure no null references
        //            .Select(emp => new MEmployeeContractModels
        //            {
        //                ContractFlag = emp!.ContractFlag?? false, // Use null-forgiving operator
        //                EmployeeId = emp.EmployeeId,
        //                EmployeeCode = emp.EmployeeCode,
        //                NameTh = emp.NameTh,
        //                NameEn = emp.NameEn,
        //                FirstNameTh = emp.FirstNameTh,
        //                FirstNameEn = emp.FirstNameEn,
        //                LastNameTh = emp.LastNameTh,
        //                LastNameEn = emp.LastNameEn,
        //                Email = emp.Email,
        //                Mobile = emp.Mobile,
        //                EmploymentDate = emp.EmploymentDate,
        //                TerminationDate = emp.TerminationDate,
        //                EmployeeType = emp.EmployeeType,
        //                EmployeeStatus = emp.EmployeeStatus,
        //                SupervisorId = emp.SupervisorId,
        //                CompanyId = emp.CompanyId,
        //                BusinessUnitId = emp.BusinessUnitId,
        //                PositionId = emp.PositionId,
        //                //Salary = emp.Salary,
        //                //IdCard = emp.IdCard,
        //                // ▼ Salary: ถอดรหัสก่อน; ถ้าไม่ได้ลองพาร์สเป็นตัวเลขดิบ; สุดท้ายให้ 0
        //                Salary = DecryptBase64String(emp.Salary, keyString)
        //                 ?? emp.Salary,

        //                // ▼ IdCard: ถอดรหัสก่อน; ถ้าไม่ได้เก็บค่าดิบไว้
        //                IdCard = DecryptBase64String(emp.IdCard, keyString)
        //                 ?? emp.IdCard,
        //                PassportNo = emp.PassportNo
        //            }).ToList();

        //        await _eContractDao.InsertOrUpdateEmployeeContractsAsync(contractModels);

        //        return Ok(new { message = "Sync and Save Complete", total = contractModels.Count });
        //    }
        //    catch (Exception ex)
        //    {
        //        return StatusCode(500, new
        //        {
        //            message = "Internal Server Error",
        //            error = ex.Message,
        //            inner = ex.InnerException?.Message,
        //            stack = ex.StackTrace
        //        });
        //    }
        //}


    [HttpGet("GetEmpContract")]
    public async Task<IActionResult> GetEmpContract([FromQuery] string employmentDate)
    {
        try
        {
            var jsonOpt = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };

            // (1) config: employee-contract
            var apiListContract = await _repositoryApi.GetAllAsync(new MapiInformationModels { ServiceNameCode = "employee-contract" });
            var apiParamContract = apiListContract.Select(x => new MapiInformationModels
            {
                ServiceNameCode = x.ServiceNameCode,
                ApiKey = x.ApiKey,
                AuthorizationType = x.AuthorizationType,
                ContentType = x.ContentType,
                CreateDate = x.CreateDate,
                Id = x.Id,
                MethodType = x.MethodType,
                ServiceNameTh = x.ServiceNameTh,
                Urldevelopment = x.Urldevelopment,
                Urlproduction = x.Urlproduction,
                Username = x.Username,
                Password = x.Password,
                UpdateDate = x.UpdateDate
            }).FirstOrDefault();

            if (apiParamContract == null)
                return BadRequest("API (employee-contract) not found.");

            // (2) call employee-contract
            var contractRaw = await _serviceApi.GetDataByParamApiAsync(apiParamContract, employmentDate);
            var contractResp = JsonSerializer.Deserialize<ApiListEmployeeContractResponse>(contractRaw?.ToString() ?? "", jsonOpt);
            if (contractResp?.Results == null || !contractResp.Results.Any())
                return BadRequest("No employee contract data found.");

            // (3) employee ids ที่จะไปดึงที่อยู่
            var employeeIds = contractResp.Results
                .Where(e => e != null && !string.IsNullOrWhiteSpace(e.EmployeeId))
                .Select(e => (e.EmployeeId ?? "").Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            // (4) config: employee-profile
            var apiListProfile = await _repositoryApi.GetAllAsync(new MapiInformationModels { ServiceNameCode = "employee-profile" });
            var apiParamProfile = apiListProfile.Select(x => new MapiInformationModels
            {
                ServiceNameCode = x.ServiceNameCode,
                ApiKey = x.ApiKey,
                AuthorizationType = x.AuthorizationType,
                ContentType = x.ContentType,
                CreateDate = x.CreateDate,
                Id = x.Id,
                MethodType = x.MethodType,
                ServiceNameTh = x.ServiceNameTh,
                Urldevelopment = x.Urldevelopment,
                Urlproduction = x.Urlproduction,
                Username = x.Username,
                Password = x.Password,
                UpdateDate = x.UpdateDate
            }).FirstOrDefault();

            // (4.x) ดึง MailingAddrTh แบบ path param: employee-profile/{empId}
            var addressByEmpId = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            if (apiParamProfile != null && employeeIds.Count > 0)
            {
                foreach (var empIdRaw in employeeIds)
                {
                    var empId = (empIdRaw ?? "").Trim();
                    if (string.IsNullOrEmpty(empId)) continue;

                    try
                    {
                        // 👇 ส่ง empId ตรง ๆ เพื่อให้กลายเป็น employee-profile/{empId}
                        var singleRaw = await _serviceApi.GetDataByParamApiAsync(apiParamProfile, empId);

                        var single = JsonSerializer.Deserialize<ApiEmployeeProfileResponse>(
                            singleRaw?.ToString() ?? "",
                            jsonOpt
                        );

                        // ❗ response ไม่มี employeeId ใน payload ให้ใช้ empId จาก path เป็นคีย์
                        var addr = single?.Results?.MailingAddrTh?.Trim();  // ใช้ mailingAddrTh ตรง ๆ

                        if (!string.IsNullOrWhiteSpace(addr))
                            addressByEmpId[empId] = addr!;
                    }
                    catch
                    {
                        // ข้ามแล้วไปต่อรายถัดไป
                    }
                }
            }

            // (5) map -> MEmployeeContractModels (Address = mailingAddrTh)
            string keyString = "5EU6l0SddrsT5HI6MhIxFkwT8JHRUyxz";

            var contractModels = contractResp.Results
                .Where(emp => emp != null)
                .Select(emp =>
                {
                    var key = (emp!.EmployeeId ?? "").Trim();
                    addressByEmpId.TryGetValue(key, out var addr);

                    return new MEmployeeContractModels
                    {
                        ContractFlag = emp.ContractFlag ?? false,
                        EmployeeId = emp.EmployeeId,
                        EmployeeCode = emp.EmployeeCode,
                        NameTh = emp.NameTh,
                        NameEn = emp.NameEn,
                        FirstNameTh = emp.FirstNameTh,
                        FirstNameEn = emp.FirstNameEn,
                        LastNameTh = emp.LastNameTh,
                        LastNameEn = emp.LastNameEn,
                        Email = emp.Email,
                        Mobile = emp.Mobile,
                        EmploymentDate = emp.EmploymentDate,
                        TerminationDate = emp.TerminationDate,
                        EmployeeType = emp.EmployeeType,
                        EmployeeStatus = emp.EmployeeStatus,
                        SupervisorId = emp.SupervisorId,
                        CompanyId = emp.CompanyId,
                        BusinessUnitId = emp.BusinessUnitId,
                        PositionId = emp.PositionId,

                        Salary = DecryptBase64String(emp.Salary, keyString) ?? emp.Salary,
                        IdCard = DecryptBase64String(emp.IdCard, keyString) ?? emp.IdCard,
                        PassportNo = emp.PassportNo,

                        Address = addr ?? ""   // กัน null
                    };
                })
                .ToList();

            // (6) upsert
            await _eContractDao.InsertOrUpdateEmployeeContractsAsync(contractModels);

            return Ok(new
            {
                message = "Sync and Save Complete",
                total = contractModels.Count,
                totalWithAddress = contractModels.Count(m => !string.IsNullOrWhiteSpace(m.Address))
            });
        }
        catch (Exception ex)
        {
            return StatusCode(500, new
            {
                message = "Internal Server Error",
                error = ex.Message,
                inner = ex.InnerException?.Message,
                stack = ex.StackTrace
            });
        }
    }


    static string NormalizeId(string s) => (s ?? string.Empty).Trim();


    [HttpGet("GetJuristicPerson")]
        public async Task<IActionResult> GetJuristicPerson([FromQuery] string? organizationJuristicID)
        {
            try
            {
                var apiList = await _repositoryApi.GetAllAsync(new MapiInformationModels { ServiceNameCode = "Juristic-Person" });

                var apiParam = apiList.Select(x => new MapiInformationModels
                {
                    ServiceNameCode = x.ServiceNameCode,
                    ApiKey = x.ApiKey,
                    AuthorizationType = x.AuthorizationType,
                    ContentType = x.ContentType,
                    CreateDate = x.CreateDate,
                    Id = x.Id,
                    MethodType = x.MethodType,
                    ServiceNameTh = x.ServiceNameTh,
                    Urldevelopment = x.Urldevelopment,
                    Urlproduction = x.Urlproduction,
                    Username = x.Username,
                    Password = x.Password,
                    UpdateDate = x.UpdateDate
                }).FirstOrDefault();

                if (apiParam == null)
                    return BadRequest("API info not found.");

                var options = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };

                if (!string.IsNullOrWhiteSpace(organizationJuristicID))
                {
                    // ✅ กรณีระบุเลขนิติบุคคล
                    var result = await _serviceApi.GetDataByParamApiAsync(apiParam, organizationJuristicID);
                    var root = JsonSerializer.Deserialize<OrganizationJuristicRoot>(result, options);

                    var person = root?.Data?.Person;
                    if (person == null)
                        return NotFound("ไม่พบข้อมูลนิติบุคคล");

                    var model = new MContractPartyModels
                    {
                        ContractPartyName = person.ContractPartyName,
                        RegType = person.RegType,
                        RegIden = person.RegIden,
                        RegDetail = person.RegDetail,
                        AddressNo = person.Address?.AddressType?.AddressNo,
                        SubDistrict = person.Address?.AddressType?.CitySubDivision?.SubDistrict,
                        District = person.Address?.AddressType?.City?.District,
                        Province = person.Address?.AddressType?.CountrySubDivision?.Province,
                        FlagActive = person.FlagActive
                    };

                    await _eContractDao.InsertOrUpdatePartyContractsAsync(new List<MContractPartyModels> { model });
                    return Ok(model);
                }
                else
                {
                    // ✅ กรณีไม่ส่ง param → ดึงทั้งหมด
                    var result = await _serviceApi.GetDataByParamApiAsync(apiParam, "");
                    var root = JsonSerializer.Deserialize<OrganizationJuristicListRoot>(result, options);

                    if (root?.Data == null || root.Data.Count == 0)
                        return NotFound("ไม่พบข้อมูลนิติบุคคลทั้งหมด");

                    var models = root.Data.Select(person => new MContractPartyModels
                    {
                        ContractPartyName = person.ContractPartyName,
                        RegType = person.RegType,
                        RegIden = person.RegIden,
                        RegDetail = person.RegDetail,
                        AddressNo = person.Address?.AddressType?.AddressNo,
                        SubDistrict = person.Address?.AddressType?.CitySubDivision?.SubDistrict,
                        District = person.Address?.AddressType?.City?.District,
                        Province = person.Address?.AddressType?.CountrySubDivision?.Province,
                        FlagActive = person.FlagActive
                    }).ToList();

                    var resultSync = await _eContractDao.SyncAllContractPartiesAsync(models);

                    return Ok(new { message = "Sync completed", total = resultSync.Count, data = resultSync });
                }
            }
            catch (Exception ex)
            {
                return StatusCode(500, new
                {
                    message = "Internal Server Error",
                    error = ex.Message,
                    inner = ex.InnerException?.Message,
                    stack = ex.StackTrace
                });
            }
        }

        [HttpGet("ExportJointContractWord")]
        public async Task<IActionResult> ExportJointContractWord()
        {
            //var detail = await _smeDao.GetProjectDetailAsync(projectCode);
            var detail = new ConJointContractModels
            {
                ProjectName = "โครงการพัฒนาศักยภาพผู้ประกอบการ SME",
                AgencyName = "มหาวิทยาลัยเทคโนโลยีราชมงคลล้านนา",
                SMEOfficialName = "นางสาวพรทิพย์ ธรรมวัฒน์",
                SMEOfficialPosition = "ผู้อำนวยการกองยุทธศาสตร์",
                AgencyRepresentative = "ดร.วิชัย วิริยะกิจจา",
                AgencyPosition = "รองอธิการบดี",
                SignDay = "1",
                SignMonth = "กันยายน",
                SignYear = "2567",

                Objectives = new List<ObjectiveItem>
    {
        new ObjectiveItem { Number = "1.", Description = "เสริมสร้างความสามารถในการแข่งขันของผู้ประกอบการ" },
        new ObjectiveItem { Number = "2.", Description = "สนับสนุนทุนพัฒนาและต่อยอดธุรกิจ" },
        new ObjectiveItem { Number = "3.", Description = "พัฒนาระบบบริหารจัดการองค์กรให้มีประสิทธิภาพ" }
    },

                SMEDuties = new List<string>
{
        "1.1  ตกลงร่วมดำเนินการโครงการโดยสนับสนุนงบประมาณ จำนวน ................บาท (.............................บาทถ้วน) ซึ่งได้รวมภาษีมูลค่าเพิ่ม ตลอดจนค่าภาษีอากรอื่น ๆ แล้วให้กับ “ชื่อหน่วยร่วม” และการใช้จ่ายเงินให้เป็นไปตามแผนการจ่ายเงินตามเอกสารแนบท้ายสัญญา",
        "1.2  ประสานการดำเนินโครงการ เพื่อให้บรรลุวัตถุประสงค์ เป้าหมายผลผลิตและผลลัพธ์",
        "1.3  กำกับ ติดตามและประเมินผลการดำเนินงานของโครงการ"
},
                AgencyDuties = new List<string>
{
        "2.1  ตกลงที่จะร่วมดำเนินการโครงการตามวัตถุประสงค์ของการโครงการและขอบเขตการดำเนินการ ตามรายละเอียดโครงการ แผนการดำเนินการ และแผนการใช้จ่ายเงิน (และอื่น ๆ เช่น คู่มือดำเนินโครงการ) ที่แนบท้ายสัญญาฉบับนี้",
        "2.2  ต้องดำเนินโครงการ ปฏิบัติตามแผนการดำเนินงาน แผนการใช้จ่ายเงิน (หรืออาจมีคู่มือการดำเนินโครงการก็ได้) อย่างเคร่งครัดและให้แล้วเสร็จภายในระยะเวลาโครงการ หากไม่ดำเนินโครงการให้แล้วเสร็จตามที่กำหนดยินยอมชำระค่าปรับให้แก่ สสว. ในอัตราร้อยละ 0.1 ของจำนวนงบประมาณที่ได้รับการสนับสนุนทั้งหมดต่อวัน นับถัดจากวันที่กำหนดแล้วเสร็จ และถ้าหากเห็นว่า “ชื่อหน่วยร่วม” ไม่อาจปฏิบัติตามสัญญาต่อไปได้ “ชื่อหน่วยร่วม” ยินยอมให้ สสว. ใช้สิทธิบอกเลิกสัญญาได้ทันที",
        "2.3  ต้องประสานการดำเนินโครงการ เพื่อให้บรรลุวัตถุประสงค์ เป้าหมายผลผลิตและผลลัพธ์",
        "2.4  ต้องให้ความร่วมมือกับ สสว. ในการกำกับ ติดตามและประเมินผลการดำเนินงานของโครงการ"
},
                OtherTerms = new List<string>
{
        "3.1  หากคู่สัญญาฝ่ายใดฝ่ายหนึ่งประสงค์จะขอแก้ไข เปลี่ยนแปลง ขยายระยะเวลาของโครงการ จะต้องแจ้งล่วงหน้าให้อีกฝ่ายหนึ่งได้ทราบเป็นลายลักษณ์อักษร และต้องได้รับความยินยอมเป็นลายลักษณ์อักษรจากอีกฝ่ายหนึ่ง และต้องทำเอกสารแก้ไข เปลี่ยนแปลง ขยายระยะเวลา เพื่อลงนามยินยอม ทั้งสองฝ่าย",
        "3.2  หากคู่สัญญาฝ่ายใดฝ่ายหนึ่งประสงค์จะขอบอกเลิกสัญญาก่อนครบกำหนดระยะเวลาดำเนินโครงการ จะต้องแจ้งล่วงหน้าให้อีกฝ่ายหนึ่งได้ทราบเป็นลายลักษณ์อักษรไม่น้อยกว่า 30 วัน และต้องได้รับความยินยอมเป็นลายลักษณ์อักษรจากอีกฝ่ายหนึ่ง และ “ชื่อหน่วยร่วม” จะต้องคืนเงินในส่วนที่ยังไม่ได้ใช้จ่ายหรือส่วนที่เหลือทั้งหมดพร้อมดอกผล (ถ้ามี) ให้แก่ สสว. ภายใน 15 วัน นับจากวันที่ได้รับหนังสือของฝ่ายที่ยินยอมให้บอกเลิก",
        "3.3  สสว. อาจบอกเลิกสัญญาได้ทันที หากตรวจสอบ หรือปรากฏข้อเท็จจริงว่า การใช้จ่ายเงินของ “ชื่อหน่วยร่วม” ไม่เป็นไปตามวัตถุประสงค์ของโครงการ แผนการดำเนินงาน และแผนการใช้จ่ายเงิน (และอื่น ๆ เช่น คู่มือดำเนินโครงการ) ทั้งมีสิทธิเรียกเงินคงเหลือคืนทั้งหมดพร้อมดอกผล (ถ้ามี) ได้ทันที",
        "3.4  ทรัพย์สินใด ๆ และ/หรือ สิทธิใด ๆ ที่ได้มาจากเงินสนับสนุนตามสัญญาร่วมดำเนินการฉบับนี้ เมื่อสิ้นสุดโครงการให้ตกได้แก่ สสว. ทั้งสิ้น เว้นแต่ สสว. จะกำหนดให้เป็นอย่างอื่น",
        "3.5  “ชื่อหน่วยร่วม” ต้องไม่ดำเนินการในลักษณะการจ้างเหมา กับหน่วยงาน องค์กร หรือบุคคลอื่น ๆ ยกเว้นกรณีการจัดหา จัดจ้าง เป็นกิจกรรมหรือเป็นเรื่อง ๆ",
        "3.6  ในกรณีที่การดำเนินการตามสัญญาฉบับนี้ เกี่ยวข้องกับข้อมูลส่วนบุคคล และการคุ้มครองทรัพย์สินทางปัญญา “ชื่อหน่วยร่วม” จะต้องปฏิบัติตามกฎหมายว่าด้วยการคุ้มครองข้อมูลส่วนบุคคลและการคุ้มครองทรัพย์สินทางปัญญาอย่างเคร่งครัด และหากเกิดความเสียหายหรือมีการฟ้องร้องใด ๆ “ชื่อหน่วยร่วม” จะต้องเป็นผู้รับผิดชอบต่อการละเมิดบทบัญญัติแห่งกฎหมายดังกล่าวแต่เพียงฝ่ายเดียวโดยสิ้นเชิง"
}
            };

            if (detail == null)
                return NotFound("ไม่พบข้อมูลโครงการ");

            var wordBytes = _serviceWord.GenJointContractAgreement(detail);
            var pdfBytes = _serviceWord.ConvertWordToPdf(wordBytes);
            return File(wordBytes,
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                $"JointContract.docx");
        }

        [HttpGet("Project")]
        [Produces("application/json")]
        public async Task<IActionResult> GetProject([FromQuery] string projectCode)
        {
            if (string.IsNullOrWhiteSpace(projectCode))
            {
                return BadRequest(new
                {
                    responseCode = "400",
                    responseMsg = "Missing required query param 'projectCode'",
                    data = Array.Empty<object>()
                });
            }

            string json;
            try
            {
                json = await _eContractDao.GetProjectByProjectCodeAsync(projectCode);
            }
            catch (Exception ex)
            {
                var err = new { responseCode = "500", responseMsg = "Database error: " + ex.Message, data = Array.Empty<object>() };
                return Content(System.Text.Json.JsonSerializer.Serialize(err), "application/json", Encoding.UTF8);
            }

            // sanity check แบบเบาๆ ว่าสตริงหน้าตาเป็น JSON
            if (string.IsNullOrWhiteSpace(json) || !(json.TrimStart().StartsWith("{") || json.TrimStart().StartsWith("[")))
            {
                var err = new { responseCode = "500", responseMsg = "Stored procedure returned invalid JSON.", data = Array.Empty<object>() };
                return Content(System.Text.Json.JsonSerializer.Serialize(err), "application/json", Encoding.UTF8);
            }

            // สำคัญ: ส่งเป็น application/json ตรง ๆ — ไม่ Ok(string) (จะถูก escape)
            return Content(json, "application/json", Encoding.UTF8);
        }
        [HttpGet("convertSaraly")]
        public IActionResult convertSaraly()
        {
            string decryptedText = "";

            try
            {

                string encryptedBase64 = "7GtvfN8LB5aEOtY0c7didw=="; //8EgYChBAnJaZTdXJNarZng==
                string encryptionKey = "5EU6l0SddrsT5HI6MhIxFkwT8JHRUyxz";

                decryptedText = DecryptBase64String(encryptedBase64, encryptionKey);

            }
            catch (Exception ex)
            {
                decryptedText = $"Decryption failed: {ex.Message}";
            }

            // ViewBag.DecryptedText = decryptedText;
            return Ok(decryptedText);
        }
        // ตัวอย่างการถอดรหัส AES-256

        public static string DecryptBase64String(string base64String, string keyString)
        {
            // 1. แปลง Base64 String เป็น byte array
            byte[] encryptedBytes = Convert.FromBase64String(base64String);

            // 2. แปลง Key String เป็น byte array
            byte[] keyBytes = Encoding.UTF8.GetBytes(keyString);

            // 3. สร้าง AES engine และตั้งค่า (initialization)
            var aesEngine = new AesEngine();
            var cbcBlockCipher = new CbcBlockCipher(aesEngine); // Bouncy Castle ใช้ CbcBlockCipher เพื่อรองรับ padding และการใช้งานที่หลากหลาย
            var keyParameter = new KeyParameter(keyBytes);

            // 4. ตั้งค่าเพื่อการถอดรหัส
            cbcBlockCipher.Init(false, keyParameter); // 'false' สำหรับ Decryption

            // 5. สร้าง PaddedBufferedBlockCipher
            var decryptor = new PaddedBufferedBlockCipher(cbcBlockCipher, new Pkcs7Padding());
            decryptor.Init(false, keyParameter);

            // 6. ถอดรหัส
            byte[] decryptedBytes = new byte[decryptor.GetOutputSize(encryptedBytes.Length)];
            int length = decryptor.ProcessBytes(encryptedBytes, 0, encryptedBytes.Length, decryptedBytes, 0);
            decryptor.DoFinal(decryptedBytes, length);

            // 7. แปลงผลลัพธ์เป็น string UTF-8
            string decryptedString = Encoding.UTF8.GetString(decryptedBytes).TrimEnd('\0'); // trim null characters if any

            return decryptedString;
        }

        [HttpGet("ExportImportContract")]
        public async Task<IActionResult> ExportImportContract([FromQuery] string? contractNumber = null)
        {
            var details = await _eContractDao.FindImportContractsAsync(contractNumber);

            if (details == null || !details.Any())
                return NotFound("ไม่พบข้อมูลโครงการ");

            var generator = _serviceWord.GenImportContract(details);
            var excelBytes = generator; // Assuming `GenImportContract` already returns a byte array.

            return File(excelBytes,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                $"ExportContract_{contractNumber}.xlsx");
        }

        [HttpGet("ExportImportContractLoanEncrypt")]
        public async Task<IActionResult> ExportImportContractLoanEncrypt([FromQuery] string? contractNumber = null)
        {
            var details = await _eContractDao.FindImportContractsAsync(contractNumber);

            if (details == null || !details.Any())
                return NotFound("ไม่พบข้อมูลโครงการ");

            var generator = _serviceWord.GenImportContractLoan(details);
            var excelBytes = generator; // Assuming `GenImportContract` already returns a byte array.

            // Get password from appsettings.json
            var password = _configuration["Password:PaswordPDF"];

            if (!string.IsNullOrEmpty(password))
            {
                using var inputStream = new MemoryStream(excelBytes);
                using var package = new OfficeOpenXml.ExcelPackage(inputStream);
                package.Encryption.IsEncrypted = true;
                package.Encryption.Password = password;

                using var outputStream = new MemoryStream();
                package.SaveAs(outputStream);
                outputStream.Position = 0;
                return File(outputStream.ToArray(),
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    $"ExportContractLoan_ProtectedEncypt_{contractNumber}.xlsx");
            }
            else
            {
                return File(excelBytes,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                $"ExportContractLoan_{contractNumber}.xlsx");
            }
        }

        [HttpGet("ExportImportContractLoan")]
        public async Task<IActionResult> ExportImportContractLoan([FromQuery] string? contractNumber = null)
        {
            var details = await _eContractDao.FindImportContractsAsync(contractNumber);

            if (details == null || !details.Any())
                return NotFound("ไม่พบข้อมูลโครงการ");

            var generator = _serviceWord.GenImportContractLoan(details);
            var excelBytes = generator; // Assuming `GenImportContract` already returns a byte array.

        
            return File(excelBytes,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                $"ExportContractLoan_{contractNumber}.xlsx");
         
        }
    }
}

