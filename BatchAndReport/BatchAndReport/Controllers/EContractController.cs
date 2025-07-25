﻿using BatchAndReport.DAO;
using BatchAndReport.Models;
using BatchAndReport.Repository;
using BatchAndReport.Services;
using Microsoft.AspNetCore.Mvc;
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

        public EContractController(
            EContractDAO eContractDao,
            IApiInformationRepository repositoryApi,
            ICallAPIService serviceApi,
            IWordEContractService serviceWord)
        {
            _eContractDao = eContractDao;
            _repositoryApi = repositoryApi;
            _serviceApi = serviceApi;
            _serviceWord = serviceWord;
        }

        [HttpGet("GetEmpContract")]
        public async Task<IActionResult> GetEmpContract([FromQuery] string employmentDate)
        {
            try
            {
                var apiList = await _repositoryApi.GetAllAsync(new MapiInformationModels { ServiceNameCode = "employee-contract" });
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

                // Prepare filter model
                var filterModel = new
                {
                    employmentDate = employmentDate
                };

                var result = await _serviceApi.GetDataByParamApiAsync(apiParam, employmentDate);

                var employees = JsonSerializer.Deserialize<ApiListEmployeeContractResponse>(result.ToString());

                if (employees?.Results == null || !employees.Results.Any())
                    return BadRequest("No employee contract data found.");

                // Map to model
                var contractModels = employees.Results
                    .Where(emp => emp != null) // Ensure no null references
                    .Select(emp => new MEmployeeContractModels
                    {
                        ContractFlag = emp!.ContractFlag, // Use null-forgiving operator
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
                        Salary = emp.Salary,
                        IdCard = emp.IdCard,
                        PassportNo = emp.PassportNo
                    }).ToList();

                await _eContractDao.InsertOrUpdateEmployeeContractsAsync(contractModels);

                return Ok(new { message = "Sync and Save Complete", total = contractModels.Count });
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

    }
}