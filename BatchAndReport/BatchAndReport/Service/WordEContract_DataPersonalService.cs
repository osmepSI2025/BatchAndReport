using BatchAndReport.DAO;
using BatchAndReport.Models;
using DinkToPdf.Contracts;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.AspNetCore.Http;
using System.Text;

public class WordEContract_DataPersonalService
{
    private readonly WordServiceSetting _w;
    private readonly Econtract_Report_PDSADAO _e;
    private readonly IConverter _pdfConverter;
    private readonly E_ContractReportDAO _eContractReportDAO;
    public WordEContract_DataPersonalService(WordServiceSetting ws, Econtract_Report_PDSADAO e
        , IConverter pdfConverter
        , E_ContractReportDAO eContractReportDAO
        )
    {
        _w = ws;
        _e = e;
        _pdfConverter = pdfConverter;
        _eContractReportDAO = eContractReportDAO;
    }

    #region  4.1.1.2.6.บันทึกข้อตกลงการแบ่งปันข้อมูลส่วนบุคคล PDSA
  

    #endregion  4.1.1.2.6.บันทึกข้อตกลงการแบ่งปันข้อมูลส่วนบุคคล

    #region PDSA
    public async Task<string> OnGetWordContact_DataPersonalService_ToPDF(string id,string typeContact="PDSA")
    {
        try
        {
            var result = await _e.GetPDSAAsync(id);
            if (result == null)
                throw new Exception("PDSA data not found.");
            var rLe = await _e.GetPDSA_LegalBasisSharingAsync(id);
            var rSd = await _e.GetPDSA_Shared_DataAsync(id);
           // var fontPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "font", "THSarabun.ttf").Replace("\\", "/");
            var fontPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "font", "THSarabun.ttf");
            string fontBase64 = "";
            if (File.Exists(fontPath))
            {
                var bytes = File.ReadAllBytes(fontPath);
                fontBase64 = Convert.ToBase64String(bytes);
            }

            var logoPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "images", "logo_SME.jpg");
            string contractLogoHtml;
            if (!string.IsNullOrEmpty(result.Organization_Logo) && result.Organization_Logo.Contains("<content>"))
            {
                try
                {
                    // ตัดเอาเฉพาะ Base64 ในแท็ก <content>...</content>
                    var contentStart = result.Organization_Logo.IndexOf("<content>") + "<content>".Length;
                    var contentEnd = result.Organization_Logo.IndexOf("</content>");
                    var contractlogoBase64 = result.Organization_Logo.Substring(contentStart, contentEnd - contentStart);

                    contractLogoHtml = $@"<div style='display:inline-block; padding:20px; font-size:32pt;'>
                 <img src='data:image/jpeg;base64,{contractlogoBase64}' width='240' height='80' />
                </div>";
                }
                catch
                {
                    contractLogoHtml = "";
                }
            }
            else
            {
                contractLogoHtml = "";
            }
            string datestring = CommonDAO.ToThaiDateStringCovert(result.Master_Contract_Sign_Date ?? DateTime.Now);            string logoBase64 = "";
            if (System.IO.File.Exists(logoPath))
            {
                var bytes = System.IO.File.ReadAllBytes(logoPath);
                logoBase64 = Convert.ToBase64String(bytes);
            }

            #region signlist 

            var signlist = await _eContractReportDAO.GetSignNameAsync(id, typeContact);
            var signatoryTableHtml = "";
            if (signlist.Count > 0)
            {
                signatoryTableHtml = await _eContractReportDAO.RenderSignatory(signlist);

            }

            var signatoryTableHtmlWitnesses = "";

            if (signlist.Count > 0)
            {
                signatoryTableHtmlWitnesses = await _eContractReportDAO.RenderSignatory_Witnesses(signlist);
            }

            #endregion signlist




            var html = $@"
<html>
<head>
    <meta charset='utf-8'>
      <style>
        @font-face {{
            font-family: 'TH Sarabun PSK';
              src: url('data:font/truetype;charset=utf-8;base64,{fontBase64}') format('truetype');
            font-weight: normal;
            font-style: normal;
        }}
        body {{
            font-size: 22px;
            font-family: 'TH Sarabun PSK', Arial, sans-serif !important;
        }}
        /* แก้การตัดคำไทย: ไม่หั่นกลางคำ, ตัดเมื่อจำเป็น */
        body, p, div {{
            word-break: keep-all;            /* ห้ามตัดกลางคำ */
            overflow-wrap: break-word;       /* ตัดเฉพาะเมื่อจำเป็น (ยาวจนล้นบรรทัด) */
            -webkit-line-break: after-white-space; /* ช่วย WebKit เก่าจัดบรรทัด */
            hyphens: none;
        }}
         .t-12 {{ font-size: 1em; }}
             .t-14 {{ font-size: 1.1em; }}
        .t-15 {{ font-size: 1.2em; }}
        .t-16 {{ font-size: 1.5em; }}
        .t-18 {{ font-size: 1.7em; }}
        .t-20 {{ font-size: 1.9em; }}
        .t-22 {{ font-size: 2.1em; }}
  .tab0 {{ text-indent: 0px;     }}
           .tab1 {{ text-indent: 48px; text-align: justify;  }}
        .tab2 {{ text-indent: 96px;  text-align: left; }}
        .tab3 {{ text-indent: 144px; text-align: left; }}
        .tab4 {{ text-indent: 192px;  text-align: left;}}
       .normal {{text-align: justify;
        text-align-last: justify;
        width: 100%;
        display: block;
        min-width: 100%;
  letter-spacing: 0.1em; /* เพิ่มช่องไฟเล็กน้อย */
    }}
        .d-flex {{ display: flex; }}
        .w-100 {{ width: 100%; }}
        .w-40 {{ width: 40%; }}
        .w-50 {{ width: 50%; }}
        .w-60 {{ width: 60%; }}
        .text-center {{ text-align: center; }}
         .sign-single-right {{
            display: flex;
            flex-direction: column;
            position: relative;
            left: 20%;
        }}
        .sign-double {{ display: flex; }}
        .text-center-right-brake {{
            margin-left: 50%;
            word-break: break-all;
        }}
        .text-right {{ text-align: right; }}
        .contract, .section {{
            margin: 12px 0;
            line-height: 1.7;
        }}
        .section {{
            font-weight: bold;
            font-size: 1.2em;
            text-align: left;
            margin-top: 24px;
        }}
        .signature-table {{
            width: 100%;
            margin-top: 32px;
            border-collapse: collapse;
        }}
        .signature-table td {{
            padding: 16px;
            text-align: center;
            vertical-align: top;
           font-size: 1.4em;
        }}
   .table {{ width: 100%; border-collapse: collapse; margin-top: 20px; font-size: 1em; }}
    .table th, .table td {{ border: 1px solid #000; padding: 8px; }}
.logo-table {{ width: 100%; border-collapse: collapse; margin-top: 40px; }}
        .logo-table td {{ border: none; }}
        p {{
            margin: 0;
            padding: 0;
        }}

    </style>

</head>
<body>
<table style='width:100%; border-collapse:collapse; margin-top:40px;'>
    <tr>
        <!-- Left: SME logo -->
        <td style='width:60%; text-align:left; vertical-align:top;'>
        <div style='display:inline-block;  padding:20px; font-size:32pt;'>
             <img src='data:image/jpeg;base64,{logoBase64}' width='240' height='80' />
            </div>
        </td>
        <!-- Right: Contract code box (replace with your actual contract code if needed) -->
        <!-- Right: Contract code box (replace with your actual contract code if needed) -->
        <td style='width:40%; text-align:center; vertical-align:top;'>
            {contractLogoHtml}
        </td>
    </tr>
</table>
</br>

    <!-- Titles -->
    <div class='t-12 text-center'><b>ข้อตกลงการแบ่งปันข้อมูลส่วนบุคคล</b></div>
    <div class='t-12 text-center'><b>(Personal Data Sharing Agreement)</b></div>
    <div class='t-12 text-center'><b>ระหว่าง</b></div>
    <div class='t-12 text-center'><b>สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม กับ {result.ContractPartyName}</b></div>
    <div class='t-12 text-center'>---------------------------------------------</div>
    <!-- Paragraphs -->
    <p class='tab2 t-12'>ข้อตกลงการแบ่งปันข้อมูลส่วนบุคคล (“ข้อตกลง”) ฉบับนี้ทำขึ้น เมื่อ {datestring} ณ สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม</p>
<p class='tab2 t-12'>
    โดยที่ สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม ซึ่งต่อไปในข้อตกลงฉบับนี้เรียกว่า “สสว.” ฝ่ายหนึ่ง ได้ตกลงใน {result.ContractPartyName} สัญญาเลขที่ {result.Master_Contract_Number} ฉบับลงวันที่ {datestring} ซึ่งต่อไปในข้อตกลงฉบับนี้เรียกว่า “สัญญาหลัก” กับ {result.ContractPartyName} ซึ่งต่อไปในข้อตกลงฉบับนี้เรียกว่า “{result.ContractPartyName}” อีกฝ่ายหนึ่ง รวมเรียกว่า “คู่สัญญา”
</p>
<p class='tab2 t-12'>เพื่อให้บรรลุวัตถุประสงค์ภายใต้ความตกลงของสัญญาหลัก คู่สัญญามีความจำเป็นต้อง
แบ่งปัน โอน แลกเปลี่ยน หรือเปิดเผย (รวมเรียกว่า “แบ่งปัน”) ข้อมูลส่วนบุคคลที่ตนเก็บรักษาแก่อีกฝ่าย
ซึ่งข้อมูลส่วนบุคคลที่แต่ละฝ่าย เก็บรวมรวม ใช้หรือเปิดเผย (รวมเรียกว่า “ประมวลผล”) นั้น แต่ละฝ่ายต่าง
เป็นผู้ควบคุมข้อมูลส่วนบุคคล ตามกฎหมายที่เกี่ยวข้องกับการคุ้มครองข้อมูลส่วนบุคคล กล่าวคือแต่ละฝ่าย
ต่างเป็นผู้มีอำนาจตัดสินใจ กำหนดรูปแบบ และกำหนดวัตถุประสงค์ ในการประมวลผลข้อมูลส่วนบุคคล
ในข้อมูลที่ตนต้องแบ่งปัน ภายใต้ข้อตกลงนี้
</p>
<p class='tab2 t-12'>ด้วยเหตุนี้ คู่สัญญาจึงตกลงจัดทำข้อตกลงฉบับนี้ และให้ถือเป็นส่วนหนึ่งของสัญญาหลัก
เพื่อเป็นหลักฐานการแบ่งปันข้อมูลส่วนบุคคลระหว่างคู่สัญญาและเพื่อดำเนินการให้เป็นไปตามพระราชบัญ
ญัติคุ้มครองข้อมูลส่วนบุคคล พ.ศ. ๒๕๖๒ และกฎหมายอื่น ๆ ที่ออกตามความในพระราชบัญญัติคุ้มครอง
ข้อมูลส่วนบุคคล พ.ศ. ๒๕๖๒ ซึ่งต่อไปในข้อตกลงฉบับนี้ รวมเรียกว่า “กฎหมายคุ้มครองข้อมูลส่วนบุคคล”
ทั้งที่มีผลใช้บังคับอยู่ ณ วันทำข้อตกลงฉบับนี้ และที่จะมีการเพิ่มเติมหรือแก้ไขเปลี่ยนแปลงในภายหลัง โดยมีรายละเอียดดังนี้
</p>
    <p class='tab2 t-12'>๑. คู่สัญญารับทราบว่า ข้อมูลส่วนบุคคล หมายถึง ข้อมูลเกี่ยวกับบุคคลธรรมดา ซึ่งทำให้
สามารถระบุตัวบุคคลนั้นได้ไม่ว่าทางตรงหรือทางอ้อม โดยคู่สัญญาแต่ละฝ่าย จะดำเนินการตามที่กฎหมาย
คุ้มครองข้อมูลส่วนบุคคลกำหนด เพื่อคุ้มครองให้การประมวลผลข้อมูลส่วนบุคคลเป็นไปอย่างเหมาะสมและ
ถูกต้องตามกฎหมาย</p>
    <p class='tab2 t-12'>๒. ข้อมูลส่วนบุคคลที่คู่สัญญาแบ่งปันกัน คู่สัญญาแต่ละฝ่ายตกลงแบ่งปันข้อมูลส่วนบุคคลดัง
รายการต่อไปนี้แก่คู่สัญญาอีกฝ่าย</p>
    <!-- Table: ข้อมูลส่วนบุคคลที่แบ่งปันโดย สสว. -->
    <table class='table t-12'>
        <tr>
            <th>ข้อมูลส่วนบุคคลที่แบ่งปันโดย สสว.</th>
            <th>วัตถุประสงค์ในการแบ่งปันข้อมูลส่วนบุคคล</th>
        </tr>
{string.Join("", rSd.Where(e => e.Owner == "OSMEP").Select(item => $"<tr><td>{item.Detail ?? "-"}</td><td>{item.Objective ?? "-"}</td></tr>"))}
    </table>
    <!-- Table: ข้อมูลส่วนบุคคลที่แบ่งปันโดยคู่สัญญา -->
   <table class='table t-12'>
        <tr>
            <th>ข้อมูลส่วนบุคคลที่แบ่งปันโดย {result.ContractPartyName}</th>
            <th>วัตถุประสงค์ในการแบ่งปันข้อมูลส่วนบุคคล</th>
        </tr>
      {string.Join("", rSd.Where(e => e.Owner == "CP").Select(item => $"<tr><td>{item.Detail ?? "-"}</td><td>{item.Objective ?? "-"}</td></tr>"))}
    </table>
    <!-- Table: ฐานกฎหมายในการแบ่งปันข้อมูลส่วนบุคคล -->

    <p class='tab2 t-12'>๓. ฐานกฎหมายในการแบ่งปันข้อมูลส่วนบุคคล ภายใต้วัตถุประสงค์ที่ระบุในข้อ ๒ คู่สัญญา
แต่ละฝ่ายมีฐานกฎหมายตามกฎหมายคุ้มครองข้อมูลส่วนบุคคลดังต่อไปนี้ ในการแบ่งปันข้อมูลส่วนบุคคลแก่
คู่สัญญาอีกฝ่าย (แต่ละฝ่ายอาจใช้ฐานกฎหมายที่ต่างกันในการแบ่งปันข้อมูลส่วนบุคคล)</p>
   <table class='table t-12'>
        <tr>
            <th>ฐานกฎหมายของ สสว.</th>
        </tr>
   {string.Join("", rLe.Where(e => e.Owner == "OSMEP").Select(item => $"<tr><td>{item.Detail ?? "-"}</td></tr>"))}
    </table>
   <table class='table t-12'>
        <tr>
            <th>ฐานกฎหมายของ {result.ContractPartyName}</th>
        </tr>
    {string.Join("", rLe.Where(e => e.Owner == "CP").Select(item => $"<tr><td>{item.Detail ?? "-"}</td></tr>"))}
    </table>
<!-- No file path since this is a template snippet -->
 <p class='tab2 t-12'>๔. คู่สัญญารับทราบและตกลงว่า แต่ละฝ่ายต่างเป็นผู้ควบคุมข้อมูลส่วนบุคคลในส่วนของ
ข้อมูลส่วนบุคคลที่ตนประมวลผล และต่างอยู่ภายใต้บังคับในการปฏิบัติตามกฎหมายคุ้มครองข้อมูลส่วนบุคคล
ในบทบัญญัติที่เกี่ยวข้องกับผู้ควบคุมข้อมูลส่วนบุคคลต่างหากจากกัน</p>
 <p class='tab2 t-12'>๕. คู่สัญญารับรองและยืนยันว่า ก่อนการแบ่งปันข้อมูลส่วนบุคคลแก่อีกฝ่าย ตนได้ดำเนิน
การแจ้งข้อมูลที่จำเป็นเกี่ยวกับการแบ่งปันข้อมูลและขอความยินยอมจากเจ้าของข้อมูลส่วนบุคคลและ/หรือ
มีฐานกฎหมายหรืออำนาจหน้าที่โดยชอบด้วยกฎหมายให้สามารถเปิดเผยข้อมูลส่วนบุคคลให้อีกฝ่าย และให้
อีกฝ่ายสามารถทำการประมวลผลข้อมูลส่วนบุคคลที่ได้รับนั้นตามวัตถุประสงค์ที่ได้ตกลงกันอย่างถูกต้องตาม
กฎหมายคุ้มครองข้อมูลส่วนบุคคลแล้ว</p>
 <p class='tab2 t-12'>๖. คู่สัญญารับรองว่า คู่สัญญาฝ่ายที่แบ่งปันข้อมูลส่วนบุคคล จะไม่ถูกจำกัดสิทธิ ยับยั้งหรือมีข้อห้ามใด ๆ ในการ</p>
 <p class='tab2 t-12'>๖.๑.ประมวลผลข้อมูลส่วนบุคคลที่ตนเป็นฝ่ายแบ่งปัน ภายใต้วัตถุประสงค์ที่กำหนดในข้อ
ตกลงฉบับนี้</p>
 <p class='tab2 t-12'>๖.๒.แบ่งปันส่วนบุคคลไปยังคู่สัญญาอีกฝ่ายเพื่อการปฏิบัติหน้าที่ตามข้อตกลงฉบับนี้</p>
 <p class='tab2 t-12'>๗. คู่สัญญาจะทำการประมวลผลข้อมูลส่วนบุคคลที่รับมาจากอีกฝ่ายเพียงเท่าที่จำเป็น เพื่อให้
บรรลุวัตถุประสงค์ที่ได้กำหนดในข้อ ๒ ของข้อตกลงฉบับนี้และแต่ละฝ่ายจะไม่ประมวลผลข้อมูล เพื่อ
วัตถุประสงค์อื่นเว้นแต่ได้รับความยินยอมจากเจ้าของข้อมูลส่วนบุคคล หรือเป็นความจำเป็นเพื่อปฏิบัติตาม
กฎหมายเท่านั้น</p>
 <p class='tab2 t-12'>๘. คู่สัญญารับรองว่าจะควบคุมดูแลให้เจ้าหน้าที่ และ/หรือลูกจ้าง ตัวแทนหรือบุคคลใด ๆ
ที่ปฏิบัติหน้าที่ในการประมวลผลข้อมูลส่วนบุคคลที่ได้รับจากอีกฝ่ายภายใต้ข้อตกลงฉบับนี้รักษาความลับ
และปฏิบัติตามกฎหมายคุ้มครองข้อมูลส่วนบุคคลอย่างเคร่งครัด และดำเนินการประมวลผลข้อมูลส่วน
บุคคลเพื่อวัตถุประสงค์ตามข้อตกลงฉบับนี้เท่านั้น โดยจะไม่ทำซ้ำ คัดลอก ทำสำเนา บันทึกภาพข้อมูล
ส่วนบุคคลไม่ว่าทั้งหมดหรือแต่บางส่วนเป็นอันขาดเว้นแต่เป็นไปตามเงื่อนไขของสัญญาหลัก หรือ
กฎหมายที่เกี่ยวข้องจะระบุหรือบัญญัติไว้เป็นประการอื่น</p>

 <p class='tab2 t-12'>๙. คู่สัญญารับรองว่าจะกำหนดให้การเข้าถึงข้อมูลส่วนบุคคลภายใต้ข้อตกลงฉบับนี้ ถูกจำกัด
เฉพาะเจ้าหน้าที่ และ/หรือลูกจ้าง ตัวแทนหรือบุคคลใด ๆ ที่ได้รับมอบหมาย มีหน้าที่เกี่ยวข้องหรือมีความ
จำเป็นในการเข้าถึงข้อมูลส่วนบุคคลภายใต้ข้อตกลงฉบับนี้เท่านั้น</p>

 <p class='tab2 t-12'>๑๐. คู่สัญญาฝ่ายที่รับข้อมูลจะไม่เปิดเผยข้อมูลส่วนบุคคลที่ได้รับจากฝ่ายที่โอนข้อมูลแก่
บุคคลของคู่สัญญาฝ่ายที่รับข้อมูลที่ไม่มีอำนาจหน้าที่ที่เกี่ยวข้องในการประมวลผล หรือบุคคลภายนอกใด ๆ
เว้นแต่ที่มีความจำเป็นต้องกระทำตามหน้าที่ในสัญญาหลัก ข้อตกลงฉบับนี้หรือเพื่อปฏิบัติตามกฎหมาย
ที่ใช้บังคับ หรือ ที่ได้รับความยินยอมเป็นลายลักษณ์อักษรจากคู่สัญญาฝ่ายที่โอนข้อมูลก่อน</p>

 <p class='tab2 t-12'>๑๑. คู่สัญญาจัดให้มีและคงไว้ซึ่งมาตรการรักษาความปลอดภัยสำหรับการประมวลผล ข้อมูล
ที่มีความเหมาะสม ทั้งในเชิงองค์กรและเชิงเทคนิคตามที่คณะกรรมการคุ้มครองข้อมูลส่วนบุคคลได้ประกาศ
กำหนดและหรือตามมาตรฐานสากล โดยคำนึงถึงลักษณะ ขอบเขต และวัตถุประสงค์ของการประมวลผล
ข้อมูลเพื่อคุ้มครองข้อมูลส่วนบุคคลจากความเสี่ยงอันเกี่ยวเนื่องกับการประมวลผลข้อมูลส่วนบุคคลเช่น 
ความเสียหายอันเกิดจากการละเมิด อุบัติเหตุ ลบ ทำลาย สูญหาย เปลี่ยนแปลง แก้ไข เข้าถึง ใช้เปิดเผย
หรือโอนข้อมูลส่วนบุคคลโดยไม่ชอบด้วยกฎหมาย</p>

 <p class='tab2 t-12'>๑๒. กรณีที่คู่สัญญาฝ่ายหนึ่งฝ่ายใด พบพฤติการณ์ที่มีลักษณะที่กระทบต่อการรักษาความ
ปลอดภัยของข้อมูลส่วนบุคคลที่แบ่งปันกันภายใต้ข้อตกลงฉบับนี้ ซึ่งอาจก่อให้เกิดความเสียหายจากการ
ละเมิด อุบัติเหตุ ลบ ทำลาย สูญหาย เปลี่ยนแปลง แก้ไข เข้าถึง ใช้เปิดเผยหรือโอนข้อมูลส่วนบุคคล
โดยไม่ชอบด้วยกฎหมาย คู่สัญญาฝ่ายที่พบเหตุดังกล่าวจะดำเนินการแจ้งให้คู่สัญญาอีกฝ่ายทราบโดยทันที
ภายในเวลาไม่เกิน {result.IncidentNotifyPeriod} ชั่วโมง</p>

 <p class='tab2 t-12'>๑๓ การแจ้งถึงเหตุการละเมิดข้อมูลส่วนบุคคลที่เกิดขึ้นภายใต้ข้อตกลงนี้ คู่สัญญาแต่ละฝ่าย
จะใช้มาตรการตามที่เห็นสมควรในการระบุถึงสาเหตุของการละเมิดและป้องกันปัญหาดังกล่าวมิให้เกิดซ้ำ
และจะให้ข้อมูลแก่อีกฝ่ายภายใต้ขอบเขตที่กฎหมายคุ้มครองข้อมูลส่วนบุคคลได้กำหนด ดังต่อไปนี้</p>
 <p class='tab2 t-12'>๑๓.๑ รายละเอียดของลักษณะและผลกระทบที่อาจเกิดขึ้นของการละเมิด</p>
 <p class='tab2 t-12'>๑๓.๒ มาตรการที่ถูกใช้เพื่อลดผลกระทบของการละเมิด</p>
 <p class='tab2 t-12'>๑๓.๓ ประเภทของข้อมูลส่วนบุคคลและเจ้าของข้อมูลส่วนบุคคลที่ถูกละเมิด หากมีปรากฏ</p>
 <p class='tab2 t-12'>๑๓.๔ ข้อมูลอื่น ๆ เกี่ยวข้องกับการละเมิด</p>
 <p class='tab2 t-12'>๑๔. คู่สัญญาตกลงจะให้ความช่วยเหลืออย่างสมเหตุสมผลแก่อีกฝ่าย เพื่อปฏิบัติตามกฎหมาย
คุ้มครองข้อมูลที่ใช้บังคับในการตอบสนองต่อข้อเรียกร้องใด ๆ ที่สมเหตุสมผลจากการใช้สิทธิต่างๆ ภายใต้
กฎหมายคุ้มครองข้อมูลส่วนบุคคลโดยเจ้าของข้อมูลส่วนบุคคล โดยพิจารณาถึงลักษณะการประมวลผล ภาระหน้าที่ภายใต้กฎหมายคุ้มครองข้อมูลส่วนบุคคลที่ใช้บังคับ และข้อมูลส่วนบุคคลที่แต่ละฝ่ายประมวลผล</p>
 <p class='tab2 t-12'>ทั้งนี้ กรณีที่เจ้าของข้อมูลส่วนบุคคลยื่นคำร้องขอใช้สิทธิดังกล่าวต่อคู่สัญญาฝ่ายหนึ่งฝ่ายใด
เพื่อใช้สิทธิในข้อมูลส่วนบุคคลที่อยู่ในความรับผิดชอบหรือได้รับมาจากอีกฝ่าย คู่สัญญาฝ่ายที่ได้รับคำร้องจะ
ต้องดำเนินการแจ้งและส่งคำร้องดังกล่าวให้แก่คู่สัญญาซึ่งเป็นฝ่ายโอนข้อมูลโดยทันที โดยคู่สัญญาฝ่ายที่รับ
คำร้องนั้นจะต้องแจ้งให้เจ้าของข้อมูลส่วนบุคคลทราบถึงการจัดการตามคำขอหรือข้อร้องเรียนของเจ้าของ
ข้อมูลส่วนบุคคลนั้นด้วย</p>
 <p class='tab2 t-12'>๑๕. หากคู่สัญญาฝ่ายหนึ่งฝ่ายใดมีความจำเป็นจะต้องเปิดเผยข้อมูลส่วนบุคคลที่ได้รับจาก
อีกฝ่ายไปยังต่างประเทศ การส่งออกซึ่งข้อมูลส่วนบุคคลดังกล่าวจะต้องได้รับปกป้องตามมาตรฐานการส่ง
ข้อมูลระหว่างประเทศตามที่กฎหมายคุ้มครองข้อมูลส่วนบุคคลของประเทศที่ส่งข้อมูลไปนั้นกำหนด ทั้งนี้
คู่สัญญาทั้งสองฝ่ายตกลงที่จะเข้าทำสัญญาใด ๆ ที่จำเป็นต่อการปฏิบัติตามกฎหมายที่ใช้บังคับกับการ
โอนข้อมูล</p>
 <p class='tab2 t-12'>๑๖. คู่สัญญาแต่ละฝ่ายอาจใช้ผู้ประมวลผลข้อมูลส่วนบุคคล เพื่อทำการประมวลผลข้อมูล
ส่วนบุคคลที่โอนและรับโอน โดยคู่สัญญาฝ่ายนั้นจะต้องทำสัญญากับผู้ประมวลผลข้อมูลเป็นลายลักษณ์
อักษรซึ่งสัญญาดังกล่าวจะต้องมีเงื่อนไขในการคุ้มครองข้อมูลส่วนบุคคลที่โอนและรับโอนไม่น้อยไปกว่า
เงื่อนไขที่ได้ระบุไว้ในข้อตกลงฉบับนี้ และเงื่อนไขทั้งหมดต้องเป็นไปตามที่กฎหมายคุ้มครองข้อมูลส่วน
บุคคลกำหนด เพื่อหลีกเลี่ยงข้อสงสัย หากคู่สัญญาฝ่ายหนึ่งฝ่ายใดได้ว่าจ้างหรือมอบหมายผู้ประมวลผล
ข้อมูลส่วนบุคคล คู่สัญญาฝ่ายนั้นยังคงต้องมีความรับผิดต่ออีกฝ่ายสำหรับการกระทำการหรือละเว้นกระทำ
การใด ๆ ของผู้ประมวลผลข้อมูลส่วนบุคคลนั้น</p>
 <p class='tab2 t-12'>๑๗. เว้นแต่กฎหมายที่เกี่ยวข้องจะบัญญัติไว้เป็นประการอื่นคู่สัญญาจะทำการลบหรือทำลาย
ข้อมูลส่วนบุคคลที่ตนได้รับจากอีกฝ่ายภายใต้ข้อตกลงฉบับนี้ภายใน {result.RetentionPeriodDays} วัน นับแต่วันที่ดำเนินการประมวล
ผลตามวัตถุประสงค์ภายใต้ข้อตกลงฉบับนี้เสร็จสิ้น หรือวันที่คู่สัญญาได้ตกลงเป็นลายลักษณ์อักษรให้ยกเลิก
สัญญาหลักแล้วแต่กรณีใดจะเกิดขึ้นก่อน</p>
 <p class='tab2 t-12'>๑๘. คู่สัญญาแต่ละฝ่ายจะต้องชดใช้ความเสียหายให้แก่อีกฝ่ายในค่าปรับ ความสูญหายหรือ
เสียหายใด ๆ ที่เกิดขึ้นกับฝ่ายที่ไม่ได้ผิดเงื่อนไข อันเนื่องมาจากการฝ่าฝืนข้อตกลงฉบับนี้ แม้ว่าจะมีข้อจำกัด
ความรับผิดภายใต้สัญญาหลักก็ตาม</p>
 <p class='tab2 t-12'>๑๙. หน้าที่และความรับผิดของคู่สัญญาในการปฏิบัติตามข้อตกลงฉบับนี้จะสิ้นสุดลงนับ
แต่วันที่การดำเนินการตามสัญญาหลักเสร็จสิ้นลง หรือ วันที่คู่สัญญาได้ตกลงเป็นลายลักษณ์อักษรให้ยกเลิก
สัญญาหลักแล้วแต่กรณีใดจะเกิดขึ้นก่อน อย่างไรก็ดี การสิ้นผลลงของข้อตกลงฉบับนี้ ไม่กระทบต่อหน้าที่
ของคู่สัญญาแต่ละฝ่ายในการลบหรือทำลายข้อมูลส่วนบุคคลตามที่ได้กำหนดในข้อ ๑๗ ของข้อตกลงฉบับนี้</p>
 <p class='tab2 t-12'>ในกรณีที่ข้อตกลง คำรับรอง การเจรจา หรือข้อผูกพันใดที่คู่สัญญามีต่อกันไม่ว่าด้วย
วาจาหรือเป็นลายลักษณ์อักษรใดขัดหรือแย้งกับข้อตกลงที่ระบุในข้อตกลงฉบับนี้ ให้ใช้ข้อความตามข้อตกลง
ฉบับนี้บังคับ</p>
 <p class='tab2 t-12'>บันทึกข้อตกลงนี้ทำขึ้นเป็นบันทึกข้อตกลงอิเล็กทรอนิกส์ คู่ตกลงได้อ่าน เข้าใจเงื่อนไข และยอมรับเงื่อนไข และได้ยืนยันว่าเป็นผู้มีอำนาจลงนามในบันทึกข้อตกลง จึงได้ลงลายมืออิเล็กทรอนิกส์พร้อมทั้งประทับตรา (ถ้ามี) ในบันทึกข้อตกลงไว้ และต่างฝ่ายต่างยึดถือไว้ฝ่ายละหนึ่งฉบับในระบบของตน </p>
    <!-- Signature Table -->

</br>
</br>
{signatoryTableHtml}
 <P class='t-12 tab2'>ข้าพเจ้าขอรับรองว่า ทั้งสองฝ่ายได้ลงนามในบันทึกข้อตกลงโดยวิธีการอิเล็กทรอนิกส์ เพื่อแสดงเจตนาของคู่ตกลงแล้ว ข้าพเจ้าจึงได้ลงลายมือชื่ออิเล็กทรอนิกส์รับรองเป็นพยานในบันทึกข้อตกลงพร้อมนี้</P>

{signatoryTableHtmlWitnesses}
</body>
</html>
";
           
            return html;
        }
        catch (Exception ex)
        {
            throw new Exception("Error in WordEContract_DataPersonalService.OnGetWordContact_DataPersonalService_ToPDF: " + ex.Message);
        }
    }
    #endregion
}