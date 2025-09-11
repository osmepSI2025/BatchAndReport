using BatchAndReport.DAO;
using BatchAndReport.Models;
using DinkToPdf.Contracts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Spire.Doc.Documents;
using System.Text;
using System.Threading.Tasks;

public class WordEContract_ControlDataService
{
    private readonly WordServiceSetting _w;
    private readonly E_ContractReportDAO _eContractReportDAO;
    private readonly IConverter _pdfConverter;
    public WordEContract_ControlDataService(WordServiceSetting ws
          , E_ContractReportDAO eContractReportDAO
         , IConverter pdfConverter
        )
    {
        _w = ws;
        _eContractReportDAO = eContractReportDAO;
        _pdfConverter = pdfConverter;
    }
    #region 4.1.1.2.5.บันทึกข้อตกลงการเป็นผู้ควบคุมข้อมูลส่วนบุคคลร่วมตัวอย่างหน้าจอ
  
    public async Task<string> OnGetWordContact_ControlDataServiceHtmlToPdf(string id,string typeContact)
    {
        var result = await _eContractReportDAO.GetJDCAAsync(id);
        if (result == null)
        {
            throw new Exception("ไม่พบข้อมูลสัญญา");
        }
        // Logo
        var logoPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "images", "logo_SME.jpg");
        string logoBase64 = "";
        if (System.IO.File.Exists(logoPath))
        {
            var bytes = System.IO.File.ReadAllBytes(logoPath);
            logoBase64 = Convert.ToBase64String(bytes);
        }
        var purplist = await _eContractReportDAO.GetJDCA_JointPurpAsync(id);
        var dtActivitySME = await _eContractReportDAO.GetJDCA_SubProcessActivitiesAsync(id);

        var activityListOSMEP = dtActivitySME?.Where(x => x.Owner == "OSMEP").ToList() ?? new List<E_ConReport_JDCA_SubProcessActivitiesModels>();
        var activityListCP = dtActivitySME?.Where(x => x.Owner == "CP").ToList() ?? new List<E_ConReport_JDCA_SubProcessActivitiesModels>();
        var fontPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "font", "THSarabunNew.ttf").Replace("\\", "/");
        var strDateTH = CommonDAO.ToThaiDateStringCovert(result.Master_Contract_Sign_Date ?? DateTime.Now);


        #region signlist 

        var signlist = await _eContractReportDAO.GetSignNameAsync(id, typeContact);
        var signatoryTableHtml = "";
        if (signlist.Count > 0)
        {
            signatoryTableHtml = await _eContractReportDAO.RenderSignatory(signlist);

        }
        #endregion signlist


        var html = $@"
<html>
<head>
    <meta charset='utf-8'>
     <style>
        @font-face {{
            font-family: 'THSarabunNew';
            src: url('file:///{fontPath}') format('truetype');
            font-weight: normal;
            font-style: normal;
        }}
        body, p, div, table, th, td {{
            font-family: 'THSarabunNew', Arial, sans-serif !important;
            font-size: 22px;
        }}
        /* แก้การตัดคำไทย: ไม่หั่นกลางคำ, ตัดเมื่อจำเป็น */
        body, p, div {{
            word-break: keep-all;
            overflow-wrap: break-word;
            -webkit-line-break: after-white-space;
            hyphens: none;
        }}
        .t-16 {{
            font-size: 1.5em;
        }}
        .t-18 {{
            font-size: 1.7em;
        }}
        .t-22 {{
            font-size: 1.9em;
        }}
        .tab1 {{ text-indent: 48px; }}
        .tab2 {{ text-indent: 96px; }}
        .tab3 {{ text-indent: 144px; }}
        .tab4 {{ text-indent: 192px; }}
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
        .table {{
            width: 100%;
            border-collapse: collapse;
            margin-top: 20px;
            font-size: 28pt;
        }}
        .table th, .table td {{
            border: 1px solid #000;
            padding: 8px;
        }}
        .sign-double {{ display: flex; }}
        .text-center-right-brake {{
            margin-left: 50%;
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
            font-size: 1.1em;
            font-family: 'THSarabunNew', Arial, sans-serif !important;
        }}
        .logo-table {{
            width: 100%;
            border-collapse: collapse;
            margin-top: 40px;
        }}
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
        <td style='width:40%; text-align:center; vertical-align:top;'>
            <div style='display:inline-block; padding:20px; font-size:32pt;'>
             <img src='data:image/jpeg;base64,{logoBase64}' width='240' height='80' />
            </div>
        </td>
    </tr>
</table>

</br>
    <div class='t-22 text-center'><b>ข้อตกลงการเป็นผู้ควบคุมข้อมูลส่วนบุคคลร่วม</b></div>
    <div class='t-18 text-center'><b>(Joint Controller Agreement)</b></div>
    <div class='t-18 text-center'><b>ระหว่าง</b></div>
    <div class='t-18 text-center'><b>สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม (สสว.)</b></div>
    <div class='t-18 text-center'><b>กับ {result.Contract_Party_Name ?? ""}</b></div>
</br>
   <p class='t-16 tab3'>ข้อตกลงการเป็นผู้ควบคุมข้อมูลส่วนบุคคลร่วม (“ข้อตกลง”) ฉบับนี้ ทำขึ้นเมื่อ {strDateTH} ณ สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม</P>
   <p class='t-16 tab3'>โดยที่ สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม ซึ่งต่อไปในข้อตกลงฉบับนี้เรียกว่า “สสว.” ฝ่ายหนึ่ง ได้ตกลงใน {result.MOU_Name ?? ""} ฉบับลงวันที่ {strDateTH} ซึ่งต่อไปในข้อตกลงฉบับนี้เรียกว่า “สัญญาหลัก” กับ  {result.Contract_Party_Name ?? ""}  ซึ่งต่อไปในข้อตกลงฉบับนี้เรียกว่า “(ชื่อเรียกคู่สัญญา)” อีกฝ่ายหนึ่ง รวมทั้งสองฝ่ายว่า “คู่สัญญา”</P>
   <p class='t-16 tab3'>เพื่อให้บรรลุตามวัตถุประสงค์ที่คู่สัญญาได้ตกลงกันภายใต้สัญญาหลัก คู่สัญญามีความจำเป็นต้องร่วมกันเก็บ รวบรวม ใช้ หรือเปิดเผย (รวมเรียกว่า “ประมวลผล”) ข้อมูลส่วนบุคคลตามพระราชบัญญัติคุ้มครองข้อมูลส่วนบุคคล พ.ศ. ๒๕๖๒ โดยที่คู่สัญญามีอำนาจตัดสินใจ กำหนดรูปแบบ รวมถึงวัตถุประสงค์ในการประมวลผลข้อมูลส่วนบุคคลนั้นร่วมกัน ในลักษณะของผู้ควบคุมข้อมูลส่วนบุคคลร่วม</P>
   <p class='t-16 tab3'>คู่สัญญาจึงตกลงจัดทำข้อตกลงฉบับนี้ และให้ถือเป็นส่วนหนึ่งของสัญญาหลัก เพื่อกำหนด
</br>ขอบเขตอำนาจหน้าที่และความรับผิดชอบของคู่สัญญาในการร่วมกันประมวลผลข้อมูลส่วนบุคคล โดยข้อ
</br>ตกลงนี้ใช้บังคับกับกิจกรรมการประมวลผลข้อมูลส่วนบุคคลทั้งสิ้นที่ดำเนินการโดยคู่สัญญา รวมถึงผู้
</br>ประมวลผลข้อมูลส่วนบุคคลซึ่งถูกหรืออาจถูกมอบหมายให้ประมวลผลข้อมูลส่วนบุคคลโดยคู่สัญญา
ทั้งนี้ เพื่อดำเนินการให้เป็นไปตามพระราชบัญญัติคุ้มครองข้อมูลส่วนบุคคล พ.ศ.๒๕๖๒ รวมถึงกฎหมายอื่น ๆ 
ที่ออกตามความของพระราชบัญญัติคุ้มครองข้อมูลส่วนบุคคล พ.ศ.๒๕๖๒ ซึ่งต่อไปในข้อตกลงฉบับนี้เรียกว่า 
“กฎหมายคุ้มครองข้อมูลส่วนบุคคล” ทั้งที่มีผลใช้บังคับอยู่ ณ วันที่ทำข้อตกลงฉบับนี้และที่อาจมีเพิ่มเติม
</br>หรือแก้ไขเปลี่ยนแปลงในภายหลัง โดยมีรายละเอียดดังต่อไปนี้</P>
   <p class='t-16 tab3'><b>ข้อ ๑ วัตถุประสงค์และวิธีการประมวลผล</b></P>
   <p class='t-16 tab3'>คู่สัญญาร่วมกันกำหนดวัตถุประสงค์และวิธีการในการประมวลผลข้อมูลดังรายการกิจกรรม
</br>การประมวลผลข้อมูลส่วนบุคคล (“กิจกรรมการประมวลผลข้อมูลส่วนบุคคลหลัก”) ดังต่อไปนี้ (ระบุวัตถุประสงค์ตามสัญญาหลักที่คู่สัญญาจะต้องดำเนินการร่วมกัน)</P>


<p class='t-16 tab3'>วัตถุประสงค์</P>
{(purplist != null && purplist.Count > 0
    ? string.Join("", purplist.Select(p => $"<p class='tab4 t-16'>{p.Objective_Description}</P>"))
    : "<p class='t-16 tab3'>- ไม่มีข้อมูลวัตถุประสงค์ -</P>")}

   <p class='t-16 tab3'>ซึ่งจากรายการกิจกรรมการประมวลผลหลักที่คู่สัญญาร่วมกันกำหนดวัตถุประสงค์ข้างต้น คู่สัญญาแต่ละฝ่ายมีการประมวลผลข้อมูลส่วนบุคคล (“กิจกรรมการประมวลผลข้อมูลส่วนบุคคลย่อย”) ดังรายละเอียดต่อไปนี้</P>
   <p class='t-16 tab3'><b>(๑) กิจกรรมการประมวลผลข้อมูลส่วนบุคคลย่อยที่ดำเนินการโดย สสว.</b></P>
    <table class='table t-16'>
        <tr>
            <th>รายการกิจกรรมการประมวลผล</th>
            <th>ฐานกฎหมายที่ใช้ในการประมวลผล</th>
            <th>รายการข้อมูลส่วนบุคคลที่ใช้ประมวลผล</th>
        </tr>
       {(activityListOSMEP != null && activityListOSMEP.Count > 0
    ? string.Join("", activityListOSMEP.Select(x => $@"
        <tr>
            <td>{x.Activity}</td>
            <td>{x.LegalBasis}</td>
            <td>{x.PersonalData}</td>
        </tr>"))
    : @"<tr><td colspan='3'>ไม่พบข้อมูลกิจกรรมของ สสว.</td></tr>")}
    </table>
   <p class='t-16 tab3'><b>(๒) กิจกรรมการประมวลผลข้อมูลส่วนบุคคลย่อยซึ่งดำเนินการโดย ({result.Contract_Party_Name ?? ""})</b></P>
    <table class='table t-16 '>
        <tr>
            <th>รายการกิจกรรมการประมวลผล</th>
            <th>ฐานกฎหมายที่ใช้ในการประมวลผล</th>
            <th>รายการข้อมูลส่วนบุคคลที่ใช้ประมวลผล</th>
        </tr>
        {string.Join("", activityListCP.Select(x => $@"
        <tr>
            <td>{x.Activity}</td>
            <td>{x.LegalBasis}</td>
            <td>{x.PersonalData}</td>
        </tr>"))}
    </table>
    <!-- Add more sections as needed, following your Word structure -->
   <p class='t-16 tab3'><b>ข้อ ๒ หน้าที่และความรับผิดชอบของคู่สัญญา</b></P>


<p class='t-16 tab3'>๒.๑ คู่สัญญารับรองว่าจะควบคุมดูแลให้เจ้าหน้าที่ พนักงาน และ/หรือลูกจ้างตัวแทน 
หรือบุคคลใด ๆ ที่ปฏิบัติหน้าที่ในการประมวลผล ข้อมูลส่วนบุคคลภายใต้ข้อตกลงฉบับนี้รักษาความลับ
และปฏิบัติตามกฎหมายคุ้มครองข้อมูลส่วนบุคคลอย่างเคร่งครัด และดำเนินการ ประมวลผล ข้อมูลส่วนบุคคลเพื่อวัตถุประสงค์ตามข้อตกลงฉบับนี้เท่านั้น โดยจะไม่ทำซ้ำ 
คัดลอก ทำสำเนา บันทึกภาพข้อมูลส่วนบุคคลไม่ว่าทั้งหมดหรือแต่บาง ส่วนเป็นอันขาด 
เว้นแต่ เป็นไปตามเงื่อนไข ของสัญญาหลัก หรือกฎหมายที่เกี่ยวข้องจะระบุหรือบัญญัติ ไว้เป็นประการอื่น

</P>
<p class='t-16 tab3'>๒.๒ คู่สัญญารับรองว่าจะกำหนดให้การเข้าถึงข้อมูลส่วนบุคคลภายใต้ข้อตกลงฉบับนี้ 
ถูกจำกัดเฉพาะเจ้าหน้าที่ พนักงาน และ/หรือลูกจ้าง ตัวแทนหรือบุคคลใด ๆ ที่ได้รับมอบหมาย 
มีหน้าที่เกี่ยวข้องหรือมีความจำเป็นในการเข้าถึงข้อมูลส่วนบุคคลภายใต้ข้อตกลงฉบับนี้ เท่านั้น</P>


<p class='t-16 tab3'>๒.๓ คู่สัญญาจะไม่เปิดเผยข้อมูลส่วนบุคคลภายใต้ข้อตกลงนี้แก่บุคคลที่ไม่มีอำนาจหน้าที่ 
เกี่ยวข้องในการประมวลผล หรือบุคคลภายนอก เว้นแต่ กรณีที่มีความจำเป็นต้องกระทำตามหน้าที่ใน 
สัญญาหลัก ของข้อตกลงฉบับนี้ หรือเพื่อปฏิบัติตามกฎหมายที่ใช้บังคับหรือที่ได้รับความยินยอม จากคู่สัญญาอีกฝ่ายก่อน</P>

<p class='t-16 tab3'>๒.๔ คู่สัญญาแต่ละฝ่ายมีหน้าที่ต้องแจ้งรายละเอียดของการประมวลผลข้อมูลส่วนบุคคล 
แก่เจ้าของข้อมูลส่วนบุคคลซึ่งถูกประมวลผลข้อมูลก่อนหรือขณะเก็บรวบรวมข้อมูลส่วนบุคคลทั้งนี้รายการ 
รายละเอียดที่ต้องแจ้งให้เป็นไปตามที่กำหนดในมาตรา ๒๓ แห่งพระราชบัญญัติคุ้มครองข้อมูลส่วนบุคคล พ.ศ. ๒๕๖๒</P>

<p class='t-16 tab3'>๒.๕ กรณีที่คู่สัญญาฝ่ายหนึ่งฝ่ายใด พบพฤติการณ์ที่มีลักษณะที่กระทบต่อการรักษาความ 
ปลอดภัยของข้อมูลส่วนบุคคลที่ประมวลผลภายใต้ข้อตกลงฉบับนี้ ซึ่งอาจก่อให้เกิดความเสียหายจาก 
การละเมิด อุบัติเหตุ การลบ ทำลาย สูญหาย เปลี่ยนแปลง แก้ไข เข้าถึง ใช้เปิดเผยหรือโอนข้อมูล 
ส่วนบุคคลโดยไม่ชอบด้วยกฎหมาย คู่สัญญาฝ่ายที่พบเหตุดังกล่าวจะดำเนินการแจ้งให้คู่สัญญาอีกฝ่ายทราบ 
พร้อมรายละเอียดของเหตุการณ์โดยไม่ชักช้าภายใน ๗๒ ชั่วโมง นับแต่ผู้ประมวลข้อมูลทราบเหตุเท่าที่จะ 
สามารถกระทำได้ทั้งนี้ คู่สัญญาแต่ละฝ่ายต่างมีหน้าที่ต้องแจ้งเหตุดังกล่าวแก่สำนักงานคณะกรรม 
การคุ้มครองข้อมูลส่วนบุคคล หรือเจ้าของข้อมูลส่วนบุคคล ตามแต่กรณีที่กฎหมายคุ้มครองข้อมูลส่วนบุคคลกำหนดไว้</P>

<p class='t-16 tab3'>๒.๖ คู่สัญญาตกลงจะให้ความช่วยเหลืออย่างสมเหตุสมผลแก่อีกฝ่ายในการตอบสนองต่อ 
ข้อเรียกร้องใด ๆ ที่สมเหตุสมผลจากการใช้สิทธิต่างๆ ภายใต้กฎหมายคุ้มครองข้อมูลส่วนบุคคลโดยเจ้าของ
ข้อมูลส่วนบุคคล โดยพิจารณาถึงลักษณะการประมวลผลภาระหน้าที่ภายใต้กฎหมายคุ้มครองข้อมูลที่ 
ใช้บังคับและข้อมูลส่วนบุคคลที่ประมวลผล ทั้งนี้ คู่สัญญาทราบว่าเจ้าของข้อมูลส่วนบุคคลอาจยื่นคำร้องขอ
ใช้สิทธิดังกล่าวต่อคู่สัญญาฝ่ายหนึ่งฝ่ายใดก็ได้ ซึ่งคู่สัญญาฝ่ายที่ได้รับคำร้องจะต้องดำเนินการแจ้งถึงคำร้อง
ดังกล่าวแก่คู่สัญญาอีกฝ่ายโดยทันทีโดยคู่สัญญาฝ่ายที่รับคำร้องนั้นจะต้องแจ้งให้เจ้าของข้อมูลทราบถึงการจัดการ
ตามคำขอหรือข้อร้องเรียนของเจ้าของข้อมูลนั้นด้วย</P>

<p class='t-16 tab3'>๒.๗ ในกรณีที่มีการใช้ผู้ประมวลผลข้อมูลส่วนบุคคลเพื่อทำการประมวลผลข้อมูลส่วน
บุคคลภายใต้ข้อตกลงนี้ให้ดำเนินการแจ้งต่อคู่สัญญาอีกฝ่ายก่อน ทั้งนี้ คู่สัญญาฝ่ายที่ใช้ผู้ประมวลผล 
ข้อมูลส่วนบุคคลจะต้องทำสัญญากับผู้ประมวลผลข้อมูลเป็นลายลักษณ์อักษรตามเงื่อนไขที่กฎหมาย
คุ้มครองข้อมูลกำหนดเพื่อหลีกเลี่ยงข้อสงสัย หากคู่สัญญาฝ่ายหนึ่งฝ่ายใดได้ว่าจ้างหรือมอบหมาย 
ผู้ประมวลผลข้อมูลส่วนบุคคลคู่สัญญาฝ่ายนั้นยังคงต้องมีความรับผิดต่ออีกฝ่ายสำหรับการกระทำการ 
หรือละเว้นกระทำการใด ๆ ของผู้ประมวลผลข้อมูลส่วนบุคคลนั้น</P>

<p class='t-16 tab3'><b>๓ การชดใช้ค่าเสียหาย</b></P>
<p class='t-16 tab3'>๓.๑ คู่สัญญาแต่ละฝ่ายจะต้องชดใช้ความเสียหายให้แก่อีกฝ่ายในค่าปรับ ความสูญหาย
หรือเสียหายใด ๆ ที่เกิดขึ้นกับฝ่ายที่ไม่ได้ผิดเงื่อนไข อันเนื่องมาจากการฝ่าฝืน 
ข้อตกลงฉบับนี้ แม้ว่าจะมีข้อจำกัดความรับผิดภายใต้สัญญาหลักก็ตาม</P>

<p class='t-16 tab3'>(๑) สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม (สสว.) ร้อยละ ๕๐</P>
<p class='t-16 tab3'>ทั้งนี้การตกลงกันของคู่สัญญานี้ ไม่มีอำนาจเหนือไปกว่าคำพิพากษาหรือคำสั่งถึงที่สุดของ
ศาลหรือหน่วยงานผู้มีอำนาจที่กำหนดให้คู่สัญญาหรือคู่สัญญาฝ่ายหนึ่งฝ่ายใดต้องถูกปรับหรือชดใช้ค่าเสียหาย</P>

<p class='t-16 tab3'><b>๔ ระยะเวลาตามข้อตกลง</b></P>
<p class='t-16 tab3'>หน้าที่และความรับผิดของคู่สัญญาในการปฏิบัติตามข้อตกลงฉบับนี้จะสิ้นสุดลงนับแต่วันที่
การดำเนินการตามสัญญาหลักเสร็จสิ้นลง หรือ วันที่คู่สัญญาได้ตกลงเป็นลายลักษณ์อักษรให้ยกเลิกสัญญา
หลัก แล้วแต่กรณีใดจะเกิดขึ้นก่อน</P>


    <p class='t-16 tab3'><b>๕ ผู้แทนของคู่สัญญาแต่ละฝ่าย</b></P>
   <p class='t-16 tab3'><b>(๑) สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม (สสว.)</b></P>
   <p class='t-16 tab3'><b>ผู้แทน</b> : {result.OSMEP_ContRep ?? ""}</P>
   <p class='t-16 tab3'><b>ติดต่อได้ที่</b> : {result.OSMEP_ContRep_Contact ?? ""}</P>
   <p class='t-16 tab3'><b>เจ้าหน้าที่คุ้มครองข้อมูลส่วนบุคคล</b> : {result.OSMEP_DPO ?? ""}</P>
   <p class='t-16 tab3'><b>ติดต่อได้ที่ </b>: {result.OSMEP_DPO_Contact ?? ""}</P>
   <p class='t-16 tab3'><b>(๒) {result.Contract_Party_Name ?? ""}</b></P>
   <p class='t-16 tab3'><b>ผู้แทน</b> : {result.CP_ContRep ?? ""}</P>
   <p class='t-16 tab3'><b>ติดต่อได้ที่</b> : {result.CP_ContRep_Contact ?? ""}</P>
   <p class='t-16 tab3'><b>เจ้าหน้าที่คุ้มครองข้อมูลส่วนบุคคล (ถ้ามี)</b> : {result.CP_DPO ?? ""}</P>
   <p class='t-16 tab3'><b>ติดต่อได้ที่</b> : {result.CP_DPO_Contact ?? ""}</P>

<p class='t-16 tab2'><b>๖.การบังคับใช้</b></P>
<p class='t-16 tab3'>ในกรณีที่ข้อตกลง คำรับรอง การเจรจาหรือข้อผูกพันใดที่คู่สัญญามีต่อกันไม่ว่าด้วย 
วาจาหรือเป็นลายลักษณ์อักษรก็ดี ขัดหรือแย้งกับข้อความที่ระบุในข้อตกลงฉบับนี้ ให้ใช้ข้อความตาม 
ข้อตกลงฉบับนี้บังคับ</P>
<p class='t-16 tab3'>ข้อตกลงฉบับนี้ทำขึ้นเป็นสองฉบับ มีข้อความถูกต้องตรงกัน คู่สัญญาทั้งสองฝ่าย 
ได้อ่าน และเข้าใจข้อความในข้อตกลงโดยละเอียดตลอดแล้ว เห็นว่าตรงตามเจตนารมณ์ทุกประการ เพื่อเป็นหลักฐาน
แห่งการนี้ทั้งสองฝ่ายจึงได้ลงลายมือชื่อพร้อมทั้งประทับตราสำคัญผูกพันนิติบุคคล (ถ้ามี) ไว้เป็นหลักฐาน ณ วัน เดือน ปี ที่ระบุข้างต้น และคู่สัญญาต่างยึดถือไว้ฝ่ายละหนึ่งฉบับ</P>

</br>
</br>
{signatoryTableHtml}

</body>
</html>
";

        return html;
    }

    #endregion 4.1.1.2.5.บันทึกข้อตกลงการเป็นผู้ควบคุมข้อมูลส่วนบุคคลร่วมตัวอย่างหน้าจอ
}
