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
        // Read CSS file content
        var cssPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "css", "contract.css");
        string contractCss = "";
        if (File.Exists(cssPath))
        {
            contractCss = File.ReadAllText(cssPath, Encoding.UTF8);
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

        var fontPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "font", "THSarabun.ttf");
        string fontBase64 = "";
        if (File.Exists(fontPath))
        {
            var bytes = File.ReadAllBytes(fontPath);
            fontBase64 = Convert.ToBase64String(bytes);
        }
        var strDateTH = CommonDAO.ToThaiDateStringCovert(result.Master_Contract_Sign_Date ?? DateTime.Now);


        #region signlist 

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
             <img src='data:image/jpeg;base64,{contractlogoBase64}' height='80' />
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
{contractCss}
   
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
            {contractLogoHtml}
        </td>
    </tr>
</table>

</br>
    <div class='t-14 text-center'><b>ข้อตกลงการเป็นผู้ควบคุมข้อมูลส่วนบุคคลร่วม</b></div>
    <div class='t-12 text-center'><b>(Joint Controller Agreement)</b></div>
    <div class='t-12 text-center'><b>ระหว่าง</b></div>
    <div class='t-12 text-center'><b>สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม (สสว.)</b></div>
    <div class='t-12 text-center'><b>กับ</b></div>
<div class='t-12 text-center'><b>{result.Contract_Party_Name ?? ""}</b></div>
</br>
   <p class='t-12 tab2'>ข้อตกลงการเป็นผู้ควบคุมข้อมูลส่วนบุคคลร่วม (“{result.Contract_Number ?? "-"}”) ฉบับนี้ ทำขึ้นเมื่อ {strDateTH} ณ สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม</P>
   <p class='t-12 tab2'>โดยที่ สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม ซึ่งต่อไปในข้อตกลงฉบับนี้เรียกว่า “สสว.” ฝ่ายหนึ่ง ได้ตกลงใน {result.MOU_Name ?? ""} ฉบับลงวันที่ {strDateTH} ซึ่งต่อไปในข้อตกลงฉบับนี้เรียกว่า “สัญญาหลัก” กับ  {result.Contract_Party_Name ?? ""}  ซึ่งต่อไปในข้อตกลงฉบับนี้เรียกว่า “{result.Contract_Party_Abb_Name ?? ""}” อีกฝ่ายหนึ่ง รวมทั้งสองฝ่ายว่า “คู่สัญญา”</P>
   <p class='t-12 tab2'>เพื่อให้บรรลุตามวัตถุประสงค์ที่คู่สัญญาได้ตกลงกันภายใต้สัญญาหลัก คู่สัญญามีความจำเป็นต้องร่วมกันเก็บ รวบรวม ใช้ หรือเปิดเผย (รวมเรียกว่า “ประมวลผล”) ข้อมูลส่วนบุคคลตามพระราชบัญญัติคุ้มครองข้อมูลส่วนบุคคล พ.ศ. ๒๕๖๒ โดยที่คู่สัญญามีอำนาจตัดสินใจ กำหนดรูปแบบ รวมถึงวัตถุประสงค์ในการประมวลผลข้อมูลส่วนบุคคลนั้นร่วมกัน ในลักษณะของผู้ควบคุมข้อมูลส่วนบุคคลร่วม</P>
   <p class='t-12 tab2'>คู่สัญญาจึงตกลงจัดทำข้อตกลงฉบับนี้ และให้ถือเป็นส่วนหนึ่งของสัญญาหลัก เพื่อกำหนด
ขอบเขตอำนาจหน้าที่และความรับผิดชอบของคู่สัญญาในการร่วมกันประมวลผลข้อมูลส่วนบุคคล โดยข้อ
ตกลงนี้ใช้บังคับกับกิจกรรมการประมวลผลข้อมูลส่วนบุคคลทั้งสิ้นที่ดำเนินการโดยคู่สัญญา รวมถึงผู้
ประมวลผลข้อมูลส่วนบุคคลซึ่งถูกหรืออาจถูกมอบหมายให้ประมวลผลข้อมูลส่วนบุคคลโดยคู่สัญญา
ทั้งนี้ เพื่อดำเนินการให้เป็นไปตามพระราชบัญญัติคุ้มครองข้อมูลส่วนบุคคล พ.ศ.๒๕๖๒ รวมถึงกฎหมายอื่น ๆ 
ที่ออกตามความของพระราชบัญญัติคุ้มครองข้อมูลส่วนบุคคล พ.ศ.๒๕๖๒ ซึ่งต่อไปในข้อตกลงฉบับนี้เรียกว่า 
“กฎหมายคุ้มครองข้อมูลส่วนบุคคล” ทั้งที่มีผลใช้บังคับอยู่ ณ วันที่ทำข้อตกลงฉบับนี้และที่อาจมีเพิ่มเติม
หรือแก้ไขเปลี่ยนแปลงในภายหลัง โดยมีรายละเอียดดังต่อไปนี้</P>
   <p class='t-12 tab2'><b>ข้อ ๑ วัตถุประสงค์และวิธีการประมวลผล</b></P>
   <p class='t-12 tab2'>คู่สัญญาร่วมกันกำหนดวัตถุประสงค์และวิธีการในการประมวลผลข้อมูลดังรายการกิจกรรม
การประมวลผลข้อมูลส่วนบุคคล (“กิจกรรมการประมวลผลข้อมูลส่วนบุคคลหลัก”) ดังต่อไปนี้ (ระบุวัตถุประสงค์ตามสัญญาหลักที่คู่สัญญาจะต้องดำเนินการร่วมกัน)</P>


<p class='t-12 tab2'>วัตถุประสงค์</P>
{(purplist != null && purplist.Count > 0
    ? string.Join("", purplist.Select(p => $"<p class='tab3 t-12'>{p.Detail}</P>"))
    : "<p class='t-12 tab2'>- ไม่มีข้อมูลวัตถุประสงค์ -</P>")}

   <p class='t-12 tab2'>ซึ่งจากรายการกิจกรรมการประมวลผลหลักที่คู่สัญญาร่วมกันกำหนดวัตถุประสงค์ข้างต้น คู่สัญญาแต่ละฝ่ายมีการประมวลผลข้อมูลส่วนบุคคล (“กิจกรรมการประมวลผลข้อมูลส่วนบุคคลย่อย”) ดังรายละเอียดต่อไปนี้</P>
   <p class='t-12 tab2'><b>(๑) กิจกรรมการประมวลผลข้อมูลส่วนบุคคลย่อยที่ดำเนินการโดย สสว.</b></P>
    <table class='table t-12'>
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
   <p class='t-12 tab2'><b>(๒) กิจกรรมการประมวลผลข้อมูลส่วนบุคคลย่อยซึ่งดำเนินการโดย ({result.Contract_Party_Name ?? ""})</b></P>
    <table class='table t-12 '>
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
   <p class='t-12 tab2'>ทั้งนี้ คู่สัญญาแต่ละฝ่ายรับรองว่าจะดำเนินการประมวลผลข้อมูลส่วนบุคคลดังรายละเอียดข้างต้นให้เป็นไปตามที่กฎหมายคุ้มครองข้อมูลส่วนบุคคลกำหนด โดยเฉพาะอย่างยิ่งในเรื่องความชอบ
ด้วยกฎหมายของการประมวลผลข้อมูลภายใต้ความเป็นผู้ควบคุมข้อมูลส่วนบุคคลร่วม โดยคู่สัญญา
แต่ละฝ่ายจะจัดให้มีและคงไว้ซึ่งมาตรการรักษาความปลอดภัยสำหรับการประมวลผลข้อมูลที่มีความเหมาะสมทั้งในมาตรการเชิงองค์กร มาตรการเชิงเทคนิค และมาตรการเชิงกายภาพ ตามที่คณะกรรมการคุ้มครองข้อมูลส่วนบุคคลได้ประกาศกำหนดและ/หรือตามมาตรฐานสากล โดยคำนึงถึงลักษณะ ขอบเขต และวัตถุประสงค์ของการประมวลผลข้อมูล เพื่อคุ้มครองข้อมูลส่วนบุคคลจากความเสี่ยงอันเกี่ยวเนื่องกับการประมวลผล
ข้อมูลส่วนบุคคล เช่น ความเสียหายอันเกิดจากการละเมิด อุบัติเหตุ การลบ ทำลาย สูญหาย เปลี่ยนแปลง แก้ไข เข้าถึง ใช้ เปิดเผยหรือโอนข้อมูลส่วนบุคคลโดยไม่ชอบด้วยกฎหมาย เป็นต้น</P>
   <p class='t-12 tab2'><b>ข้อ ๒ หน้าที่และความรับผิดชอบของคู่สัญญา</b></P>


<p class='t-12 tab2'>๒.๑ คู่สัญญารับรองว่าจะควบคุมดูแลให้เจ้าหน้าที่ พนักงาน และ/หรือลูกจ้างตัวแทน 
หรือบุคคลใด ๆ ที่ปฏิบัติหน้าที่ในการประมวลผล ข้อมูลส่วนบุคคลภายใต้ข้อตกลงฉบับนี้รักษาความลับ
และปฏิบัติตามกฎหมายคุ้มครองข้อมูลส่วนบุคคลอย่างเคร่งครัด และดำเนินการ ประมวลผล ข้อมูลส่วนบุคคลเพื่อวัตถุประสงค์ตามข้อตกลงฉบับนี้เท่านั้น โดยจะไม่ทำซ้ำ 
คัดลอก ทำสำเนา บันทึกภาพข้อมูลส่วนบุคคลไม่ว่าทั้งหมดหรือแต่บาง ส่วนเป็นอันขาด 
เว้นแต่ เป็นไปตามเงื่อนไข ของสัญญาหลัก หรือกฎหมายที่เกี่ยวข้องจะระบุหรือบัญญัติ ไว้เป็นประการอื่น

</P>
<p class='t-12 tab2'>๒.๒ คู่สัญญารับรองว่าจะกำหนดให้การเข้าถึงข้อมูลส่วนบุคคลภายใต้ข้อตกลงฉบับนี้ 
ถูกจำกัดเฉพาะเจ้าหน้าที่ พนักงาน และ/หรือลูกจ้าง ตัวแทนหรือบุคคลใด ๆ ที่ได้รับมอบหมาย 
มีหน้าที่เกี่ยวข้องหรือมีความจำเป็นในการเข้าถึงข้อมูลส่วนบุคคลภายใต้ข้อตกลงฉบับนี้ เท่านั้น</P>


<p class='t-12 tab2'>๒.๓ คู่สัญญาจะไม่เปิดเผยข้อมูลส่วนบุคคลภายใต้ข้อตกลงนี้แก่บุคคลที่ไม่มีอำนาจหน้าที่ 
เกี่ยวข้องในการประมวลผล หรือบุคคลภายนอก เว้นแต่ กรณีที่มีความจำเป็นต้องกระทำตามหน้าที่ใน 
สัญญาหลัก ของข้อตกลงฉบับนี้ หรือเพื่อปฏิบัติตามกฎหมายที่ใช้บังคับหรือที่ได้รับความยินยอม จากคู่สัญญาอีกฝ่ายก่อน</P>

<p class='t-12 tab2'>๒.๔ คู่สัญญาแต่ละฝ่ายมีหน้าที่ต้องแจ้งรายละเอียดของการประมวลผลข้อมูลส่วนบุคคล 
แก่เจ้าของข้อมูลส่วนบุคคลซึ่งถูกประมวลผลข้อมูลก่อนหรือขณะเก็บรวบรวมข้อมูลส่วนบุคคลทั้งนี้รายการ 
รายละเอียดที่ต้องแจ้งให้เป็นไปตามที่กำหนดในมาตรา ๒๓ แห่งพระราชบัญญัติคุ้มครองข้อมูลส่วนบุคคล พ.ศ. ๒๕๖๒</P>

<p class='t-12 tab2'>๒.๕ กรณีที่คู่สัญญาฝ่ายหนึ่งฝ่ายใด พบพฤติการณ์ที่มีลักษณะที่กระทบต่อการรักษาความ 
ปลอดภัยของข้อมูลส่วนบุคคลที่ประมวลผลภายใต้ข้อตกลงฉบับนี้ ซึ่งอาจก่อให้เกิดความเสียหายจาก 
การละเมิด อุบัติเหตุ การลบ ทำลาย สูญหาย เปลี่ยนแปลง แก้ไข เข้าถึง ใช้เปิดเผยหรือโอนข้อมูล 
ส่วนบุคคลโดยไม่ชอบด้วยกฎหมาย คู่สัญญาฝ่ายที่พบเหตุดังกล่าวจะดำเนินการแจ้งให้คู่สัญญาอีกฝ่ายทราบ 
พร้อมรายละเอียดของเหตุการณ์โดยไม่ชักช้าภายใน ๗๒ ชั่วโมง นับแต่ผู้ประมวลข้อมูลทราบเหตุเท่าที่จะ 
สามารถกระทำได้ทั้งนี้ คู่สัญญาแต่ละฝ่ายต่างมีหน้าที่ต้องแจ้งเหตุดังกล่าวแก่สำนักงานคณะกรรม 
การคุ้มครองข้อมูลส่วนบุคคล หรือเจ้าของข้อมูลส่วนบุคคล ตามแต่กรณีที่กฎหมายคุ้มครองข้อมูลส่วนบุคคลกำหนดไว้</P>

<p class='t-12 tab2'>๒.๖ คู่สัญญาตกลงจะให้ความช่วยเหลืออย่างสมเหตุสมผลแก่อีกฝ่ายในการตอบสนองต่อ 
ข้อเรียกร้องใด ๆ ที่สมเหตุสมผลจากการใช้สิทธิต่างๆ ภายใต้กฎหมายคุ้มครองข้อมูลส่วนบุคคลโดยเจ้าของ
ข้อมูลส่วนบุคคล โดยพิจารณาถึงลักษณะการประมวลผลภาระหน้าที่ภายใต้กฎหมายคุ้มครองข้อมูลที่ 
ใช้บังคับและข้อมูลส่วนบุคคลที่ประมวลผล ทั้งนี้ คู่สัญญาทราบว่าเจ้าของข้อมูลส่วนบุคคลอาจยื่นคำร้องขอ
ใช้สิทธิดังกล่าวต่อคู่สัญญาฝ่ายหนึ่งฝ่ายใดก็ได้ ซึ่งคู่สัญญาฝ่ายที่ได้รับคำร้องจะต้องดำเนินการแจ้งถึงคำร้อง
ดังกล่าวแก่คู่สัญญาอีกฝ่ายโดยทันทีโดยคู่สัญญาฝ่ายที่รับคำร้องนั้นจะต้องแจ้งให้เจ้าของข้อมูลทราบถึงการจัดการ
ตามคำขอหรือข้อร้องเรียนของเจ้าของข้อมูลนั้นด้วย</P>

<p class='t-12 tab2'>๒.๗ ในกรณีที่มีการใช้ผู้ประมวลผลข้อมูลส่วนบุคคลเพื่อทำการประมวลผลข้อมูลส่วน
บุคคลภายใต้ข้อตกลงนี้ให้ดำเนินการแจ้งต่อคู่สัญญาอีกฝ่ายก่อน ทั้งนี้ คู่สัญญาฝ่ายที่ใช้ผู้ประมวลผล 
ข้อมูลส่วนบุคคลจะต้องทำสัญญากับผู้ประมวลผลข้อมูลเป็นลายลักษณ์อักษรตามเงื่อนไขที่กฎหมาย
คุ้มครองข้อมูลกำหนดเพื่อหลีกเลี่ยงข้อสงสัย หากคู่สัญญาฝ่ายหนึ่งฝ่ายใดได้ว่าจ้างหรือมอบหมาย 
ผู้ประมวลผลข้อมูลส่วนบุคคลคู่สัญญาฝ่ายนั้นยังคงต้องมีความรับผิดต่ออีกฝ่ายสำหรับการกระทำการ 
หรือละเว้นกระทำการใด ๆ ของผู้ประมวลผลข้อมูลส่วนบุคคลนั้น</P>

<p class='t-12 tab2'><b>๓ การชดใช้ค่าเสียหาย</b></P>
<p class='t-12 tab2'>๓.๑ คู่สัญญาแต่ละฝ่ายจะต้องชดใช้ความเสียหายให้แก่อีกฝ่ายในค่าปรับ ความสูญหาย
หรือเสียหายใด ๆ ที่เกิดขึ้นกับฝ่ายที่ไม่ได้ผิดเงื่อนไข อันเนื่องมาจากการฝ่าฝืน 
ข้อตกลงฉบับนี้ แม้ว่าจะมีข้อจำกัดความรับผิดภายใต้สัญญาหลักก็ตาม</P>
<p class='t-12 tab2'>๓.๒ ในกรณีที่คู่สัญญาต้องรับผิดร่วมกันในค่าปรับหรือการชดใช้ความเสียหายตามกฎหมายคุ้มครองข้อมูลส่วนบุคคล โดยไม่สามารถพิจารณาเป็นที่ประจักษ์ได้ว่าฝ่ายหนึ่งฝ่ายใดการทำการเป็นเหตุ
ให้เกิดความเสียหายแต่เพียงผู้เดียว หรือจากการถูกศาลหรือหน่วยงานผู้มีอำนาจมีคำพิพากษาหรือคำสั่ง
ถึงที่สุดให้คู่สัญญาร่วมกันรับผิดดังกล่าว คู่สัญญาตกลงกันแบ่งความรับผิดเป็นสัดส่วนดังต่อไปนี้</P>

<p class='t-12 tab2'>(๑) สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม (สสว.) ร้อยละ ๕๐</P>
<p class='t-12 tab2'>(๒) {result.Contract_Party_Name} ร้อยละ ๕๐</P>
<p class='t-12 tab2'>ทั้งนี้การตกลงกันของคู่สัญญานี้ ไม่มีอำนาจเหนือไปกว่าคำพิพากษาหรือคำสั่งถึงที่สุดของ
ศาลหรือหน่วยงานผู้มีอำนาจที่กำหนดให้คู่สัญญาหรือคู่สัญญาฝ่ายหนึ่งฝ่ายใดต้องถูกปรับหรือชดใช้ค่าเสียหาย</P>

<p class='t-12 tab2'><b>๔ ระยะเวลาตามข้อตกลง</b></P>
<p class='t-12 tab2'>หน้าที่และความรับผิดของคู่สัญญาในการปฏิบัติตามข้อตกลงฉบับนี้จะสิ้นสุดลงนับแต่วันที่
การดำเนินการตามสัญญาหลักเสร็จสิ้นลง หรือ วันที่คู่สัญญาได้ตกลงเป็นลายลักษณ์อักษรให้ยกเลิกสัญญา
หลัก แล้วแต่กรณีใดจะเกิดขึ้นก่อน</P>


    <p class='t-12 tab2'><b>๕ ผู้แทนของคู่สัญญาแต่ละฝ่าย</b></P>
   <p class='t-12 tab2'>คู่สัญญาตกลงแต่งตั้งผู้แทนของแต่ละฝ่าย ดังรายการต่อไปนี้</P>
   <p class='t-12 tab2'><b>(๑) สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม (สสว.)</b></P>
   <p class='t-12 tab3'><b>ผู้แทน</b> : {result.OSMEP_ContRep ?? ""}</P>
   <p class='t-12 tab3'><b>ติดต่อได้ที่</b> : {result.OSMEP_ContRep_Contact ?? ""}</P>
   <p class='t-12 tab3'><b>เจ้าหน้าที่คุ้มครองข้อมูลส่วนบุคคล</b> : {result.OSMEP_DPO ?? ""}</P>
   <p class='t-12 tab3'><b>ติดต่อได้ที่ </b>: {result.OSMEP_DPO_Contact ?? ""}</P>
   <p class='t-12 tab2'><b>(๒) {result.Contract_Party_Name ?? ""}</b></P>
   <p class='t-12 tab3'><b>ผู้แทน</b> : {result.CP_ContRep ?? ""}</P>
   <p class='t-12 tab3'><b>ติดต่อได้ที่</b> : {result.CP_ContRep_Contact ?? ""}</P>
   <p class='t-12 tab3'><b>เจ้าหน้าที่คุ้มครองข้อมูลส่วนบุคคล (ถ้ามี)</b> : {result.CP_DPO ?? ""}</P>
   <p class='t-12 tab3'><b>ติดต่อได้ที่</b> : {result.CP_DPO_Contact ?? ""}</P>

<p class='t-12 tab2'><b>๖.การบังคับใช้</b></P>
<p class='t-12 tab2'>ในกรณีที่ข้อตกลง คำรับรอง การเจรจาหรือข้อผูกพันใดที่คู่สัญญามีต่อกันไม่ว่าด้วย 
วาจาหรือเป็นลายลักษณ์อักษรก็ดี ขัดหรือแย้งกับข้อความที่ระบุในข้อตกลงฉบับนี้ ให้ใช้ข้อความตาม 
ข้อตกลงฉบับนี้บังคับ</P>
<p class='t-12 tab2'>บันทึกข้อตกลงนี้ทำขึ้นเป็นบันทึกข้อตกลงอิเล็กทรอนิกส์ คู่ตกลงได้อ่าน เข้าใจเงื่อนไข และยอมรับเงื่อนไข และได้ยืนยันว่าเป็นผู้มีอำนาจลงนามในบันทึกข้อตกลง จึงได้ลงลายมืออิเล็กทรอนิกส์พร้อมทั้งประทับตรา (ถ้ามี) ในบันทึกข้อตกลงไว้ และต่างฝ่ายต่างยึดถือไว้ฝ่ายละหนึ่งฉบับในระบบของตน </P>

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

    #endregion 4.1.1.2.5.บันทึกข้อตกลงการเป็นผู้ควบคุมข้อมูลส่วนบุคคลร่วมตัวอย่างหน้าจอ
}
