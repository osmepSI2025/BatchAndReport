using BatchAndReport.DAO;
using DinkToPdf.Contracts;
using System.Text;
public class WordEContract_AMJOAService
{
    private readonly WordServiceSetting _w;
    private readonly E_ContractReportDAO _eContractReportDAO;
    private readonly IConverter _pdfConverter; // เพิ่ม DI สำหรับ PDF Converter

    public WordEContract_AMJOAService(
        WordServiceSetting ws,
        E_ContractReportDAO eContractReportDAO
      , IConverter pdfConverter
    )
    {
        _w = ws;
        _eContractReportDAO = eContractReportDAO;
        _pdfConverter = pdfConverter;
    }


    public async Task<string> OnGetWordContact_AMJOAServiceHtmlToPDF(string conId)
    {
        var dataResult = await _eContractReportDAO.GetJOAAsync(conId);
        if (dataResult == null)
            throw new Exception("JOA data not found.");
        var fontPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "font", "THSarabunNew.ttf").Replace("\\", "/");
        var cssPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "css", "contract.css").Replace("\\", "/");
        var logoPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "images", "logo_SME.jpg");
        string logoBase64 = "";
        if (System.IO.File.Exists(logoPath))
        {
            var bytes = System.IO.File.ReadAllBytes(logoPath);
            logoBase64 = Convert.ToBase64String(bytes);
        }
        string contractLogoHtml;
        if (!string.IsNullOrEmpty(dataResult.Organization_Logo) && dataResult.Organization_Logo.Contains("<content>"))
        {
            try
            {
                // ตัดเอาเฉพาะ Base64 ในแท็ก <content>...</content>
                var contentStart = dataResult.Organization_Logo.IndexOf("<content>") + "<content>".Length;
                var contentEnd = dataResult.Organization_Logo.IndexOf("</content>");
                var contractlogoBase64 = dataResult.Organization_Logo.Substring(contentStart, contentEnd - contentStart);

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
        #region checkมอบอำนาจ
        string strAttorneyLetterDate = CommonDAO.ToArabicDateStringCovert(dataResult.Grant_Date ?? DateTime.Now);
        string strAttorney = "";
        var HtmlAttorney = new StringBuilder();
        if (dataResult.AttorneyFlag == true)
        {
            strAttorney = "ผู้มีอำนาจ กระทำการแทน ปรากฏตามเอกสารแต่งตั้ง และ/หรือ มอบอำนาจ ฉบับลงวันที่ " + strAttorneyLetterDate + "";

        }
        else
        {
            strAttorney = "";
        }
        #endregion

        // data mock 6. ตัวชี้วัดความสำเร็จของโครงการ
        var mockIndicators = GenerateMockData();

        // Build the HTML table for section 6 using mockIndicators
        var indicatorTable = new StringBuilder();
        indicatorTable.AppendLine("<table style='width:100%; border-collapse:collapse; margin-top:10px; font-size:1.2em;'>");
        indicatorTable.AppendLine("<tr>");
        indicatorTable.AppendLine("<th style='border:1px solid #000;'>ผลผลิต</th>");
        indicatorTable.AppendLine("<th style='border:1px solid #000;'>หน่วยนับ</th>");
        indicatorTable.AppendLine("<th style='border:1px solid #000;'>เป้าหมาย</th>");
        indicatorTable.AppendLine("<th style='border:1px solid #000;'>วิธีการวัด</th>");
        indicatorTable.AppendLine("</tr>");

        // Outputs
        for (int i = 0; i < mockIndicators.Outputs.Count; i++)
        {
            var o = mockIndicators.Outputs[i];
            indicatorTable.AppendLine("<tr>");
            indicatorTable.AppendLine($"<td style='border:1px solid #000;'>{i + 1}. {o.Description}</td>");
            indicatorTable.AppendLine($"<td style='border:1px solid #000;'>{o.UnitOfMeasurement}</td>");
            indicatorTable.AppendLine($"<td style='border:1px solid #000;'>{o.Target}</td>");
            indicatorTable.AppendLine($"<td style='border:1px solid #000;'>{o.MeasurementMethod}</td>");
            indicatorTable.AppendLine("</tr>");
        }

        // Outcomes header
        indicatorTable.AppendLine("<tr>");

        indicatorTable.AppendLine("<th style='border:1px solid #000;'>ผลลัพธ์</th>");
        indicatorTable.AppendLine("<th style='border:1px solid #000;'>หน่วยนับ</th>");
        indicatorTable.AppendLine("<th style='border:1px solid #000;'>เป้าหมาย</th>");
        indicatorTable.AppendLine("<th style='border:1px solid #000;'>วิธีการวัด</th>");
        indicatorTable.AppendLine("</tr>");

        // Outcomes
        for (int i = 0; i < mockIndicators.Outcomes.Count; i++)
        {
            var o = mockIndicators.Outcomes[i];
            indicatorTable.AppendLine("<tr>");
            indicatorTable.AppendLine($"<td style='border:1px solid #000;'>{i + 1}. {o.Description}</td>");
            indicatorTable.AppendLine($"<td style='border:1px solid #000;'>{o.UnitOfMeasurement}</td>");
            indicatorTable.AppendLine($"<td style='border:1px solid #000;'>{o.Target}</td>");
            indicatorTable.AppendLine($"<td style='border:1px solid #000;'>{o.MeasurementMethod}</td>");
            indicatorTable.AppendLine("</tr>");
        }
        indicatorTable.AppendLine("</table>");

        var strDateTH = CommonDAO.ToThaiDateString(dataResult.Contract_SignDate ?? DateTime.Now);
        var purposeList = await _eContractReportDAO.GetJOAPoposeAsync(conId);

        var signatoryHtml = new StringBuilder();
        var companySealHtml = new StringBuilder();
        bool sealAdded = false; // กันซ้ำ

        foreach (var signer in dataResult.Signatories)
        {
            string signatureHtml;
            string companySeal = ""; // กัน warning

            // ► ลายเซ็นรายบุคคล (เดิม)
            if (!string.IsNullOrEmpty(signer?.DS_FILE) && signer.DS_FILE.Contains("<content>"))
            {
                try
                {
                    var contentStart = signer.DS_FILE.IndexOf("<content>") + "<content>".Length;
                    var contentEnd = signer.DS_FILE.IndexOf("</content>");
                    var base64 = signer.DS_FILE.Substring(contentStart, contentEnd - contentStart);

                    signatureHtml = $@"<div class='t-16 text-center tab1'>
    <img src='data:image/png;base64,{base64}' alt='signature' style='max-height: 80px;' />
</div>";
                }
                catch
                {
                    signatureHtml = "<div class='t-16 text-center tab1'>(ลงชื่อ....................)</div>";
                }
            }
            else
            {
                signatureHtml = "<div class='t-16 text-center tab1'>(ลงชื่อ....................)</div>";
            }
            // ► ตราประทับ: ให้พิจารณาเมื่อเจอ CP_S เท่านั้น (ไม่เช็ค null/empty ตรง if ชั้นนอก)
            if (!sealAdded && signer?.Signatory_Type == "CP_S")
            {
                if (!string.IsNullOrEmpty(signer.Company_Seal) && signer.Company_Seal.Contains("<content>"))
                {
                    try
                    {
                        var contentStart = signer.Company_Seal.IndexOf("<content>") + "<content>".Length;
                        var contentEnd = signer.Company_Seal.IndexOf("</content>");
                        var base64 = signer.Company_Seal.Substring(contentStart, contentEnd - contentStart);

                        companySeal = $@"
<div class='t-16 text-center tab1'>
    <img src='data:image/png;base64,{base64}' alt='signature' style='max-height: 80px;' />
</div>";

                        companySealHtml.AppendLine($@"
<div class='text-center'>
    {companySeal}
</div>
");
                        sealAdded = true;
                    }
                    catch
                    {
                        
                        sealAdded = true;
                    }
                }
                else
                {
                    // ไม่มีไฟล์ตรา/ไม่มี <content> ⇒ ใส่ placeholder ครั้งเดียว
                    
                    sealAdded = true;
                }
            }

            signatoryHtml.AppendLine($@"
<div class='sign-single-right'>
    {signatureHtml}
    <div class='t-16 text-center tab1'>({signer?.Signatory_Name})</div>
    <div class='t-16 text-center tab1'>{signer?.BU_UNIT}</div>
</div>");
        }

        // ► Fallback: ถ้าจบลูปแล้วยังไม่มีตราประทับ แต่คุณ “ต้องการให้มีอย่างน้อย placeholder 1 ครั้ง”
        if (!sealAdded)
        {
            
            sealAdded = true;
        }

        // ► ประกอบผลลัพธ์
        var signatoryWithLogoHtml = new StringBuilder();
        if (companySealHtml.Length > 0) signatoryWithLogoHtml.Append(companySealHtml);
        signatoryWithLogoHtml.Append(signatoryHtml);



        var html = $@"<html>
<head>
    <meta charset='utf-8'>
  
     <style>
        @font-face {{
            font-family: 'THSarabunNew';
            src: url('file:///{fontPath}') format('truetype');
            font-weight: normal;
            font-style: normal;
        }}
         body {{
            font-size: 22px;
            font-family: 'THSarabunNew', Arial, sans-serif;
        }}
        /* แก้การตัดคำไทย: ไม่หั่นกลางคำ, ตัดเมื่อจำเป็น */
        body, p, div {{
            word-break: keep-all;            /* ห้ามตัดกลางคำ */
            overflow-wrap: break-word;       /* ตัดเฉพาะเมื่อจำเป็น (ยาวจนล้นบรรทัด) */
            -webkit-line-break: after-white-space; /* ช่วย WebKit เก่าจัดบรรทัด */
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
        .tab0 {{ text-indent: 0px;     }}
        .tab1 {{ text-indent: 48px;     }}
        .tab2 {{ text-indent: 96px;    }}
        .tab3 {{ text-indent: 144px;    }}
        .tab4 {{ text-indent: 192px;   }}
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
        .table {{ width: 100%; border-collapse: collapse; margin-top: 20px; font-size: 28pt; }}
        .table th, .table td {{ border: 1px solid #000; padding: 8px; }}

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
        }}
     .logo-table {{ width: 100%; border-collapse: collapse; margin-top: 40px; }}
        .logo-table td {{ border: none; }}
        p {{
            margin: 0;
            padding: 0;
        }}
    </style>
</head><body>

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
          
        </td>
    </tr>
</table>
</br>

</br>
    <div class='t-22 text-center'><b>แนวทางการจัดทำ</b></div>
    <div class='t-22 text-center'><b>เอกสารแนบท้ายบันทึกข้อตกลงความร่วมมือและสัญญาร่วมดำเนินการ</b></div>
    <div class='t-18 text-center'><b>ข้อกำหนดของการดำเนินงาน</b></div>
  <div class='t-18 text-center'><b>โครงการ………………………………………………………………….………………………………………..</b></div>
  <div class='t-18 text-center'><b>ระหว่าง</b></div>
   <div class='t-18 text-center'><b>สำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม ( สสว. )</b></div>
 <div class='t-18 text-center'><b>กับ</b></div>
<div class='t-18 text-center'><b>ชื่อหน่วยร่วมดำเนินการ………………….………………( ชื่อย่อ )………..</b></div>
</br>
<p class='t-16 tab0'><b>๑. หลักการและเหตุผล</b></p>
    <P class='t-16 tab3'>
       ………………………………………………………………….……………………………………….. 
 ………………………………………………………………….……………………………………….. 
………………………………………………………………….……………………………………….. 
………………………………………………………………….……………………………………….. 
………………………………………………………………….………………………………………..
    </P>
<p class='t-16 tab0'><b>๒.๒. วัตถุประสงค์</b></p>
  <P class='t-16 tab1'>๒.๑ .........................................</P>
    
 <P class='t-16 tab1'>๒.๒ .........................................</P>
 <P class='t-16 tab1'>๒.๓ .........................................</P>
 <P class='t-16 tab1'>ความใดในเอกสารแนบท้ายบันทึกข้อตกลงที่ขัดแย้งกับข้อความในบันทึกข้อตกลงนี้ให้
ใช้ข้อความในบันทึกข้อตกลงนี้บังคับในกรณีที่เอกสารแนบท้ายบันทึกข้อตกลงขัดแย้งกันเอง
ผู้รับจ้างจะต้องปฏิบัติตามคำวินิจฉัยของผู้ว่าจ้าง
ทั้งนี้ ผู้รับจ้างไม่มีสิทธิเรียกร้องค่าเสียหายหรือค่าใช้จ่ายใดๆ ทั้งสิ้น
</P>
    <P class='t-16 tab0'><B>๓. กลุ่มเป้าหมาย</B></P>

    <p class='t-16 tab1'>๓.๑ .........................................................</P>
 <p class='t-16 tab1'>๓.๒ .........................................................</P>

  <P class='t-16 tab1'><b>คุณสมบัติผู้เข้าร่วมโครงการ</b></P>
    <p class='t-16 tab1'>1 .........................................................</P>
  <p class='t-16 tab1'>2 .........................................................</P>
    <p class='t-16 tab1'>3 .........................................................</P>

<P class='t-16 tab0'><b>๔. พื้นที่ดำเนินการ</b></P>
 <p class='t-16 tab1'> .........................................................</P>
 <p class='t-16 tab1'> .........................................................</P>
 <p class='t-16 tab1'> .........................................................</P>

<P class='t-16 tab0'><b>๕. ขอบเขตการดำเนินงาน และกิจกรรมโครงการ</b></P>
<P class='t-16 tab1'>สสว. จะร่วมดำเนินการกับ…..(ชื่อย่อหน่วยร่วม)...... เพื่อการพัฒนาผู้ประกอบการกลุ่มเป้าหมายให้ตรงตามวัตถุประสงค์
และได้ตัวชี้วัดตามผลผลิตและผลลัพธ์ของ “…………………………………………….” ปี......... โดยมีขอบเขตการดำเนินงานและกิจกรรมโครงการ  ดังนี้</P>

<P class='t-16 tab2'><b>๕.๑  ขอบเขตการดำเนินงาน</b></P>
 <p class='t-16 tab3'> ๕.๑.๑ จัดทำแผนดำเนินงานและรายละเอียดของโครงการ.........ปีงบประมาณ ......... โดยให้เป็นไปตามวัตถุประสงค์
เป้าหมาย ผลผลิตและผลลัพธ์ของโครงการ ตามที่ สสว. กำหนด  ดังนี้</P>
 <p class='t-16 tab4'>• รายละเอียดโครงการ (แบบฟอร์ม สสว. ๑๐๐)</P>
<p class='t-16 tab4'>• เป้าหมายผลผลิตและผลลัพธ์ของโครงการ (แบบฟอร์ม สสว. ๑๐๐/๑)</P>
<p class='t-16 tab4'>• แผนการดำเนินโครงการ (แบบฟอร์ม สสว. ๑๐๐/๒)</P>
<p class='t-16 tab4'>• แผนการใช้จ่ายเงินของโครงการ (แบบฟอร์ม สสว. ๑๐๐/๓)</P>
<p class='t-16 tab3'>๕.๑.๒ จัดให้มีหัวหน้าทีมงาน เพื่อทำหน้าที่บริหารโครงการ และเจ้าหน้าที่ทีมงานที่มีความสามารถในการจัดทำข้อมูล เอกสาร รายงาน ที่มีประสิทธิภาพ และกรณีเปลี่ยนแปลงหัวหน้าทีมงาน 
ต้องเสนอรายชื่อเพื่อให้หน่วยงานบริหารโครงการของ สสว. พิจารณาให้ความเห็นชอบก่อน
</P>
<p class='t-16 tab3'>๕.๑.๓ การประชาสัมพันธ์โครงการในรูปแบบต่าง ๆ ต้องจัดให้มีสัญลักษณ์ สสว. (LOGO) 
ตรงตามอัตลักษณ์องค์กรของ สสว. ทุกกิจกรรมของการดำเนินงานโครงการ และทุกสื่อประชาสัมพันธ์ รวมถึงการบันทึกวีดีโอ Facebook Live การอบรม/สัมมนาและการประชาสัมพันธ์ผ่าน SME CONNEXT และ  SME ONE  โดยความเห็นชอบของ สสว.

</P>
<p class='t-16 tab3'>๕.๑.๔ จัดหาที่ปรึกษา/ผู้เชี่ยวชาญในการพัฒนาผู้ประกอบการแต่ละขั้นตอนการดำเนินงาน และดูแลให้คำแนะนำแก่ผู้ประกอบการ
SMEs ที่เข้าร่วมโครงการอย่างใกล้ชิด โดยที่ปรึกษา/ผู้เชี่ยวชาญจะต้อง ไม่ซ้ำซ้อนกับของหน่วยร่วมอื่น ๆ
ที่ดำเนินการอยู่ภายใต้โครงการเดียวกันและจะต้องปฏิบัติตามจรรยาบรรณที่ปรึกษาหน่วยงานที่ปรึกษาสังกัด
จรรยาบรรณที่ปรึกษาของศูนย์ข้อมูลที่ปรึกษากระทรวงการคลัง และต้องสมัครเข้าเป็นที่ปรึกษาของ สสว. โดยลงทะเบียนผ่าน  www.thesmecoach.com (ถ้ามี) และนำส่งรายชื่อที่ปรึกษา/ผู้เชี่ยวชาญ ที่ลงทะเบียนแล้วให้ สสว. ก่อนวันสิ้นสุดโครงการ
</P>

<p class='t-16 tab3'>นอกจากนี้ ที่ปรึกษาและ/หรือผู้เชี่ยวชาญเฉพาะด้าน จะต้องไม่ให้ช่วงงานทั้งหมดที่มอบหมาย หรือโอนงาน
หรือละทิ้งงานให้ผู้อื่นเป็นผู้ปฏิบัติงานแทนโดยไม่ได้รับความยินยอมจาก สสว.
และกรณีมีการเปลี่ยนที่ปรึกษา/ผู้เชี่ยวชาญ จะต้องเสนอรายชื่อและประวัติการทำงานเพื่อให้หน่วยงานบริหารโครงการของ สสว. พิจารณาให้ความเห็นชอบก่อน
</P>
<P class='t-16 tab3'>๕.๑.๕ จัดเตรียมวิทยากรผู้บรรยาย หรือผู้ทรงคุณวุฒิที่มีความรู้และประสบการณ์ให้เหมาะสมสอดคล้องกับหัวข้อการจัดกิจกรรมหรือWorkshop
หรือกิจกรรมอบรมสัมมนา โดยแนบประวัติวิทยากรผู้บรรยาย หรือผู้ทรงคุณวุฒิ มาพร้อมแผนการดำเนินงานโครงการ
และจัดให้มีบันทึกคลิปวีดีโอในการอบรมตามหลักสูตร เพื่อให้ผู้ประกอบการสามารถเรียนย้อนหลัง หรือนำไปเผยแพร่ใน Social Media และหรือสื่อ  อื่น ๆ โดยต้องได้รับความยินยอมจากวิทยากร (ถ้ามี)</P>
<P class='t-16 tab3'>๕.๑.๖ คัดเลือกกลุ่มเป้าหมายที่เข้าร่วมโครงการจะต้องไม่ซ้ำซ้อนกับกลุ่มเป้าหมายของหน่วยร่วมอื่น ๆ ที่ดำเนินการอยู่ภายใต้โครงการเดียวกันนี้</P>
<P class='t-16 tab3'>๕.๑.๗ การบันทึกข้อมูลผู้ประกอบการที่เข้าร่วมโครงการ ต้องดำเนินการ  ดังนี้ </P>
<P class='t-16 tab4'>๑) ลงทะเบียน SME ผู้รับบริการภาครัฐในระบบทะเบียน SME (หากไม่มีเลขทะเบียน SME ผู้รับบริการภาครัฐ) </P>
<P class='t-16 tab4'>๒) จัดทำข้อมูลผู้ประกอบการให้สมบูรณ์ตามแบบฟอร์มที่ สสว. กำหนด</P>
 <P class='t-16 tab4'> - บันทึกข้อมูลในระบบติดตามประเมินผลโครงการส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม และจัดทำหรือบันทึกข้อมูลผู้เข้าร่วมโครงการตามแบบฟอร์มที่ สสว. กำหนด และให้ความร่วมมือกับ สสว. ในการเก็บข้อมูลผู้เข้าร่วมโครงการเพื่อใช้ในการประเมินผลสำเร็จของโครงการ</P>
 <P class='t-16 tab4'> - ข้อมูลผู้ประกอบการ โดยใส่รหัสประเภทมาตรฐานอุตสาหกรรม (ประเทศไทย) Thailand Standard Industrial Classification (TSIC) ให้ระบุเป็นจำนวน ๕ หลัก  </P>
 <P class='t-16 tab4'>ทั้งนี้ รายชื่อผู้ประกอบการ SME ที่เข้าร่วมโครงการทั้งหมด เป็นกรรมสิทธิ์ของ สสว. 
ไม่สามารถนำไปใช้หรือเผยแพร่ที่อื่น
</P>
 <P class='t-16 tab3'>๕.๑.๘ (ชื่อย่อหน่วยร่วมดำเนินการ) จะต้องดำเนินงานซึ่งก่อให้เกิดผลสัมฤทธิ์ที่มีมูลค่าทางเศรษฐกิจของประเทศ (เช่น ยอดขายเพิ่มขึ้น การลงทุนเพิ่มขึ้น มูลค่าการลงทุน ตัวเลขการจ้างงานเพิ่มขึ้น) สามารถสืบค้น วัดผล และคำนวณผลเชิงปริมาณ/คุณภาพได้ หรือเอกสารแสดงที่มาของการเกิดผลสัมฤทธิ์
(ถ้ามี) และจัดทำผลลัพธ์เชิงประจักษ์ของโครงการตามแบบรายงานผลสัมฤทธิ์ที่ส่งผลต่อมูลค่าทางเศรษฐกิจ 
</P>
 <P class='t-16 tab0'><b>๕.๒  กิจกรรมโครงการ</b></P>
 <P class='t-16 tab3'>๕.๒.๑ .......................</P>
  <P class='t-16 tab3'>๕.๒.๒ .......................</P>
 <P class='t-16 tab3'>๕.๒.๓ .......................</P>
 <P class='t-16 tab3'>๕.๒.๔ .......................</P>
 <P class='t-16 tab3'>๕.๒.๕ .......................</P>
 <P class='t-16 tab3'>๕.๒.๖ .......................</P>
 <P class='t-16 tab3'>๕.๒.๗ .......................</P>
 <P class='t-16 tab3'>๕.๒.๘ .......................</P>

  <P class='t-16 tab0'><b>๖. ตัวชี้วัดความสำเร็จของโครงการ</b></P>
   {indicatorTable.ToString()}

   <P class='t-16 tab0'><b>๗. ระยะเวลาการดำเนินงาน</b></P>
    <P class='t-16 tab1'>ระยะเวลา</P>
  
  <P class='t-16 tab0'><b>๘. การรายงานรายเดือนตามแบบฟอร์ม สสว.</b></P>
    <P class='t-16 tab1'>๘.๑ รายงานความคืบหน้าประจำเดือน ตามแบบฟอร์ม สสว. ๒๐๐/๑, สสว. ๒๐๐/๑, สสว. ๒๐๐/๓ ประจำทุกสิ้นเดือน พร้อมไฟล์อิเล็กทรอนิกส์ สสว. ภายในวันที่ ๕ ของเดือนถัดไป จนกว่าจะครบกำหนดระยะเวลาการดำเนินงาน</P>
     <P class='t-16 tab1'>๘.๒ รายงานข้อมูลผู้รับบริการ (ผู้เข้าร่วมโครงการ) ในรูปแบบและแบบฟอร์มที่ สสว.กำหนด โดยให้นำส่งข้อมูลเป็นประจำทุกเดือน หรือ สิ้นสุดในแต่ละกิจกรรม</P>

<P class='t-16 tab0'><b>๙. การเก็บรักษาเงิน  การส่งมอบงาน และการเบิกจ่ายเงิน</b></P>
    <P class='t-16 tab1'><b>๙.๑ การเก็บรักษาเงิน</b></P>
    <P class='t-16 tab2'>การเก็บรักษาเงิน  สสว. จะจ่ายเงินให้หน่วยร่วมในชื่อบัญชีหลักของหน่วยร่วมเท่านั้น  โดย หน่วยร่วมต้องส่งสำเนาหน้าสมุดคู่ฝากเงินธนาคารให้ สสว. พร้อมเอกสารเบิกจ่ายเงินงวดแรก  และให้หน่วยร่วมดำเนินการดังนี้</P>
        <P class='t-16 tab2'>- เปิดบัญชีธนาคารประเภทออมทรัพย์แยกเฉพาะโครงการ โดยใช้ชื่อบัญชีว่า “(ชื่อหน่วยร่วม) เพื่อโครงการ/กิจกรรม ..................................................ประจำปี 25…. ” (ควรระบุแบบย่อ หรือสั้น ๆ แต่ต้องมีความหมายชัดเจนหรือสื่อได้ว่าเป็นโครงการใด) เพื่อให้สะดวกต่อการตรวจสอบการใช้จ่ายเงินและเพื่อให้เกิดความคล่องตัวในการปิดบัญชีและสามารถคืนเงินต้นพร้อมดอกผล(ถ้ามี) ได้ทันทีเมื่อสิ้นสุดโครงการ </P>
    <P class='t-16 tab2'>- ดอกเบี้ยจากการฝากเงินกับธนาคารถือว่าเป็นเงินของ สสว. หน่วยร่วมดำเนินการจะต้องนำส่งเงินคงเหลือและดอกเบี้ย พร้อมสำเนาหลักฐานการปิดบัญชีแยกเฉพาะโครงการให้กับ สสว. เมื่อสิ้นสุดสัญญาหรือปิดโครงการ</P>
    <P class='t-16 tab2'>- เมื่อสิ้นสุดโครงการต้องจัดทำหนังสือแจ้งปิดโครงการ พร้อมสรุปผลการดำเนินการ รายงาน
การใช้จ่ายเงิน ให้แก่ สสว. ภายใน ๓๐ วัน  และนำส่งเงินคงเหลือพร้อมดอกผล (ถ้ามี) ให้แก่ สสว. ภายใน ๑๕ วันนับถัดจากวันที่ได้รับเงินสนับสนุนงวดสุดท้าย
</P>
<P class='t-16 tab1'><b>๙.๒ การส่งมอบงาน </b></P>
     <P class='t-16 tab2'> การขอเบิกจ่ายงบประมาณในแต่ละงวดนั้น จะต้องจัดทำเป็นหนังสือถึงผู้อำนวยการสำนักงานส่งเสริมวิสาหกิจขนาดกลางและขนาดย่อม เพื่อขอเบิกจ่ายเงินตามงวดงาน พร้อมแนบเอกสารตามที่ สสว.กำหนด ดังนี้
</P>
    <P class='t-16 tab2'>• หนังสือแต่งตั้งผู้มีอำนาจลงนาม</P>
 <P class='t-16 tab2'>•	สำเนาบัตรข้าราชการหรือบัตรประชาชนของผู้มีอำนาจลงนาม</P>
 <P class='t-16 tab2'>•	กรณีมอบอำนาจแทน ผู้มีอำนาจลงนาม ต้องมีหนังสือมอบอำนาจและสำเนาบัตรประชาชนของผู้มีอำนาจลงนามและผู้รับมอบอำนาจลงนาม</P>
    <P class='t-16 tab1'><b>๙.๓ การจ่ายเงิน</b></P>
    <P class='t-16 tab2'>สสว. จะสนับสนุนค่าใช้จ่ายในการดำเนินโครงการเป็นเงินทั้งสิ้น ................................ บาท (.........................................บาทถ้วน)  โดย สสว. จะเบิกจ่ายเงินให้เมื่อ ………..….. ส่งมอบผลงานพร้อมทั้งรายงานและเอกสารที่เกี่ยวข้องครบสมบูรณ์ตามที่ระบุไว้ในรายละเอียดงวดงาน และ สสว. ได้ตรวจรับงานเรียบร้อยแล้ว โดยได้แบ่งการชำระเงินเป็น ๓ งวด  ดังนี้</P>
 <P class='t-16 tab2'>งวดที่ ๑ สนับสนุนงบประมาณ ร้อยละ ........... ของวงเงินสัญญาร่วมดำเนินการ 
จำนวน ................................ บาท (..................................บาทถ้วน) 
ภายใน .............วันหลังจากลงนามในสัญญาร่วมดำเนินการ เมื่อ ....... ดำเนินการส่งมอบงานดังนี้</P>
 <P class='t-16 tab2'>-.......................................</P>
  <P class='t-16 tab2'>-.......................................</P>
 <P class='t-16 tab2'>-.......................................</P>
   <P class='t-16 tab2'>งวดที่ ๒ สนับสนุนงบประมาณ ร้อยละ ………..ของวงเงินสัญญาร่วมดำเนินการ
จำนวน ................................ บาท (.....................................บาทถ้วน)
ภายใน …… วันหลังจากลงนามในสัญญาร่วมดำเนินการ เมื่อ ……. ได้ดำเนินงานตามแผนงานและส่งมอบงานดังนี้</P>
 <P class='t-16 tab2'>-.......................................</P>
  <P class='t-16 tab2'>-.......................................</P>
 <P class='t-16 tab2'>-.......................................</P>
   <P class='t-16 tab2'>งวดที่ ๓ (งวดสุดท้าย) สนับสนุนงบประมาณร้อยละ 
ของวงเงินสัญญาร่วมดำเนินการ จำนวน ................................ บาท (..............................บาทถ้วน)
ภายในวันที่ …………….…. เมื่อ ....................... ดำเนินกิจกรรมเสร็จสิ้นเรียบร้อยแล้ว </P>
 <P class='t-16 tab2'>-.......................................</P>
  <P class='t-16 tab2'>-.......................................</P>
 <P class='t-16 tab2'>-.......................................</P>
<P class='t-16 tab2'><b>พร้อมแนบเอกสารตามข้อ ๕.๑.๘</b></P>

<P class='t-16 tab0'><b>๑๐. เงื่อนไขอื่น ๆ</b></P>

<P class='t-16 tab1'>๑๐.๑ (ชื่อย่อหน่วยร่วมดำเนินการ) ต้องปฏิบัติตามคู่มือการดำเนินโครงการ ของ สสว. โดยเคร่งครัด</P>
<P class='t-16 tab1'>๑๐.๒ (ชื่อย่อหน่วยร่วมดำเนินการ) ต้องประสานงานกับ สสว. อย่างต่อเนื่องและใกล้ชิด  ต้องอำนวยความสะดวกให้ สสว. หรือเจ้าหน้าที่ของ สสว. ในการประสาน กำกับ บริหารจัดการ และประเมินผลการดำเนินโครงการ</P>
<P class='t-16 tab1'>๑๐.๓ (ชื่อย่อหน่วยร่วมดำเนินการ) ต้องรับผิดชอบประสานงานกับผู้ประกอบการที่เข้าร่วมโครงการ อย่างต่อเนื่องใกล้ชิด  ต้องอำนวยความสะดวกให้ผู้ประกอบการที่เข้าร่วมโครงการและการดำเนินกิจกรรมตามโครงการ  รวมถึงสนับสนุนค่าใช้จ่ายในส่วนที่เกี่ยวข้องกับกิจกรรมในโครงการให้แก่ผู้ประกอบการตามที่ สสว. กำหนดไว้ (ถ้ามี)</P>
<P class='t-16 tab1'>๑๐.๔ กรณีค่าใช้จ่ายต่าง ๆ ในการดำเนินการและอื่น ๆ อันเกิดขึ้นจากการดำเนินกิจกรรมกับ สสว. (ชื่อย่อหน่วยร่วมดำเนินการ) ต้องเป็นผู้รับผิดชอบค่าใช้จ่ายทั้งสิ้น จะใช้สิทธิเรียกร้องค่าเสียหายใด ๆ จาก สสว. ไม่ได้</P>
<P class='t-16 tab1'>๑๐.๕ (ชื่อย่อหน่วยร่วมดำเนินการ) มีหน้าที่รับผิดชอบบริหารจัดการบัญชีและการจัดเก็บเอกสารที่เกี่ยวข้องกับการดำเนินโครงการ และเกี่ยวกับการรับเงิน การจ่ายเงินหรือการก่อหนี้ผูกพันทางการเงิน จัดทำบัญชีรายรับรายจ่าย รวมถึงหนังสือและเอกสารอื่นที่เกี่ยวข้องกับการดำเนินโครงการ เพื่อให้หน่วยงานตรวจสอบ เช่น สำนักงานการตรวจเงินแผ่นดิน สำนักงานป้องกันและปราบปรามการทุจริตแห่งชาติ เป็นต้น สามารถใช้ตรวจสอบและอ้างอิงได้  ทั้งนี้ ระยะเวลาการจัดเก็บเอกสารให้เป็นไปตามระเบียบของราชการ</P>

<P class='t-16 tab0'><b>๑๑. เอกสารประกอบการจัดทำบันทึกข้อตกลงความร่วมมือและสัญญาร่วมดำเนินการ</b></P>
<P class='t-16 tab1'>สสว. (รับรองสำเนาถูกต้อง ทุกสำเนาเอกสาร)</P>
<P class='t-16 tab2'>๑. สำเนาหนังสือแต่งตั้งผู้มีอำนาจลงนามของ สสว. จำนวน ๒ ชุด </P>
<P class='t-16 tab2'>๒. สำเนาบัตรข้าราชการหรือบัตรประชาชนของผู้มีอำนาจลงนามของ สสว. จำนวน ๒ ชุด</P>
<P class='t-16 tab2'>๓. กรณีมอบอำนาจแทนผู้มีอำนาจลงนาม ต้องมีหนังสือมอบอำนาจ และสำเนาบัตรประชาชน
ผู้มีอำนาจลงนามและผู้รับมอบอำนาจลงนาม เพิ่มจำนวน ๑ ชุด</P>
<P class='t-16 tab1'>หน่วยร่วมดำเนินการ (รับรองสำเนาถูกต้อง ทุกสำเนาเอกสาร)</P>
<P class='t-16 tab2'>๑. สำเนาเอกสารแสดงการจดทะเบียนเป็นนิติบุคคล หรือแสดงการจัดตั้งหน่วยงาน จำนวน ๒ ชุด </P>
<P class='t-16 tab2'>๒. สำเนาหนังสือแต่งตั้งผู้มีอำนาจลงนาม จำนวน ๒ ชุด</P>
<P class='t-16 tab2'>๔. กรณีมอบอำนาจแทนผู้มีอำนาจลงนาม ต้องมีหนังสือมอบอำนาจ และสำเนาบัตรประชาชน
ผู้มีอำนาจลงนามและผู้รับมอบอำนาจลงนาม เพิ่มจำนวน ๑ ชุด
</P>

</div>
{signatoryWithLogoHtml}
</body>
</html>
";


        return html;
    }

    public static ProjectData GenerateMockData()
    {
        var projectData = new ProjectData();

        // Mock data for "Outputs"
        projectData.Outputs = new List<ProjectIndicator>
        {
            new ProjectIndicator
            {
                Description = "เอกสารคู่มือการใช้งาน (User manual documents)",
                UnitOfMeasurement = "ฉบับ (Copies)",
                Target = "5",
                MeasurementMethod = "การตรวจสอบเอกสาร (Document review)"
            },
            new ProjectIndicator
            {
                Description = "จำนวนผู้เข้าอบรม (Number of trainees)",
                UnitOfMeasurement = "คน (People)",
                Target = "100",
                MeasurementMethod = "รายชื่อผู้เข้าอบรม (Trainee list)"
            },
            new ProjectIndicator
            {
                Description = "การจัดกิจกรรม Workshop (Workshop activities)",
                UnitOfMeasurement = "ครั้ง (Times)",
                Target = "2",
                MeasurementMethod = "สรุปผลการจัดกิจกรรม (Activity summary report)"
            },
            new ProjectIndicator
            {
                Description = "รายงานผลการดำเนินงาน (Project progress reports)",
                UnitOfMeasurement = "ฉบับ (Copies)",
                Target = "3",
                MeasurementMethod = "การส่งมอบเอกสารตามกำหนด (On-time document submission)"
            }
        };

        // Mock data for "Outcomes"
        projectData.Outcomes = new List<ProjectIndicator>
        {
            new ProjectIndicator
            {
                Description = "ผู้ใช้งานมีความเข้าใจในการใช้ระบบ (User understanding of the system)",
                UnitOfMeasurement = "เปอร์เซ็นต์ (%)",
                Target = "85%",
                MeasurementMethod = "แบบสอบถามประเมินความพึงพอใจ (Satisfaction survey)"
            },
            new ProjectIndicator
            {
                Description = "ลดระยะเวลาในการทำงาน (Reduced work duration)",
                UnitOfMeasurement = "เปอร์เซ็นต์ (%)",
                Target = "20%",
                MeasurementMethod = "การเปรียบเทียบระยะเวลาก่อนและหลังใช้งาน (Before and after usage time comparison)"
            }
        };

        return projectData;
    }
    public class ProjectIndicator
    {
        public string Description { get; set; }
        public string UnitOfMeasurement { get; set; }
        public string Target { get; set; }
        public string MeasurementMethod { get; set; }
    }

    // Defines a container for the project's success indicators
    public class ProjectData
    {
        public List<ProjectIndicator> Outputs { get; set; }
        public List<ProjectIndicator> Outcomes { get; set; }
    }
}
