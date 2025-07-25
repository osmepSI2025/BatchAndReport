using BatchAndReport.Services;
using iText.IO.Font.Constants;
using iText.Kernel.Font;
using iText.Kernel.Pdf;

//using iText.Krnel.Pdf;
using iText.Layout;
using iText.Layout.Element;

public class PdfService : IPdfService
{
    public async Task<SMEProjectDetailModels?> GetProjectDetailAsync(string projectCode)
    {
        // Implementation for fetching project details
        return await Task.FromResult<SMEProjectDetailModels?>(null);
    }
    public byte[] GeneratePdf(SMEProjectDetailModels model)
    {
        using var stream = new MemoryStream();
        var writer = new PdfWriter(stream);
        var pdf = new PdfDocument(writer);
        var doc = new Document(pdf);

        var boldFont = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD);
        var normalFont = PdfFontFactory.CreateFont(StandardFonts.HELVETICA);

        doc.SetFont(normalFont);

        doc.Add(new Paragraph("แบบฟอร์มรายงานผล").SetFont(boldFont).SetFontSize(16));
        doc.Add(new Paragraph($"ชื่อโครงการ: {model.ProjectName}"));
        doc.Add(new Paragraph($"รหัสโครงการ: {model.ProjectCode}"));
        doc.Add(new Paragraph($"หน่วยงาน: {model.MinistryName}"));
        doc.Add(new Paragraph($"ปีงบประมาณ: {model.FiscalYear}"));
        doc.Add(new Paragraph($"งบประมาณที่ขอ: {model.BudgetAmount:N0}"));
        doc.Add(new Paragraph($"งบประมาณที่อนุมัติ: {model.BudgetAmountApprove?.ToString("N0") ?? "-"}"));
        doc.Add(new Paragraph($"สถานะโครงการ: {model.ProjectStatusName}"));
        doc.Add(new Paragraph($"กลยุทธ์: {model.StrategyDesc}"));
        doc.Add(new Paragraph($"พื้นที่ดำเนินงาน: {model.OperationArea}"));
        doc.Add(new Paragraph($"คะแนนประเมิน: {model.Score}"));
        doc.Add(new Paragraph($"ช่วงเวลาดำเนินการ: {model.StartDate:dd/MM/yyyy} - {model.EndDate:dd/MM/yyyy}"));

        doc.Add(new Paragraph("วัตถุประสงค์:").SetFont(boldFont));
        doc.Add(new Paragraph(model.ProjectObjective ?? "-"));

        doc.Add(new Paragraph("หลักการและเหตุผล:").SetFont(boldFont));
        doc.Add(new Paragraph(model.ProjectRationale ?? "-"));

        doc.Add(new Paragraph("กลุ่มเป้าหมาย:").SetFont(boldFont));
        doc.Add(new Paragraph(model.TargetGroup ?? "-"));

        doc.Close();
        return stream.ToArray();
    }
}
