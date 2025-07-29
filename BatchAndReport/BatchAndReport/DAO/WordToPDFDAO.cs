using DocumentFormat.OpenXml.Packaging;
using Microsoft.AspNetCore.Mvc;

using Spire.Doc;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;

namespace BatchAndReport.DAO
{
    public class WordToPDFDAO : Controller // Added inheritance from Controller to fix CS0103
    {
        public IActionResult OnGetPdfWithInterop()
        {
            try {
                var doc = new Document();
                doc.LoadFromFile("wwwroot/Document/PDSA/PDSA_1.docx");

                var stream = new MemoryStream();
                doc.SaveToStream(stream, FileFormat.PDF);
                stream.Position = 0;

                return File(stream, "application/pdf", "PDSA_1.pdf");
            }
            catch (Exception ex)
            {
                // Log the exception or handle it as needed
                return StatusCode(500, "Internal server error: " + ex.Message);
            }
          
        }
        public byte[] ConvertWordBytesToPdf(byte[] wordBytes)
        {
            using var wordStream = new MemoryStream(wordBytes);
            var doc = new Document();
            doc.LoadFromStream(wordStream, FileFormat.Docx);

            using var pdfStream = new MemoryStream();
            doc.SaveToStream(pdfStream, FileFormat.PDF);
            pdfStream.Position = 0;
            return pdfStream.ToArray();
        }
        public byte[] ConvertWordBytesToPdf_OpenXml(byte[] wordBytes)
        {
            // Open XML SDK can read and manipulate Word documents, but cannot convert to PDF directly.
            // You must use a library like Spire.Doc, Syncfusion, or a cloud API for PDF conversion.
            // This is a placeholder to show how to open a Word document with Open XML SDK.
            using var wordStream = new MemoryStream(wordBytes);
            using var wordDoc = WordprocessingDocument.Open(wordStream, false);
            // Manipulate document if needed...
            // For PDF conversion, use Spire.Doc or another library as shown in your existing code.
            throw new NotSupportedException("Open XML SDK does not support direct Word to PDF conversion.");
        }
        //public bool ConvertWordToPdfWithLibreOffice(string inputDocxPath, string outputPdfPath)
        //{
        //    var process = new System.Diagnostics.Process();
        //    process.StartInfo.FileName = "soffice"; // LibreOffice executable
        //    process.StartInfo.Arguments = $"--headless --convert-to pdf \"{inputDocxPath}\" --outdir \"{Path.GetDirectoryName(outputPdfPath)}\"";
        //    process.StartInfo.UseShellExecute = false;
        //    process.StartInfo.CreateNoWindow = true;
        //    process.Start();

        //    process.WaitForExit();

        //    // Check if PDF was created
        //    return System.IO.File.Exists(outputPdfPath);
        //}.

        public bool ConvertWordToPdfWithLibreOffice(string inputDocxPath, string outputPdfPath)
        {
            var process = new System.Diagnostics.Process();
            // Use the full path to the LibreOffice executable
            process.StartInfo.FileName = @"C:\Program Files\LibreOffice_25.2_SDK\sdk\setsdkenv_windows.exe"; // Update this path as needed
            process.StartInfo.Arguments = $"--headless --convert-to pdf \"{inputDocxPath}\" --outdir \"{Path.GetDirectoryName(outputPdfPath)}\"";
            process.StartInfo.UseShellExecute = false;
            process.StartInfo.CreateNoWindow = true;
            process.Start();

            process.WaitForExit();

            // Check if PDF was created
            return System.IO.File.Exists(outputPdfPath);
        }

        public byte[] ConvertWordBytesToPdf_Syncfusion(byte[] wordBytes)
        {
            using var wordStream = new MemoryStream(wordBytes);
            using var document = new WordDocument(wordStream, Syncfusion.DocIO.FormatType.Docx);
            using var renderer = new DocIORenderer();
            using var pdfDocument = renderer.ConvertToPDF(document);
            using var pdfStream = new MemoryStream();
            pdfDocument.Save(pdfStream);
            pdfStream.Position = 0;
            return pdfStream.ToArray();
        }
    }
}
