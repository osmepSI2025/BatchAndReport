using Microsoft.AspNetCore.Mvc;

using Spire.Doc;

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

    }
}
