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
                doc.LoadFromFile("wwwroot/document/Testcontract.docx");

                var stream = new MemoryStream();
                doc.SaveToStream(stream, FileFormat.PDF);
                stream.Position = 0;

                return File(stream, "application/pdf", "Testcontract.pdf");
            }
            catch (Exception ex)
            {
                // Log the exception or handle it as needed
                return StatusCode(500, "Internal server error: " + ex.Message);
            }
          
        }
    }
}
