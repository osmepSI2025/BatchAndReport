using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;

namespace BatchAndReport.Pages
{
    public class testUploadModel : PageModel
    {
        [BindProperty]
        public List<IFormFile> PostedFiles { get; set; }

        public string StatusMessage { get; set; }
        public bool IsSuccess { get; set; }

        public void OnGet()
        {
        }
    }
}
