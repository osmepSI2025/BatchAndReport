using System.Diagnostics;
using BatchAndReport.Models;
using Microsoft.AspNetCore.Mvc;

namespace BatchAndReport.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }

        public IActionResult LoadEmpInfo()
        {
            return View();
        }
        public IActionResult LoadSMEInfo()
        {
            return View();
        }
        public IActionResult LoadWFInfo()
        {
            return View();
        }
        public IActionResult LoadEcontractInfo()
        {
            return View();
        }
        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
