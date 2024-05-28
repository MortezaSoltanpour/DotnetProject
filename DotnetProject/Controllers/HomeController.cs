using DotnetProject.Models;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;

namespace DotnetProject.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index(string id = "1.jpg")
        {
            ViewBag.id = id;
            return View();
        }

        public IActionResult GetFile(string id)
        {
            string fileName = $"{id}";
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/img", fileName);

            var fileBytes = System.IO.File.ReadAllBytes(filePath);
            var contentType = "application/octet-stream";

            return new FileContentResult(fileBytes, contentType)
            {
                FileDownloadName = fileName
            };
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}