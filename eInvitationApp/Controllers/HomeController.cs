using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using eInvitationApp.Models;
using System.Data;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Microsoft.AspNetCore.Hosting;
using Syncfusion.DocToPDFConverter;
using Syncfusion.Pdf;
using System.IO;
using GemBox.Document;

namespace eInvitationApp.Controllers
{
    public class HomeController : Controller
    {

        private readonly IHostingEnvironment _hostingEnvironment;
        

        public HomeController(IHostingEnvironment hostingEnvironment)
        {
            _hostingEnvironment = hostingEnvironment;
        }

        [HttpGet]
        public IActionResult Index()
        {
            ViewBag.ShowLinkbtn = false;
            return View();
        }

        [HttpPost]
        public IActionResult Index(InviteModel obj)
        {
            ViewBag.ShowLinkbtn = false;
            if (obj.Password.ToLower() == "20moola20")
            {
                ViewBag.ShowLinkbtn = true;
            }

            return View();
        }

        [HttpGet]
        public IActionResult Register(string id)
        {

            return View();
        }

        [HttpPost]
        public IActionResult Register(AttendeeModel obj)
        {
            return Download(obj.FirstName, obj.LastName);
        }

        [HttpGet]
        public IActionResult Download(string firstName, string lastName)
        {

            // If using Professional version, put your serial key below.
            ComponentInfo.SetLicense("FREE-LIMITED-KEY");

            var document = DocumentModel.Load("Resources/MOOLA2020.docx");
            document.Content.Replace("%Fullname%", firstName + " " + lastName, new CharacterFormat() { FontColor = Color.Black, Size = 15, FontName = "Arial" });
            byte[] fileContents;
            MemoryStream data_stream;

            //var filePath = @"wwwroot/resources/" + firstName + "_" + lastName + "_invite.pdf";

            //document.Save(filePath);

            var options = GemBox.Document.SaveOptions.PdfDefault;

            // Save document to DOCX format in byte array.
            using (var stream = new MemoryStream())
            {
                document.Save(stream, options);
                data_stream = stream;
                fileContents = stream.ToArray();
            }

            //Download the PDF document in the browser
            //FileStreamResult fileStreamResult = new FileStreamResult(data_stream, "application/pdf");

            //fileStreamResult.FileDownloadName = "Sample.pdf";

            //return fileStreamResult;

            // Stream document to browser in DOCX format.
            //return File(fileContents, options.ContentType, Path.GetFileName(firstName + "_" + lastName + "_invite.pdf"));
            return File(fileContents, options.ContentType, firstName + "_" + lastName + "_invite.pdf");
        }

        public IActionResult Privacy()
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
