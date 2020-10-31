using eInvitationApp.Models;
using GemBox.Document;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using System.IO;

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
            if (obj.Password == "20moola20")
            {
                ViewBag.ShowLinkbtn = true;
            }

            return View();
        }

        [HttpGet]
        public IActionResult Register(string id)
        {
            //string command_id = (command != null) ? command.ToString() : "";

            //if (command_id.Equals("Download Invitation"))
            //{
            //    if(obj.FirstName != null || obj.FirstName == "")
            //    {
            //        var options = GemBox.Document.SaveOptions.PdfDefault;

            //        DocumentModel document = this.Download(obj.FirstName, obj.LastName);
            //        // Save document to DOCX format in byte array.
            //        using (var stream = new MemoryStream())
            //        {
            //            document.Save(stream, options);

            //            return View(); // return File(stream.ToArray(), options.ContentType, obj.FirstName + "_" + obj.LastName + "_invite.pdf");
            //        }
            //    }
            //    else
            //    {
            //        return View();
            //    }
            //}
            //else
            //{
            //    return View();
            //}
            return View();

        }

        [HttpPost]
        public IActionResult Register(AttendeeModel obj)
        {
            if (!ModelState.IsValid)
                return View(obj);
            var options = GemBox.Document.SaveOptions.PdfDefault;

            DocumentModel document = this.Download(obj.FirstName, obj.LastName);
            // Save document to DOCX format in byte array.
            using (var stream = new MemoryStream())
            {
                document.Save(stream, options);
                stream.Position = 0;
                return File(stream.ToArray(), options.ContentType, obj.FirstName + "_" + obj.LastName + "_invite.pdf");
            }

        }


        private DocumentModel Download(string firstName, string lastName)
        {

            // If using Professional version, put your serial key below.
            ComponentInfo.SetLicense("FREE-LIMITED-KEY");
            var fullName = (firstName + " " + lastName).ToUpper();
            var document = DocumentModel.Load("Resources/MOOLA2020.docx");
            document.Content.Replace("%Fullname%", fullName, new CharacterFormat() { FontColor = Color.Black, Size = 16, FontName = "Tahoma" });
            return document;
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
