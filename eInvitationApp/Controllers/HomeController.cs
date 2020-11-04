using eInvitationApp.Models;
using GemBox.Document;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using RestSharp;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace eInvitationApp.Controllers
{
    public class HomeController : Controller
    {

        public class Statistics
        {
            public string total_links { get; set; }
            public string total_not_registered { get; set; }
            public string total_registered { get; set; }
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

            var uri = "/macros/s/AKfycbwilW4_LJrUVzziROPO6gdL3wM8cWnN-hke729NaQ67_vSfG_g/exec";
            if (obj.Password == "20moola20")
            {
                ViewBag.ShowLinkbtn = true;
            
                var client = new RestClient("https://script.google.com");
                var request = new RestRequest(uri, Method.POST);

                request.AddJsonBody(new
                {
                    action = "get_stats"
                });

                IRestResponse response = client.Execute(request);
                var content = response.Content;

                Dictionary<string, object> statistics = JsonConvert.DeserializeObject<Dictionary<string, object>>(content);

                var data = JsonConvert.SerializeObject(statistics["description"]);
                var datas = JsonConvert.DeserializeObject<Statistics>(data);

                ViewBag.TotalLinkGenerated = datas.total_links;
                ViewBag.TotalNotRegistered = datas.total_not_registered;
                ViewBag.TotalRegistered = datas.total_registered;
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
            if (!ModelState.IsValid)
                return View(obj);
            //var options = SaveOptions.PdfDefault;
            var options = new PdfSaveOptions() { ImageDpi = 220 };

            DocumentModel document = Download(obj.FirstName, obj.LastName);
            //Save document to DOCX format in byte array.
            using (var stream = new MemoryStream())
            {
                document.Save(stream, options);
                return File(stream.ToArray(), "application/pdf", obj.FirstName + "_" + obj.LastName + "_invite.pdf");
            }

        }


        private DocumentModel Download(string firstName, string lastName)
        {
            // If using Professional version, put your serial key below.
            ComponentInfo.SetLicense("FREE-LIMITED-KEY");
            string fullName = (firstName + " " + lastName).ToUpper();

            var document = DocumentModel.Load("wwwroot/MOOLA2020.docx");
            document.Content.Replace("%Fullname%", fullName);
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
