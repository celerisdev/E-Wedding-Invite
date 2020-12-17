using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using eInvitationApp.Models;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using RestSharp;

// For more information on enabling MVC for empty projects, visit https://go.microsoft.com/fwlink/?LinkID=397860

namespace eInvitationApp.Controllers
{
    public class AttendanceController : Controller
    {
        //// GET: /<controller>/
        //public IActionResult Index()
        //{
        //    return View();
        //}

        // GET: Home

        public class Attendees
        {  
            public string ID { get; set; }
            public string FIRST_NAME { get; set; }
            public string LAST_NAME { get; set; }
            public string EMAIL { get; set; }
            public string REG_DATE { get; set; }
            public string PHONE_NUMBER { get; set; }
            public string STATUS { get; set; }
            public Newtonsoft.Json.Linq.JArray infodata { get; set; }
        }

        [HttpGet]
        public IActionResult Index()
        {
            var uri = "/macros/s/AKfycbwilW4_LJrUVzziROPO6gdL3wM8cWnN-hke729NaQ67_vSfG_g/exec";

            var client = new RestClient("https://script.google.com");
            var request = new RestRequest(uri, Method.POST);

            request.AddJsonBody(new
            {
                action = "get_users"
            });

            IRestResponse response = client.Execute(request);
            var content = response.Content;

            Dictionary<string, object> statistics = JsonConvert.DeserializeObject<Dictionary<string, object>>(content);

            var data = JsonConvert.SerializeObject(statistics["description"]);
            List<Attendees> dataList = JsonConvert.DeserializeObject<List<Attendees>>(data);
            //var datas = JsonConvert.DeserializeObject<Dictionary<string, object>>(data);

            var array = JArray.FromObject(dataList);
            //ViewBag.data = statistics["description"];
            //return View();
            return View(new AttendanceModel
            {
                infodata = array
            });
        }
        [HttpPost]
        public IActionResult Index(string IDD)
        {
            var uri = "/macros/s/AKfycbwilW4_LJrUVzziROPO6gdL3wM8cWnN-hke729NaQ67_vSfG_g/exec";

            var client = new RestClient("https://script.google.com");
            var request = new RestRequest(uri, Method.POST);
            request.AddJsonBody(new
            {
                action = "mark_attendee",
                id = IDD
            });

            IRestResponse response = client.Execute(request);
            var content = response.Content;



            //var request_ = new RestRequest(uri, Method.POST);
            //request_.AddJsonBody(new
            //{
            //    action = "get_users"
            //});

            //IRestResponse response_ = client.Execute(request_);
            //var content_ = response_.Content;

            //Dictionary<string, object> statistics = JsonConvert.DeserializeObject<Dictionary<string, object>>(content_);

            //var data = JsonConvert.SerializeObject(statistics["description"]);
            //List<Attendees> dataList = JsonConvert.DeserializeObject<List<Attendees>>(data);
            ////var datas = JsonConvert.DeserializeObject<Dictionary<string, object>>(data);

            //var array = JArray.FromObject(dataList);

            return View();
        }

        //[HttpPost]
        //public ActionResult Index(string Name, string Id)
        //{
        //    ViewBag.Message = "Name: " + Name + " CustomerId: " + Id;
        //    return View();
        //}
    }
}
