using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace eInvitationApp.Models
{
    public class AttendeeModel
    {
        public string id { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string email { get; set; }
        public string phone { get; set; }
    }

    public class AttendanceModel
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public Newtonsoft.Json.Linq.JArray infodata { get; set; }
    }
}
