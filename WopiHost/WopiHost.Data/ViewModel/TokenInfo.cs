using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WopiHost.Data.ViewModel
{
    public class TokenInfo
    {
        //Email or UserNmae
        public string UserName { get; set; }
        public string Password { get; set; }
        public int UserId { get; set; }
        public string Refreshtoken { get; set; }
    }
    public class TokenResult
    {
        public string Access_token { get; set; }
        public DateTime? Expiration { get; set; }
        public string UserEmail { get; set; }
        public int StatusCode { get; set; }
        public string Message { get; set; }
    }
}
