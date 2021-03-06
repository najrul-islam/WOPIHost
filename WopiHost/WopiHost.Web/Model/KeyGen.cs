using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Web;


namespace WopiHost.Web
{
    public interface IKeyGen
    {
        string GetHash(string value);

        bool Validate(string name, string access_token);
    }
    public class KeyGen : IKeyGen
    {
        byte[] _key;
        int _saltLength = 8;

        static RNGCryptoServiceProvider s_rngCsp = new RNGCryptoServiceProvider();

        public KeyGen()
        {
            var key = "Dbh1zG6OvhISvBNHqLDDWIF7Nvwdf2iTJEkv04sH8rTjzfiROd4WJunI0yP+Amd3K2MqMa4rphNYGTd1XzV8tA==";
            if (string.IsNullOrEmpty(key))
                throw new ArgumentNullException("must supply a HmacKey - check config");
            _key = Encoding.UTF8.GetBytes(key);
        }

        public string GetHash(string value)
        {
            byte[] salt = new byte[_saltLength];
            s_rngCsp.GetBytes(salt);

            var saltStr = Convert.ToBase64String(salt);
            return GetHash(value, saltStr);
        }

        internal string GetHash(string value, string saltStr)
        {
            HMACSHA256 hmac = new HMACSHA256(_key);
            var hash = hmac.ComputeHash(Encoding.UTF8.GetBytes(saltStr + value));
            var rv = Convert.ToBase64String(hash);
            return saltStr + rv;
        }


        public bool Validate(string name, string access_token)
        {
            var targetHash = GetHash(name, access_token.Substring(0,_saltLength + 4));  //hack for base64 form
            return String.Equals(access_token, targetHash);
        }


    }
}