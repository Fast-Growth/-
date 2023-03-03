using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace Расписание_занятий.Class
{
    internal class Encryption
    {
        public static string GetHash(string connStr)
        {
            var sha = new SHA1Managed();
            byte[] hash = sha.ComputeHash(Encoding.UTF8.GetBytes(connStr));
            return Convert.ToBase64String(hash);
        }
    }
}
