using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Dashboard
{
    public static class DbConfig
    {
        public static string ConnectionString()
        {
            string constring = "Server=localhost;Database=estatistica;Uid=root;Pwd='Programacao2021'";

            return constring;
        }
    }
}
