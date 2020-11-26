using System;
using System.Collections.Generic;
using System.Text;

namespace SendMail
{
    class CSV
    {
        public string Escola { get; set; }
        public string EscolaEmail { get; set; }
        public string GestorEmail { get; set; }
        public string SecretarioEmail { get; set; }

        public override string ToString()
        {
            return $"[{Escola}, {EscolaEmail}, {GestorEmail}, {SecretarioEmail}]";
        }
    }
}
