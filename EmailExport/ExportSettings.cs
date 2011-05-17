using System;
using System.Collections.Generic;
using System.Text;
using Kofax.Eclipse.Base;

namespace EmailExport
{
    public class ExportSettings
    {
        public string   EmailToAddress      { get; set; }
        public string   EmailCcAddress      { get; set; }
        public string   EmailBccAddress     { get; set; }
        public string   EmailFromAddress    { get; set; }
        public string   SmtpServer          { get; set; }

        public string   UserName            { get; set; }
        public string   Password            { get; set; }

        public Guid         FileTypeId      { get; set; }
        public ReleaseMode  ReleaseMode     { get; set; }

        private int     _portNumber;
        public int PortNumber 
        { 
            get 
            {
                if (_portNumber == 0)
                    return 587;
                else
                    return _portNumber;
            }
            set
            { 
                _portNumber = value; 
            }
        }
    }
}
