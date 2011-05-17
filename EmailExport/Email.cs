using System;
using System.Collections.Generic;
using System.Text;
using System.Net.Mail;
using System.Net.Mime;

namespace EmailExport
{
    internal class Email
    {
        ExportSettings  _settings   = new ExportSettings();
        MailMessage     _mail       = new MailMessage();
        SmtpClient      _smtpClient;

        public Email(ExportSettings settings)
        {
            _settings = settings;

            _smtpClient = new SmtpClient(_settings.SmtpServer);
            
            _mail.From = new MailAddress(_settings.EmailFromAddress);
            _mail.To.Add(_settings.EmailDestination);

            _smtpClient.Port = _settings.PortNumber; ;
            _smtpClient.Credentials = new System.Net.NetworkCredential(_settings.UserName, _settings.Password);
            //SmtpServer.EnableSsl = true;
            
        }

        public void SendEmail(string fileName, string subject)
        {
            _mail.Subject = subject;
            _mail.Body = "";

            _mail.Attachments.Add(CreateAttachment(fileName));

            _smtpClient.Send(_mail);
        }

        private Attachment CreateAttachment(string fileName)
        {
            Attachment attachment           = new Attachment(fileName, MediaTypeNames.Application.Octet);
            ContentDisposition disposition  = attachment.ContentDisposition;
            
            disposition.CreationDate        = System.IO.File.GetCreationTime(fileName);
            disposition.ModificationDate    = System.IO.File.GetLastWriteTime(fileName);
            disposition.ReadDate            = System.IO.File.GetLastAccessTime(fileName);

            return attachment;
        }
    }
}
