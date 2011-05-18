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

            AddSender();

            _smtpClient.Port = _settings.PortNumber;
            _smtpClient.Credentials = new System.Net.NetworkCredential(_settings.UserName, _settings.Password);
            //SmtpServer.EnableSsl = true;  
        }

        public void SendEmail(List<string> files, string subject)
        {
            _mail.Subject = subject;
            _mail.Body = "";

            foreach (var file in files)
                _mail.Attachments.Add(CreateAttachment(file));

            _smtpClient.Send(_mail);
        }

        private void AddSender()
        {
            if (!string.IsNullOrEmpty(_settings.EmailToAddress))
            {
                var emails = Common.ParseEmailAddresses(_settings.EmailToAddress);

                foreach (var email in emails)
                    _mail.To.Add(email);
            }

            if (!string.IsNullOrEmpty(_settings.EmailCcAddress))
            {
                var emails = Common.ParseEmailAddresses(_settings.EmailCcAddress);

                foreach (var email in emails)
                    _mail.CC.Add(email);
            }

            if (!string.IsNullOrEmpty(_settings.EmailBccAddress))
            {
                var emails = Common.ParseEmailAddresses(_settings.EmailBccAddress);

                foreach (var email in emails)
                    _mail.Bcc.Add(email);
            }
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
