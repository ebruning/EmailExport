using System;
using System.IO;
using System.Collections.Generic;
using System.Text;
using Kofax.Eclipse.Base;

namespace EmailExport
{
    public class EmailExport : IReleaseScript2
    {
        IDocumentOutputConverter    _documentConverter  = null;
        IPageOutputConverter        _pageConverter      = null;
        ExportSettings              _settings           = new ExportSettings();
        Email                       _email              = null;

        String                      _batchName          = string.Empty;
        List<string>                _fileList           = new List<string>();

        public void CustomizeSettings(IList<IExporter> exporters, IJob job, IApplication licenseData)
        {
            
        }

        public string Description
        {
            get { return "Exports the files to an email message"; }
        }

        public void DeserializeSettings(Stream input)
        {
            using (BinaryReader reader = new BinaryReader(input))
            {
                try
                {
                    _settings.EmailDestination = reader.ReadString();
                    _settings.EmailFromAddress = reader.ReadString();
                    _settings.SmtpServer = reader.ReadString();
                    _settings.UserName = reader.ReadString();
                    _settings.Password = reader.ReadString();
                    _settings.FileTypeId = new Guid(reader.ReadString());
                    _settings.ReleaseMode = (ReleaseMode)Enum.Parse(typeof(ReleaseMode), reader.ReadString());
                    _settings.PortNumber = reader.ReadInt32();
                }
                catch
                {
                    _settings = new ExportSettings();
                }
            }
            
        }

        public void EndBatch(IBatch batch, object handle, ReleaseResult result)
        {
            
        }

        public void EndDocument(IDocument doc, object handle, ReleaseResult result)
        {
           
        }

        public void EndRelease(object handle, ReleaseResult result)
        {
            _email = null;  
        }

        public Guid Id
        {
            get { return new Guid("{2898303F-65F7-4c63-8F1C-A5B818E52812}"); }
        }

        public bool IsSupported(ReleaseMode mode)
        {
            return true;
        }

        public string Name
        {
            get { return "Email Export"; }
        }

        public string OutFolderName(string documentNumber)
        {
            return Path.Combine(Path.GetTempPath(), Path.Combine("EmailExport", Path.Combine(_batchName, documentNumber)));
        }

        public void Release(IPage page)
        {
           
        }

        public void Release(IDocument doc)
        {
            var outputFolder    = OutFolderName(doc.Number.ToString());
            var outputFileName  = Path.ChangeExtension(doc.Number.ToString(), _documentConverter.DefaultExtension);

            if (Directory.Exists(outputFolder))
                Directory.Delete(outputFolder, true);

            _documentConverter.Convert(doc,Path.Combine(outputFolder, outputFileName));

            _email.SendEmail(Path.Combine(outputFolder, outputFileName), _batchName);

        }

        public void SerializeSettings(Stream output)
        {
            using (BinaryWriter writer = new BinaryWriter(output))
            {
                writer.Write(_settings.EmailDestination);
                writer.Write(_settings.EmailFromAddress);
                writer.Write(_settings.SmtpServer);
                writer.Write(_settings.UserName);
                writer.Write(_settings.Password);
                writer.Write(_settings.FileTypeId.ToString());
                writer.Write(_settings.ReleaseMode.ToString());
                writer.Write(_settings.PortNumber);
            }
        }

        public void Setup(IList<IExporter> exporters, IJob job, IDictionary<string, string> releaseData, IApplication licenseData)
        {
            Setup form = new Setup(_settings, exporters);
            if (form.ShowDialog() != System.Windows.Forms.DialogResult.OK) return;
        }

        public void Setup(IList<IExporter> exporters, IIndexField[] indexFields, IDictionary<string, string> releaseData) {}

        public object StartBatch(IBatch batch)
        {
            _batchName = batch.Name;

            return null;
        }

        public object StartDocument(IDocument doc)
        {
            return null;
        }

        public object StartRelease(IList<IExporter> exporters, IIndexField[] indexFields, IDictionary<string, string> releaseData, IApplication licenseData)
        {
            _email = new Email(_settings);

            foreach (IExporter exporter in exporters)
            {
                if (exporter.Id == _settings.FileTypeId)
                {
                    if (_settings.ReleaseMode == ReleaseMode.SinglePage)
                        _pageConverter = exporter as IPageOutputConverter;
                    else
                        _documentConverter = exporter as IDocumentOutputConverter;
                }
            }

            return null;
        }

        public object StartRelease(IList<IExporter> exporters, IIndexField[] indexFields, IDictionary<string, string> releaseData) { return null; }

        public ReleaseMode WorkingMode
        {
            get { return ReleaseMode.MultiPage; }
        }
    }
}
