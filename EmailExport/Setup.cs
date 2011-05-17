﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Kofax.Eclipse.Base;

namespace EmailExport
{
    public partial class Setup : Form
    {
        private readonly List<IPageOutputConverter>     _pageConverters     = new List<IPageOutputConverter>();
        private readonly List<IDocumentOutputConverter> _docConverters      = new List<IDocumentOutputConverter>();
        private readonly ExportSettings                 _settings           = new ExportSettings();

        public Setup(ExportSettings settings, IEnumerable<IExporter> exporters)
        {
            _settings = settings;

            InitializeComponent();

            UpdateUI();

            foreach (IExporter exporter in exporters)
            {
                if (exporter is IPageOutputConverter)
                    _pageConverters.Add(exporter as IPageOutputConverter);
                if (exporter is IDocumentOutputConverter)
                    _docConverters.Add(exporter as IDocumentOutputConverter);
            }

            cbSinglePage.Checked = _settings.ReleaseMode == ReleaseMode.SinglePage;
            cbMultipage.Checked = _settings.ReleaseMode == ReleaseMode.MultiPage;
        }

        private void UpdateUI()
        {
            txtEmailDestination.Text = _settings.EmailDestination;
            txtFromAddress.Text = _settings.EmailFromAddress;
            txtSmtpServer.Text = _settings.SmtpServer;
            txtPortNumber.Text = _settings.PortNumber.ToString();
            txtUserName.Text = _settings.UserName;
            txtPassword.Text = _settings.Password;
        }

        private bool ValidateServerSettings()
        {
            if (!ValidateEmailAddress(txtEmailDestination.Text))
                return false;

            if (!ValidateEmailAddress(txtFromAddress.Text))
                return false;

            if (String.IsNullOrEmpty(txtSmtpServer.Text))
                return false;

            int x; //users to validate the port
            if (!int.TryParse(txtPortNumber.Text, out x))
                return false;

            return true;

        }

        private bool ValidateEmailAddress(string emailAddress)
        {
            return System.Text.RegularExpressions.Regex.IsMatch(
                                 emailAddress,
                                 @"^(?("")("".+?""@)|(([0-9a-zA-Z]((\.(?!\.))|[-!#\$%&'\*\+/=\?\^`\{\}\|~\w])*)(?<=[0-9a-zA-Z])@))" +
                                 @"(?(\[)(\[(\d{1,3}\.){3}\d{1,3}\])|(([0-9a-zA-Z][-\w]*[0-9a-zA-Z]\.)+[a-zA-Z]{2,6}))$"
                                 );
        }

        private void cboFileType_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cboFileType.SelectedIndex < 0)
                return;

            if (cbSinglePage.Checked && cboFileType.SelectedIndex < _pageConverters.Count)
            {
                btnSetup.Enabled = _pageConverters[cboFileType.SelectedIndex].IsCustomizable;
                _settings.FileTypeId = _pageConverters[cboFileType.SelectedIndex].Id;
            }

            if (cbMultipage.Checked && cboFileType.SelectedIndex < _docConverters.Count)
            {
                btnSetup.Enabled = _docConverters[cboFileType.SelectedIndex].IsCustomizable;
                _settings.FileTypeId = _docConverters[cboFileType.SelectedIndex].Id;
            }
        }

        private void btnSetupFileType_Click(object sender, EventArgs e)
        {
            if (cboFileType.SelectedIndex < 0)
                return;

            if (cbSinglePage.Checked && cboFileType.SelectedIndex < _pageConverters.Count)
                _pageConverters[cboFileType.SelectedIndex].Setup(new Dictionary<string, string>());

            if (cbMultipage.Checked && cboFileType.SelectedIndex < _docConverters.Count)
                _docConverters[cboFileType.SelectedIndex].Setup(new Dictionary<string, string>());
        }

        private void optionSingleMulti_CheckedChanged(object sender, EventArgs e)
        {
            cboFileType.Items.Clear();

            if (cbSinglePage.Checked)
                foreach (IPageOutputConverter pageConverter in _pageConverters)
                {
                    cboFileType.Items.Add(string.Format("{0} - {1}", pageConverter.DefaultExtension, pageConverter.Name));
                    if (pageConverter.Id == _settings.FileTypeId)
                        cboFileType.SelectedIndex = _pageConverters.IndexOf(pageConverter);
                }

            if (cbMultipage.Checked)
                foreach (IDocumentOutputConverter docConverter in _docConverters)
                {
                    cboFileType.Items.Add(string.Format("{0} - {1}", docConverter.DefaultExtension, docConverter.Name));
                    if (docConverter.Id == _settings.FileTypeId)
                        cboFileType.SelectedIndex = _docConverters.IndexOf(docConverter);
                }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (!ValidateServerSettings())
            {
                MessageBox.Show("Invalid settings", "Invalid Settings", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            _settings.EmailDestination = txtEmailDestination.Text;
            _settings.EmailFromAddress = txtFromAddress.Text;
            _settings.SmtpServer = txtSmtpServer.Text;
            _settings.PortNumber = Convert.ToInt32(txtPortNumber.Text);
            _settings.UserName = txtUserName.Text;
            _settings.Password = txtPassword.Text;

            DialogResult = DialogResult.OK;                
        }
    }
}