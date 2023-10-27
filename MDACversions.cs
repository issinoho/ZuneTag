//------------------------------------------------------------------
// Zune Meta Tag Editor
// MDAC versioning Class
//
// <copyright file="MDACversions.cs" company="The Drunken Bakery">
//     Copyright (c) 2009 The Drunken Bakery. All rights reserved.
// </copyright>
//
// Editor to update WMV meta tags for the Zune
// Supplies installed MDAC versions.
//
// Author: IRS
// $Revision: 1.1 $
//------------------------------------------------------------------

namespace DrunkenBakery.ZuneTag
{
    using System;
    using System.Windows.Forms;

    using Microsoft.Win32;

    /// <summary>
    /// Reports on installed MDAC versions
    /// </summary>
    public partial class MDACversions : Form
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="MDACversions"/> class.
        /// </summary>
        public MDACversions()
        {
            this.InitializeComponent();

            // Clear list
            this.lvStatus.Columns.Add("Major Version", this.lvStatus.Width / 2, HorizontalAlignment.Left);
            this.lvStatus.Columns.Add("Revision", this.lvStatus.Width / 2 - 3, HorizontalAlignment.Left);
            this.lvStatus.Items.Clear();

            // Now get the versions from the reg
            this.ScrapeRegistry();
        }

        /// <summary>
        /// Scrapes the registry for .NET keys and lists them
        /// </summary>
        private void ScrapeRegistry()
        {
            var regKey = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\DataAccess\", false);
            if (regKey != null)
            {
                var verVal = (string)regKey.GetValue("Version");
                var revVal = (string)regKey.GetValue("FullInstallVer");
                this.AddEntry(verVal, revVal);
            }
        }

        /// <summary>
        /// Adds an entry to the list of versions.
        /// </summary>
        /// <param name="newEntry">The new entry.</param>
        /// <param name="subEntry">The sub entry.</param>
        private void AddEntry(string newEntry, string subEntry)
        {
            var itmX = new ListViewItem(newEntry, 0);
            this.lvStatus.Items.Add(itmX);
            var i = this.lvStatus.Items.Count - 1;
            this.lvStatus.Items[i].SubItems.Add(subEntry);
        }

        /// <summary>
        /// Handles the Click event of the cmdOK control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="System.EventArgs"/> instance containing the event data.</param>
        private void cmdOK_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
