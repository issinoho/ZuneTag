//------------------------------------------------------------------
// Zune Meta Tag Editor
// NET Versioning Class
//
// <copyright file="NETversions.cs" company="The Drunken Bakery">
//     Copyright (c) 2009 The Drunken Bakery. All rights reserved.
// </copyright>
//
// Editor to update WMV meta tags for the Zune
// Supplies installed NET versions.
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
    /// Reports on installed .NET versions
    /// </summary>
    public partial class NETversions : Form
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="NETversions"/> class.
        /// </summary>
        public NETversions()
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
            var regKey = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\NET Framework Setup\NDP\", false);
            if (regKey != null)
                foreach (var keyname in regKey.GetSubKeyNames())
                {
                    var revKey = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\NET Framework Setup\NDP\" + keyname + @"\", false);
                    if (revKey != null)
                    {
                        var revVal = (string) revKey.GetValue("Version");
                        this.AddEntry(keyname, revVal);
                    }
                }
        }

        /// <summary>
        /// Adds an entry to the list of versions.
        /// </summary>
        /// <param name="newEntry">The new entry.</param>
        /// <param name="subEntry">The sub entry.</param>
        private void AddEntry(string newEntry, string subEntry)
        {
            ListViewItem itmX = null;

            itmX = new ListViewItem(newEntry, 0);
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
