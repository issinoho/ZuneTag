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
            InitializeComponent();

            // Clear list
            lvStatus.Columns.Add("Major Version", (lvStatus.Width / 2), HorizontalAlignment.Left);
            lvStatus.Columns.Add("Revision", (lvStatus.Width / 2) - 3, HorizontalAlignment.Left);
            lvStatus.Items.Clear();

            // Now get the versions from the reg
            ScrapeRegistry();
        }

        /// <summary>
        /// Scrapes the registry for .NET keys and lists them
        /// </summary>
        private void ScrapeRegistry()
        {
            Microsoft.Win32.RegistryKey regKey;
            Microsoft.Win32.RegistryKey revKey;

            regKey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\NET Framework Setup\NDP\", false);
            foreach ( string Keyname in regKey.GetSubKeyNames())
            {
                revKey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\NET Framework Setup\NDP\" + Keyname + @"\", false);
                string revVal = (string)revKey.GetValue("Version");
                AddEntry(Keyname, revVal);
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
            lvStatus.Items.Add(itmX);
            int i = (lvStatus.Items.Count - 1);
            lvStatus.Items[i].SubItems.Add(subEntry);
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
