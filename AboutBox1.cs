//------------------------------------------------------------------
// Zune Meta Tag Editor
// About Class
//
// <copyright file="AboutBox1.cs" company="The Drunken Bakery">
//     Copyright (c) 2009 The Drunken Bakery. All rights reserved.
// </copyright>
//
// Editor to update WMV meta tags for the Zune
// Standard About Us.
//
// Author: IRS
// $Revision: 1.1 $
//------------------------------------------------------------------

namespace DrunkenBakery.ZuneTag
{
    using System;
    using System.IO;
    using System.Reflection;
    using System.Runtime.Versioning;
    using System.Security;
    using System.Windows.Forms;

    using Microsoft.Win32;

    /// <summary>
    ///     Standard Cygnet About box.
    /// </summary>
    internal partial class AboutBox1 : Form
    {
        /// <summary>
        ///     Initializes a new instance of the <see cref="AboutBox1" /> class.
        /// </summary>
        public AboutBox1()
        {
            this.InitializeComponent();

            var asm = Assembly.GetExecutingAssembly();

            // Initialize the AboutBox to display the product information from the assembly information.
            // Change assembly information settings for your application through either:
            // - Project->Properties->Application->Assembly Information
            // - AssemblyInfo.cs
            this.Text = string.Format("About {0}", this.AssemblyTitle);
            this.labelProductName.Text = this.AssemblyProduct + " - " + this.AssemblyTitle;
            this.labelVersion.Text = string.Format("Version {0}", this.AssemblyVersion);
            this.labelCopyright.Text = this.AssemblyCopyright + DateTime.Today.Year;
            this.labelCompanyName.Text = this.AssemblyCompany;
            this.textBoxDescription.Text = this.AssemblyDescription + Environment.NewLine + Environment.NewLine + "Compiled on " + GetCompiledFrameworkName(asm) + Environment.NewLine + "Running on .NET Framework " + GetInstalledFrameworkVersion() + Environment.NewLine;

            // Use Reflection to get a list of depenedent assemblies
            this.textBoxDescription.AppendText(Environment.NewLine + "Dependent Assemblies:");
            var refs = asm.GetReferencedAssemblies();
            foreach (var myRef in refs)
            {
                this.textBoxDescription.AppendText(Environment.NewLine + myRef.Name + " v" + myRef.Version);
            }
        }

        /// <summary>
        ///     Gets the assembly title.
        /// </summary>
        /// <value>The assembly title.</value>
        public string AssemblyTitle
        {
            get
            {
                // Get all Title attributes on this assembly
                var attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyTitleAttribute), false);

                // If there is at least one Title attribute
                if (attributes.Length > 0)
                {
                    // Select the first one
                    var titleAttribute = (AssemblyTitleAttribute)attributes[0];

                    // If it is not an empty string, return it
                    if (titleAttribute.Title != string.Empty)
                    {
                        return titleAttribute.Title;
                    }
                }

                // If there was no Title attribute, or if the Title attribute was the empty string, return the .exe name
                return Path.GetFileNameWithoutExtension(Assembly.GetExecutingAssembly().CodeBase);
            }
        }

        /// <summary>
        ///     Gets the assembly version.
        /// </summary>
        /// <value>The assembly version.</value>
        public string AssemblyVersion => Assembly.GetExecutingAssembly().GetName().Version.ToString();

        /// <summary>
        ///     Gets the assembly description.
        /// </summary>
        /// <value>The assembly description.</value>
        public string AssemblyDescription
        {
            get
            {
                // Get all Description attributes on this assembly
                var attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyDescriptionAttribute), false);

                // If there aren't any Description attributes, return an empty string
                if (attributes.Length == 0)
                {
                    return string.Empty;
                }

                // If there is a Description attribute, return its value
                return ((AssemblyDescriptionAttribute)attributes[0]).Description;
            }
        }

        /// <summary>
        ///     Gets the assembly product.
        /// </summary>
        /// <value>The assembly product.</value>
        public string AssemblyProduct
        {
            get
            {
                // Get all Product attributes on this assembly
                var attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyProductAttribute), false);

                // If there aren't any Product attributes, return an empty string
                if (attributes.Length == 0)
                {
                    return string.Empty;
                }

                // If there is a Product attribute, return its value
                return ((AssemblyProductAttribute)attributes[0]).Product;
            }
        }

        /// <summary>
        ///     Gets the assembly copyright.
        /// </summary>
        /// <value>The assembly copyright.</value>
        public string AssemblyCopyright
        {
            get
            {
                // Get all Copyright attributes on this assembly
                var attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyCopyrightAttribute), false);

                // If there aren't any Copyright attributes, return an empty string
                if (attributes.Length == 0)
                {
                    return string.Empty;
                }

                // If there is a Copyright attribute, return its value
                return ((AssemblyCopyrightAttribute)attributes[0]).Copyright;
            }
        }

        /// <summary>
        ///     Gets the assembly company.
        /// </summary>
        /// <value>The assembly company.</value>
        public string AssemblyCompany
        {
            get
            {
                // Get all Company attributes on this assembly
                var attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyCompanyAttribute), false);

                // If there aren't any Company attributes, return an empty string
                if (attributes.Length == 0)
                {
                    return string.Empty;
                }

                // If there is a Company attribute, return its value
                return ((AssemblyCompanyAttribute)attributes[0]).Company;
            }
        }

        /// <summary>
        ///     Gets the human-readable .NET Framework version this assembly was compiled against.
        /// </summary>
        /// <param name="assembly">The assembly to inspect.</param>
        /// <returns>A display name such as ".NET Framework 4.8".</returns>
        /// <remarks>
        ///     <see cref="Assembly.ImageRuntimeVersion" /> only reports the CLR version (e.g.
        ///     "v4.0.30319"), which is identical for every .NET Framework release from 2.0
        ///     through 4.8. The <see cref="TargetFrameworkAttribute" /> MSBuild embeds from
        ///     TargetFrameworkVersion is the only reliable way to recover the actual target.
        /// </remarks>
        private static string GetCompiledFrameworkName(Assembly assembly)
        {
            var targetFramework = assembly.GetCustomAttribute<TargetFrameworkAttribute>();
            if (!string.IsNullOrEmpty(targetFramework?.FrameworkDisplayName))
            {
                return targetFramework.FrameworkDisplayName;
            }

            return targetFramework?.FrameworkName ?? ".NET " + assembly.ImageRuntimeVersion;
        }

        /// <summary>
        ///     Gets the .NET Framework version currently installed and running on this machine.
        /// </summary>
        /// <returns>A display name such as "4.8".</returns>
        /// <remarks>
        ///     <see cref="Environment.Version" /> only reports the CLR build number, which does
        ///     not distinguish between in-place 4.x updates. Microsoft's documented detection
        ///     method reads the "Release" value from the registry instead.
        /// </remarks>
        private static string GetInstalledFrameworkVersion()
        {
            try
            {
                using (var ndpKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry32)
                    .OpenSubKey(@"SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full\"))
                {
                    if (ndpKey?.GetValue("Release") is int releaseKey)
                    {
                        return CheckFor45PlusVersion(releaseKey);
                    }
                }
            }
            catch (SecurityException)
            {
            }
            catch (IOException)
            {
            }

            return "unknown (CLR v" + Environment.Version + ")";
        }

        /// <summary>
        ///     Maps a .NET Framework 4.5+ registry "Release" value to its version number.
        /// </summary>
        /// <param name="releaseKey">The release value read from the registry.</param>
        /// <returns>The corresponding .NET Framework version.</returns>
        /// <remarks>
        ///     Thresholds are from
        ///     https://learn.microsoft.com/dotnet/framework/migration-guide/how-to-determine-which-versions-are-installed.
        /// </remarks>
        private static string CheckFor45PlusVersion(int releaseKey)
        {
            if (releaseKey >= 533320)
            {
                return "4.8.1 or later";
            }

            if (releaseKey >= 528040)
            {
                return "4.8";
            }

            if (releaseKey >= 461808)
            {
                return "4.7.2";
            }

            if (releaseKey >= 461308)
            {
                return "4.7.1";
            }

            if (releaseKey >= 460798)
            {
                return "4.7";
            }

            if (releaseKey >= 394802)
            {
                return "4.6.2";
            }

            if (releaseKey >= 394254)
            {
                return "4.6.1";
            }

            if (releaseKey >= 393295)
            {
                return "4.6";
            }

            if (releaseKey >= 379893)
            {
                return "4.5.2";
            }

            if (releaseKey >= 378675)
            {
                return "4.5.1";
            }

            if (releaseKey >= 378389)
            {
                return "4.5";
            }

            return "unknown (release " + releaseKey + ")";
        }

        /// <summary>
        ///     Handles the Click event of the okButton control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="System.EventArgs" /> instance containing the event data.</param>
        private void OkButton_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}