//------------------------------------------------------------------
// Zune Meta Tag Editor
// About Class
//
// <copyright file="TagEditor.cs" company="The Drunken Bakery">
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
    using System.Reflection;
    using System.Windows.Forms;

    /// <summary>
    /// Standard Cygnet About box.
    /// </summary>
    partial class AboutBox1 : Form
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="AboutBox1"/> class.
        /// </summary>
        public AboutBox1()
        {
            this.InitializeComponent();

            Assembly asm = Assembly.GetExecutingAssembly();

            // Initialize the AboutBox to display the product information from the assembly information.
            // Change assembly information settings for your application through either:
            // - Project->Properties->Application->Assembly Information
            // - AssemblyInfo.cs
            this.Text = String.Format("About {0}", this.AssemblyTitle);
            this.labelProductName.Text = this.AssemblyProduct + " - " + this.AssemblyTitle;
            this.labelVersion.Text = String.Format("Version {0}", this.AssemblyVersion);
            this.labelCopyright.Text = this.AssemblyCopyright;
            this.labelCompanyName.Text = this.AssemblyCompany;
            this.textBoxDescription.Text = this.AssemblyDescription +
                                            System.Environment.NewLine +
                                            System.Environment.NewLine +
                                            "Compiled on .NET " + asm.ImageRuntimeVersion.ToString() +
                                            System.Environment.NewLine +
                                            "Running on .NET v" + Environment.Version.ToString() + 
                                            System.Environment.NewLine;

            // Use Reflection to get a list of depenedent assemblies
            this.textBoxDescription.AppendText(System.Environment.NewLine + "Dependent Assemblies:");
            AssemblyName[] refs = asm.GetReferencedAssemblies();
            foreach (AssemblyName myRef in refs)
            {
                this.textBoxDescription.AppendText(System.Environment.NewLine + myRef.Name + " v" + myRef.Version.ToString());
            }
        }

        #region Assembly Attribute Accessors

        /// <summary>
        /// Gets the assembly title.
        /// </summary>
        /// <value>The assembly title.</value>
        public string AssemblyTitle
        {
            get
            {
                // Get all Title attributes on this assembly
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyTitleAttribute), false);

                // If there is at least one Title attribute
                if (attributes.Length > 0)
                {
                    // Select the first one
                    AssemblyTitleAttribute titleAttribute = (AssemblyTitleAttribute)attributes[0];

                    // If it is not an empty string, return it
                    if (titleAttribute.Title != string.Empty)
                    {
                        return titleAttribute.Title;
                    }
                }

                // If there was no Title attribute, or if the Title attribute was the empty string, return the .exe name
                return System.IO.Path.GetFileNameWithoutExtension(Assembly.GetExecutingAssembly().CodeBase);
            }
        }

        /// <summary>
        /// Gets the assembly version.
        /// </summary>
        /// <value>The assembly version.</value>
        public string AssemblyVersion
        {
            get
            {
                return Assembly.GetExecutingAssembly().GetName().Version.ToString();
            }
        }

        /// <summary>
        /// Gets the assembly description.
        /// </summary>
        /// <value>The assembly description.</value>
        public string AssemblyDescription
        {
            get
            {
                // Get all Description attributes on this assembly
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyDescriptionAttribute), false);

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
        /// Gets the assembly product.
        /// </summary>
        /// <value>The assembly product.</value>
        public string AssemblyProduct
        {
            get
            {
                // Get all Product attributes on this assembly
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyProductAttribute), false);

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
        /// Gets the assembly copyright.
        /// </summary>
        /// <value>The assembly copyright.</value>
        public string AssemblyCopyright
        {
            get
            {
                // Get all Copyright attributes on this assembly
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyCopyrightAttribute), false);

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
        /// Gets the assembly company.
        /// </summary>
        /// <value>The assembly company.</value>
        public string AssemblyCompany
        {
            get
            {
                // Get all Company attributes on this assembly
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyCompanyAttribute), false);

                // If there aren't any Company attributes, return an empty string
                if (attributes.Length == 0)
                {
                    return string.Empty;
                }

                // If there is a Company attribute, return its value
                return ((AssemblyCompanyAttribute)attributes[0]).Company;
            }
        }
        #endregion

        /// <summary>
        /// Handles the Click event of the okButton control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="System.EventArgs"/> instance containing the event data.</param>
        private void OkButton_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
