//------------------------------------------------------------------
// Zune Meta Tag Editor
// System Information Class
//
// <copyright file="SysInfo.cs" company="The Drunken Bakery">
//     Copyright (c) 2009 The Drunken Bakery. All rights reserved.
// </copyright>
//
// Editor to update WMV meta tags for the Zune
// Displays current system information.
//
// Author: IRS
// $Revision: 1.1 $
//------------------------------------------------------------------

namespace DrunkenBakery.ZuneTag
{
    using System;
    using System.Management;
    using System.Text.RegularExpressions;
    using System.Windows.Forms;

    /// <summary>
    /// Reports on System Information
    /// </summary>
    public partial class SysInfo : Form
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="SysInfo"/> class.
        /// </summary>
        public SysInfo()
        {
            InitializeComponent();

            // Empty trees
            tvOptions.Nodes.Clear();
            tvCheat.Nodes.Clear();

            // Make the dummy one visible while we build the real tree
            tvOptions.Visible = false;
            tvCheat.Visible = true;

            // Please wait...
            TreeNode newNode = new TreeNode("Gathering data, please wait...");
            newNode.ImageIndex = 23;
            newNode.SelectedImageIndex = 23;
            tvCheat.Nodes.Add(newNode);

            // Wait and then gather data
            timer1.Enabled = true;
        }

        /// <summary>
        /// Builds the tree.
        /// </summary>
        private void BuildTree()
        {
            TreeNode newNode;

            // Empty tree
            this.SuspendLayout();
            tvOptions.Nodes.Clear();

            // Top level branches
            newNode = new TreeNode("Operating System");
            newNode.ImageIndex = 10;
            newNode.SelectedImageIndex = 10;
            tvOptions.Nodes.Add(newNode);
            // OS children
            GetOS(newNode);

            newNode = new TreeNode("Computer");
            newNode.ImageIndex = 0;
            newNode.SelectedImageIndex = 0;
            tvOptions.Nodes.Add(newNode);
            // Computer children
            GetComputer(newNode);

            newNode = new TreeNode("Owner");
            newNode.ImageIndex = 12;
            newNode.SelectedImageIndex = 12;
            tvOptions.Nodes.Add(newNode);
            GetOwner(newNode);

            newNode = new TreeNode("Network");
            newNode.ImageIndex = 11;
            newNode.SelectedImageIndex = 11;
            tvOptions.Nodes.Add(newNode);
            GetNetwork(newNode);

            newNode = new TreeNode("Storage");
            newNode.ImageIndex = 6;
            newNode.SelectedImageIndex = 6;
            tvOptions.Nodes.Add(newNode);
            GetStorage(newNode);

            this.ResumeLayout();
        }

        /// <summary>
        /// Gets the storage information.
        /// </summary>
        /// <param name="newNode">The new node.</param>
        private void GetStorage(TreeNode newNode)
        {
            try
            {
                ManagementObjectSearcher query1 = new ManagementObjectSearcher("select FreeSpace,Size,Name from Win32_LogicalDisk where DriveType=3");
                ManagementObjectCollection queryCollection1 = query1.Get();
                foreach (ManagementObject mo in queryCollection1)
                {
                    TreeNode childNode;
                    UInt64 FreeSpace = System.Convert.ToUInt64(mo["FreeSpace"]);
                    UInt64 Size = System.Convert.ToUInt64(mo["Size"]);
                    childNode = new TreeNode(mo["Name"].ToString() + ": " + (Size / 1073741824) + " Gb (" + (FreeSpace / 1073741824) + " Gb free)");
                    childNode.ImageIndex = 15;
                    childNode.SelectedImageIndex = 15;
                    newNode.Nodes.Add(childNode);
                }
            }
            catch (Exception)
            {
            }
        }

        /// <summary>
        /// Gets the OS details.
        /// </summary>
        /// <param name="newNode">The new node.</param>
        private void GetOS(TreeNode newNode)
        {
            try
            {
                ManagementObjectSearcher query1 = new ManagementObjectSearcher("SELECT * FROM Win32_OperatingSystem");
                ManagementObjectCollection queryCollection1 = query1.Get();
                foreach( ManagementObject mo in queryCollection1 ) 
                {
                    TreeNode childNode;
                    childNode = new TreeNode(mo["Caption"].ToString());
                    childNode.ImageIndex = 7;
                    childNode.SelectedImageIndex = 7;
                    newNode.Nodes.Add(childNode);
                    childNode = new TreeNode(mo["CSDVersion"].ToString());
                    childNode.ImageIndex = 8;
                    childNode.SelectedImageIndex = 8;
                    newNode.Nodes.Add(childNode);
                }   
            }
            catch(Exception)
            {
            }
        }

        /// <summary>
        /// Gets the network information.
        /// </summary>
        /// <param name="newNode">The new node.</param>
        private void GetNetwork(TreeNode newNode)
        {
            try
            {
                ManagementObjectSearcher query1;
                ManagementObjectCollection queryCollection1;

                // Domain stuff
                query1 = new ManagementObjectSearcher("SELECT * FROM Win32_OperatingSystem");
                queryCollection1 = query1.Get();
                foreach (ManagementObject mo in queryCollection1)
                {
                    TreeNode childNode;
                    childNode = new TreeNode(mo["CSName"].ToString());
                    childNode.ImageIndex = 22;
                    childNode.SelectedImageIndex = 22;
                    newNode.Nodes.Add(childNode);
                }
                // Domain stuff
                query1 = new ManagementObjectSearcher("SELECT * FROM Win32_ComputerSystem");
                queryCollection1 = query1.Get();
                foreach (ManagementObject mo in queryCollection1)
                {
                    TreeNode childNode;
                    childNode = new TreeNode(mo["UserName"].ToString());
                    childNode.ImageIndex = 2;
                    childNode.SelectedImageIndex = 2;
                    newNode.Nodes.Add(childNode);
                    childNode = new TreeNode(mo["Domain"].ToString());
                    childNode.ImageIndex = 21;
                    childNode.SelectedImageIndex = 21;
                    newNode.Nodes.Add(childNode);
                }
                // IP Address
                string myHost = System.Net.Dns.GetHostName();
                string myIP = System.Net.Dns.GetHostEntry(myHost).AddressList[0].ToString();
                TreeNode ipNode;
                ipNode = new TreeNode(myIP);
                ipNode.ImageIndex = 20;
                ipNode.SelectedImageIndex = 20;
                newNode.Nodes.Add(ipNode);
            }
            catch (Exception)
            {
            }
        }

        /// <summary>
        /// Gets the owner information.
        /// </summary>
        /// <param name="newNode">The new node.</param>
        private void GetOwner(TreeNode newNode)
        {
            try
            {
                ManagementObjectSearcher query1 = new ManagementObjectSearcher("SELECT * FROM Win32_OperatingSystem");
                ManagementObjectCollection queryCollection1 = query1.Get();
                foreach (ManagementObject mo in queryCollection1)
                {
                    TreeNode childNode;
                    childNode = new TreeNode(mo["RegisteredUser"].ToString());
                    childNode.ImageIndex = 3;
                    childNode.SelectedImageIndex = 3;
                    newNode.Nodes.Add(childNode);
                    childNode = new TreeNode(mo["Organization"].ToString());
                    childNode.ImageIndex = 4;
                    childNode.SelectedImageIndex = 4;
                    newNode.Nodes.Add(childNode);
                    childNode = new TreeNode(mo["SerialNumber"].ToString());
                    childNode.ImageIndex = 5;
                    childNode.SelectedImageIndex = 5;
                    newNode.Nodes.Add(childNode);
                }
            }
            catch (Exception)
            {
            }
        }

        /// <summary>
        /// Gets the computer details.
        /// </summary>
        /// <param name="newNode">The new node.</param>
        private void GetComputer(TreeNode newNode)
        {
            try
            {
                ManagementObjectSearcher query1;
                ManagementObjectCollection queryCollection1;

                // Manufacturer details
                query1 = new ManagementObjectSearcher("SELECT * FROM Win32_ComputerSystem");
                queryCollection1 = query1.Get();
                foreach (ManagementObject mo in queryCollection1)
                {
                    TreeNode childNode;
                    childNode = new TreeNode(mo["Manufacturer"].ToString());
                    childNode.ImageIndex = 14;
                    childNode.SelectedImageIndex = 14;
                    newNode.Nodes.Add(childNode);
                    childNode = new TreeNode(mo["Model"].ToString());
                    childNode.ImageIndex = 13;
                    childNode.SelectedImageIndex = 13;
                    newNode.Nodes.Add(childNode);
                }
                // Processor details
                query1 = new ManagementObjectSearcher("SELECT * FROM Win32_Processor");
                queryCollection1 = query1.Get();
                int count = 1;
                foreach (ManagementObject mo in queryCollection1)
                {
                    TreeNode childNode;
                    childNode = new TreeNode("CPU " + count++ + ": " + Regex.Replace(mo["Name"].ToString(), @"^\s+|\s+$", "") + " (" + mo["AddressWidth"].ToString() + " bit)");
                    childNode.ImageIndex = 17;
                    childNode.SelectedImageIndex = 17;
                    newNode.Nodes.Add(childNode);
                }
                // Memory
                UInt64 totalCapacity = 0;
                query1 = new ManagementObjectSearcher("SELECT * FROM Win32_PhysicalMemory");
                queryCollection1 = query1.Get();
                foreach (ManagementObject mo in queryCollection1)
                {
                    totalCapacity += System.Convert.ToUInt64(mo["Capacity"]);
                }
                TreeNode memNode;
                memNode = new TreeNode("Memory: " + (totalCapacity / 1073741824) + " Gb");
                memNode.ImageIndex = 19;
                memNode.SelectedImageIndex = 19;
                newNode.Nodes.Add(memNode);
            }
            catch (Exception)
            {
            }
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

        private void timer1_Tick(object sender, EventArgs e)
        {
            // Stop re-entrancy
            timer1.Enabled = false;

            // Tree
            BuildTree();
            tvOptions.SelectedNode = tvOptions.Nodes[0];

            // Now switch the trees
            tvCheat.Visible = false;
            tvOptions.Visible = true;
        }
    }
}
