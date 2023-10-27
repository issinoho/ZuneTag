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
    using System.Linq;
    using System.Management;
    using System.Net;
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
            this.InitializeComponent();

            // Empty trees
            this.tvOptions.Nodes.Clear();
            this.tvCheat.Nodes.Clear();

            // Make the dummy one visible while we build the real tree
            this.tvOptions.Visible = false;
            this.tvCheat.Visible = true;

            // Please wait...
            var newNode = new TreeNode("Gathering data, please wait...")
            {
                ImageIndex = 23,
                SelectedImageIndex = 23
            };
            this.tvCheat.Nodes.Add(newNode);

            // Wait and then gather data
            this.timer1.Enabled = true;
        }

        /// <summary>
        /// Builds the tree.
        /// </summary>
        private void BuildTree()
        {
            // Empty tree
            this.SuspendLayout();
            this.tvOptions.Nodes.Clear();

            // Top level branches
            var newNode = new TreeNode("Operating System")
            {
                ImageIndex = 10,
                SelectedImageIndex = 10
            };
            this.tvOptions.Nodes.Add(newNode);
            // OS children
            GetOs(newNode);

            newNode = new TreeNode("Computer")
            {
                ImageIndex = 0,
                SelectedImageIndex = 0
            };
            this.tvOptions.Nodes.Add(newNode);
            // Computer children
            GetComputer(newNode);

            newNode = new TreeNode("Owner")
            {
                ImageIndex = 12,
                SelectedImageIndex = 12
            };
            this.tvOptions.Nodes.Add(newNode);
            GetOwner(newNode);

            newNode = new TreeNode("Network")
            {
                ImageIndex = 11,
                SelectedImageIndex = 11
            };
            this.tvOptions.Nodes.Add(newNode);
            GetNetwork(newNode);

            newNode = new TreeNode("Storage")
            {
                ImageIndex = 6,
                SelectedImageIndex = 6
            };
            this.tvOptions.Nodes.Add(newNode);
            GetStorage(newNode);

            this.ResumeLayout();
        }

        /// <summary>
        /// Gets the storage information.
        /// </summary>
        /// <param name="newNode">The new node.</param>
        private static void GetStorage(TreeNode newNode)
        {
            try
            {
                var query1 = new ManagementObjectSearcher("select FreeSpace,Size,Name from Win32_LogicalDisk where DriveType=3");
                var queryCollection1 = query1.Get();
                foreach (var mo in queryCollection1.Cast<ManagementObject>())
                {
                    var freeSpace = Convert.ToUInt64(mo["FreeSpace"]);
                    var size = Convert.ToUInt64(mo["Size"]);
                    var childNode = new TreeNode(mo["Name"] + ": " + size / 1073741824 + " Gb (" + freeSpace / 1073741824 + " Gb free)")
                    {
                        ImageIndex = 15,
                        SelectedImageIndex = 15
                    };
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
        private static void GetOs(TreeNode newNode)
        {
            try
            {
                var query1 = new ManagementObjectSearcher("SELECT * FROM Win32_OperatingSystem");
                var queryCollection1 = query1.Get();
                foreach (var mo in queryCollection1.Cast<ManagementObject>())
                {
                    var childNode = new TreeNode(mo["Caption"].ToString())
                    {
                        ImageIndex = 7,
                        SelectedImageIndex = 7
                    };
                    newNode.Nodes.Add(childNode);
                    childNode = new TreeNode(mo["CSDVersion"].ToString())
                    {
                        ImageIndex = 8,
                        SelectedImageIndex = 8
                    };
                    newNode.Nodes.Add(childNode);
                }
            }
            catch (Exception)
            {
            }
        }

        /// <summary>
        /// Gets the network information.
        /// </summary>
        /// <param name="newNode">The new node.</param>
        private static void GetNetwork(TreeNode newNode)
        {
            try
            {
                // Domain stuff
                var query1 = new ManagementObjectSearcher("SELECT * FROM Win32_OperatingSystem");
                var queryCollection1 = query1.Get();
                foreach (var mo in queryCollection1.Cast<ManagementObject>())
                {
                    var childNode = new TreeNode(mo["CSName"].ToString())
                    {
                        ImageIndex = 22,
                        SelectedImageIndex = 22
                    };
                    newNode.Nodes.Add(childNode);
                }
                // Domain stuff
                query1 = new ManagementObjectSearcher("SELECT * FROM Win32_ComputerSystem");
                queryCollection1 = query1.Get();
                foreach (var mo in queryCollection1.Cast<ManagementObject>())
                {
                    var childNode = new TreeNode(mo["UserName"].ToString())
                    {
                        ImageIndex = 2,
                        SelectedImageIndex = 2
                    };
                    newNode.Nodes.Add(childNode);
                    childNode = new TreeNode(mo["Domain"].ToString())
                    {
                        ImageIndex = 21,
                        SelectedImageIndex = 21
                    };
                    newNode.Nodes.Add(childNode);
                }
                // IP Address
                var myHost = Dns.GetHostName();
                var myIp = Dns.GetHostEntry(myHost).AddressList[0].ToString();
                var ipNode = new TreeNode(myIp)
                {
                    ImageIndex = 20,
                    SelectedImageIndex = 20
                };
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
        private static void GetOwner(TreeNode newNode)
        {
            try
            {
                var query1 = new ManagementObjectSearcher("SELECT * FROM Win32_OperatingSystem");
                var queryCollection1 = query1.Get();
                foreach (var mo in queryCollection1.Cast<ManagementObject>())
                {
                    var childNode = new TreeNode(mo["RegisteredUser"].ToString())
                    {
                        ImageIndex = 3,
                        SelectedImageIndex = 3
                    };
                    newNode.Nodes.Add(childNode);
                    childNode = new TreeNode(mo["Organization"].ToString())
                    {
                        ImageIndex = 4,
                        SelectedImageIndex = 4
                    };
                    newNode.Nodes.Add(childNode);
                    childNode = new TreeNode(mo["SerialNumber"].ToString())
                    {
                        ImageIndex = 5,
                        SelectedImageIndex = 5
                    };
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
        private static void GetComputer(TreeNode newNode)
        {
            try
            {
                // Manufacturer details
                var query1 =
                    new ManagementObjectSearcher("SELECT * FROM Win32_ComputerSystem");
                var queryCollection1 = query1.Get();
                foreach (var mo in queryCollection1.Cast<ManagementObject>())
                {
                    var childNode = new TreeNode(mo["Manufacturer"].ToString())
                    {
                        ImageIndex = 14,
                        SelectedImageIndex = 14
                    };
                    newNode.Nodes.Add(childNode);
                    childNode = new TreeNode(mo["Model"].ToString())
                    {
                        ImageIndex = 13,
                        SelectedImageIndex = 13
                    };
                    newNode.Nodes.Add(childNode);
                }

                // Processor details
                query1 = new ManagementObjectSearcher("SELECT * FROM Win32_Processor");
                queryCollection1 = query1.Get();
                var count = 1;
                foreach (var mo in queryCollection1.Cast<ManagementObject>())
                {
                    var childNode = new TreeNode("CPU " + count++ + ": " + Regex.Replace(mo["Name"].ToString(), @"^\s+|\s+$", "") + " (" + mo["AddressWidth"] + " bit)")
                    {
                        ImageIndex = 17,
                        SelectedImageIndex = 17
                    };
                    newNode.Nodes.Add(childNode);
                }

                // Memory
                query1 = new ManagementObjectSearcher("SELECT * FROM Win32_PhysicalMemory");
                queryCollection1 = query1.Get();
                var totalCapacity = queryCollection1.Cast<ManagementObject>().Aggregate<ManagementObject, ulong>(0, (current, mo) => current + Convert.ToUInt64(mo["Capacity"]));
                var memNode = new TreeNode("Memory: " + totalCapacity / 1073741824 + " Gb")
                {
                    ImageIndex = 19,
                    SelectedImageIndex = 19
                };
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
            this.timer1.Enabled = false;

            // Tree
            this.BuildTree();
            this.tvOptions.SelectedNode = this.tvOptions.Nodes[0];

            // Now switch the trees
            this.tvCheat.Visible = false;
            this.tvOptions.Visible = true;
        }
    }
}
