//------------------------------------------------------------------
// Zune Meta Tag Editor
// Main Form
//
// <copyright file="Form1.cs" company="The Drunken Bakery">
//     Copyright (c) 2009 The Drunken Bakery. All rights reserved.
// </copyright>
//
// Editor to update WMV meta tags for the Zune
// Main application form which drives all functionality.
//
// Author: IRS
// $Revision: 1.11 $
//------------------------------------------------------------------

namespace DrunkenBakery.ZuneTag
{
    using System;
    using System.Collections.Generic;
    using System.Drawing;
    using System.Drawing.Drawing2D;
    using System.IO;
    using System.Net;
    using System.Runtime.InteropServices;
    using System.ServiceModel;
    using System.Windows.Forms;
    using Amazon.ECS;
    using JockerSoft.Media;
    using WMFSDKWrapper;

    /// <summary>
    /// Main application form which drives all functionality.
    /// </summary>
    public partial class Form1 : Form
    {
        #region Enums
        /// <summary>
        /// Severity of logging entry
        /// </summary>
        private enum LogType { Success, Fail, Info }
        private enum MediaType {Unknown, Video, Movie, Music, TV}
        #endregion

        #region Constants
        const int ScreenRefresh = 1;
        const int ScreenLines = 1000;
        const string ThisApp = "Zune Tag Editor";
        const string ThisPublisher = "The Drunken Bakery";
        const ushort Stream = 65535;
        const ushort newStream = 0;
        const ushort Language = 0;
        const string TypeVideo = "BD-30-98-DB-B3-3A-AB-4F-8A-37-1A-99-5F-7F-F7-4B";
        const string TypeMovie = "C9-7F-B8-A9-47-BD-F0-4B-AC-4F-65-5B-89-F7-D8-68";
        const string TypeMusic = "E2-89-E6-E3-8C-BA-30-43-96-DF-A0-EE-EF-FA-68-76";
        const string TypeTV = "8A-25-7F-BA-F7-62-A9-47-B2-1F-46-51-C4-2A-00-0E";
        const string AmazonID = "AKIAI344DCI3P6HJXGOA";
        const string AmazonKey = "dM23mxiQzaMEy0O2I1qIacqJijV/JyqIlFRgTP+Y";
        #endregion

        #region Class Variables
        private Form frmNET;
        private Form frmMDAC;
        private Form frmInfo;
        private Form frmAbout;
        List<ListViewItem> lvitems = new List<ListViewItem>();
        private System.Threading.Timer _ScreenLogTimer;
        private System.Threading.TimerCallback _ScreenLogTimerCallback;
        List<Attribute> _attributes = new List<Attribute>();
        private ushort indexPrimaryVideo;
        private ushort indexSecondaryVideo;
        private int pageNumber = 0;
        private int safePageNumber = 0;
        #endregion

        /// <summary>
        /// Initializes a new instance of the <see cref="Form1"/> class.
        /// </summary>
        public Form1()
        {
            InitializeComponent();

            // Upgrade settings from older version
            System.Reflection.Assembly a = System.Reflection.Assembly.GetExecutingAssembly();
            Version appVersion = a.GetName().Version;
            string appVersionString = appVersion.ToString();

            if (Properties.Settings.Default.ApplicationVersion != appVersion.ToString())
            {
                Properties.Settings.Default.Upgrade();
                Properties.Settings.Default.ApplicationVersion = appVersionString;
            }

            // Form title bar
            this.Text = ThisApp + " freshly baked at " + ThisPublisher;

            // Tooltips
            toolTip1.SetToolTip(cmdSave, "Change video type");
            toolTip2.SetToolTip(cmdModify, "Save this attribute back to the file");
            toolTip3.SetToolTip(cmdMovieSave, "Update Movie with these tags");
            toolTip4.SetToolTip(cmdMusicSave, "Update Music Video with these tags");
            toolTip5.SetToolTip(cmdTVSave, "Update TV Show with these tags");
            toolTip6.SetToolTip(cmdVideoSave, "Update Video with these tags");
            toolTip7.SetToolTip(cmdReset, "Hard reset the video type attributes");
            toolTip8.SetToolTip(cmdCopyAz, "Copy the tags back ready for saving");
            toolTip9.SetToolTip(cmdAmazonSearch, "Search Amazon for videos that match");
            toolTip10.SetToolTip(cmdBrowse, "Open a WMV media file");

            // Initialise Event Views
            InitEventView(lvStatus);

            // Initialise Media Types
            InitMediaTypes();

            // Logging
            AddLogEntry("--------------------------------------------------", LogType.Info);
            AddLogEntry("Welcome to the " + ThisApp + " v" + appVersionString, LogType.Info);
            AddLogEntry("Ready.");

            // Start Timers
            _ScreenLogTimerCallback = new System.Threading.TimerCallback(_ScreenLogTimer_Elapsed);
            _ScreenLogTimer = new System.Threading.Timer(_ScreenLogTimerCallback, null, (Convert.ToInt32(ScreenRefresh) * 1000), System.Threading.Timeout.Infinite);
        }

        /// <summary>
        /// Inits the media types.
        /// </summary>
        private void InitMediaTypes()
        {
            cbMediaType.Items.Clear();
            cbMediaType.Items.Add("Unknown");
            cbMediaType.Items.Add("Video");
            cbMediaType.Items.Add("Movies");
            cbMediaType.Items.Add("Music Videos");
            cbMediaType.Items.Add("TV Shows");
            cbMediaType.SelectedIndex = 0;
        }

        /// <summary>
        /// Updates the status view by flushing the buffer.
        /// </summary>
        /// <param name="sender">The sender.</param>
        private void _ScreenLogTimer_Elapsed(object sender)
        {
            CycleStatusView();
            _ScreenLogTimer.Change((Convert.ToInt32(ScreenRefresh) * 1000), System.Threading.Timeout.Infinite);
        }

        /// <summary>
        /// Cycles the status view.
        /// </summary>
        private void CycleStatusView()
        {
            PauseOutput(lvStatus);
            FlushOutput(lvStatus);
            ResumeOutput(lvStatus);
        }

        private delegate void FlushOutputDelegate(ListView lv);

        /// <summary>
        /// Flushes the output.
        /// </summary>
        /// <param name="lv">The lv.</param>
        private void FlushOutput(ListView lv)
        {
            if (this.InvokeRequired)
            {
                this.BeginInvoke(new FlushOutputDelegate(FlushOutput), new object[] { lv });
                return;
            }

            if (lvitems.Count > 0)
            {
                if (lv.Items.Count >= Convert.ToInt32(ScreenLines)) lv.Items.Clear();
                lv.BeginUpdate();
                lv.Items.AddRange(lvitems.ToArray());
                lv.EnsureVisible(lv.Items.Count - 1);
                lv.EndUpdate();
                lvitems.Clear();
            }
        }

        private delegate void PauseOutputDelegate(ListView lv);

        /// <summary>
        /// Pauses the output.
        /// </summary>
        /// <param name="lv">The lv.</param>
        private void PauseOutput(ListView lv)
        {
            if (this.InvokeRequired)
            {
                this.BeginInvoke(new PauseOutputDelegate(PauseOutput), new object[] { lv });
                return;
            }

            lv.BeginUpdate();
        }

        private delegate void ResumeOutputDelegate(ListView lv);
        /// <summary>
        /// Resumes the output.
        /// </summary>
        /// <param name="lv">The lv.</param>
        private void ResumeOutput(ListView lv)
        {
            if (this.InvokeRequired)
            {
                this.BeginInvoke(new ResumeOutputDelegate(ResumeOutput), new object[] { lv });
                return;
            }

            lv.EndUpdate();
        }

        /// <summary>
        /// Inits the event view.
        /// </summary>
        /// <param name="lvX">The lv X.</param>
        private void InitEventView(ListView lvX)
        {
            lvX.Columns.Add("Time", 140, HorizontalAlignment.Left);
            lvX.Columns.Add("Event Details", 1000, HorizontalAlignment.Left);
            lvX.Items.Clear();
        }

        /// <summary>
        /// Adds the log entry.
        /// </summary>
        /// <param name="newEntry">The new entry.</param>
        private void AddLogEntry(string newEntry)
        {
            AddLogEntry(newEntry, LogType.Success);
        }

        /// <summary>
        /// Adds the log entry.
        /// </summary>
        /// <param name="newEntry">The new entry.</param>
        /// <param name="whichLog">The which log.</param>
        private void AddLogEntry(string newEntry, LogType whichLog)
        {
            switch (whichLog)
            {
                case LogType.Success:
                    lvitems.Add(new ListViewItem(DateTime.Now.ToString(), 0));
                    break;

                case LogType.Fail:
                    lvitems.Add(new ListViewItem(DateTime.Now.ToString(), 1));
                    break;

                case LogType.Info:
                    lvitems.Add(new ListViewItem(DateTime.Now.ToString(), 2));
                    break;
            }

            int i = (lvitems.Count - 1);
            lvitems[i].SubItems.Add(newEntry);
            slStatus.Text = newEntry;
        }

        /// <summary>
        /// Handles the Click event of the cmdBrowse control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="System.EventArgs"/> instance containing the event data.</param>
        private void cmdBrowse_Click(object sender, EventArgs e)
        {
            // Get file
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.InitialDirectory = Environment.CurrentDirectory;
            openFileDialog1.Filter = "Windows Media Video (*.wmv)|*.wmv"; 
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.RestoreDirectory = false;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                lblMediaFile.Text = openFileDialog1.FileName;
                RegisterNewMediaFile();
            }

        }

        /// <summary>
        /// Registers the new media file.
        /// </summary>
        private void RegisterNewMediaFile()
        {
            if (!System.IO.File.Exists(lblMediaFile.Text))
            {
                AddLogEntry("Can't load - media file not found", LogType.Fail);
                return;
            }
            else
            {
                // Grab still frame, if possible
                try
                {
                    Size size = new Size();
                    size.Height = pictureBox1.Height;
                    size.Width = pictureBox1.Width;
                    this.pictureBox1.Image = FrameGrabber.GetFrameFromVideo(lblMediaFile.Text, 0.2d, size);
                    SetImage(pictureBox1);
                }
                catch (InvalidVideoFileException ex)
                {
                    AddLogEntry(ex.Message, LogType.Fail);
                }
                catch (StackOverflowException)
                {
                    AddLogEntry("The target image size is too big", LogType.Fail);
                }

                // Make sure all supported attributes are defined
                AddMissingAttributes();

                // Refresh attributes from file
                InspectFile();

                // Logging
                AddLogEntry(lblMediaFile.Text + " successfully loaded", LogType.Success);
            }
        }

        /// <summary>
        /// Generates the image dimensions.
        /// </summary>
        /// <param name="currW">The curr W.</param>
        /// <param name="currH">The curr H.</param>
        /// <param name="destW">The dest W.</param>
        /// <param name="destH">The dest H.</param>
        /// <returns></returns>
        public Size GenerateImageDimensions(int currW, int currH, int destW, int destH)
        {
            //double to hold the final multiplier to use when scaling the image
            double multiplier = 0;

            //string for holding layout
            string layout;

            //determine if it's Portrait or Landscape
            if (currH > currW) layout = "portrait";
            else layout = "landscape";

            switch (layout.ToLower())
            {
                case "portrait":
                    //calculate multiplier on heights
                    if (destH > destW)
                    {
                        multiplier = (double)destW / (double)currW;
                    }

                    else
                    {
                        multiplier = (double)destH / (double)currH;
                    }
                    break;
                case "landscape":
                    //calculate multiplier on widths
                    if (destH > destW)
                    {
                        multiplier = (double)destW / (double)currW;
                    }

                    else
                    {
                        multiplier = (double)destH / (double)currH;
                    }
                    break;
            }

            //return the new image dimensions
            return new Size((int)(currW * multiplier), (int)(currH * multiplier));
        }

        /// <summary>
        /// Sets the image.
        /// </summary>
        /// <param name="pb">The pb.</param>
        private void SetImage(PictureBox pb)
        {
            try
            {
                //create a temp image
                System.Drawing.Image img = pb.Image;

                //calculate the size of the image
                Size imgSize = GenerateImageDimensions(img.Width, img.Height, this.pictureBox1.Width, this.pictureBox1.Height);

                //create a new Bitmap with the proper dimensions
                Bitmap finalImg = new Bitmap(img, imgSize.Width, imgSize.Height);

                //create a new Graphics object from the image
                Graphics gfx = Graphics.FromImage(img);

                //clean up the image (take care of any image loss from resizing)
                gfx.InterpolationMode = InterpolationMode.HighQualityBicubic;

                //empty the PictureBox
                pb.Image = null;

                //center the new image
                pb.SizeMode = PictureBoxSizeMode.CenterImage;

                //set the new image
                pb.Image = finalImg;
            }
            catch (System.Exception e)
            {
                MessageBox.Show(e.Message);
            }
        }

        /// <summary>
        /// Handles the Click event of the nETVersionsToolStripMenuItem control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="System.EventArgs"/> instance containing the event data.</param>
        private void nETVersionsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (frmNET == null) frmNET = new NETversions();
            frmNET.ShowDialog();
        }

        /// <summary>
        /// Handles the Click event of the mDACVersionsToolStripMenuItem control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="System.EventArgs"/> instance containing the event data.</param>
        private void mDACVersionsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (frmMDAC == null) frmMDAC = new MDACversions();
            frmMDAC.ShowDialog();
        }

        /// <summary>
        /// Handles the Click event of the systemInformationToolStripMenuItem control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="System.EventArgs"/> instance containing the event data.</param>
        private void systemInformationToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (frmInfo == null) frmInfo = new SysInfo();
            frmInfo.ShowDialog();
        }

        /// <summary>
        /// Handles the Click event of the aboutToolStripMenuItem control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="System.EventArgs"/> instance containing the event data.</param>
        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (frmAbout == null) frmAbout = new AboutBox1();
            frmAbout.ShowDialog();
        }

        /// <summary>
        /// Handles the Click event of the exitToolStripMenuItem1 control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="System.EventArgs"/> instance containing the event data.</param>
        private void exitToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        /// <summary>
        /// Inspects the file.
        /// </summary>
        private void InspectFile()
        {
            bool isVideo = false;
            int thisIndex = 0;

            // Destroy existing attribute list
            _attributes.Clear();
            cbAttributes.Items.Clear();

            // Update from file
            if (ShowAttributes3(lblMediaFile.Text, Stream))
            {
                for (int i = 0; i < _attributes.Count; i++)
                {
                    // Create new instance
                    Attribute u = _attributes[i];

                    // Add to combo box
                    cbAttributes.Items.Add(u);

                    // Is this a video
                    if (u.Name == "WM/MediaClassPrimaryID")
                    {
                        indexPrimaryVideo = u.Index;
                        isVideo = (u.Value == TypeVideo);
                        if (isVideo) thisIndex = 1;
                    }

                    // Type of video?
                    if (u.Name == "WM/MediaClassSecondaryID")
                    {
                        indexSecondaryVideo = u.Index;

                        if (isVideo)
                        {
                            switch (u.Value)
                            {
                                case TypeMovie:
                                    thisIndex = 2;
                                    break;

                                case TypeMusic:
                                    thisIndex = 3;
                                    break;

                                case TypeTV:
                                    thisIndex = 4;
                                    break;
                            }
                        }
                    }
                }

                cbAttributes.SelectedIndex = 0;
                cbMediaType.SelectedIndex = thisIndex;

                // Update fields
                txtSearchCriteria.Text = "";
                switch (thisIndex)
                {
                    case 1:
                        LoadVideoAttributes();
                        break;

                    case 2:
                        LoadMovieAttributes();
                        break;

                    case 3:
                        LoadMusicAttributes();
                        break;

                    case 4:
                        LoadTVAttributes();
                        break;
                }

                // If no title has been established then look at filename
                if (txtSearchCriteria.Text.Length == 0)
                {
                    string filename = Path.GetFileNameWithoutExtension(lblMediaFile.Text);
                    filename = filename.Replace(".", " ");
                    filename = filename.Replace("_", " ");
                    filename = filename.Replace("-", " ");
                    txtSearchCriteria.Text = filename;
                }
            }
        }

        /// <summary>
        /// Creates a metadata editor and opens the file.
        /// </summary>
        /// <param name="pwszInFile">The PWSZ in file.</param>
        /// <param name="ppEditor">The pp editor.</param>
        /// <returns></returns>
        public bool EditorOpenFile(string pwszInFile, out IWMMetadataEditor ppEditor)
        {
            ppEditor = null;

            try
            {
                WMFSDKFunctions.WMCreateEditor(out ppEditor);

                ppEditor.Open(pwszInFile);
            }
            catch (System.Runtime.InteropServices.COMException e)
            {
                AddLogEntry(e.Message, LogType.Fail);
                return (false);
            }

            return (true);
        }

        /// <summary>
        /// Displays the specified attribute.
        /// </summary>
        /// <param name="wIndex">Index of the w.</param>
        /// <param name="wStream">The w stream.</param>
        /// <param name="pwszName">Name of the PWSZ.</param>
        /// <param name="AttribDataType">Type of the attrib data.</param>
        /// <param name="wLangID">The w lang ID.</param>
        /// <param name="pbValue">The pb value.</param>
        /// <param name="dwValueLen">The dw value len.</param>
        public void PrintAttribute(ushort wIndex,
                             ushort wStream,
                             string pwszName,
                             WMT_ATTR_DATATYPE AttribDataType,
                             ushort wLangID,
                             byte[] pbValue,
                             uint dwValueLen)
        {
            string pwszValue = String.Empty;

            //
            // Make the data type string
            //
            string pwszType = "Unknown";
            string[] pTypes = { "DWORD", "STRING", "BINARY", "BOOL", "QWORD", "WORD", "GUID" };

            if (pTypes.Length > Convert.ToInt32(AttribDataType))
            {
                pwszType = pTypes[Convert.ToInt32(AttribDataType)];
            }

            //
            // The attribute value.
            //
            switch (AttribDataType)
            {
                // String
                case WMFSDKWrapper.WMT_ATTR_DATATYPE.WMT_TYPE_STRING:

                    if (0 == dwValueLen)
                    {
                        pwszValue = "***** NULL *****";
                    }
                    else
                    {
                        if ((0xFE == Convert.ToInt16(pbValue[0])) &&
                             (0xFF == Convert.ToInt16(pbValue[1])))
                        {
                            pwszValue = "\"UTF-16LE BOM+\"";

                            if (4 <= dwValueLen)
                            {
                                for (int i = 0; i < pbValue.Length - 2; i += 2)
                                {
                                    pwszValue += Convert.ToString(BitConverter.ToChar(pbValue, i));
                                }
                            }

                            pwszValue = pwszValue + "\"";
                        }
                        else if ((0xFF == Convert.ToInt16(pbValue[0])) &&
                                  (0xFE == Convert.ToInt16(pbValue[1])))
                        {
                            pwszValue = "\"UTF-16BE BOM+\"";
                            if (4 <= dwValueLen)
                            {
                                for (int i = 0; i < pbValue.Length - 2; i += 2)
                                {
                                    pwszValue += Convert.ToString(BitConverter.ToChar(pbValue, i));
                                }
                            }

                            pwszValue = pwszValue + "\"";
                        }
                        else
                        {
                            pwszValue = "\"";
                            if (2 <= dwValueLen)
                            {
                                for (int i = 0; i < pbValue.Length - 2; i += 2)
                                {
                                    pwszValue += Convert.ToString(BitConverter.ToChar(pbValue, i));
                                }
                            }

                            pwszValue += "\"";
                        }
                    }
                    break;

                // Binary
                case WMFSDKWrapper.WMT_ATTR_DATATYPE.WMT_TYPE_BINARY:

                    pwszValue = "[" + dwValueLen.ToString() + " bytes]";
                    break;

                // Boolean
                case WMFSDKWrapper.WMT_ATTR_DATATYPE.WMT_TYPE_BOOL:

                    if (BitConverter.ToBoolean(pbValue, 0))
                    {
                        pwszValue = "True";
                    }
                    else
                    {
                        pwszValue = "False";
                    }
                    break;

                // DWORD
                case WMFSDKWrapper.WMT_ATTR_DATATYPE.WMT_TYPE_DWORD:

                    uint dwValue = BitConverter.ToUInt32(pbValue, 0);
                    pwszValue = dwValue.ToString();
                    break;

                // QWORD
                case WMFSDKWrapper.WMT_ATTR_DATATYPE.WMT_TYPE_QWORD:

                    ulong qwValue = BitConverter.ToUInt64(pbValue, 0);
                    pwszValue = qwValue.ToString();
                    break;

                // WORD
                case WMFSDKWrapper.WMT_ATTR_DATATYPE.WMT_TYPE_WORD:

                    uint wValue = BitConverter.ToUInt16(pbValue, 0);
                    pwszValue = wValue.ToString();
                    break;

                // GUID
                case WMFSDKWrapper.WMT_ATTR_DATATYPE.WMT_TYPE_GUID:

                    pwszValue = BitConverter.ToString(pbValue, 0, pbValue.Length);
                    break;

                default:

                    break;
            }

            // Add to attribute list
            Attribute _attribute = new Attribute(wIndex, pwszName.Substring(0, pwszName.Length - 1), pwszValue, AttribDataType);

            // Add to list
            _attributes.Add(_attribute);
        }

        /// <summary>
        /// Displays all attributes for the specified stream.
        /// </summary>
        /// <param name="pwszFileName">Name of the PWSZ file.</param>
        /// <param name="wStreamNum">The w stream num.</param>
        /// <returns></returns>
        public bool ShowAttributes(string pwszFileName, ushort wStreamNum)
        {
            try
            {
                IWMMetadataEditor MetadataEditor;
                IWMHeaderInfo3 HeaderInfo3;
                ushort wAttributeCount;

                WMFSDKFunctions.WMCreateEditor(out MetadataEditor);

                MetadataEditor.Open(pwszFileName);

                HeaderInfo3 = (IWMHeaderInfo3)MetadataEditor;

                HeaderInfo3.GetAttributeCount(wStreamNum, out wAttributeCount);

                for (ushort wAttribIndex = 0; wAttribIndex < wAttributeCount; wAttribIndex++)
                {
                    WMT_ATTR_DATATYPE wAttribType;
                    string pwszAttribName = null;
                    byte[] pbAttribValue = null;
                    ushort wAttribNameLen = 0;
                    ushort wAttribValueLen = 0;

                    HeaderInfo3.GetAttributeByIndex(wAttribIndex,
                                                     ref wStreamNum,
                                                     pwszAttribName,
                                                     ref wAttribNameLen,
                                                     out wAttribType,
                                                     pbAttribValue,
                                                     ref wAttribValueLen);

                    pbAttribValue = new byte[wAttribValueLen];
                    pwszAttribName = new String((char)0, wAttribNameLen);

                    HeaderInfo3.GetAttributeByIndex(wAttribIndex,
                                                     ref wStreamNum,
                                                     pwszAttribName,
                                                     ref wAttribNameLen,
                                                     out wAttribType,
                                                     pbAttribValue,
                                                     ref wAttribValueLen);

                    PrintAttribute(wAttribIndex, wStreamNum, pwszAttribName, wAttribType, 0, pbAttribValue, wAttribValueLen);
                }
            }
            catch (Exception e)
            {
                AddLogEntry(e.Message, LogType.Fail);
                return (false);
            }

            return (true);
        }

        /// <summary>
        /// Displays all attributes for the specified stream, with support for GetAttributeByIndexEx.
        /// </summary>
        /// <param name="pwszFileName">Name of the PWSZ file.</param>
        /// <param name="wStreamNum">The w stream num.</param>
        /// <returns></returns>
        public bool ShowAttributes3(string pwszFileName, ushort wStreamNum)
        {
            try
            {
                IWMMetadataEditor MetadataEditor;
                IWMHeaderInfo3 HeaderInfo3;
                ushort wAttributeCount = 0;

                WMFSDKFunctions.WMCreateEditor(out MetadataEditor);

                MetadataEditor.Open(pwszFileName);

                HeaderInfo3 = (IWMHeaderInfo3)MetadataEditor;

                HeaderInfo3.GetAttributeCountEx(wStreamNum, out wAttributeCount);

                for (ushort wAttribIndex = 0; wAttribIndex < wAttributeCount; wAttribIndex++)
                {
                    WMT_ATTR_DATATYPE wAttribType;
                    ushort wLangIndex = 0;
                    string pwszAttribName = null;
                    byte[] pbAttribValue = null;
                    ushort wAttribNameLen = 0;
                    uint dwAttribValueLen = 0;

                    HeaderInfo3.GetAttributeByIndexEx(wStreamNum,
                                                       wAttribIndex,
                                                       pwszAttribName,
                                                       ref wAttribNameLen,
                                                       out wAttribType,
                                                       out wLangIndex,
                                                       pbAttribValue,
                                                       ref dwAttribValueLen);

                    pwszAttribName = new String((char)0, wAttribNameLen);
                    pbAttribValue = new byte[dwAttribValueLen];

                    HeaderInfo3.GetAttributeByIndexEx(wStreamNum,
                                                       wAttribIndex,
                                                       pwszAttribName,
                                                       ref wAttribNameLen,
                                                       out wAttribType,
                                                       out wLangIndex,
                                                       pbAttribValue,
                                                       ref dwAttribValueLen);

                    PrintAttribute(wAttribIndex, wStreamNum, pwszAttribName, wAttribType, 0, pbAttribValue, dwAttribValueLen);
                }

                // Close file
                MetadataEditor.Close();
            }
            catch (Exception e)
            {
                AddLogEntry(e.Message, LogType.Fail);
                return (false);
            }

            return (true);
        }

        /// <summary>
        /// Delete the attribute at the specified index.
        /// </summary>
        /// <param name="pwszFileName">Name of the PWSZ file.</param>
        /// <param name="wStreamNum">The w stream num.</param>
        /// <param name="wAttribIndex">Index of the w attrib.</param>
        /// <returns></returns>
        public bool DeleteAttrib(string pwszFileName, ushort wStreamNum, ushort wAttribIndex)
        {
            try
            {
                IWMMetadataEditor MetadataEditor;
                IWMHeaderInfo3 HeaderInfo3;

                WMFSDKFunctions.WMCreateEditor(out MetadataEditor);

                MetadataEditor.Open(pwszFileName);

                HeaderInfo3 = (IWMHeaderInfo3)MetadataEditor;

                HeaderInfo3.DeleteAttribute(wStreamNum, wAttribIndex);

                MetadataEditor.Flush();

                MetadataEditor.Close();
            }
            catch (Exception e)
            {
                AddLogEntry(e.Message, LogType.Fail);
                return (false);
            }

            return (true);
        }

        /// <summary>
        /// Converts attributes to byte arrays.
        /// </summary>
        /// <param name="AttribDataType">Type of the attrib data.</param>
        /// <param name="pwszValue">The PWSZ value.</param>
        /// <param name="pbValue">The pb value.</param>
        /// <param name="nValueLength">Length of the n value.</param>
        /// <returns></returns>
        public bool TranslateAttrib(WMT_ATTR_DATATYPE AttribDataType, string pwszValue, out byte[] pbValue, out int nValueLength)
        {
            switch (AttribDataType)
            {
                case WMFSDKWrapper.WMT_ATTR_DATATYPE.WMT_TYPE_DWORD:

                    nValueLength = 4;
                    uint[] pdwAttribValue = new uint[1] { Convert.ToUInt32(pwszValue) };

                    pbValue = new Byte[nValueLength];
                    Buffer.BlockCopy(pdwAttribValue, 0, pbValue, 0, nValueLength);

                    return (true);

                case WMFSDKWrapper.WMT_ATTR_DATATYPE.WMT_TYPE_WORD:

                    nValueLength = 2;
                    ushort[] pwAttribValue = new ushort[1] { Convert.ToUInt16(pwszValue) };

                    pbValue = new Byte[nValueLength];
                    Buffer.BlockCopy(pwAttribValue, 0, pbValue, 0, nValueLength);

                    return (true);

                case WMFSDKWrapper.WMT_ATTR_DATATYPE.WMT_TYPE_QWORD:

                    nValueLength = 8;
                    ulong[] pqwAttribValue = new ulong[1] { Convert.ToUInt64(pwszValue) };

                    pbValue = new Byte[nValueLength];
                    Buffer.BlockCopy(pqwAttribValue, 0, pbValue, 0, nValueLength);

                    return (true);

                case WMFSDKWrapper.WMT_ATTR_DATATYPE.WMT_TYPE_STRING:

                    nValueLength = (ushort)((pwszValue.Length + 1) * 2);
                    pbValue = new Byte[nValueLength];

                    Buffer.BlockCopy(pwszValue.ToCharArray(), 0, pbValue, 0, pwszValue.Length * 2);
                    pbValue[nValueLength - 2] = 0;
                    pbValue[nValueLength - 1] = 0;

                    return (true);

                case WMFSDKWrapper.WMT_ATTR_DATATYPE.WMT_TYPE_BOOL:

                    nValueLength = 4;
                    pdwAttribValue = new uint[1] { Convert.ToUInt32(pwszValue) };
                    if (pdwAttribValue[0] != 0)
                    {
                        pdwAttribValue[0] = 1;
                    }

                    pbValue = new Byte[nValueLength];
                    Buffer.BlockCopy(pdwAttribValue, 0, pbValue, 0, nValueLength);

                    return (true);

                case WMFSDKWrapper.WMT_ATTR_DATATYPE.WMT_TYPE_GUID:

                    int discarded;
                    pbValue = HexEncoding.GetBytes(pwszValue, out discarded);
                    nValueLength = HexEncoding.GetByteCount(pwszValue);

                    return (true);

                default:

                    pbValue = null;
                    nValueLength = 0;

                    return (false);
            }
        }

        /// <summary>
        /// Set the specified attribute.
        /// </summary>
        /// <param name="pwszFileName">Name of the PWSZ file.</param>
        /// <param name="wStreamNum">The w stream num.</param>
        /// <param name="pwszAttribName">Name of the PWSZ attrib.</param>
        /// <param name="wAttribType">Type of the w attrib.</param>
        /// <param name="pwszAttribValue">The PWSZ attrib value.</param>
        /// <returns></returns>
        public bool SetAttrib(string pwszFileName, ushort wStreamNum, string pwszAttribName,
                        ushort wAttribType, string pwszAttribValue)
        {
            try
            {
                IWMMetadataEditor MetadataEditor;
                IWMHeaderInfo3 HeaderInfo3;
                byte[] pbAttribValue;
                int nAttribValueLen;
                WMT_ATTR_DATATYPE AttribDataType = (WMT_ATTR_DATATYPE)wAttribType;

                if (!TranslateAttrib(AttribDataType, pwszAttribValue, out pbAttribValue, out nAttribValueLen))
                {
                    return false;
                }

                WMFSDKFunctions.WMCreateEditor(out MetadataEditor);

                MetadataEditor.Open(pwszFileName);

                HeaderInfo3 = (IWMHeaderInfo3)MetadataEditor;

                HeaderInfo3.SetAttribute(wStreamNum,
                                          pwszAttribName,
                                          AttribDataType,
                                          pbAttribValue,
                                          (ushort)nAttribValueLen);

                MetadataEditor.Flush();

                MetadataEditor.Close();
            }
            catch (Exception e)
            {
                AddLogEntry(e.Message, LogType.Fail);
                return (false);
            }

            return (true);
        }

        /// <summary>
        /// Add an attribute with the specifed language index.
        /// </summary>
        /// <param name="pwszFileName">Name of the PWSZ file.</param>
        /// <param name="wStreamNum">The w stream num.</param>
        /// <param name="pwszAttribName">Name of the PWSZ attrib.</param>
        /// <param name="wAttribType">Type of the w attrib.</param>
        /// <param name="pwszAttribValue">The PWSZ attrib value.</param>
        /// <param name="wLangIndex">Index of the w lang.</param>
        /// <returns></returns>
        public bool AddAttrib(string pwszFileName, ushort wStreamNum, string pwszAttribName,
                        ushort wAttribType, string pwszAttribValue, ushort wLangIndex)
        {
            IWMMetadataEditor MetadataEditor = null;
            IWMHeaderInfo3 HeaderInfo3;
            byte[] pbAttribValue;
            int nAttribValueLen;
            WMT_ATTR_DATATYPE AttribDataType = (WMT_ATTR_DATATYPE)wAttribType;
            ushort wAttribIndex = 0;

            try
            {
                if (!TranslateAttrib(AttribDataType, pwszAttribValue, out pbAttribValue, out nAttribValueLen))
                {
                    return false;
                }

                WMFSDKFunctions.WMCreateEditor(out MetadataEditor);

                MetadataEditor.Open(pwszFileName);

                HeaderInfo3 = (IWMHeaderInfo3)MetadataEditor;

                HeaderInfo3.AddAttribute(wStreamNum,
                                          pwszAttribName,
                                          out wAttribIndex,
                                          AttribDataType,
                                          wLangIndex,
                                          pbAttribValue,
                                          (uint)nAttribValueLen);
            }
            catch (Exception)
            {
                // AddLogEntry(e.Message, LogType.Fail);
                return (false);
            }
            finally
            {
                MetadataEditor.Flush();
                MetadataEditor.Close();
            }

            return (true);
        }

        /// <summary>
        /// Modifies the value of the specified attribute.
        /// </summary>
        /// <param name="pwszFileName">Name of the PWSZ file.</param>
        /// <param name="wStreamNum">The w stream num.</param>
        /// <param name="wAttribIndex">Index of the w attrib.</param>
        /// <param name="wAttribType">Type of the w attrib.</param>
        /// <param name="pwszAttribValue">The PWSZ attrib value.</param>
        /// <param name="wLangIndex">Index of the w lang.</param>
        /// <returns></returns>
        public bool ModifyAttrib(string pwszFileName, ushort wStreamNum, ushort wAttribIndex,
                           ushort wAttribType, string pwszAttribValue, ushort wLangIndex)
        {
            IWMMetadataEditor MetadataEditor = null;
            IWMHeaderInfo3 HeaderInfo3;
            byte[] pbAttribValue;
            int nAttribValueLen;
            WMT_ATTR_DATATYPE AttribDataType = (WMT_ATTR_DATATYPE)wAttribType;

            try
            {
                if (!TranslateAttrib(AttribDataType, pwszAttribValue, out pbAttribValue, out nAttribValueLen))
                {
                    return false;
                }

                WMFSDKFunctions.WMCreateEditor(out MetadataEditor);

                MetadataEditor.Open(pwszFileName);

                HeaderInfo3 = (IWMHeaderInfo3)MetadataEditor;

                HeaderInfo3.ModifyAttribute(wStreamNum,
                                             wAttribIndex,
                                             AttribDataType,
                                             wLangIndex,
                                             pbAttribValue,
                                             (uint)nAttribValueLen);
            }
            catch (Exception e)
            {
                AddLogEntry(e.Message, LogType.Fail);
                return (false);
            }
            finally
            {
                MetadataEditor.Flush();
                MetadataEditor.Close();
            }

            return (true);
        }

        /// <summary>
        /// Attribs the exists.
        /// </summary>
        /// <param name="pwszFileName">Name of the PWSZ file.</param>
        /// <param name="wStreamNum">The w stream num.</param>
        /// <param name="searchAttrib">The search attrib.</param>
        /// <returns></returns>
        public bool AttribExists(string pwszFileName, ushort wStreamNum, string searchAttrib)
        {
            bool isFound = false;

            try
            {
                IWMMetadataEditor MetadataEditor;
                IWMHeaderInfo3 HeaderInfo3;
                ushort wAttributeCount = 0;

                WMFSDKFunctions.WMCreateEditor(out MetadataEditor);

                MetadataEditor.Open(pwszFileName);

                HeaderInfo3 = (IWMHeaderInfo3)MetadataEditor;

                HeaderInfo3.GetAttributeCountEx(wStreamNum, out wAttributeCount);

                for (ushort wAttribIndex = 0; (wAttribIndex < wAttributeCount) && !isFound; wAttribIndex++)
                {
                    WMT_ATTR_DATATYPE wAttribType;
                    ushort wLangIndex = 0;
                    string pwszAttribName = null;
                    byte[] pbAttribValue = null;
                    ushort wAttribNameLen = 0;
                    uint dwAttribValueLen = 0;

                    HeaderInfo3.GetAttributeByIndexEx(wStreamNum,
                                                       wAttribIndex,
                                                       pwszAttribName,
                                                       ref wAttribNameLen,
                                                       out wAttribType,
                                                       out wLangIndex,
                                                       pbAttribValue,
                                                       ref dwAttribValueLen);

                    pwszAttribName = new String((char)0, wAttribNameLen);
                    pbAttribValue = new byte[dwAttribValueLen];

                    HeaderInfo3.GetAttributeByIndexEx(wStreamNum,
                                                       wAttribIndex,
                                                       pwszAttribName,
                                                       ref wAttribNameLen,
                                                       out wAttribType,
                                                       out wLangIndex,
                                                       pbAttribValue,
                                                       ref dwAttribValueLen);

                    if (pwszAttribName.Substring(0, pwszAttribName.Length - 1) == searchAttrib)
                    {
                        isFound = true;
                    }
                }

                // Close file
                MetadataEditor.Close();
            }
            catch (Exception e)
            {
                AddLogEntry(e.Message, LogType.Fail);
                return (false);
            }

            return (isFound);
        }

        /// <summary>
        /// Handles the 1 event of the cmdModify_Click control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="System.EventArgs"/> instance containing the event data.</param>
        private void cmdModify_Click_1(object sender, EventArgs e)
        {
            if (!System.IO.File.Exists(lblMediaFile.Text))
            {
                return;
            }

            if (cbAttributes.SelectedItem == null)
            {
                return;
            }

            if (txtNewValue.Text.Length == 0)
            {
                return;
            }

            Attribute _attribute = (Attribute)cbAttributes.SelectedItem;
            if (_attribute == null)
            {
                return;
            }

            if (ModifyAttrib(lblMediaFile.Text, Stream, _attribute.Index, Convert.ToUInt16(_attribute.Type), txtNewValue.Text, Language))
            {
                AddLogEntry(_attribute.Name + " successfully modified");
                InspectFile();
            }
        }

        /// <summary>
        /// Handles the SelectedIndexChanged event of the cbAttributes control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="System.EventArgs"/> instance containing the event data.</param>
        private void cbAttributes_SelectedIndexChanged(object sender, EventArgs e)
        {
            Attribute _attribute = (Attribute)cbAttributes.SelectedItem;
            if (_attribute == null)
            {
                return;
            }

            txtNewValue.Text = _attribute.Value;
        }

        /// <summary>
        /// Handles the Click event of the cmdSave control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="System.EventArgs"/> instance containing the event data.</param>
        private void cmdSave_Click(object sender, EventArgs e)
        {
            string newType = "";

            if (!System.IO.File.Exists(lblMediaFile.Text))
            {
                return;
            }

            switch (cbMediaType.SelectedIndex)
            {
                case 1:
                    newType = "Generic Video";
                    EditAttribute("WM/MediaClassPrimaryID", TypeVideo);
                    //ModifyAttrib(lblMediaFile.Text, Stream, indexPrimaryVideo, Convert.ToUInt16(WMT_ATTR_DATATYPE.WMT_TYPE_GUID), TypeVideo, Language);
                    break;

                case 2:
                    newType = "Movie";
                    EditAttribute("WM/MediaClassPrimaryID", TypeVideo);
                    EditAttribute("WM/MediaClassSecondaryID", TypeMovie);
                    //ModifyAttrib(lblMediaFile.Text, Stream, indexPrimaryVideo, Convert.ToUInt16(WMT_ATTR_DATATYPE.WMT_TYPE_GUID), TypeVideo, Language);
                    //ModifyAttrib(lblMediaFile.Text, Stream, indexSecondaryVideo, Convert.ToUInt16(WMT_ATTR_DATATYPE.WMT_TYPE_GUID), TypeMovie, Language);
                    break;

                case 3:
                    newType = "Music Video";
                    EditAttribute("WM/MediaClassPrimaryID", TypeVideo);
                    EditAttribute("WM/MediaClassSecondaryID", TypeMusic);
                    //ModifyAttrib(lblMediaFile.Text, Stream, indexPrimaryVideo, Convert.ToUInt16(WMT_ATTR_DATATYPE.WMT_TYPE_GUID), TypeVideo, Language);
                    //ModifyAttrib(lblMediaFile.Text, Stream, indexSecondaryVideo, Convert.ToUInt16(WMT_ATTR_DATATYPE.WMT_TYPE_GUID), TypeMusic, Language);
                    break;

                case 4:
                    newType = "TV Show";
                    EditAttribute("WM/MediaClassPrimaryID", TypeVideo);
                    EditAttribute("WM/MediaClassSecondaryID", TypeTV);
                    //ModifyAttrib(lblMediaFile.Text, Stream, indexPrimaryVideo, Convert.ToUInt16(WMT_ATTR_DATATYPE.WMT_TYPE_GUID), TypeVideo, Language);
                    //ModifyAttrib(lblMediaFile.Text, Stream, indexSecondaryVideo, Convert.ToUInt16(WMT_ATTR_DATATYPE.WMT_TYPE_GUID), TypeTV, Language);
                    break;
            }

            // Refresh screen
            if (newType.Length > 0)
            {
                AddLogEntry("File modified to be a " + newType, LogType.Success);
                InspectFile();
            }
        }

        /// <summary>
        /// Adds the missing attributes.
        /// </summary>
        private void AddMissingAttributes()
        {
            if (!AttribExists(lblMediaFile.Text, Stream, "WM/MediaClassPrimaryID"))
            {
                AddAttrib(lblMediaFile.Text, newStream, "WM/MediaClassPrimaryID", Convert.ToUInt16(WMT_ATTR_DATATYPE.WMT_TYPE_GUID), TypeVideo, Language);
            }
            if (!AttribExists(lblMediaFile.Text, Stream, "WM/MediaClassSecondaryID"))
                AddAttrib(lblMediaFile.Text, newStream, "WM/MediaClassSecondaryID", Convert.ToUInt16(WMT_ATTR_DATATYPE.WMT_TYPE_GUID), TypeVideo, Language);
            if (!AttribExists(lblMediaFile.Text, Stream, "Title"))
                AddAttrib(lblMediaFile.Text, newStream, "Title", Convert.ToUInt16(WMT_ATTR_DATATYPE.WMT_TYPE_STRING), "Unknown", Language);
            if (!AttribExists(lblMediaFile.Text, Stream, "WM/SubTitle")) 
                AddAttrib(lblMediaFile.Text, newStream, "WM/SubTitle", Convert.ToUInt16(WMT_ATTR_DATATYPE.WMT_TYPE_STRING), "Unknown", Language);
            if (!AttribExists(lblMediaFile.Text, Stream, "WM/SubTitleDescription")) 
                AddAttrib(lblMediaFile.Text, newStream, "WM/SubTitleDescription", Convert.ToUInt16(WMT_ATTR_DATATYPE.WMT_TYPE_STRING), "Unknown", Language);
            if (!AttribExists(lblMediaFile.Text, Stream, "Author")) 
                AddAttrib(lblMediaFile.Text, newStream, "Author", Convert.ToUInt16(WMT_ATTR_DATATYPE.WMT_TYPE_STRING), "Unknown", Language);
            if (!AttribExists(lblMediaFile.Text, Stream, "WM/Year")) 
                AddAttrib(lblMediaFile.Text, newStream, "WM/Year", Convert.ToUInt16(WMT_ATTR_DATATYPE.WMT_TYPE_STRING), "Unknown", Language);
            if (!AttribExists(lblMediaFile.Text, Stream, "WM/OriginalBroadcastDateTime")) 
                AddAttrib(lblMediaFile.Text, newStream, "WM/OriginalBroadcastDateTime", Convert.ToUInt16(WMT_ATTR_DATATYPE.WMT_TYPE_STRING), "Unknown", Language);
            if (!AttribExists(lblMediaFile.Text, Stream, "WM/ParentalRating")) 
                AddAttrib(lblMediaFile.Text, newStream, "WM/ParentalRating", Convert.ToUInt16(WMT_ATTR_DATATYPE.WMT_TYPE_STRING), "Unknown", Language);
            if (!AttribExists(lblMediaFile.Text, Stream, "WM/TVNetworkAffiliation")) 
                AddAttrib(lblMediaFile.Text, newStream, "WM/TVNetworkAffiliation", Convert.ToUInt16(WMT_ATTR_DATATYPE.WMT_TYPE_STRING), "Unknown", Language);
            if (!AttribExists(lblMediaFile.Text, Stream, "WM/Genre")) 
                AddAttrib(lblMediaFile.Text, newStream, "WM/Genre", Convert.ToUInt16(WMT_ATTR_DATATYPE.WMT_TYPE_STRING), "Unknown", Language);
            if (!AttribExists(lblMediaFile.Text, Stream, "Description")) 
                AddAttrib(lblMediaFile.Text, newStream, "Description", Convert.ToUInt16(WMT_ATTR_DATATYPE.WMT_TYPE_STRING), "Unknown", Language);
            if (!AttribExists(lblMediaFile.Text, Stream, "WM/TrackNumber")) 
                AddAttrib(lblMediaFile.Text, newStream, "WM/TrackNumber", Convert.ToUInt16(WMT_ATTR_DATATYPE.WMT_TYPE_DWORD), "01", Language);
        }

        /// <summary>
        /// Edits the picture.
        /// </summary>
        private void EditPicture(PictureBox myPicture)
        {
            IWMMetadataEditor MetadataEditor;
            IWMHeaderInfo3 HeaderInfo3;

            try
            {
                ImageConverter imageConverter = new ImageConverter();
                WMPicture picture = new WMPicture();

                picture.pwszMIMEType = Marshal.StringToCoTaskMemUni("image/jpeg\0");
                picture.pwszDescription = Marshal.StringToCoTaskMemUni("AlbumArt\0");
                picture.bPictureType = 3;

                byte[] data = (byte[])imageConverter.ConvertTo(myPicture.Image, typeof(byte[]));
                picture.dwDataLen = data.Length;
                picture.pbData = Marshal.AllocCoTaskMem(picture.dwDataLen);
                Marshal.Copy(data, 0, picture.pbData, picture.dwDataLen);
                IntPtr pictureParam = Marshal.AllocCoTaskMem(Marshal.SizeOf(picture));
                Marshal.StructureToPtr(picture, pictureParam, false);

                WMFSDKFunctions.WMCreateEditor(out MetadataEditor);
                MetadataEditor.Open(lblMediaFile.Text);
                HeaderInfo3 = (IWMHeaderInfo3)MetadataEditor;
                HeaderInfo3.SetPicAttribute(0, "WM/Picture", WMT_ATTR_DATATYPE.WMT_TYPE_BINARY, pictureParam, (ushort)Marshal.SizeOf(picture));
                MetadataEditor.Flush();
                MetadataEditor.Close();
            }
            catch (Exception e)
            {
                AddLogEntry(e.Message, LogType.Fail);
            }
        }

        /// <summary>
        /// Handles the SelectedIndexChanged event of the cbMediaType control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="System.EventArgs"/> instance containing the event data.</param>
        private void cbMediaType_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (cbMediaType.SelectedIndex)
            {
                case 0:
                    gbMovie.Visible = false;
                    gbMusic.Visible = false;
                    gbTV.Visible = false;
                    gbVideo.Visible = false;
                    break;

                case 1:
                    gbMovie.Visible = false;
                    gbMusic.Visible = false;
                    gbTV.Visible = false;
                    LoadVideoAttributes();
                    gbVideo.Visible = true;
                    break;

                case 2:
                    gbMusic.Visible = false;
                    gbTV.Visible = false;
                    gbVideo.Visible = false;
                    LoadMovieAttributes();
                    gbMovie.Visible = true;
                    break;

                case 3:
                    gbMovie.Visible = false;
                    gbTV.Visible = false;
                    gbVideo.Visible = false;
                    LoadMusicAttributes();
                    gbMusic.Visible = true;
                    break;

                case 4:
                    gbMovie.Visible = false;
                    gbMusic.Visible = false;
                    gbVideo.Visible = false;
                    LoadTVAttributes();
                    gbTV.Visible = true;
                    break;
            }
        }

        /// <summary>
        /// Loads the TV attributes.
        /// </summary>
        private void LoadTVAttributes()
        {
            foreach (Attribute _attribute in _attributes)
            {
                string myValue = _attribute.Value.Replace("\"", "");

                switch (_attribute.Name)
                {
                    case "Title":
                        txtTVTitle.Text = myValue;
                        txtSearchCriteria.Text = txtTVTitle.Text;
                        break;

                    case "WM/SubTitle":
                        txtTVSubTitle.Text = myValue;
                        break;

                    case "WM/SubTitleDescription":
                        txtTVDescription.Text = myValue;
                        break;

                    case "Author":
                        txtTVAuthor.Text = myValue;
                        break;

                    case "WM/Year":
                        txtTVYear.Text = myValue;
                        break;

                    case "WM/OriginalBroadcastDateTime":
                        txtTVDate.Text = myValue;
                        break;

                    case "WM/ParentalRating":
                        txtTVRating.Text = myValue;
                        break;

                    case "WM/TVNetworkAffiliation":
                        txtTVNetwork.Text = myValue;
                        break;

                    case "WM/Genre":
                        txtTVGenre.Text = myValue;
                        break;

                    case "WM/TrackNumber":
                        txtTVTrack.Text = myValue;
                        break;
                }
            }
        }

        /// <summary>
        /// Loads the movie attributes.
        /// </summary>
        private void LoadMovieAttributes()
        {
            foreach (Attribute _attribute in _attributes)
            {
                string myValue = _attribute.Value.Replace("\"", "");

                switch (_attribute.Name)
                {
                    case "Title":
                        txtMovieTitle.Text = myValue;
                        txtSearchCriteria.Text = txtMovieTitle.Text;
                        break;

                    case "WM/SubTitleDescription":
                        txtMovieDescription.Text = myValue;
                        break;

                    case "Author":
                        txtMovieAuthor.Text = myValue;
                        break;

                    case "WM/Year":
                        txtMovieYear.Text = myValue;
                        break;

                    case "WM/OriginalBroadcastDateTime":
                        txtMovieDate.Text = myValue;
                        break;

                    case "WM/ParentalRating":
                        txtMovieRating.Text = myValue;
                        break;

                    case "WM/Genre":
                        txtMovieGenre.Text = myValue;
                        break;
                }
            }
        }

        /// <summary>
        /// Loads the video attributes.
        /// </summary>
        private void LoadVideoAttributes()
        {
            foreach (Attribute _attribute in _attributes)
            {
                string myValue = _attribute.Value.Replace("\"", "");

                switch (_attribute.Name)
                {
                    case "Title":
                        txtVideoTitle.Text = myValue;
                        txtSearchCriteria.Text = txtVideoTitle.Text;
                        break;

                    case "WM/SubTitleDescription":
                        txtVideoDescription.Text = myValue;
                        break;

                    case "Author":
                        txtVideoAuthor.Text = myValue;
                        break;

                    case "WM/Year":
                        txtVideoYear.Text = myValue;
                        break;

                    case "WM/Genre":
                        txtVideoGenre.Text = myValue;
                        break;
                }
            }
        }

        /// <summary>
        /// Loads the music attributes.
        /// </summary>
        private void LoadMusicAttributes()
        {
            foreach (Attribute _attribute in _attributes)
            {
                string myValue = _attribute.Value.Replace("\"", "");

                switch (_attribute.Name)
                {
                    case "Title":
                        txtMusicTitle.Text = myValue;
                        txtSearchCriteria.Text = txtMusicTitle.Text;
                        break;

                    case "WM/SubTitleDescription":
                        txtMusicDescription.Text = myValue;
                        break;

                    case "Author":
                        txtMusicAuthor.Text = myValue;
                        break;

                    case "WM/Year":
                        txtMusicYear.Text = myValue;
                        break;

                    case "WM/OriginalBroadcastDateTime":
                        txtMusicDate.Text = myValue;
                        break;

                    case "WM/ParentalRating":
                        txtMusicRating.Text = myValue;
                        break;

                    case "WM/Genre":
                        txtMusicGenre.Text = myValue;
                        break;
                }
            }
        }

        /// <summary>
        /// Handles the Click event of the cmdTVSave control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="System.EventArgs"/> instance containing the event data.</param>
        private void cmdTVSave_Click(object sender, EventArgs e)
        {
            EditAttribute("Title", txtTVTitle.Text);
            EditAttribute("WM/SubTitle", txtTVSubTitle.Text);
            EditAttribute("WM/SubTitleDescription", txtTVDescription.Text);
            EditAttribute("Author", txtTVAuthor.Text);
            EditAttribute("WM/Year", txtTVYear.Text);
            EditAttribute("WM/OriginalBroadcastDateTime", txtTVDate.Text);
            EditAttribute("WM/ParentalRating", txtTVRating.Text);
            EditAttribute("WM/TVNetworkAffiliation", txtTVNetwork.Text);
            EditAttribute("WM/Genre", txtTVGenre.Text);
            EditAttribute("WM/TrackNumber", txtTVTrack.Text);
            //EditPicture(pictureBox1);

            AddLogEntry(lblMediaFile.Text + " successfully modified", LogType.Success);
            InspectFile();
        }

        /// <summary>
        /// Modifies the attribute.
        /// </summary>
        /// <param name="theAttribute">The attribute.</param>
        /// <param name="newValue">The new value.</param>
        private void EditAttribute(string theAttribute, string newValue)
        {
            foreach (Attribute _attribute in _attributes)
            {
                if (_attribute.Name == theAttribute)
                {
                    ModifyAttrib(lblMediaFile.Text, Stream, _attribute.Index, Convert.ToUInt16(_attribute.Type), newValue, Language);
                    break;
                }
            }
        }

        /// <summary>
        /// Handles the Click event of the cmdMovieSave control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="System.EventArgs"/> instance containing the event data.</param>
        private void cmdMovieSave_Click(object sender, EventArgs e)
        {
            EditAttribute("Title", txtMovieTitle.Text);
            EditAttribute("WM/SubTitleDescription", txtMovieDescription.Text);
            EditAttribute("Author", txtMovieAuthor.Text);
            EditAttribute("WM/Year", txtMovieYear.Text);
            EditAttribute("WM/OriginalBroadcastDateTime", txtMovieDate.Text);
            EditAttribute("WM/ParentalRating", txtMovieRating.Text);
            EditAttribute("WM/Genre", txtMovieGenre.Text);
            //EditPicture(pictureBox1);

            AddLogEntry(lblMediaFile.Text + " successfully modified", LogType.Success);
            InspectFile();
        }

        /// <summary>
        /// Handles the Click event of the cmdMusicSave control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="System.EventArgs"/> instance containing the event data.</param>
        private void cmdMusicSave_Click(object sender, EventArgs e)
        {
            EditAttribute("Title", txtMusicTitle.Text);
            EditAttribute("WM/SubTitleDescription", txtMusicDescription.Text);
            EditAttribute("Author", txtMusicAuthor.Text);
            EditAttribute("WM/Year", txtMusicYear.Text);
            EditAttribute("WM/OriginalBroadcastDateTime", txtMusicDate.Text);
            EditAttribute("WM/ParentalRating", txtMusicRating.Text);
            EditAttribute("WM/Genre", txtMusicGenre.Text);
            //EditPicture(pictureBox1);

            AddLogEntry(lblMediaFile.Text + " successfully modified", LogType.Success);
            InspectFile();
        }

        /// <summary>
        /// Handles the Click event of the cmdVideoSave control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="System.EventArgs"/> instance containing the event data.</param>
        private void cmdVideoSave_Click(object sender, EventArgs e)
        {
            EditAttribute("Title", txtVideoTitle.Text);
            EditAttribute("WM/SubTitleDescription", txtVideoDescription.Text);
            EditAttribute("Author", txtVideoAuthor.Text);
            EditAttribute("WM/Year", txtVideoYear.Text);
            EditAttribute("WM/Genre", txtVideoGenre.Text);
            //EditPicture(pictureBox1);

            AddLogEntry(lblMediaFile.Text + " successfully modified", LogType.Success);
            InspectFile();
        }

        /// <summary>
        /// Handles the Click event of the cmdAmazonSearch control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="System.EventArgs"/> instance containing the event data.</param>
        private void cmdAmazonSearch_Click(object sender, EventArgs e)
        {
            pageNumber = 1;
            AmazonSearch();
        }

        /// <summary>
        /// Amazons the search.
        /// </summary>
        private void AmazonSearch()
        {
            // Set up visuals pre-search
            AddLogEntry("Searching Amazon for " + txtSearchCriteria.Text + "...", LogType.Info);
            SetButtons(false);
            progressBar1.Visible = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            // Do the search
            backgroundWorker1.RunWorkerAsync();
        }

        /// <summary>
        /// Sets the buttons.
        /// </summary>
        /// <param name="isON">if set to <c>true</c> [is ON].</param>
        private void SetButtons(bool isON)
        {
            cmdAmazonSearch.Enabled = isON;
            cmdNext.Enabled = isON;
            cmdPrev.Enabled = isON;
        }

        /// <summary>
        /// Performs the search.
        /// </summary>
        private void PerformSearch()
        {
            if (txtSearchCriteria.Text.Length == 0)
            {
                return;
            }

            try
            {
                // create a WCF Amazon ECS client
                BasicHttpBinding basicBinding = new BasicHttpBinding(BasicHttpSecurityMode.Transport);
                basicBinding.MaxReceivedMessageSize = 2147483647;
                basicBinding.ReaderQuotas.MaxStringContentLength = int.MaxValue;
                AWSECommerceServicePortTypeClient client = new AWSECommerceServicePortTypeClient(
                    basicBinding,
                    new EndpointAddress("https://webservices.amazon.com/onca/soap?Service=AWSECommerceService"));

                // add authentication to the ECS client  
                client.ChannelFactory.Endpoint.Behaviors.Add(new AmazonSigningEndpointBehavior(AmazonID, AmazonKey));

                ItemSearchRequest request = new ItemSearchRequest();
                request.Availability = ItemSearchRequestAvailability.Available;
                request.Condition = Condition.All;
                request.ItemPage = pageNumber.ToString();
                request.MerchantId = "Amazon";
                request.ResponseGroup = new string[] { "Medium", "Reviews", "ItemAttributes", "BrowseNodes" };
                request.SearchIndex = "DVD";
                request.Title = txtSearchCriteria.Text;

                ItemSearch itemSearch = new ItemSearch();
                itemSearch.Request = new ItemSearchRequest[] { request };
                itemSearch.AWSAccessKeyId = AmazonID;
                ItemSearchResponse response = client.ItemSearch(itemSearch);

                if (response.Items[0].Item != null)
                {
                    lbResults.Items.Clear();
                    foreach (Items resultItem in response.Items)
                    {
                        foreach (Item displayItem in resultItem.Item)
                        {
                            ItemAttributes item = displayItem.ItemAttributes;

                            // Work out which review (if any) we can use
                            string myReview;
                            if (displayItem.EditorialReviews == null)
                            {
                                if (displayItem.CustomerReviews == null)
                                {
                                    myReview = "";
                                }
                                else
                                {
                                    myReview = orNull(displayItem.CustomerReviews.Review[0].Content);
                                }
                            }
                            else
                            {
                                myReview = orNull(displayItem.EditorialReviews[0].Content);
                            }

                            AmazonEntry _entry = new AmazonEntry(
                                    orNull(displayItem.ASIN),
                                    orNull(item.Title),
                                    item.Director == null ? "" : orNull(item.Director[0]),
                                    orNull(item.TheatricalReleaseDate),
                                    myReview,
                                    orNull(displayItem.DetailPageURL),
                                    orNull(GetGenre(displayItem)),
                                    orNull(item.AudienceRating),
                                    displayItem.MediumImage == null ? "" : orNull(displayItem.MediumImage.URL));
                            lbResults.Items.Add(_entry);
                        }
                    }
                    lblPage.Text = "Page " + pageNumber.ToString();
                    safePageNumber = pageNumber;
                }
                else
                {
                    pageNumber = safePageNumber;
                }

                AddLogEntry("Search Complete", LogType.Info);
            }
            catch (Exception ex)
            {
                pageNumber = safePageNumber;
                AddLogEntry(ex.Message, LogType.Fail);
            }
        }

        /// <summary>
        /// Shows the cover art.
        /// </summary>
        /// <param name="item">The item.</param>
        private void ShowCoverArt(string coverUrl)
        {
            if (coverUrl != "")
            {
                WebClient client = new WebClient();
                client.Headers["User-Agent"] = "Mozilla/4.0";
                byte[] bytes = client.DownloadData(coverUrl);
                MemoryStream stream = new MemoryStream(bytes);
                pbCover.Image = System.Drawing.Image.FromStream(stream);
            }
        }

        /// <summary>
        /// Gets the genre.
        /// </summary>
        /// <param name="item">The item.</param>
        /// <returns></returns>
        static string GetGenre(Item item)
        {
            if (item.BrowseNodes == null) return null;
            string genre = string.Empty;
            foreach (BrowseNode node in item.BrowseNodes.BrowseNode)
            {
                if (node.Name.ToLower() == "general")
                    if (node.Ancestors != null)
                        return node.Ancestors[0].Name;
            }
            return "<unspecified>";
        }

        /// <summary>
        /// Ors the null.
        /// </summary>
        /// <param name="inputString">The input string.</param>
        /// <returns></returns>
        private string orNull(string inputString)
        {
            return (string.IsNullOrEmpty(inputString) ? "" : inputString);
        }

        /// <summary>
        /// Handles the TextChanged event of the txtSearchCriteria control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="System.EventArgs"/> instance containing the event data.</param>
        private void txtSearchCriteria_TextChanged(object sender, EventArgs e)
        {
            pageNumber = 0;
        }

        /// <summary>
        /// Handles the KeyDown event of the txtSearchCriteria control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="System.Windows.Forms.KeyEventArgs"/> instance containing the event data.</param>
        private void txtSearchCriteria_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                cmdAmazonSearch_Click(sender, e);
            }
        }

        /// <summary>
        /// Handles the SelectedIndexChanged event of the lbResults control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="System.EventArgs"/> instance containing the event data.</param>
        private void lbResults_SelectedIndexChanged(object sender, EventArgs e)
        {
            AmazonEntry _entry = (AmazonEntry)lbResults.SelectedItem;
            if (_entry != null)
            {
                txtAzTitle.Text = _entry.Title;
                txtAzYear.Text = _entry.Year;
                txtAzDirector.Text = _entry.Director;
                txtAzDescription.Text = _entry.Description;
                ShowCoverArt(_entry.Cover);
            }
        }

        /// <summary>
        /// Handles the DoubleClick event of the lbResults control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="System.EventArgs"/> instance containing the event data.</param>
        private void lbResults_DoubleClick(object sender, EventArgs e)
        {
            if (lbResults.SelectedItem == null) return;

            AmazonEntry _entry = (AmazonEntry)lbResults.SelectedItem;
            if (_entry != null)
            {
                if (_entry.URL.Length > 0)
                {
                    System.Diagnostics.Process.Start(_entry.URL);
                }
            }
        }

        /// <summary>
        /// Handles the Click event of the cmdCopyAz control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="System.EventArgs"/> instance containing the event data.</param>
        private void cmdCopyAz_Click(object sender, EventArgs e)
        {
            if (lbResults.SelectedItem == null) return;
            if (lblMediaFile.Text.Length == 0) return;

            AmazonEntry _entry = (AmazonEntry)lbResults.SelectedItem;
            if (_entry != null)
            {
                switch (cbMediaType.SelectedIndex)
                {
                    case 0:
                        break;

                    case 1:
                        AddLogEntry("Copying details to current media file");
                        txtVideoAuthor.Text = _entry.Director;
                        txtVideoDescription.Text = _entry.Description;
                        txtVideoGenre.Text = _entry.Genre;
                        txtVideoTitle.Text = _entry.Title;
                        txtVideoYear.Text = _entry.Year;
                        break;

                    case 2:
                        AddLogEntry("Copying details to current media file");
                        txtMovieAuthor.Text = _entry.Director;
                        txtMovieDate.Text = _entry.Date;
                        txtMovieDescription.Text = _entry.Description;
                        txtMovieGenre.Text = _entry.Genre;
                        txtMovieRating.Text = _entry.Rating;
                        txtMovieTitle.Text = _entry.Title;
                        txtMovieYear.Text = _entry.Year;
                        break;

                    case 3:
                        AddLogEntry("Copying details to current media file");
                        txtMusicAuthor.Text = _entry.Director;
                        txtMusicDate.Text = _entry.Date;
                        txtMusicDescription.Text = _entry.Description;
                        txtMusicGenre.Text = _entry.Genre;
                        txtMusicRating.Text = _entry.Rating;
                        txtMusicTitle.Text = _entry.Title;
                        txtMusicYear.Text = _entry.Year;
                        break;

                    case 4:
                        AddLogEntry("Copying details to current media file");
                        txtTVAuthor.Text = _entry.Director;
                        txtTVDate.Text = _entry.Date;
                        txtTVDescription.Text = _entry.Description;
                        txtTVGenre.Text = _entry.Genre;
                        txtTVRating.Text = _entry.Rating;
                        txtTVTitle.Text = _entry.Title;
                        txtTVYear.Text = _entry.Year;
                        break;
                }

                // Screenshot
                //pictureBox1.Image = pbCover.Image;

                // Switch back to the first tab to see the results
                tabControl1.SelectedIndex = 0;
            }
        }

        /// <summary>
        /// Handles the Click event of the cmdNext control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="System.EventArgs"/> instance containing the event data.</param>
        private void cmdNext_Click(object sender, EventArgs e)
        {
            pageNumber++;
            AmazonSearch();
        }

        /// <summary>
        /// Handles the Click event of the cmdPrev control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="System.EventArgs"/> instance containing the event data.</param>
        private void cmdPrev_Click(object sender, EventArgs e)
        {
            pageNumber--;
            if (pageNumber <= 0)
            {
                pageNumber = 1;
            }
            else
            {
                AmazonSearch();
            }
        }

        /// <summary>
        /// Handles the Tick event of the timer1 control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="System.EventArgs"/> instance containing the event data.</param>
        private void timer1_Tick(object sender, EventArgs e)
        {
            progressBar1.Value += 5;
            if (progressBar1.Value > 120)
                progressBar1.Value = 0;
        }

        /// <summary>
        /// Handles the DoWork event of the backgroundWorker1 control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="System.ComponentModel.DoWorkEventArgs"/> instance containing the event data.</param>
        private void backgroundWorker1_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            PerformSearch();
        }

        /// <summary>
        /// Handles the RunWorkerCompleted event of the backgroundWorker1 control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="System.ComponentModel.RunWorkerCompletedEventArgs"/> instance containing the event data.</param>
        private void backgroundWorker1_RunWorkerCompleted(object sender, System.ComponentModel.RunWorkerCompletedEventArgs e)
        {
            this.Cursor = Cursors.Default;
            progressBar1.Visible = false;
            SetButtons(true);
        }

        /// <summary>
        /// Handles the DragDrop event of the lblMediaFile control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="System.Windows.Forms.DragEventArgs"/> instance containing the event data.</param>
        private void lblMediaFile_DragDrop(object sender, DragEventArgs e)
        {
            string[] s = (string[])e.Data.GetData(DataFormats.FileDrop, false);

            string ext = Path.GetExtension(s[0]).ToLower();
            if (ext == ".wmv")
            {
                lblMediaFile.Text = s[0];
                RegisterNewMediaFile();
            }
            else
            {
                AddLogEntry("Invalid media file format", LogType.Fail);
            }
        }

        /// <summary>
        /// Handles the DragEnter event of the lblMediaFile control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="System.Windows.Forms.DragEventArgs"/> instance containing the event data.</param>
        private void lblMediaFile_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.All;
            else
                e.Effect = DragDropEffects.None;
        }

        /// <summary>
        /// Handles the Click event of the cmdReset control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="System.EventArgs"/> instance containing the event data.</param>
        private void cmdReset_Click(object sender, EventArgs e)
        {
            // Delete the main keys
            foreach (Attribute _attribute in _attributes)
            {
                if (_attribute.Name == "WM/MediaClassPrimaryID")
                {
                    DeleteAttrib(lblMediaFile.Text, Stream, _attribute.Index);
                    break;
                }

                if (_attribute.Name == "WM/MediaClassSecondaryID")
                {
                    DeleteAttrib(lblMediaFile.Text, Stream, _attribute.Index);
                    break;
                }
            }

            // Now re-add the keys
            AddAttrib(lblMediaFile.Text, newStream, "WM/MediaClassPrimaryID", Convert.ToUInt16(WMT_ATTR_DATATYPE.WMT_TYPE_GUID), TypeVideo, Language);
            AddAttrib(lblMediaFile.Text, newStream, "WM/MediaClassSecondaryID", Convert.ToUInt16(WMT_ATTR_DATATYPE.WMT_TYPE_GUID), TypeVideo, Language);
        }
    }
}