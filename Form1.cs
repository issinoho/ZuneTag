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
    using System.ComponentModel;
    using System.Diagnostics;
    using System.Drawing;
    using System.Globalization;
    using System.IO;
    using System.Linq;
    using System.Net;
    using System.Reflection;
    using System.Runtime.ExceptionServices;
    using System.Runtime.InteropServices;
    using System.Threading;
    using System.Threading.Tasks;
    using System.Windows.Forms;

    using DrunkenBakery.ZuneTag.Properties;

    using TMDbLib.Client;
    using TMDbLib.Objects.General;
    using TMDbLib.Objects.Search;

    using WMFSDKWrapper;

    using Xabe.FFmpeg;
    using Xabe.FFmpeg.Downloader;

    using Timer = System.Threading.Timer;

    /// <summary>
    ///     Main application form which drives all functionality.
    /// </summary>
    public partial class Form1 : Form
    {
        private const int ScreenRefresh = 1;

        private const int ScreenLines = 1000;

        private const string ThisApp = "Zune Tag Editor";

        private const string ThisPublisher = "The Drunken Bakery";

        private const ushort Stream = 65535;

        private const ushort NewStream = 0;

        private const ushort Language = 0;

        private const string TypeVideo = "BD-30-98-DB-B3-3A-AB-4F-8A-37-1A-99-5F-7F-F7-4B";

        private const string TypeMovie = "C9-7F-B8-A9-47-BD-F0-4B-AC-4F-65-5B-89-F7-D8-68";

        private const string TypeMusic = "E2-89-E6-E3-8C-BA-30-43-96-DF-A0-EE-EF-FA-68-76";

        private const string TypeTv = "8A-25-7F-BA-F7-62-A9-47-B2-1F-46-51-C4-2A-00-0E";

        private readonly List<Attribute> attributes = new List<Attribute>();

        private readonly List<ListViewItem> lvitems = new List<ListViewItem>();

        private readonly Timer screenLogTimer;

        private readonly TimerCallback screenLogTimerCallback;

        private Form frmAbout;

        private ushort indexPrimaryVideo;

        private ushort indexSecondaryVideo;

        /// <summary>
        ///     Initializes a new instance of the <see cref="Form1" /> class.
        /// </summary>
        public Form1()
        {
            this.InitializeComponent();

            // Upgrade settings from older version
            var a = Assembly.GetExecutingAssembly();
            var appVersion = a.GetName().Version;
            var appVersionString = appVersion.ToString();

            if (Settings.Default.ApplicationVersion != appVersion.ToString())
            {
                Settings.Default.Upgrade();
                Settings.Default.ApplicationVersion = appVersionString;
            }

            // Form title bar
            this.Text = ThisApp + @" freshly baked at " + ThisPublisher;

            // Tooltips
            this.toolTip1.SetToolTip(this.cmdSave, "Change video type");
            this.toolTip2.SetToolTip(this.cmdModify, "Save this attribute back to the file");
            this.toolTip3.SetToolTip(this.cmdMovieSave, "Update Movie with these tags");
            this.toolTip4.SetToolTip(this.cmdMusicSave, "Update Music Video with these tags");
            this.toolTip5.SetToolTip(this.cmdTVSave, "Update TV Show with these tags");
            this.toolTip6.SetToolTip(this.cmdVideoSave, "Update Video with these tags");
            this.toolTip7.SetToolTip(this.cmdReset, "Hard reset the video type attributes");
            this.toolTip8.SetToolTip(this.cmdCopyResult, "Copy the tags back ready for saving");
            this.toolTip9.SetToolTip(this.cmdSearchTmdb, "Search TMDB for videos that match");
            this.toolTip10.SetToolTip(this.cmdBrowse, "Open a WMV media file");

            // Initialise Event Views
            this.InitEventView(this.lvStatus);

            // Initialise Media Types
            this.InitMediaTypes();

            // Ensure FFmpeg is available
            FFmpegDownloader.GetLatestVersion(FFmpegVersion.Official);

            // Logging
            this.AddLogEntry("--------------------------------------------------", LogType.Info);
            this.AddLogEntry("Welcome to the " + ThisApp + " v" + appVersionString, LogType.Info);
            this.AddLogEntry("Ready.");

            // Start Timers
            this.screenLogTimerCallback = this.ScreenLogTimer_Elapsed;
            this.screenLogTimer = new Timer(this.screenLogTimerCallback, null, Convert.ToInt32(ScreenRefresh) * 1000, Timeout.Infinite);
        }

        private delegate void FlushOutputDelegate(ListView lv);

        private delegate void PauseOutputDelegate(ListView lv);

        private delegate void ResumeOutputDelegate(ListView lv);

        /// <summary>
        ///     Severity of logging entry.
        /// </summary>
        private enum LogType
        {
            Success,

            Fail,

            Info,
        }

        /// <inheritdoc />
        public sealed override string Text
        {
            get => base.Text;
            set => base.Text = value;
        }

        /// <summary>
        ///     Generates the image dimensions.
        /// </summary>
        /// <param name="currW">The curr W.</param>
        /// <param name="currH">The curr H.</param>
        /// <param name="destW">The dest W.</param>
        /// <param name="destH">The dest H.</param>
        /// <returns>Image size.</returns>
        public Size GenerateImageDimensions(int currW, int currH, int destW, int destH)
        {
            // Double to hold the final multiplier to use when scaling the image
            double multiplier = 0;

            // String for holding layout

            // Determine if it's Portrait or Landscape
            var layout = currH > currW ? "portrait" : "landscape";

            switch (layout.ToLower())
            {
                case "portrait":
                    // Calculate multiplier on heights
                    if (destH > destW)
                    {
                        multiplier = destW / (double)currW;
                    }
                    else
                    {
                        multiplier = destH / (double)currH;
                    }

                    break;

                case "landscape":
                    // Calculate multiplier on widths
                    if (destH > destW)
                    {
                        multiplier = destW / (double)currW;
                    }
                    else
                    {
                        multiplier = destH / (double)currH;
                    }

                    break;
            }

            // Return the new image dimensions
            return new Size((int)(currW * multiplier), (int)(currH * multiplier));
        }

        /// <summary>
        ///     Creates a metadata editor and opens the file.
        /// </summary>
        /// <param name="pwszInFile">The PWSZ in file.</param>
        /// <param name="ppEditor">The pp editor.</param>
        /// <returns>Success of operation.</returns>
        public bool EditorOpenFile(string pwszInFile, out IWMMetadataEditor ppEditor)
        {
            ppEditor = null;

            try
            {
                WMFSDKFunctions.WMCreateEditor(out ppEditor);

                ppEditor.Open(pwszInFile);
            }
            catch (COMException e)
            {
                this.AddLogEntry(e.Message, LogType.Fail);
                return false;
            }

            return true;
        }

        /// <summary>
        ///     Displays the specified attribute.
        /// </summary>
        /// <param name="wIndex">Index of the w.</param>
        /// <param name="wStream">The w stream.</param>
        /// <param name="pwszName">Name of the PWSZ.</param>
        /// <param name="attribDataType">Type of the attrib data.</param>
        /// <param name="wLangId">The w lang ID.</param>
        /// <param name="pbValue">The pb value.</param>
        /// <param name="dwValueLen">The dw value len.</param>
        public void PrintAttribute(ushort wIndex, ushort wStream, string pwszName, WMT_ATTR_DATATYPE attribDataType, ushort wLangId, byte[] pbValue, uint dwValueLen)
        {
            var pwszValue = string.Empty;

            // Make the data type string
            string[] pTypes = { "DWORD", "STRING", "BINARY", "BOOL", "QWORD", "WORD", "GUID" };

            if (pTypes.Length > Convert.ToInt32(attribDataType))
            {
            }

            // The attribute value.
            switch (attribDataType)
            {
                // String
                case WMT_ATTR_DATATYPE.WMT_TYPE_STRING:

                    if (dwValueLen == 0)
                    {
                        pwszValue = "***** NULL *****";
                    }
                    else
                    {
                        if (Convert.ToInt16(pbValue[0]) == 0xFE && Convert.ToInt16(pbValue[1]) == 0xFF)
                        {
                            pwszValue = "\"UTF-16LE BOM+\"";

                            if (dwValueLen >= 4)
                            {
                                for (var i = 0; i < pbValue.Length - 2; i += 2)
                                {
                                    pwszValue += Convert.ToString(BitConverter.ToChar(pbValue, i));
                                }
                            }

                            pwszValue = pwszValue + "\"";
                        }
                        else if (Convert.ToInt16(pbValue[0]) == 0xFF && Convert.ToInt16(pbValue[1]) == 0xFE)
                        {
                            pwszValue = "\"UTF-16BE BOM+\"";
                            if (dwValueLen >= 4)
                            {
                                for (var i = 0; i < pbValue.Length - 2; i += 2)
                                {
                                    pwszValue += Convert.ToString(BitConverter.ToChar(pbValue, i));
                                }
                            }

                            pwszValue = pwszValue + "\"";
                        }
                        else
                        {
                            pwszValue = "\"";
                            if (dwValueLen >= 2)
                            {
                                for (var i = 0; i < pbValue.Length - 2; i += 2)
                                {
                                    pwszValue += Convert.ToString(BitConverter.ToChar(pbValue, i));
                                }
                            }

                            pwszValue += "\"";
                        }
                    }

                    break;

                // Binary
                case WMT_ATTR_DATATYPE.WMT_TYPE_BINARY:

                    pwszValue = "[" + dwValueLen + " bytes]";
                    break;

                // Boolean
                case WMT_ATTR_DATATYPE.WMT_TYPE_BOOL:

                    pwszValue = BitConverter.ToBoolean(pbValue, 0) ? "True" : "False";

                    break;

                // DWORD
                case WMT_ATTR_DATATYPE.WMT_TYPE_DWORD:

                    var dwValue = BitConverter.ToUInt32(pbValue, 0);
                    pwszValue = dwValue.ToString();
                    break;

                // QWORD
                case WMT_ATTR_DATATYPE.WMT_TYPE_QWORD:

                    var qwValue = BitConverter.ToUInt64(pbValue, 0);
                    pwszValue = qwValue.ToString();
                    break;

                // WORD
                case WMT_ATTR_DATATYPE.WMT_TYPE_WORD:

                    uint wValue = BitConverter.ToUInt16(pbValue, 0);
                    pwszValue = wValue.ToString();
                    break;

                // GUID
                case WMT_ATTR_DATATYPE.WMT_TYPE_GUID:

                    pwszValue = BitConverter.ToString(pbValue, 0, pbValue.Length);
                    break;
            }

            // Add to attribute list
            var attribute = new Attribute(wIndex, pwszName.Substring(0, pwszName.Length - 1), pwszValue, attribDataType);

            // Add to list
            this.attributes.Add(attribute);
        }

        /// <summary>
        ///     Displays all attributes for the specified stream.
        /// </summary>
        /// <param name="pwszFileName">Name of the PWSZ file.</param>
        /// <param name="wStreamNum">The w stream num.</param>
        /// <returns>Success of operation.</returns>
        public bool ShowAttributes(string pwszFileName, ushort wStreamNum)
        {
            try
            {
                WMFSDKFunctions.WMCreateEditor(out var metadataEditor);

                metadataEditor.Open(pwszFileName);

                var headerInfo3 = (IWMHeaderInfo3)metadataEditor;

                headerInfo3.GetAttributeCount(wStreamNum, out var wAttributeCount);

                for (ushort wAttribIndex = 0; wAttribIndex < wAttributeCount; wAttribIndex++)
                {
                    string pwszAttribName = null;
                    byte[] pbAttribValue = null;
                    ushort wAttribNameLen = 0;
                    ushort wAttribValueLen = 0;

                    headerInfo3.GetAttributeByIndex(wAttribIndex, ref wStreamNum, pwszAttribName, ref wAttribNameLen, out var wAttribType, pbAttribValue, ref wAttribValueLen);

                    pbAttribValue = new byte[wAttribValueLen];
                    pwszAttribName = new string((char)0, wAttribNameLen);

                    headerInfo3.GetAttributeByIndex(wAttribIndex, ref wStreamNum, pwszAttribName, ref wAttribNameLen, out wAttribType, pbAttribValue, ref wAttribValueLen);

                    this.PrintAttribute(wAttribIndex, wStreamNum, pwszAttribName, wAttribType, 0, pbAttribValue, wAttribValueLen);
                }
            }
            catch (Exception e)
            {
                this.AddLogEntry(e.Message, LogType.Fail);
                return false;
            }

            return true;
        }

        /// <summary>
        ///     Displays all attributes for the specified stream, with support for GetAttributeByIndexEx.
        /// </summary>
        /// <param name="pwszFileName">Name of the PWSZ file.</param>
        /// <param name="wStreamNum">The w stream num.</param>
        /// <returns>Success of operation.</returns>
        public bool ShowAttributes3(string pwszFileName, ushort wStreamNum)
        {
            try
            {
                WMFSDKFunctions.WMCreateEditor(out var metadataEditor);

                metadataEditor.Open(pwszFileName);

                var headerInfo3 = (IWMHeaderInfo3)metadataEditor;

                headerInfo3.GetAttributeCountEx(wStreamNum, out var wAttributeCount);

                for (ushort wAttribIndex = 0; wAttribIndex < wAttributeCount; wAttribIndex++)
                {
                    string pwszAttribName = null;
                    byte[] pbAttribValue = null;
                    ushort wAttribNameLen = 0;
                    uint dwAttribValueLen = 0;

                    headerInfo3.GetAttributeByIndexEx(wStreamNum, wAttribIndex, pwszAttribName, ref wAttribNameLen, out var wAttribType, out _, pbAttribValue, ref dwAttribValueLen);

                    pwszAttribName = new string((char)0, wAttribNameLen);
                    pbAttribValue = new byte[dwAttribValueLen];

                    headerInfo3.GetAttributeByIndexEx(wStreamNum, wAttribIndex, pwszAttribName, ref wAttribNameLen, out wAttribType, out _, pbAttribValue, ref dwAttribValueLen);

                    this.PrintAttribute(wAttribIndex, wStreamNum, pwszAttribName, wAttribType, 0, pbAttribValue, dwAttribValueLen);
                }

                // Close file
                metadataEditor.Close();
            }
            catch (Exception e)
            {
                this.AddLogEntry(e.Message, LogType.Fail);
                return false;
            }

            return true;
        }

        /// <summary>
        ///     Delete the attribute at the specified index.
        /// </summary>
        /// <param name="pwszFileName">Name of the PWSZ file.</param>
        /// <param name="wStreamNum">The w stream num.</param>
        /// <param name="wAttribIndex">Index of the w attrib.</param>
        /// <returns>Success of operation.</returns>
        public bool DeleteAttrib(string pwszFileName, ushort wStreamNum, ushort wAttribIndex)
        {
            try
            {
                WMFSDKFunctions.WMCreateEditor(out var metadataEditor);

                metadataEditor.Open(pwszFileName);

                var headerInfo3 = (IWMHeaderInfo3)metadataEditor;

                headerInfo3.DeleteAttribute(wStreamNum, wAttribIndex);

                metadataEditor.Flush();

                metadataEditor.Close();
            }
            catch (Exception e)
            {
                this.AddLogEntry(e.Message, LogType.Fail);
                return false;
            }

            return true;
        }

        /// <summary>
        ///     Converts attributes to byte arrays.
        /// </summary>
        /// <param name="attribDataType">Type of the attrib data.</param>
        /// <param name="pwszValue">The PWSZ value.</param>
        /// <param name="pbValue">The pb value.</param>
        /// <param name="nValueLength">Length of the n value.</param>
        /// <returns>Success of operation.</returns>
        public bool TranslateAttrib(WMT_ATTR_DATATYPE attribDataType, string pwszValue, out byte[] pbValue, out int nValueLength)
        {
            switch (attribDataType)
            {
                case WMT_ATTR_DATATYPE.WMT_TYPE_DWORD:

                    nValueLength = 4;
                    var pdwAttribValue = new[] { Convert.ToUInt32(pwszValue) };

                    pbValue = new byte[nValueLength];
                    Buffer.BlockCopy(pdwAttribValue, 0, pbValue, 0, nValueLength);

                    return true;

                case WMT_ATTR_DATATYPE.WMT_TYPE_WORD:

                    nValueLength = 2;
                    var pwAttribValue = new[] { Convert.ToUInt16(pwszValue) };

                    pbValue = new byte[nValueLength];
                    Buffer.BlockCopy(pwAttribValue, 0, pbValue, 0, nValueLength);

                    return true;

                case WMT_ATTR_DATATYPE.WMT_TYPE_QWORD:

                    nValueLength = 8;
                    var pqwAttribValue = new[] { Convert.ToUInt64(pwszValue) };

                    pbValue = new byte[nValueLength];
                    Buffer.BlockCopy(pqwAttribValue, 0, pbValue, 0, nValueLength);

                    return true;

                case WMT_ATTR_DATATYPE.WMT_TYPE_STRING:

                    nValueLength = (ushort)((pwszValue.Length + 1) * 2);
                    pbValue = new byte[nValueLength];

                    Buffer.BlockCopy(pwszValue.ToCharArray(), 0, pbValue, 0, pwszValue.Length * 2);
                    pbValue[nValueLength - 2] = 0;
                    pbValue[nValueLength - 1] = 0;

                    return true;

                case WMT_ATTR_DATATYPE.WMT_TYPE_BOOL:

                    nValueLength = 4;
                    pdwAttribValue = new[] { Convert.ToUInt32(pwszValue) };
                    if (pdwAttribValue[0] != 0)
                    {
                        pdwAttribValue[0] = 1;
                    }

                    pbValue = new byte[nValueLength];
                    Buffer.BlockCopy(pdwAttribValue, 0, pbValue, 0, nValueLength);

                    return true;

                case WMT_ATTR_DATATYPE.WMT_TYPE_GUID:

                    pbValue = HexEncoding.GetBytes(pwszValue, out _);
                    nValueLength = HexEncoding.GetByteCount(pwszValue);

                    return true;

                default:

                    pbValue = null;
                    nValueLength = 0;

                    return false;
            }
        }

        /// <summary>
        ///     Set the specified attribute.
        /// </summary>
        /// <param name="pwszFileName">Name of the PWSZ file.</param>
        /// <param name="wStreamNum">The w stream num.</param>
        /// <param name="pwszAttribName">Name of the PWSZ attrib.</param>
        /// <param name="wAttribType">Type of the w attrib.</param>
        /// <param name="pwszAttribValue">The PWSZ attrib value.</param>
        /// <returns>Success of operation.</returns>
        public bool SetAttrib(string pwszFileName, ushort wStreamNum, string pwszAttribName, ushort wAttribType, string pwszAttribValue)
        {
            try
            {
                var attribDataType = (WMT_ATTR_DATATYPE)wAttribType;

                if (!this.TranslateAttrib(attribDataType, pwszAttribValue, out var pbAttribValue, out var nAttribValueLen))
                {
                    return false;
                }

                WMFSDKFunctions.WMCreateEditor(out var metadataEditor);

                metadataEditor.Open(pwszFileName);

                var headerInfo3 = (IWMHeaderInfo3)metadataEditor;

                headerInfo3.SetAttribute(wStreamNum, pwszAttribName, attribDataType, pbAttribValue, (ushort)nAttribValueLen);

                metadataEditor.Flush();

                metadataEditor.Close();
            }
            catch (Exception e)
            {
                this.AddLogEntry(e.Message, LogType.Fail);
                return false;
            }

            return true;
        }

        /// <summary>
        ///     Add an attribute with the specifed language index.
        /// </summary>
        /// <param name="pwszFileName">Name of the PWSZ file.</param>
        /// <param name="wStreamNum">The w stream num.</param>
        /// <param name="pwszAttribName">Name of the PWSZ attrib.</param>
        /// <param name="wAttribType">Type of the w attrib.</param>
        /// <param name="pwszAttribValue">The PWSZ attrib value.</param>
        /// <param name="wLangIndex">Index of the w lang.</param>
        /// <returns>Success of operation.</returns>
        public bool AddAttrib(string pwszFileName, ushort wStreamNum, string pwszAttribName, ushort wAttribType, string pwszAttribValue, ushort wLangIndex)
        {
            IWMMetadataEditor metadataEditor = null;
            IWMHeaderInfo3 headerInfo3;
            var attribDataType = (WMT_ATTR_DATATYPE)wAttribType;

            try
            {
                if (!this.TranslateAttrib(attribDataType, pwszAttribValue, out var pbAttribValue, out var nAttribValueLen))
                {
                    return false;
                }

                WMFSDKFunctions.WMCreateEditor(out metadataEditor);

                metadataEditor.Open(pwszFileName);

                headerInfo3 = (IWMHeaderInfo3)metadataEditor;

                headerInfo3.AddAttribute(wStreamNum, pwszAttribName, out _, attribDataType, wLangIndex, pbAttribValue, (uint)nAttribValueLen);
            }
            catch (Exception)
            {
                // AddLogEntry(e.Message, LogType.Fail);
                return false;
            }
            finally
            {
                if (metadataEditor != null)
                {
                    metadataEditor.Flush();
                    metadataEditor.Close();
                }
            }

            return true;
        }

        /// <summary>
        ///     Modifies the value of the specified attribute.
        /// </summary>
        /// <param name="pwszFileName">Name of the PWSZ file.</param>
        /// <param name="wStreamNum">The w stream num.</param>
        /// <param name="wAttribIndex">Index of the w attrib.</param>
        /// <param name="wAttribType">Type of the w attrib.</param>
        /// <param name="pwszAttribValue">The PWSZ attrib value.</param>
        /// <param name="wLangIndex">Index of the w lang.</param>
        /// <returns>Success of operation.</returns>
        public bool ModifyAttrib(string pwszFileName, ushort wStreamNum, ushort wAttribIndex, ushort wAttribType, string pwszAttribValue, ushort wLangIndex)
        {
            IWMMetadataEditor metadataEditor = null;
            IWMHeaderInfo3 headerInfo3;
            var attribDataType = (WMT_ATTR_DATATYPE)wAttribType;

            try
            {
                if (!this.TranslateAttrib(attribDataType, pwszAttribValue, out var pbAttribValue, out var nAttribValueLen))
                {
                    return false;
                }

                WMFSDKFunctions.WMCreateEditor(out metadataEditor);

                metadataEditor.Open(pwszFileName);

                headerInfo3 = (IWMHeaderInfo3)metadataEditor;

                headerInfo3.ModifyAttribute(wStreamNum, wAttribIndex, attribDataType, wLangIndex, pbAttribValue, (uint)nAttribValueLen);
            }
            catch (Exception e)
            {
                this.AddLogEntry(e.Message, LogType.Fail);
                return false;
            }
            finally
            {
                if (metadataEditor != null)
                {
                    metadataEditor.Flush();
                    metadataEditor.Close();
                }
            }

            return true;
        }

        /// <summary>
        ///     Attribs the exists.
        /// </summary>
        /// <param name="pwszFileName">Name of the PWSZ file.</param>
        /// <param name="wStreamNum">The w stream num.</param>
        /// <param name="searchAttrib">The search attrib.</param>
        /// <returns>Success of operation.</returns>
        public bool AttribExists(string pwszFileName, ushort wStreamNum, string searchAttrib)
        {
            var isFound = false;

            try
            {
                IWMHeaderInfo3 headerInfo3;
                ushort wAttributeCount;

                WMFSDKFunctions.WMCreateEditor(out var metadataEditor);

                metadataEditor.Open(pwszFileName);

                headerInfo3 = (IWMHeaderInfo3)metadataEditor;

                headerInfo3.GetAttributeCountEx(wStreamNum, out wAttributeCount);

                for (ushort wAttribIndex = 0; wAttribIndex < wAttributeCount && !isFound; wAttribIndex++)
                {
                    string pwszAttribName = null;
                    byte[] pbAttribValue = null;
                    ushort wAttribNameLen = 0;
                    uint dwAttribValueLen = 0;

                    headerInfo3.GetAttributeByIndexEx(wStreamNum, wAttribIndex, pwszAttribName, ref wAttribNameLen, out _, out _, pbAttribValue, ref dwAttribValueLen);

                    pwszAttribName = new string((char)0, wAttribNameLen);
                    pbAttribValue = new byte[dwAttribValueLen];

                    headerInfo3.GetAttributeByIndexEx(wStreamNum, wAttribIndex, pwszAttribName, ref wAttribNameLen, out _, out _, pbAttribValue, ref dwAttribValueLen);

                    if (pwszAttribName.Substring(0, pwszAttribName.Length - 1) == searchAttrib)
                    {
                        isFound = true;
                    }
                }

                // Close file
                metadataEditor.Close();
            }
            catch (Exception e)
            {
                this.AddLogEntry(e.Message, LogType.Fail);
                return false;
            }

            return isFound;
        }

        /// <summary>
        ///     Initialises the media types.
        /// </summary>
        private void InitMediaTypes()
        {
            this.cbMediaType.Items.Clear();
            this.cbMediaType.Items.Add("Unknown");
            this.cbMediaType.Items.Add("Video");
            this.cbMediaType.Items.Add("Movies");
            this.cbMediaType.Items.Add("Music Videos");
            this.cbMediaType.Items.Add("TV Shows");
            this.cbMediaType.SelectedIndex = 0;
        }

        /// <summary>
        ///     Updates the status view by flushing the buffer.
        /// </summary>
        /// <param name="sender">The sender.</param>
        private void ScreenLogTimer_Elapsed(object sender)
        {
            this.CycleStatusView();
            this.screenLogTimer.Change(Convert.ToInt32(ScreenRefresh) * 1000, Timeout.Infinite);
        }

        /// <summary>
        ///     Cycles the status view.
        /// </summary>
        private void CycleStatusView()
        {
            this.PauseOutput(this.lvStatus);
            this.FlushOutput(this.lvStatus);
            this.ResumeOutput(this.lvStatus);
        }

        /// <summary>
        ///     Flushes the output.
        /// </summary>
        /// <param name="lv">The lv.</param>
        private void FlushOutput(ListView lv)
        {
            if (this.InvokeRequired)
            {
                this.BeginInvoke(new FlushOutputDelegate(this.FlushOutput), lv);
                return;
            }

            if (this.lvitems.Count > 0)
            {
                if (lv.Items.Count >= Convert.ToInt32(ScreenLines))
                {
                    lv.Items.Clear();
                }

                lv.BeginUpdate();
                lv.Items.AddRange(this.lvitems.ToArray());
                lv.EnsureVisible(lv.Items.Count - 1);
                lv.EndUpdate();
                this.lvitems.Clear();
            }
        }

        /// <summary>
        ///     Pauses the output.
        /// </summary>
        /// <param name="lv">The lv.</param>
        private void PauseOutput(ListView lv)
        {
            if (this.InvokeRequired)
            {
                this.BeginInvoke(new PauseOutputDelegate(this.PauseOutput), lv);
                return;
            }

            lv.BeginUpdate();
        }

        /// <summary>
        ///     Resumes the output.
        /// </summary>
        /// <param name="lv">The lv.</param>
        private void ResumeOutput(ListView lv)
        {
            if (this.InvokeRequired)
            {
                this.BeginInvoke(new ResumeOutputDelegate(this.ResumeOutput), lv);
                return;
            }

            lv.EndUpdate();
        }

        /// <summary>
        ///     Initialises the event view.
        /// </summary>
        /// <param name="lvX">The lv X.</param>
        private void InitEventView(ListView lvX)
        {
            lvX.Columns.Add("Time", 140, HorizontalAlignment.Left);
            lvX.Columns.Add("Event Details", 1000, HorizontalAlignment.Left);
            lvX.Items.Clear();
        }

        /// <summary>
        ///     Adds the log entry.
        /// </summary>
        /// <param name="newEntry">The new entry.</param>
        /// <param name="whichLog">The which log.</param>
        private void AddLogEntry(string newEntry, LogType whichLog = LogType.Success)
        {
            switch (whichLog)
            {
                case LogType.Success:
                    this.lvitems.Add(new ListViewItem(DateTime.Now.ToString(CultureInfo.CurrentCulture), 0));
                    break;

                case LogType.Fail:
                    this.lvitems.Add(new ListViewItem(DateTime.Now.ToString(CultureInfo.CurrentCulture), 1));
                    break;

                case LogType.Info:
                    this.lvitems.Add(new ListViewItem(DateTime.Now.ToString(CultureInfo.CurrentCulture), 2));
                    break;
            }

            var i = this.lvitems.Count - 1;
            this.lvitems[i].SubItems.Add(newEntry);
            this.slStatus.Text = newEntry;
        }

        /// <summary>
        ///     Handles the Click event of the cmdBrowse control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="System.EventArgs" /> instance containing the event data.</param>
        private void CmdBrowse_Click(object sender, EventArgs e)
        {
            // Get file
            var openFileDialog1 = new OpenFileDialog
            {
                InitialDirectory = Environment.CurrentDirectory,
                Filter = @"Windows Media Video (*.wmv)|*.wmv",
                FilterIndex = 1,
                RestoreDirectory = false,
            };

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                this.lblMediaFile.Text = openFileDialog1.FileName;
                _ = this.RegisterNewMediaFileAsync();
            }
        }

        /// <summary>
        ///     Registers the new media file.
        /// </summary>
        private async Task RegisterNewMediaFileAsync()
        {
            var input = this.lblMediaFile.Text;
            var output = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".png");

            if (!File.Exists(input))
            {
                this.AddLogEntry("Can't load - media file not found", LogType.Fail);
                return;
            }

            // Grab still frame, if possible
            var conversion = await FFmpeg.Conversions.FromSnippet.Snapshot(input, output, TimeSpan.FromSeconds(1));
            var result = await conversion.Start();

            // Load frame into PB
            var fs = new FileStream(output, FileMode.Open, FileAccess.Read);
            this.pictureBox1.Image = Image.FromStream(fs);
            fs.Close();

            // Make sure all supported attributes are defined
            this.AddMissingAttributes();

            // Refresh attributes from file
            this.InspectFile();

            // Logging
            this.AddLogEntry(this.lblMediaFile.Text + " successfully loaded");
        }

        /// <summary>
        ///     Handles the Click event of the aboutToolStripMenuItem control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="System.EventArgs" /> instance containing the event data.</param>
        private void AboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (this.frmAbout == null)
            {
                this.frmAbout = new AboutBox1();
            }

            this.frmAbout.ShowDialog();
        }

        /// <summary>
        ///     Handles the Click event of the exitToolStripMenuItem1 control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="System.EventArgs" /> instance containing the event data.</param>
        private void ExitToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        /// <summary>
        ///     Inspects the file.
        /// </summary>
        private void InspectFile()
        {
            var isVideo = false;
            var thisIndex = 0;

            // Destroy existing attribute list
            this.attributes.Clear();
            this.cbAttributes.Items.Clear();

            // Get attributes
            if (!this.ShowAttributes3(this.lblMediaFile.Text, Stream))
            {
                return;
            }

            // Update from file
            foreach (var u in this.attributes)
            {
                // Add to combo box
                this.cbAttributes.Items.Add(u);

                // Is this a video
                if (u.Name == "WM/MediaClassPrimaryID")
                {
                    this.indexPrimaryVideo = u.Index;
                    isVideo = u.Value == TypeVideo;
                    if (isVideo)
                    {
                        thisIndex = 1;
                    }
                }

                // Type of video?
                if (u.Name == "WM/MediaClassSecondaryID")
                {
                    this.indexSecondaryVideo = u.Index;

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

                            case TypeTv:
                                thisIndex = 4;
                                break;
                        }
                    }
                }
            }

            this.cbAttributes.SelectedIndex = 0;
            this.cbMediaType.SelectedIndex = thisIndex;

            // Update fields
            this.txtSearchCriteria.Text = string.Empty;
            switch (thisIndex)
            {
                case 1:
                    this.LoadVideoAttributes();
                    break;

                case 2:
                    this.LoadMovieAttributes();
                    break;

                case 3:
                    this.LoadMusicAttributes();
                    break;

                case 4:
                    this.LoadTvAttributes();
                    break;
            }

            // If no title has been established then look at filename
            if (this.txtSearchCriteria.Text.Length == 0)
            {
                var filename = Path.GetFileNameWithoutExtension(this.lblMediaFile.Text);
                filename = filename.Replace(".", " ");
                filename = filename.Replace("_", " ");
                filename = filename.Replace("-", " ");
                this.txtSearchCriteria.Text = filename;
            }
        }

        /// <summary>
        ///     Handles the 1 event of the cmdModify_Click control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="System.EventArgs" /> instance containing the event data.</param>
        private void CmdModify_Click_1(object sender, EventArgs e)
        {
            if (!File.Exists(this.lblMediaFile.Text))
            {
                return;
            }

            if (this.cbAttributes.SelectedItem == null)
            {
                return;
            }

            if (this.txtNewValue.Text.Length == 0)
            {
                return;
            }

            var attribute = (Attribute)this.cbAttributes.SelectedItem;
            if (attribute == null)
            {
                return;
            }

            if (this.ModifyAttrib(this.lblMediaFile.Text, Stream, attribute.Index, Convert.ToUInt16(attribute.Type), this.txtNewValue.Text, Language))
            {
                this.AddLogEntry(attribute.Name + " successfully modified");
                this.InspectFile();
            }
        }

        /// <summary>
        ///     Handles the SelectedIndexChanged event of the cbAttributes control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="System.EventArgs" /> instance containing the event data.</param>
        private void CbAttributes_SelectedIndexChanged(object sender, EventArgs e)
        {
            var attribute = (Attribute)this.cbAttributes.SelectedItem;
            if (attribute == null)
            {
                return;
            }

            this.txtNewValue.Text = attribute.Value;
        }

        /// <summary>
        ///     Handles the Click event of the cmdSave control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="System.EventArgs" /> instance containing the event data.</param>
        private void CmdSave_Click(object sender, EventArgs e)
        {
            var newType = string.Empty;

            if (!File.Exists(this.lblMediaFile.Text))
            {
                return;
            }

            switch (this.cbMediaType.SelectedIndex)
            {
                case 1:
                    newType = "Generic Video";
                    this.EditAttribute("WM/MediaClassPrimaryID", TypeVideo);
                    this.EditAttribute("WM/MediaClassSecondaryID", TypeVideo);
                    this.ModifyAttrib(this.lblMediaFile.Text, Stream, this.indexPrimaryVideo, Convert.ToUInt16(WMT_ATTR_DATATYPE.WMT_TYPE_GUID), TypeVideo, Language);
                    this.ModifyAttrib(this.lblMediaFile.Text, Stream, this.indexSecondaryVideo, Convert.ToUInt16(WMT_ATTR_DATATYPE.WMT_TYPE_GUID), TypeVideo, Language);
                    break;

                case 2:
                    newType = "Movie";
                    this.EditAttribute("WM/MediaClassPrimaryID", TypeVideo);
                    this.EditAttribute("WM/MediaClassSecondaryID", TypeMovie);
                    this.ModifyAttrib(this.lblMediaFile.Text, Stream, this.indexPrimaryVideo, Convert.ToUInt16(WMT_ATTR_DATATYPE.WMT_TYPE_GUID), TypeVideo, Language);
                    this.ModifyAttrib(this.lblMediaFile.Text, Stream, this.indexSecondaryVideo, Convert.ToUInt16(WMT_ATTR_DATATYPE.WMT_TYPE_GUID), TypeMovie, Language);
                    break;

                case 3:
                    newType = "Music Video";
                    this.EditAttribute("WM/MediaClassPrimaryID", TypeVideo);
                    this.EditAttribute("WM/MediaClassSecondaryID", TypeMusic);
                    this.ModifyAttrib(this.lblMediaFile.Text, Stream, this.indexPrimaryVideo, Convert.ToUInt16(WMT_ATTR_DATATYPE.WMT_TYPE_GUID), TypeVideo, Language);
                    this.ModifyAttrib(this.lblMediaFile.Text, Stream, this.indexSecondaryVideo, Convert.ToUInt16(WMT_ATTR_DATATYPE.WMT_TYPE_GUID), TypeMusic, Language);
                    break;

                case 4:
                    newType = "TV Show";
                    this.EditAttribute("WM/MediaClassPrimaryID", TypeVideo);
                    this.EditAttribute("WM/MediaClassSecondaryID", TypeTv);
                    this.ModifyAttrib(this.lblMediaFile.Text, Stream, this.indexPrimaryVideo, Convert.ToUInt16(WMT_ATTR_DATATYPE.WMT_TYPE_GUID), TypeVideo, Language);
                    this.ModifyAttrib(this.lblMediaFile.Text, Stream, this.indexSecondaryVideo, Convert.ToUInt16(WMT_ATTR_DATATYPE.WMT_TYPE_GUID), TypeTv, Language);
                    break;
            }

            // Refresh screen
            if (newType.Length > 0)
            {
                this.AddLogEntry("File modified to be a " + newType);
                this.InspectFile();
            }
        }

        /// <summary>
        ///     Adds the missing attributes.
        /// </summary>
        private void AddMissingAttributes()
        {
            if (!this.AttribExists(this.lblMediaFile.Text, Stream, "WM/MediaClassPrimaryID"))
            {
                this.AddAttrib(this.lblMediaFile.Text, NewStream, "WM/MediaClassPrimaryID", Convert.ToUInt16(WMT_ATTR_DATATYPE.WMT_TYPE_GUID), TypeVideo, Language);
            }

            if (!this.AttribExists(this.lblMediaFile.Text, Stream, "WM/MediaClassSecondaryID"))
            {
                this.AddAttrib(this.lblMediaFile.Text, NewStream, "WM/MediaClassSecondaryID", Convert.ToUInt16(WMT_ATTR_DATATYPE.WMT_TYPE_GUID), TypeVideo, Language);
            }

            if (!this.AttribExists(this.lblMediaFile.Text, Stream, "Title"))
            {
                this.AddAttrib(this.lblMediaFile.Text, NewStream, "Title", Convert.ToUInt16(WMT_ATTR_DATATYPE.WMT_TYPE_STRING), "Unknown", Language);
            }

            if (!this.AttribExists(this.lblMediaFile.Text, Stream, "WM/SubTitle"))
            {
                this.AddAttrib(this.lblMediaFile.Text, NewStream, "WM/SubTitle", Convert.ToUInt16(WMT_ATTR_DATATYPE.WMT_TYPE_STRING), "Unknown", Language);
            }

            if (!this.AttribExists(this.lblMediaFile.Text, Stream, "WM/SubTitleDescription"))
            {
                this.AddAttrib(this.lblMediaFile.Text, NewStream, "WM/SubTitleDescription", Convert.ToUInt16(WMT_ATTR_DATATYPE.WMT_TYPE_STRING), "Unknown", Language);
            }

            if (!this.AttribExists(this.lblMediaFile.Text, Stream, "Author"))
            {
                this.AddAttrib(this.lblMediaFile.Text, NewStream, "Author", Convert.ToUInt16(WMT_ATTR_DATATYPE.WMT_TYPE_STRING), "Unknown", Language);
            }

            if (!this.AttribExists(this.lblMediaFile.Text, Stream, "WM/Year"))
            {
                this.AddAttrib(this.lblMediaFile.Text, NewStream, "WM/Year", Convert.ToUInt16(WMT_ATTR_DATATYPE.WMT_TYPE_STRING), "Unknown", Language);
            }

            if (!this.AttribExists(this.lblMediaFile.Text, Stream, "WM/OriginalBroadcastDateTime"))
            {
                this.AddAttrib(this.lblMediaFile.Text, NewStream, "WM/OriginalBroadcastDateTime", Convert.ToUInt16(WMT_ATTR_DATATYPE.WMT_TYPE_STRING), "Unknown", Language);
            }

            if (!this.AttribExists(this.lblMediaFile.Text, Stream, "WM/ParentalRating"))
            {
                this.AddAttrib(this.lblMediaFile.Text, NewStream, "WM/ParentalRating", Convert.ToUInt16(WMT_ATTR_DATATYPE.WMT_TYPE_STRING), "Unknown", Language);
            }

            if (!this.AttribExists(this.lblMediaFile.Text, Stream, "WM/TVNetworkAffiliation"))
            {
                this.AddAttrib(this.lblMediaFile.Text, NewStream, "WM/TVNetworkAffiliation", Convert.ToUInt16(WMT_ATTR_DATATYPE.WMT_TYPE_STRING), "Unknown", Language);
            }

            if (!this.AttribExists(this.lblMediaFile.Text, Stream, "WM/Genre"))
            {
                this.AddAttrib(this.lblMediaFile.Text, NewStream, "WM/Genre", Convert.ToUInt16(WMT_ATTR_DATATYPE.WMT_TYPE_STRING), "Unknown", Language);
            }

            if (!this.AttribExists(this.lblMediaFile.Text, Stream, "Description"))
            {
                this.AddAttrib(this.lblMediaFile.Text, NewStream, "Description", Convert.ToUInt16(WMT_ATTR_DATATYPE.WMT_TYPE_STRING), "Unknown", Language);
            }

            if (!this.AttribExists(this.lblMediaFile.Text, Stream, "WM/TrackNumber"))
            {
                this.AddAttrib(this.lblMediaFile.Text, NewStream, "WM/TrackNumber", Convert.ToUInt16(WMT_ATTR_DATATYPE.WMT_TYPE_DWORD), "01", Language);
            }
        }

        /// <summary>
        ///     Handles the SelectedIndexChanged event of the cbMediaType control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="System.EventArgs" /> instance containing the event data.</param>
        private void CbMediaType_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (this.cbMediaType.SelectedIndex)
            {
                case 0:
                    this.gbMovie.Visible = false;
                    this.gbMusic.Visible = false;
                    this.gbTV.Visible = false;
                    this.gbVideo.Visible = false;
                    break;

                case 1:
                    this.gbMovie.Visible = false;
                    this.gbMusic.Visible = false;
                    this.gbTV.Visible = false;
                    this.LoadVideoAttributes();
                    this.gbVideo.Visible = true;
                    break;

                case 2:
                    this.gbMusic.Visible = false;
                    this.gbTV.Visible = false;
                    this.gbVideo.Visible = false;
                    this.LoadMovieAttributes();
                    this.gbMovie.Visible = true;
                    break;

                case 3:
                    this.gbMovie.Visible = false;
                    this.gbTV.Visible = false;
                    this.gbVideo.Visible = false;
                    this.LoadMusicAttributes();
                    this.gbMusic.Visible = true;
                    break;

                case 4:
                    this.gbMovie.Visible = false;
                    this.gbMusic.Visible = false;
                    this.gbVideo.Visible = false;
                    this.LoadTvAttributes();
                    this.gbTV.Visible = true;
                    break;
            }
        }

        /// <summary>
        ///     Loads the TV attributes.
        /// </summary>
        private void LoadTvAttributes()
        {
            foreach (var attribute in this.attributes)
            {
                var myValue = attribute.Value.Replace("\"", string.Empty);

                switch (attribute.Name)
                {
                    case "Title":
                        this.txtTVTitle.Text = myValue;
                        this.txtSearchCriteria.Text = this.txtTVTitle.Text;
                        break;

                    case "WM/SubTitle":
                        this.txtTVSubTitle.Text = myValue;
                        break;

                    case "WM/SubTitleDescription":
                        this.txtTVDescription.Text = myValue;
                        break;

                    case "Author":
                        this.txtTVAuthor.Text = myValue;
                        break;

                    case "WM/Year":
                        this.txtTVYear.Text = myValue;
                        break;

                    case "WM/OriginalBroadcastDateTime":
                        this.txtTVDate.Text = myValue;
                        break;

                    case "WM/ParentalRating":
                        this.txtTVRating.Text = myValue;
                        break;

                    case "WM/TVNetworkAffiliation":
                        this.txtTVNetwork.Text = myValue;
                        break;

                    case "WM/Genre":
                        this.txtTVGenre.Text = myValue;
                        break;

                    case "WM/TrackNumber":
                        this.txtTVTrack.Text = myValue;
                        break;
                }
            }
        }

        /// <summary>
        ///     Loads the movie attributes.
        /// </summary>
        private void LoadMovieAttributes()
        {
            foreach (var attribute in this.attributes)
            {
                var myValue = attribute.Value.Replace("\"", string.Empty);

                switch (attribute.Name)
                {
                    case "Title":
                        this.txtMovieTitle.Text = myValue;
                        this.txtSearchCriteria.Text = this.txtMovieTitle.Text;
                        break;

                    case "WM/SubTitleDescription":
                        this.txtMovieDescription.Text = myValue;
                        break;

                    case "Author":
                        this.txtMovieAuthor.Text = myValue;
                        break;

                    case "WM/Year":
                        this.txtMovieYear.Text = myValue;
                        break;

                    case "WM/OriginalBroadcastDateTime":
                        this.txtMovieDate.Text = myValue;
                        break;

                    case "WM/ParentalRating":
                        this.txtMovieRating.Text = myValue;
                        break;

                    case "WM/Genre":
                        this.txtMovieGenre.Text = myValue;
                        break;
                }
            }
        }

        /// <summary>
        ///     Loads the video attributes.
        /// </summary>
        private void LoadVideoAttributes()
        {
            foreach (var attribute in this.attributes)
            {
                var myValue = attribute.Value.Replace("\"", string.Empty);

                switch (attribute.Name)
                {
                    case "Title":
                        this.txtVideoTitle.Text = myValue;
                        this.txtSearchCriteria.Text = this.txtVideoTitle.Text;
                        break;

                    case "WM/SubTitleDescription":
                        this.txtVideoDescription.Text = myValue;
                        break;

                    case "Author":
                        this.txtVideoAuthor.Text = myValue;
                        break;

                    case "WM/Year":
                        this.txtVideoYear.Text = myValue;
                        break;

                    case "WM/Genre":
                        this.txtVideoGenre.Text = myValue;
                        break;
                }
            }
        }

        /// <summary>
        ///     Loads the music attributes.
        /// </summary>
        private void LoadMusicAttributes()
        {
            foreach (var attribute in this.attributes)
            {
                var myValue = attribute.Value.Replace("\"", string.Empty);

                switch (attribute.Name)
                {
                    case "Title":
                        this.txtMusicTitle.Text = myValue;
                        this.txtSearchCriteria.Text = this.txtMusicTitle.Text;
                        break;

                    case "WM/SubTitleDescription":
                        this.txtMusicDescription.Text = myValue;
                        break;

                    case "Author":
                        this.txtMusicAuthor.Text = myValue;
                        break;

                    case "WM/Year":
                        this.txtMusicYear.Text = myValue;
                        break;

                    case "WM/OriginalBroadcastDateTime":
                        this.txtMusicDate.Text = myValue;
                        break;

                    case "WM/ParentalRating":
                        this.txtMusicRating.Text = myValue;
                        break;

                    case "WM/Genre":
                        this.txtMusicGenre.Text = myValue;
                        break;
                }
            }
        }

        /// <summary>
        ///     Handles the Click event of the cmdTVSave control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="System.EventArgs" /> instance containing the event data.</param>
        private void CmdTVSave_Click(object sender, EventArgs e)
        {
            this.EditAttribute("Title", this.txtTVTitle.Text);
            this.EditAttribute("WM/SubTitle", this.txtTVSubTitle.Text);
            this.EditAttribute("WM/SubTitleDescription", this.txtTVDescription.Text);
            this.EditAttribute("Author", this.txtTVAuthor.Text);
            this.EditAttribute("WM/Year", this.txtTVYear.Text);
            this.EditAttribute("WM/OriginalBroadcastDateTime", this.txtTVDate.Text);
            this.EditAttribute("WM/ParentalRating", this.txtTVRating.Text);
            this.EditAttribute("WM/TVNetworkAffiliation", this.txtTVNetwork.Text);
            this.EditAttribute("WM/Genre", this.txtTVGenre.Text);
            this.EditAttribute("WM/TrackNumber", this.txtTVTrack.Text);
            this.EditPicture(this.pictureBox1);
            this.AddLogEntry(this.lblMediaFile.Text + " successfully modified");
            this.InspectFile();
        }

        /// <summary>
        ///     Modifies the attribute.
        /// </summary>
        /// <param name="theAttribute">The attribute.</param>
        /// <param name="newValue">The new value.</param>
        private void EditAttribute(string theAttribute, string newValue)
        {
            foreach (var attribute in this.attributes.Where(attribute => attribute.Name == theAttribute))
            {
                this.ModifyAttrib(this.lblMediaFile.Text, Stream, attribute.Index, Convert.ToUInt16(attribute.Type), newValue, Language);
                break;
            }
        }

        /// <summary>
        ///     Handles the Click event of the cmdMovieSave control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="System.EventArgs" /> instance containing the event data.</param>
        private void CmdMovieSave_Click(object sender, EventArgs e)
        {
            this.EditAttribute("Title", this.txtMovieTitle.Text);
            this.EditAttribute("WM/SubTitleDescription", this.txtMovieDescription.Text);
            this.EditAttribute("Author", this.txtMovieAuthor.Text);
            this.EditAttribute("WM/Year", this.txtMovieYear.Text);
            this.EditAttribute("WM/OriginalBroadcastDateTime", this.txtMovieDate.Text);
            this.EditAttribute("WM/ParentalRating", this.txtMovieRating.Text);
            this.EditAttribute("WM/Genre", this.txtMovieGenre.Text);
            this.EditPicture(this.pictureBox1);
            this.AddLogEntry(this.lblMediaFile.Text + " successfully modified");
            this.InspectFile();
        }

        /// <summary>
        ///     Handles the Click event of the cmdMusicSave control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="System.EventArgs" /> instance containing the event data.</param>
        private void CmdMusicSave_Click(object sender, EventArgs e)
        {
            this.EditAttribute("Title", this.txtMusicTitle.Text);
            this.EditAttribute("WM/SubTitleDescription", this.txtMusicDescription.Text);
            this.EditAttribute("Author", this.txtMusicAuthor.Text);
            this.EditAttribute("WM/Year", this.txtMusicYear.Text);
            this.EditAttribute("WM/OriginalBroadcastDateTime", this.txtMusicDate.Text);
            this.EditAttribute("WM/ParentalRating", this.txtMusicRating.Text);
            this.EditAttribute("WM/Genre", this.txtMusicGenre.Text);
            this.EditPicture(this.pictureBox1);
            this.AddLogEntry(this.lblMediaFile.Text + " successfully modified");
            this.InspectFile();
        }

        /// <summary>
        ///     Handles the Click event of the cmdVideoSave control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="System.EventArgs" /> instance containing the event data.</param>
        private void CmdVideoSave_Click(object sender, EventArgs e)
        {
            this.EditAttribute("Title", this.txtVideoTitle.Text);
            this.EditAttribute("WM/SubTitleDescription", this.txtVideoDescription.Text);
            this.EditAttribute("Author", this.txtVideoAuthor.Text);
            this.EditAttribute("WM/Year", this.txtVideoYear.Text);
            this.EditAttribute("WM/Genre", this.txtVideoGenre.Text);
            this.EditPicture(this.pictureBox1);
            this.AddLogEntry(this.lblMediaFile.Text + " successfully modified");
            this.InspectFile();
        }

        /// <summary>
        ///     Handles the Click event of the cmdSearchTmdb control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="System.EventArgs" /> instance containing the event data.</param>
        private void CmdSearchTmdb_Click(object sender, EventArgs e)
        {
            // Set up visuals pre-search
            this.lbResults.Items.Clear();
            this.SetButtons(false);
            this.progressBar1.Visible = true;
            this.Cursor = Cursors.WaitCursor;
            this.AddLogEntry("Searching TMDB for " + this.txtSearchCriteria.Text + "...", LogType.Info);
            Application.DoEvents();

            // Do the search
            this.backgroundWorker1.RunWorkerAsync();
        }

        /// <summary>
        ///     Sets the buttons.
        /// </summary>
        /// <param name="isOn">if set to <c>true</c> [is ON].</param>
        private void SetButtons(bool isOn)
        {
            this.cmdSearchTmdb.Enabled = isOn;
        }

        /// <summary>
        ///     Shows the cover art.
        /// </summary>
        /// <param name="coverUrl">The URL of the cover art.</param>
        private void ShowCoverArt(string coverUrl)
        {
            if (string.IsNullOrEmpty(coverUrl))
            {
                return;
            }

            var client = new WebClient();
            client.Headers["User-Agent"] = "Mozilla/4.0";
            var bytes = client.DownloadData(Settings.Default.PosterBase + coverUrl);
            var stream = new MemoryStream(bytes);
            this.pbCover.Image = Image.FromStream(stream);
        }

        /// <summary>
        ///     Handles the KeyDown event of the txtSearchCriteria control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="System.Windows.Forms.KeyEventArgs" /> instance containing the event data.</param>
        private void TxtSearchCriteria_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                this.CmdSearchTmdb_Click(sender, e);
            }
        }

        /// <summary>
        ///     Handles the SelectedIndexChanged event of the lbResults control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="System.EventArgs" /> instance containing the event data.</param>
        private void LbResults_SelectedIndexChanged(object sender, EventArgs e)
        {
            var entry = (TMDbSearchResult)this.lbResults.SelectedItem;
            if (entry == null)
            {
                return;
            }

            // If not already retrieved, get additional data
            if (!entry.ExtraData)
            {
                var client = new TMDbClient(Settings.Default.APIkey);
                var movie = client.GetMovieAsync(entry.Movie.Id).Result;
                if (movie != null)
                {
                    entry.Genre = movie.Genres.Count > 0 ? movie.Genres[0].Name : "unknown";
                    entry.Url = movie.Homepage;
                }

                var credits = client.GetMovieCreditsAsync(entry.Movie.Id).Result;
                if (credits != null)
                {
                    foreach (var credit in credits.Crew.Where(credit => credit.Department == "Directing"))
                    {
                        entry.Director = credit.Name;
                        break;
                    }
                }

                entry.ExtraData = true;
            }

            this.txtAzTitle.Text = entry.Movie.Title;
            this.txtAzYear.Text = entry.Movie.ReleaseDate != null ? entry.Movie.ReleaseDate.Value.Year.ToString() : @"unknown";

            this.txtAzDirector.Text = entry.Director;
            this.txtAzDescription.Text = entry.Movie.Overview;
            this.ShowCoverArt(entry.Movie.PosterPath);
        }

        /// <summary>
        ///     Handles the Click event of the cmdCopyResult control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="System.EventArgs" /> instance containing the event data.</param>
        private void CmdCopyResult_Click(object sender, EventArgs e)
        {
            if (this.lbResults.SelectedItem == null)
            {
                return;
            }

            if (this.lblMediaFile.Text.Length == 0)
            {
                return;
            }

            var entry = (TMDbSearchResult)this.lbResults.SelectedItem;
            if (entry == null)
            {
                return;
            }

            switch (this.cbMediaType.SelectedIndex)
            {
                case 0:
                    break;

                case 1:
                    this.AddLogEntry("Copying details to current media file...", LogType.Info);
                    this.txtVideoAuthor.Text = entry.Director;
                    this.txtVideoDescription.Text = entry.Movie.Overview;
                    this.txtVideoGenre.Text = entry.Genre;
                    this.txtVideoTitle.Text = entry.Movie.Title;
                    if (entry.Movie.ReleaseDate != null)
                    {
                        this.txtVideoYear.Text = entry.Movie.ReleaseDate.Value.Year.ToString();
                    }

                    break;

                case 2:
                    this.AddLogEntry("Copying details to current media file...", LogType.Info);
                    this.txtMovieAuthor.Text = entry.Director;
                    if (entry.Movie.ReleaseDate != null)
                    {
                        this.txtMovieDate.Text = entry.Movie.ReleaseDate.Value.Date.ToString(CultureInfo.CurrentCulture);
                    }

                    this.txtMovieDescription.Text = entry.Movie.Overview;
                    this.txtMovieGenre.Text = entry.Genre;
                    this.txtMovieRating.Text = entry.Movie.VoteAverage.ToString(CultureInfo.CurrentCulture);
                    this.txtMovieTitle.Text = entry.Movie.Title;
                    if (entry.Movie.ReleaseDate != null)
                    {
                        this.txtMovieYear.Text = entry.Movie.ReleaseDate.Value.Year.ToString();
                    }

                    break;

                case 3:
                    this.AddLogEntry("Copying details to current media file...", LogType.Info);
                    this.txtMusicAuthor.Text = entry.Director;
                    if (entry.Movie.ReleaseDate != null)
                    {
                        this.txtMusicDate.Text = entry.Movie.ReleaseDate.Value.Date.ToString(CultureInfo.CurrentCulture);
                    }

                    this.txtMusicDescription.Text = entry.Movie.Overview;
                    this.txtMusicGenre.Text = entry.Genre;
                    this.txtMusicRating.Text = entry.Movie.VoteAverage.ToString(CultureInfo.CurrentCulture);
                    this.txtMusicTitle.Text = entry.Movie.Title;
                    if (entry.Movie.ReleaseDate != null)
                    {
                        this.txtMusicYear.Text = entry.Movie.ReleaseDate.Value.Year.ToString();
                    }

                    break;

                case 4:
                    this.AddLogEntry("Copying details to current media file...", LogType.Info);
                    this.txtTVAuthor.Text = entry.Director;
                    if (entry.Movie.ReleaseDate != null)
                    {
                        this.txtTVDate.Text = entry.Movie.ReleaseDate.Value.Date.ToString(CultureInfo.CurrentCulture);
                    }

                    this.txtTVDescription.Text = entry.Movie.Overview;
                    this.txtTVGenre.Text = entry.Genre;
                    this.txtTVRating.Text = entry.Movie.VoteAverage.ToString(CultureInfo.CurrentCulture);
                    this.txtTVTitle.Text = entry.Movie.Title;
                    if (entry.Movie.ReleaseDate != null)
                    {
                        this.txtTVYear.Text = entry.Movie.ReleaseDate.Value.Year.ToString();
                    }

                    break;
            }

            // Screen grab
            this.pictureBox1.Image = this.pbCover.Image;

            // Switch back to the first tab to see the results
            this.AddLogEntry("Media file details updated.");
            this.tabControl1.SelectedIndex = 0;
        }

        /// <summary>
        ///     Handles the DoWork event of the backgroundWorker1 control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="System.ComponentModel.DoWorkEventArgs" /> instance containing the event data.</param>
        private void BackgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            if (this.backgroundWorker1.CancellationPending)
            {
                e.Cancel = true;
                return;
            }

            if (this.txtSearchCriteria.Text.Length == 0)
            {
                e.Cancel = true;
                return;
            }

            var client = new TMDbClient(Settings.Default.APIkey);
            var results = client.SearchMovieAsync(this.txtSearchCriteria.Text).Result;
            e.Result = results;
        }

        /// <summary>
        ///     Handles the RunWorkerCompleted event of the backgroundWorker1 control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">
        ///     The <see cref="System.ComponentModel.RunWorkerCompletedEventArgs" /> instance containing the event
        ///     data.
        /// </param>
        private void BackgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            // Check for errors
            if (e.Cancelled)
            {
                return;
            }

            if (e.Error != null)
            {
                this.AddLogEntry($"Error detected - {e.Error}");
                return;
            }

            if (e.Result == null)
            {
                return;
            }

            // Display results
            var results = (SearchContainer<SearchMovie>)e.Result;
            foreach (var entry in results.Results.Select(result => new TMDbSearchResult
                     {
                         Movie = result,
                     }))
            {
                this.lbResults.Items.Add(entry);
            }

            // Visuals
            this.Cursor = Cursors.Default;
            this.progressBar1.Visible = false;
            this.SetButtons(true);
            this.slStatus.Text = string.Empty;
        }

        /// <summary>
        ///     Handles the DragDrop event of the lblMediaFile control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="System.Windows.Forms.DragEventArgs" /> instance containing the event data.</param>
        private void LblMediaFile_DragDrop(object sender, DragEventArgs e)
        {
            var s = (string[])e.Data.GetData(DataFormats.FileDrop, false);

            var ext = Path.GetExtension(s[0]).ToLower();
            if (ext == ".wmv")
            {
                this.lblMediaFile.Text = s[0];
                _ = this.RegisterNewMediaFileAsync();
            }
            else
            {
                this.AddLogEntry("Invalid media file format", LogType.Fail);
            }
        }

        /// <summary>
        ///     Handles the DragEnter event of the lblMediaFile control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="System.Windows.Forms.DragEventArgs" /> instance containing the event data.</param>
        private void LblMediaFile_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = e.Data.GetDataPresent(DataFormats.FileDrop) ? DragDropEffects.All : DragDropEffects.None;
        }

        /// <summary>
        ///     Handles the Click event of the cmdReset control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="System.EventArgs" /> instance containing the event data.</param>
        private void CmdReset_Click(object sender, EventArgs e)
        {
            // Delete the main keys
            foreach (var attribute in this.attributes)
            {
                if (attribute.Name == "WM/MediaClassPrimaryID")
                {
                    this.DeleteAttrib(this.lblMediaFile.Text, Stream, attribute.Index);
                    break;
                }

                if (attribute.Name == "WM/MediaClassSecondaryID")
                {
                    this.DeleteAttrib(this.lblMediaFile.Text, Stream, attribute.Index);
                    break;
                }
            }

            // Now re-add the keys
            this.AddAttrib(this.lblMediaFile.Text, NewStream, "WM/MediaClassPrimaryID", Convert.ToUInt16(WMT_ATTR_DATATYPE.WMT_TYPE_GUID), TypeVideo, Language);
            this.AddAttrib(this.lblMediaFile.Text, NewStream, "WM/MediaClassSecondaryID", Convert.ToUInt16(WMT_ATTR_DATATYPE.WMT_TYPE_GUID), TypeVideo, Language);
        }

        /// <summary>
        ///     Handles the double-click event of the Listbox.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="System.EventArgs" /> instance containing the event data.</param>
        private void LbResults_DoubleClick(object sender, EventArgs e)
        {
            if (this.lbResults.SelectedItem == null)
            {
                return;
            }

            var entry = (TMDbSearchResult)this.lbResults.SelectedItem;
            if (entry?.Url.Length > 0)
            {
                Process.Start(entry.Url);
            }
        }

        /// <summary>
        ///     Edits the picture.
        /// </summary>
        [HandleProcessCorruptedStateExceptions]
        private void EditPicture(PictureBox myPicture)
        {
            try
            {
                var picture = new WMPicture
                {
                    PwszMIMEType = Marshal.StringToCoTaskMemUni("image/jpeg\0"),
                    PwszDescription = Marshal.StringToCoTaskMemUni("AlbumArt\0"),
                    BPictureType = 3,
                };

                var tempFilePath = AppContext.BaseDirectory + Settings.Default.TemporaryFile;
                var data = File.ReadAllBytes(tempFilePath);
                picture.DwDataLen = data.Length;
                picture.PbData = Marshal.AllocCoTaskMem(picture.DwDataLen);
                Marshal.Copy(data, 0, picture.PbData, picture.DwDataLen);
                var pictureParam = Marshal.AllocCoTaskMem(Marshal.SizeOf(picture));
                Marshal.StructureToPtr(picture, pictureParam, false);

                WMFSDKFunctions.WMCreateEditor(out var metadataEditor);
                metadataEditor.Open(this.lblMediaFile.Text);
                var headerInfo3 = (IWMHeaderInfo3)metadataEditor;

                headerInfo3.SetPicAttribute(0, "WM/Picture", WMT_ATTR_DATATYPE.WMT_TYPE_BINARY, pictureParam, (ushort)Marshal.SizeOf(picture));
                metadataEditor.Flush();
                metadataEditor.Close();
            }
            catch (Exception e)
            {
                this.AddLogEntry(e.Message, LogType.Fail);
            }
        }

        /// <summary>
        ///     Captures the form closing event.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The instance containing the event data.</param>
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            var tempFilePath = AppContext.BaseDirectory + Settings.Default.TemporaryFile;
            try
            {
                File.Delete(tempFilePath);
            }
            catch (Exception f)
            {
                this.AddLogEntry(f.Message, LogType.Fail);
            }
        }
    }
}