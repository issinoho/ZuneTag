// ------------------------------------------------------------------
//  DrunkenBakery Zune Tag
//  ZuneTag.ZuneTag
// 
//  <copyright file="MainForm.cs" company="The Drunken Bakery">
//      Copyright (c) 2009-2012 The Drunken Bakery. All rights reserved.
//  </copyright>
// 
//  Author: IRS
// ------------------------------------------------------------------
namespace DrunkenBakery.ZuneTag
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Diagnostics;
    using System.Drawing;
    using System.Drawing.Drawing2D;
    using System.IO;
    using System.Net;
    using System.Reflection;
    using System.Runtime.InteropServices;
    using System.ServiceModel;
    using System.Threading;
    using System.Windows.Forms;

    using DrunkenBakery.ZuneTag.Amazon.ECS;
    using DrunkenBakery.ZuneTag.Properties;

    using JockerSoft.Media;

    using WMFSDKWrapper;

    using Image = System.Drawing.Image;
    using Timer = System.Threading.Timer;

    /// <summary>
    ///     Main application form which drives all functionality.
    /// </summary>
    public partial class MainForm : Form
    {
        #region Constants and Fields

        /// <summary>
        /// The amazon id.
        /// </summary>
        private const string AmazonID = "AKIAI344DCI3P6HJXGOA";

        /// <summary>
        /// The amazon key.
        /// </summary>
        private const string AmazonKey = "dM23mxiQzaMEy0O2I1qIacqJijV/JyqIlFRgTP+Y";

        /// <summary>
        /// The language.
        /// </summary>
        private const ushort Language = 0;

        /// <summary>
        /// The screen lines.
        /// </summary>
        private const int ScreenLines = 1000;

        /// <summary>
        /// The screen refresh.
        /// </summary>
        private const int ScreenRefresh = 1;

        /// <summary>
        /// The stream.
        /// </summary>
        private const ushort Stream = 65535;

        /// <summary>
        /// The this app.
        /// </summary>
        private const string ThisApp = "Zune Tag Editor";

        /// <summary>
        /// The this publisher.
        /// </summary>
        private const string ThisPublisher = "The Drunken Bakery";

        /// <summary>
        /// The type movie.
        /// </summary>
        private const string TypeMovie = "C9-7F-B8-A9-47-BD-F0-4B-AC-4F-65-5B-89-F7-D8-68";

        /// <summary>
        /// The type music.
        /// </summary>
        private const string TypeMusic = "E2-89-E6-E3-8C-BA-30-43-96-DF-A0-EE-EF-FA-68-76";

        /// <summary>
        /// The type tv.
        /// </summary>
        private const string TypeTV = "8A-25-7F-BA-F7-62-A9-47-B2-1F-46-51-C4-2A-00-0E";

        /// <summary>
        /// The type video.
        /// </summary>
        private const string TypeVideo = "BD-30-98-DB-B3-3A-AB-4F-8A-37-1A-99-5F-7F-F7-4B";

        /// <summary>
        /// The new stream.
        /// </summary>
        private const ushort newStream = 0;

        /// <summary>
        /// The _ screen log timer.
        /// </summary>
        private readonly Timer _ScreenLogTimer;

        /// <summary>
        /// The _ screen log timer callback.
        /// </summary>
        private readonly TimerCallback _ScreenLogTimerCallback;

        /// <summary>
        /// The _attributes.
        /// </summary>
        private readonly List<Attribute> _attributes = new List<Attribute>();

        /// <summary>
        /// The lvitems.
        /// </summary>
        private readonly List<ListViewItem> lvitems = new List<ListViewItem>();

        /// <summary>
        /// The frm about.
        /// </summary>
        private Form frmAbout;

        /// <summary>
        /// The index primary video.
        /// </summary>
        private ushort indexPrimaryVideo;

        /// <summary>
        /// The index secondary video.
        /// </summary>
        private ushort indexSecondaryVideo;

        /// <summary>
        /// The page number.
        /// </summary>
        private int pageNumber;

        /// <summary>
        /// The safe page number.
        /// </summary>
        private int safePageNumber;

        #endregion

        #region Constructors and Destructors

        /// <summary>
        ///     Initializes a new instance of the <see cref="MainForm" /> class.
        /// </summary>
        public MainForm()
        {
            this.InitializeComponent();

            // Upgrade settings from older version
            Assembly a = Assembly.GetExecutingAssembly();
            Version appVersion = a.GetName().Version;
            string appVersionString = appVersion.ToString();

            if (Settings.Default.ApplicationVersion != appVersion.ToString())
            {
                Settings.Default.Upgrade();
                Settings.Default.ApplicationVersion = appVersionString;
            }

            // Form title bar
            this.Text = ThisApp + " freshly baked at " + ThisPublisher;

            // Tooltips
            this.toolTip1.SetToolTip(this.cmdSave, "Change video type");
            this.toolTip2.SetToolTip(this.cmdModify, "Save this attribute back to the file");
            this.toolTip3.SetToolTip(this.cmdMovieSave, "Update Movie with these tags");
            this.toolTip4.SetToolTip(this.cmdMusicSave, "Update Music Video with these tags");
            this.toolTip5.SetToolTip(this.cmdTVSave, "Update TV Show with these tags");
            this.toolTip6.SetToolTip(this.cmdVideoSave, "Update Video with these tags");
            this.toolTip7.SetToolTip(this.cmdReset, "Hard reset the video type attributes");
            this.toolTip8.SetToolTip(this.cmdCopyAz, "Copy the tags back ready for saving");
            this.toolTip9.SetToolTip(this.cmdAmazonSearch, "Search Amazon for videos that match");
            this.toolTip10.SetToolTip(this.cmdBrowse, "Open a WMV media file");

            // Initialise Event Views
            this.InitEventView(this.lvStatus);

            // Initialise Media Types
            this.InitMediaTypes();

            // Logging
            this.AddLogEntry("--------------------------------------------------", LogType.Info);
            this.AddLogEntry("Welcome to the " + ThisApp + " v" + appVersionString, LogType.Info);
            this.AddLogEntry("Ready.");

            // Start Timers
            this._ScreenLogTimerCallback = this._ScreenLogTimer_Elapsed;
            this._ScreenLogTimer = new Timer(
                this._ScreenLogTimerCallback, null, Convert.ToInt32(ScreenRefresh) * 1000, Timeout.Infinite);
        }

        #endregion

        #region Delegates

        /// <summary>
        /// The flush output delegate.
        /// </summary>
        /// <param name="lv">
        /// The lv.
        /// </param>
        private delegate void FlushOutputDelegate(ListView lv);

        /// <summary>
        /// The pause output delegate.
        /// </summary>
        /// <param name="lv">
        /// The lv.
        /// </param>
        private delegate void PauseOutputDelegate(ListView lv);

        /// <summary>
        /// The resume output delegate.
        /// </summary>
        /// <param name="lv">
        /// The lv.
        /// </param>
        private delegate void ResumeOutputDelegate(ListView lv);

        #endregion

        #region Enums

        /// <summary>
        ///     Severity of logging entry
        /// </summary>
        private enum LogType
        {
            /// <summary>
            /// The success.
            /// </summary>
            Success, 

            /// <summary>
            /// The fail.
            /// </summary>
            Fail, 

            /// <summary>
            /// The info.
            /// </summary>
            Info
        }

        /// <summary>
        /// The media type.
        /// </summary>
        private enum MediaType
        {
            /// <summary>
            /// The unknown.
            /// </summary>
            Unknown, 

            /// <summary>
            /// The video.
            /// </summary>
            Video, 

            /// <summary>
            /// The movie.
            /// </summary>
            Movie, 

            /// <summary>
            /// The music.
            /// </summary>
            Music, 

            /// <summary>
            /// The tv.
            /// </summary>
            TV
        }

        #endregion

        #region Public Methods and Operators

        /// <summary>
        /// Add an attribute with the specifed language index.
        /// </summary>
        /// <param name="pwszFileName">
        /// Name of the PWSZ file. 
        /// </param>
        /// <param name="wStreamNum">
        /// The w stream num. 
        /// </param>
        /// <param name="pwszAttribName">
        /// Name of the PWSZ attrib. 
        /// </param>
        /// <param name="wAttribType">
        /// Type of the w attrib. 
        /// </param>
        /// <param name="pwszAttribValue">
        /// The PWSZ attrib value. 
        /// </param>
        /// <param name="wLangIndex">
        /// Index of the w lang. 
        /// </param>
        /// <returns>
        /// The <see cref="bool"/>.
        /// </returns>
        public bool AddAttrib(
            string pwszFileName, 
            ushort wStreamNum, 
            string pwszAttribName, 
            ushort wAttribType, 
            string pwszAttribValue, 
            ushort wLangIndex)
        {
            IWMMetadataEditor MetadataEditor = null;
            IWMHeaderInfo3 HeaderInfo3;
            byte[] pbAttribValue;
            int nAttribValueLen;
            var AttribDataType = (WMT_ATTR_DATATYPE)wAttribType;
            ushort wAttribIndex = 0;

            try
            {
                if (!this.TranslateAttrib(AttribDataType, pwszAttribValue, out pbAttribValue, out nAttribValueLen))
                {
                    return false;
                }

                WMFSDKFunctions.WMCreateEditor(out MetadataEditor);

                MetadataEditor.Open(pwszFileName);

                HeaderInfo3 = (IWMHeaderInfo3)MetadataEditor;

                HeaderInfo3.AddAttribute(
                    wStreamNum, 
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
                return false;
            }
            finally
            {
                MetadataEditor.Flush();
                MetadataEditor.Close();
            }

            return true;
        }

        /// <summary>
        /// Attribs the exists.
        /// </summary>
        /// <param name="pwszFileName">
        /// Name of the PWSZ file. 
        /// </param>
        /// <param name="wStreamNum">
        /// The w stream num. 
        /// </param>
        /// <param name="searchAttrib">
        /// The search attrib. 
        /// </param>
        /// <returns>
        /// The <see cref="bool"/>.
        /// </returns>
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

                    HeaderInfo3.GetAttributeByIndexEx(
                        wStreamNum, 
                        wAttribIndex, 
                        pwszAttribName, 
                        ref wAttribNameLen, 
                        out wAttribType, 
                        out wLangIndex, 
                        pbAttribValue, 
                        ref dwAttribValueLen);

                    pwszAttribName = new String((char)0, wAttribNameLen);
                    pbAttribValue = new byte[dwAttribValueLen];

                    HeaderInfo3.GetAttributeByIndexEx(
                        wStreamNum, 
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
                this.AddLogEntry(e.Message, LogType.Fail);
                return false;
            }

            return isFound;
        }

        /// <summary>
        /// Delete the attribute at the specified index.
        /// </summary>
        /// <param name="pwszFileName">
        /// Name of the PWSZ file. 
        /// </param>
        /// <param name="wStreamNum">
        /// The w stream num. 
        /// </param>
        /// <param name="wAttribIndex">
        /// Index of the w attrib. 
        /// </param>
        /// <returns>
        /// The <see cref="bool"/>.
        /// </returns>
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
                this.AddLogEntry(e.Message, LogType.Fail);
                return false;
            }

            return true;
        }

        /// <summary>
        /// Creates a metadata editor and opens the file.
        /// </summary>
        /// <param name="pwszInFile">
        /// The PWSZ in file. 
        /// </param>
        /// <param name="ppEditor">
        /// The pp editor. 
        /// </param>
        /// <returns>
        /// The <see cref="bool"/>.
        /// </returns>
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
        /// Generates the image dimensions.
        /// </summary>
        /// <param name="currW">
        /// The curr W. 
        /// </param>
        /// <param name="currH">
        /// The curr H. 
        /// </param>
        /// <param name="destW">
        /// The dest W. 
        /// </param>
        /// <param name="destH">
        /// The dest H. 
        /// </param>
        /// <returns>
        /// The <see cref="Size"/>.
        /// </returns>
        public Size GenerateImageDimensions(int currW, int currH, int destW, int destH)
        {
            // double to hold the final multiplier to use when scaling the image
            double multiplier = 0;

            // string for holding layout
            string layout;

            // determine if it's Portrait or Landscape
            if (currH > currW)
            {
                layout = "portrait";
            }
            else
            {
                layout = "landscape";
            }

            switch (layout.ToLower())
            {
                case "portrait":

                    // calculate multiplier on heights
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

                    // calculate multiplier on widths
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

            // return the new image dimensions
            return new Size((int)(currW * multiplier), (int)(currH * multiplier));
        }

        /// <summary>
        /// Modifies the value of the specified attribute.
        /// </summary>
        /// <param name="pwszFileName">
        /// Name of the PWSZ file. 
        /// </param>
        /// <param name="wStreamNum">
        /// The w stream num. 
        /// </param>
        /// <param name="wAttribIndex">
        /// Index of the w attrib. 
        /// </param>
        /// <param name="wAttribType">
        /// Type of the w attrib. 
        /// </param>
        /// <param name="pwszAttribValue">
        /// The PWSZ attrib value. 
        /// </param>
        /// <param name="wLangIndex">
        /// Index of the w lang. 
        /// </param>
        /// <returns>
        /// The <see cref="bool"/>.
        /// </returns>
        public bool ModifyAttrib(
            string pwszFileName, 
            ushort wStreamNum, 
            ushort wAttribIndex, 
            ushort wAttribType, 
            string pwszAttribValue, 
            ushort wLangIndex)
        {
            IWMMetadataEditor MetadataEditor = null;
            IWMHeaderInfo3 HeaderInfo3;
            byte[] pbAttribValue;
            int nAttribValueLen;
            var AttribDataType = (WMT_ATTR_DATATYPE)wAttribType;

            try
            {
                if (!this.TranslateAttrib(AttribDataType, pwszAttribValue, out pbAttribValue, out nAttribValueLen))
                {
                    return false;
                }

                WMFSDKFunctions.WMCreateEditor(out MetadataEditor);

                MetadataEditor.Open(pwszFileName);

                HeaderInfo3 = (IWMHeaderInfo3)MetadataEditor;

                HeaderInfo3.ModifyAttribute(
                    wStreamNum, wAttribIndex, AttribDataType, wLangIndex, pbAttribValue, (uint)nAttribValueLen);
            }
            catch (Exception e)
            {
                this.AddLogEntry(e.Message, LogType.Fail);
                return false;
            }
            finally
            {
                MetadataEditor.Flush();
                MetadataEditor.Close();
            }

            return true;
        }

        /// <summary>
        /// Displays the specified attribute.
        /// </summary>
        /// <param name="wIndex">
        /// Index of the w. 
        /// </param>
        /// <param name="wStream">
        /// The w stream. 
        /// </param>
        /// <param name="pwszName">
        /// Name of the PWSZ. 
        /// </param>
        /// <param name="AttribDataType">
        /// Type of the attrib data. 
        /// </param>
        /// <param name="wLangID">
        /// The w lang ID. 
        /// </param>
        /// <param name="pbValue">
        /// The pb value. 
        /// </param>
        /// <param name="dwValueLen">
        /// The dw value len. 
        /// </param>
        public void PrintAttribute(
            ushort wIndex, 
            ushort wStream, 
            string pwszName, 
            WMT_ATTR_DATATYPE AttribDataType, 
            ushort wLangID, 
            byte[] pbValue, 
            uint dwValueLen)
        {
            string pwszValue = string.Empty;

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
                case WMT_ATTR_DATATYPE.WMT_TYPE_STRING:

                    if (0 == dwValueLen)
                    {
                        pwszValue = "***** NULL *****";
                    }
                    else
                    {
                        if ((0xFE == Convert.ToInt16(pbValue[0])) && (0xFF == Convert.ToInt16(pbValue[1])))
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
                        else if ((0xFF == Convert.ToInt16(pbValue[0])) && (0xFE == Convert.ToInt16(pbValue[1])))
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
                case WMT_ATTR_DATATYPE.WMT_TYPE_BINARY:

                    pwszValue = "[" + dwValueLen.ToString() + " bytes]";
                    break;

                    // Boolean
                case WMT_ATTR_DATATYPE.WMT_TYPE_BOOL:

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
                case WMT_ATTR_DATATYPE.WMT_TYPE_DWORD:

                    uint dwValue = BitConverter.ToUInt32(pbValue, 0);
                    pwszValue = dwValue.ToString();
                    break;

                    // QWORD
                case WMT_ATTR_DATATYPE.WMT_TYPE_QWORD:

                    ulong qwValue = BitConverter.ToUInt64(pbValue, 0);
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

                default:

                    break;
            }

            // Add to attribute list
            var _attribute = new Attribute(
                wIndex, pwszName.Substring(0, pwszName.Length - 1), pwszValue, AttribDataType);

            // Add to list
            this._attributes.Add(_attribute);
        }

        /// <summary>
        /// Set the specified attribute.
        /// </summary>
        /// <param name="pwszFileName">
        /// Name of the PWSZ file. 
        /// </param>
        /// <param name="wStreamNum">
        /// The w stream num. 
        /// </param>
        /// <param name="pwszAttribName">
        /// Name of the PWSZ attrib. 
        /// </param>
        /// <param name="wAttribType">
        /// Type of the w attrib. 
        /// </param>
        /// <param name="pwszAttribValue">
        /// The PWSZ attrib value. 
        /// </param>
        /// <returns>
        /// The <see cref="bool"/>.
        /// </returns>
        public bool SetAttrib(
            string pwszFileName, ushort wStreamNum, string pwszAttribName, ushort wAttribType, string pwszAttribValue)
        {
            try
            {
                IWMMetadataEditor MetadataEditor;
                IWMHeaderInfo3 HeaderInfo3;
                byte[] pbAttribValue;
                int nAttribValueLen;
                var AttribDataType = (WMT_ATTR_DATATYPE)wAttribType;

                if (!this.TranslateAttrib(AttribDataType, pwszAttribValue, out pbAttribValue, out nAttribValueLen))
                {
                    return false;
                }

                WMFSDKFunctions.WMCreateEditor(out MetadataEditor);

                MetadataEditor.Open(pwszFileName);

                HeaderInfo3 = (IWMHeaderInfo3)MetadataEditor;

                HeaderInfo3.SetAttribute(
                    wStreamNum, pwszAttribName, AttribDataType, pbAttribValue, (ushort)nAttribValueLen);

                MetadataEditor.Flush();

                MetadataEditor.Close();
            }
            catch (Exception e)
            {
                this.AddLogEntry(e.Message, LogType.Fail);
                return false;
            }

            return true;
        }

        /// <summary>
        /// Displays all attributes for the specified stream.
        /// </summary>
        /// <param name="pwszFileName">
        /// Name of the PWSZ file. 
        /// </param>
        /// <param name="wStreamNum">
        /// The w stream num. 
        /// </param>
        /// <returns>
        /// The <see cref="bool"/>.
        /// </returns>
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

                    HeaderInfo3.GetAttributeByIndex(
                        wAttribIndex, 
                        ref wStreamNum, 
                        pwszAttribName, 
                        ref wAttribNameLen, 
                        out wAttribType, 
                        pbAttribValue, 
                        ref wAttribValueLen);

                    pbAttribValue = new byte[wAttribValueLen];
                    pwszAttribName = new String((char)0, wAttribNameLen);

                    HeaderInfo3.GetAttributeByIndex(
                        wAttribIndex, 
                        ref wStreamNum, 
                        pwszAttribName, 
                        ref wAttribNameLen, 
                        out wAttribType, 
                        pbAttribValue, 
                        ref wAttribValueLen);

                    this.PrintAttribute(
                        wAttribIndex, wStreamNum, pwszAttribName, wAttribType, 0, pbAttribValue, wAttribValueLen);
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
        /// Displays all attributes for the specified stream, with support for GetAttributeByIndexEx.
        /// </summary>
        /// <param name="pwszFileName">
        /// Name of the PWSZ file. 
        /// </param>
        /// <param name="wStreamNum">
        /// The w stream num. 
        /// </param>
        /// <returns>
        /// The <see cref="bool"/>.
        /// </returns>
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

                    HeaderInfo3.GetAttributeByIndexEx(
                        wStreamNum, 
                        wAttribIndex, 
                        pwszAttribName, 
                        ref wAttribNameLen, 
                        out wAttribType, 
                        out wLangIndex, 
                        pbAttribValue, 
                        ref dwAttribValueLen);

                    pwszAttribName = new String((char)0, wAttribNameLen);
                    pbAttribValue = new byte[dwAttribValueLen];

                    HeaderInfo3.GetAttributeByIndexEx(
                        wStreamNum, 
                        wAttribIndex, 
                        pwszAttribName, 
                        ref wAttribNameLen, 
                        out wAttribType, 
                        out wLangIndex, 
                        pbAttribValue, 
                        ref dwAttribValueLen);

                    this.PrintAttribute(
                        wAttribIndex, wStreamNum, pwszAttribName, wAttribType, 0, pbAttribValue, dwAttribValueLen);
                }

                // Close file
                MetadataEditor.Close();
            }
            catch (Exception e)
            {
                this.AddLogEntry(e.Message, LogType.Fail);
                return false;
            }

            return true;
        }

        /// <summary>
        /// Converts attributes to byte arrays.
        /// </summary>
        /// <param name="AttribDataType">
        /// Type of the attrib data. 
        /// </param>
        /// <param name="pwszValue">
        /// The PWSZ value. 
        /// </param>
        /// <param name="pbValue">
        /// The pb value. 
        /// </param>
        /// <param name="nValueLength">
        /// Length of the n value. 
        /// </param>
        /// <returns>
        /// The <see cref="bool"/>.
        /// </returns>
        public bool TranslateAttrib(
            WMT_ATTR_DATATYPE AttribDataType, string pwszValue, out byte[] pbValue, out int nValueLength)
        {
            switch (AttribDataType)
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

                    int discarded;
                    pbValue = HexEncoding.GetBytes(pwszValue, out discarded);
                    nValueLength = HexEncoding.GetByteCount(pwszValue);

                    return true;

                default:

                    pbValue = null;
                    nValueLength = 0;

                    return false;
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// Gets the genre.
        /// </summary>
        /// <param name="item">
        /// The item. 
        /// </param>
        /// <returns>
        /// The <see cref="string"/>.
        /// </returns>
        private static string GetGenre(Item item)
        {
            if (item.BrowseNodes == null)
            {
                return null;
            }

            string genre = string.Empty;
            foreach (BrowseNode node in item.BrowseNodes.BrowseNode)
            {
                if (node.Name.ToLower() == "general")
                {
                    if (node.Ancestors != null)
                    {
                        return node.Ancestors[0].Name;
                    }
                }
            }

            return "<unspecified>";
        }

        /// <summary>
        /// Adds the log entry.
        /// </summary>
        /// <param name="newEntry">
        /// The new entry. 
        /// </param>
        private void AddLogEntry(string newEntry)
        {
            this.AddLogEntry(newEntry, LogType.Success);
        }

        /// <summary>
        /// Adds the log entry.
        /// </summary>
        /// <param name="newEntry">
        /// The new entry. 
        /// </param>
        /// <param name="whichLog">
        /// The which log. 
        /// </param>
        private void AddLogEntry(string newEntry, LogType whichLog)
        {
            switch (whichLog)
            {
                case LogType.Success:
                    this.lvitems.Add(new ListViewItem(DateTime.Now.ToString(), 0));
                    break;

                case LogType.Fail:
                    this.lvitems.Add(new ListViewItem(DateTime.Now.ToString(), 1));
                    break;

                case LogType.Info:
                    this.lvitems.Add(new ListViewItem(DateTime.Now.ToString(), 2));
                    break;
            }

            int i = this.lvitems.Count - 1;
            this.lvitems[i].SubItems.Add(newEntry);
            this.slStatus.Text = newEntry;
        }

        /// <summary>
        ///     Adds the missing attributes.
        /// </summary>
        private void AddMissingAttributes()
        {
            if (!this.AttribExists(this.lblMediaFile.Text, Stream, "WM/MediaClassPrimaryID"))
            {
                this.AddAttrib(
                    this.lblMediaFile.Text, 
                    newStream, 
                    "WM/MediaClassPrimaryID", 
                    Convert.ToUInt16(WMT_ATTR_DATATYPE.WMT_TYPE_GUID), 
                    TypeVideo, 
                    Language);
            }

            if (!this.AttribExists(this.lblMediaFile.Text, Stream, "WM/MediaClassSecondaryID"))
            {
                this.AddAttrib(
                    this.lblMediaFile.Text, 
                    newStream, 
                    "WM/MediaClassSecondaryID", 
                    Convert.ToUInt16(WMT_ATTR_DATATYPE.WMT_TYPE_GUID), 
                    TypeVideo, 
                    Language);
            }

            if (!this.AttribExists(this.lblMediaFile.Text, Stream, "Title"))
            {
                this.AddAttrib(
                    this.lblMediaFile.Text, 
                    newStream, 
                    "Title", 
                    Convert.ToUInt16(WMT_ATTR_DATATYPE.WMT_TYPE_STRING), 
                    "Unknown", 
                    Language);
            }

            if (!this.AttribExists(this.lblMediaFile.Text, Stream, "WM/SubTitle"))
            {
                this.AddAttrib(
                    this.lblMediaFile.Text, 
                    newStream, 
                    "WM/SubTitle", 
                    Convert.ToUInt16(WMT_ATTR_DATATYPE.WMT_TYPE_STRING), 
                    "Unknown", 
                    Language);
            }

            if (!this.AttribExists(this.lblMediaFile.Text, Stream, "WM/SubTitleDescription"))
            {
                this.AddAttrib(
                    this.lblMediaFile.Text, 
                    newStream, 
                    "WM/SubTitleDescription", 
                    Convert.ToUInt16(WMT_ATTR_DATATYPE.WMT_TYPE_STRING), 
                    "Unknown", 
                    Language);
            }

            if (!this.AttribExists(this.lblMediaFile.Text, Stream, "Author"))
            {
                this.AddAttrib(
                    this.lblMediaFile.Text, 
                    newStream, 
                    "Author", 
                    Convert.ToUInt16(WMT_ATTR_DATATYPE.WMT_TYPE_STRING), 
                    "Unknown", 
                    Language);
            }

            if (!this.AttribExists(this.lblMediaFile.Text, Stream, "WM/Year"))
            {
                this.AddAttrib(
                    this.lblMediaFile.Text, 
                    newStream, 
                    "WM/Year", 
                    Convert.ToUInt16(WMT_ATTR_DATATYPE.WMT_TYPE_STRING), 
                    "Unknown", 
                    Language);
            }

            if (!this.AttribExists(this.lblMediaFile.Text, Stream, "WM/OriginalBroadcastDateTime"))
            {
                this.AddAttrib(
                    this.lblMediaFile.Text, 
                    newStream, 
                    "WM/OriginalBroadcastDateTime", 
                    Convert.ToUInt16(WMT_ATTR_DATATYPE.WMT_TYPE_STRING), 
                    "Unknown", 
                    Language);
            }

            if (!this.AttribExists(this.lblMediaFile.Text, Stream, "WM/ParentalRating"))
            {
                this.AddAttrib(
                    this.lblMediaFile.Text, 
                    newStream, 
                    "WM/ParentalRating", 
                    Convert.ToUInt16(WMT_ATTR_DATATYPE.WMT_TYPE_STRING), 
                    "Unknown", 
                    Language);
            }

            if (!this.AttribExists(this.lblMediaFile.Text, Stream, "WM/TVNetworkAffiliation"))
            {
                this.AddAttrib(
                    this.lblMediaFile.Text, 
                    newStream, 
                    "WM/TVNetworkAffiliation", 
                    Convert.ToUInt16(WMT_ATTR_DATATYPE.WMT_TYPE_STRING), 
                    "Unknown", 
                    Language);
            }

            if (!this.AttribExists(this.lblMediaFile.Text, Stream, "WM/Genre"))
            {
                this.AddAttrib(
                    this.lblMediaFile.Text, 
                    newStream, 
                    "WM/Genre", 
                    Convert.ToUInt16(WMT_ATTR_DATATYPE.WMT_TYPE_STRING), 
                    "Unknown", 
                    Language);
            }

            if (!this.AttribExists(this.lblMediaFile.Text, Stream, "Description"))
            {
                this.AddAttrib(
                    this.lblMediaFile.Text, 
                    newStream, 
                    "Description", 
                    Convert.ToUInt16(WMT_ATTR_DATATYPE.WMT_TYPE_STRING), 
                    "Unknown", 
                    Language);
            }

            if (!this.AttribExists(this.lblMediaFile.Text, Stream, "WM/TrackNumber"))
            {
                this.AddAttrib(
                    this.lblMediaFile.Text, 
                    newStream, 
                    "WM/TrackNumber", 
                    Convert.ToUInt16(WMT_ATTR_DATATYPE.WMT_TYPE_DWORD), 
                    "01", 
                    Language);
            }
        }

        /// <summary>
        ///     Amazons the search.
        /// </summary>
        private void AmazonSearch()
        {
            // Set up visuals pre-search
            this.AddLogEntry("Searching Amazon for " + this.txtSearchCriteria.Text + "...", LogType.Info);
            this.SetButtons(false);
            this.progressBar1.Visible = true;
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            // Do the search
            this.backgroundWorker1.RunWorkerAsync();
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
        /// Modifies the attribute.
        /// </summary>
        /// <param name="theAttribute">
        /// The attribute. 
        /// </param>
        /// <param name="newValue">
        /// The new value. 
        /// </param>
        private void EditAttribute(string theAttribute, string newValue)
        {
            foreach (Attribute _attribute in this._attributes)
            {
                if (_attribute.Name == theAttribute)
                {
                    this.ModifyAttrib(
                        this.lblMediaFile.Text, 
                        Stream, 
                        _attribute.Index, 
                        Convert.ToUInt16(_attribute.Type), 
                        newValue, 
                        Language);
                    break;
                }
            }
        }

        /// <summary>
        /// Edits the picture.
        /// </summary>
        /// <param name="myPicture">
        /// The my Picture.
        /// </param>
        private void EditPicture(PictureBox myPicture)
        {
            IWMMetadataEditor MetadataEditor;
            IWMHeaderInfo3 HeaderInfo3;

            try
            {
                var imageConverter = new ImageConverter();
                var picture = new WMPicture();

                picture.pwszMIMEType = Marshal.StringToCoTaskMemUni("image/jpeg\0");
                picture.pwszDescription = Marshal.StringToCoTaskMemUni("AlbumArt\0");
                picture.bPictureType = 3;

                var data = (byte[])imageConverter.ConvertTo(myPicture.Image, typeof(byte[]));
                picture.dwDataLen = data.Length;
                picture.pbData = Marshal.AllocCoTaskMem(picture.dwDataLen);
                Marshal.Copy(data, 0, picture.pbData, picture.dwDataLen);
                IntPtr pictureParam = Marshal.AllocCoTaskMem(Marshal.SizeOf(picture));
                Marshal.StructureToPtr(picture, pictureParam, false);

                WMFSDKFunctions.WMCreateEditor(out MetadataEditor);
                MetadataEditor.Open(this.lblMediaFile.Text);
                HeaderInfo3 = (IWMHeaderInfo3)MetadataEditor;
                HeaderInfo3.SetPicAttribute(
                    0, "WM/Picture", WMT_ATTR_DATATYPE.WMT_TYPE_BINARY, pictureParam, (ushort)Marshal.SizeOf(picture));
                MetadataEditor.Flush();
                MetadataEditor.Close();
            }
            catch (Exception e)
            {
                this.AddLogEntry(e.Message, LogType.Fail);
            }
        }

        /// <summary>
        /// Flushes the output.
        /// </summary>
        /// <param name="lv">
        /// The lv. 
        /// </param>
        private void FlushOutput(ListView lv)
        {
            if (this.InvokeRequired)
            {
                this.BeginInvoke(new FlushOutputDelegate(this.FlushOutput), new object[] { lv });
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
        /// Inits the event view.
        /// </summary>
        /// <param name="lvX">
        /// The lv X. 
        /// </param>
        private void InitEventView(ListView lvX)
        {
            lvX.Columns.Add("Time", 140, HorizontalAlignment.Left);
            lvX.Columns.Add("Event Details", 1000, HorizontalAlignment.Left);
            lvX.Items.Clear();
        }

        /// <summary>
        ///     Inits the media types.
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
        ///     Inspects the file.
        /// </summary>
        private void InspectFile()
        {
            bool isVideo = false;
            int thisIndex = 0;

            // Destroy existing attribute list
            this._attributes.Clear();
            this.cbAttributes.Items.Clear();

            // Update from file
            if (this.ShowAttributes3(this.lblMediaFile.Text, Stream))
            {
                for (int i = 0; i < this._attributes.Count; i++)
                {
                    // Create new instance
                    Attribute u = this._attributes[i];

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

                                case TypeTV:
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
                        this.LoadTVAttributes();
                        break;
                }

                // If no title has been established then look at filename
                if (this.txtSearchCriteria.Text.Length == 0)
                {
                    string filename = Path.GetFileNameWithoutExtension(this.lblMediaFile.Text);
                    filename = filename.Replace(".", " ");
                    filename = filename.Replace("_", " ");
                    filename = filename.Replace("-", " ");
                    this.txtSearchCriteria.Text = filename;
                }
            }
        }

        /// <summary>
        ///     Loads the movie attributes.
        /// </summary>
        private void LoadMovieAttributes()
        {
            foreach (Attribute _attribute in this._attributes)
            {
                string myValue = _attribute.Value.Replace("\"", string.Empty);

                switch (_attribute.Name)
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
        ///     Loads the music attributes.
        /// </summary>
        private void LoadMusicAttributes()
        {
            foreach (Attribute _attribute in this._attributes)
            {
                string myValue = _attribute.Value.Replace("\"", string.Empty);

                switch (_attribute.Name)
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
        ///     Loads the TV attributes.
        /// </summary>
        private void LoadTVAttributes()
        {
            foreach (Attribute _attribute in this._attributes)
            {
                string myValue = _attribute.Value.Replace("\"", string.Empty);

                switch (_attribute.Name)
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
        ///     Loads the video attributes.
        /// </summary>
        private void LoadVideoAttributes()
        {
            foreach (Attribute _attribute in this._attributes)
            {
                string myValue = _attribute.Value.Replace("\"", string.Empty);

                switch (_attribute.Name)
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
        /// Pauses the output.
        /// </summary>
        /// <param name="lv">
        /// The lv. 
        /// </param>
        private void PauseOutput(ListView lv)
        {
            if (this.InvokeRequired)
            {
                this.BeginInvoke(new PauseOutputDelegate(this.PauseOutput), new object[] { lv });
                return;
            }

            lv.BeginUpdate();
        }

        /// <summary>
        ///     Performs the search.
        /// </summary>
        private void PerformSearch()
        {
            if (this.txtSearchCriteria.Text.Length == 0)
            {
                return;
            }

            try
            {
                // create a WCF Amazon ECS client
                var basicBinding = new BasicHttpBinding(BasicHttpSecurityMode.Transport);
                basicBinding.MaxReceivedMessageSize = 2147483647;
                basicBinding.ReaderQuotas.MaxStringContentLength = int.MaxValue;
                var client = new AWSECommerceServicePortTypeClient(
                    basicBinding, 
                    new EndpointAddress("https://webservices.amazon.com/onca/soap?Service=AWSECommerceService"));

                // add authentication to the ECS client  
                client.ChannelFactory.Endpoint.Behaviors.Add(new AmazonSigningEndpointBehavior(AmazonID, AmazonKey));

                var request = new ItemSearchRequest();
                request.Availability = ItemSearchRequestAvailability.Available;
                request.Condition = Condition.All;
                request.ItemPage = this.pageNumber.ToString();
                request.MerchantId = "Amazon";
                request.ResponseGroup = new[] { "Medium", "Reviews", "ItemAttributes", "BrowseNodes" };
                request.SearchIndex = "DVD";
                request.Title = this.txtSearchCriteria.Text;

                var itemSearch = new ItemSearch();
                itemSearch.Request = new[] { request };
                itemSearch.AWSAccessKeyId = AmazonID;
                ItemSearchResponse response = client.ItemSearch(itemSearch);

                if (response.Items[0].Item != null)
                {
                    this.lbResults.Items.Clear();
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
                                    myReview = string.Empty;
                                }
                                else
                                {
                                    myReview = this.orNull(displayItem.CustomerReviews.Review[0].Content);
                                }
                            }
                            else
                            {
                                myReview = this.orNull(displayItem.EditorialReviews[0].Content);
                            }

                            var _entry = new AmazonEntry(
                                this.orNull(displayItem.ASIN), 
                                this.orNull(item.Title), 
                                item.Director == null ? string.Empty : this.orNull(item.Director[0]), 
                                this.orNull(item.TheatricalReleaseDate), 
                                myReview, 
                                this.orNull(displayItem.DetailPageURL), 
                                this.orNull(GetGenre(displayItem)), 
                                this.orNull(item.AudienceRating), 
                                displayItem.MediumImage == null ? string.Empty : this.orNull(displayItem.MediumImage.URL));
                            this.lbResults.Items.Add(_entry);
                        }
                    }

                    this.lblPage.Text = "Page " + this.pageNumber.ToString();
                    this.safePageNumber = this.pageNumber;
                }
                else
                {
                    this.pageNumber = this.safePageNumber;
                }

                this.AddLogEntry("Search Complete", LogType.Info);
            }
            catch (Exception ex)
            {
                this.pageNumber = this.safePageNumber;
                this.AddLogEntry(ex.Message, LogType.Fail);
            }
        }

        /// <summary>
        ///     Registers the new media file.
        /// </summary>
        private void RegisterNewMediaFile()
        {
            if (!File.Exists(this.lblMediaFile.Text))
            {
                this.AddLogEntry("Can't load - media file not found", LogType.Fail);
                return;
            }
            else
            {
                // Grab still frame, if possible
                try
                {
                    var size = new Size();
                    size.Height = this.pictureBox1.Height;
                    size.Width = this.pictureBox1.Width;
                    this.pictureBox1.Image = FrameGrabber.GetFrameFromVideo(this.lblMediaFile.Text, 0.2d, size);
                    this.SetImage(this.pictureBox1);
                }
                catch (InvalidVideoFileException ex)
                {
                    this.AddLogEntry(ex.Message, LogType.Fail);
                }
                catch (StackOverflowException)
                {
                    this.AddLogEntry("The target image size is too big", LogType.Fail);
                }

                // Make sure all supported attributes are defined
                this.AddMissingAttributes();

                // Refresh attributes from file
                this.InspectFile();

                // Logging
                this.AddLogEntry(this.lblMediaFile.Text + " successfully loaded", LogType.Success);
            }
        }

        /// <summary>
        /// Resumes the output.
        /// </summary>
        /// <param name="lv">
        /// The lv. 
        /// </param>
        private void ResumeOutput(ListView lv)
        {
            if (this.InvokeRequired)
            {
                this.BeginInvoke(new ResumeOutputDelegate(this.ResumeOutput), new object[] { lv });
                return;
            }

            lv.EndUpdate();
        }

        /// <summary>
        /// Sets the buttons.
        /// </summary>
        /// <param name="isON">
        /// if set to <c>true</c> [is ON]. 
        /// </param>
        private void SetButtons(bool isON)
        {
            this.cmdAmazonSearch.Enabled = isON;
            this.cmdNext.Enabled = isON;
            this.cmdPrev.Enabled = isON;
        }

        /// <summary>
        /// Sets the image.
        /// </summary>
        /// <param name="pb">
        /// The pb. 
        /// </param>
        private void SetImage(PictureBox pb)
        {
            try
            {
                // create a temp image
                Image img = pb.Image;

                // calculate the size of the image
                Size imgSize = this.GenerateImageDimensions(
                    img.Width, img.Height, this.pictureBox1.Width, this.pictureBox1.Height);

                // create a new Bitmap with the proper dimensions
                var finalImg = new Bitmap(img, imgSize.Width, imgSize.Height);

                // create a new Graphics object from the image
                Graphics gfx = Graphics.FromImage(img);

                // clean up the image (take care of any image loss from resizing)
                gfx.InterpolationMode = InterpolationMode.HighQualityBicubic;

                // empty the PictureBox
                pb.Image = null;

                // center the new image
                pb.SizeMode = PictureBoxSizeMode.CenterImage;

                // set the new image
                pb.Image = finalImg;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
        }

        /// <summary>
        /// Shows the cover art.
        /// </summary>
        /// <param name="coverUrl">
        /// The cover Url.
        /// </param>
        private void ShowCoverArt(string coverUrl)
        {
            if (coverUrl != string.Empty)
            {
                var client = new WebClient();
                client.Headers["User-Agent"] = "Mozilla/4.0";
                byte[] bytes = client.DownloadData(coverUrl);
                var stream = new MemoryStream(bytes);
                this.pbCover.Image = Image.FromStream(stream);
            }
        }

        /// <summary>
        /// Updates the status view by flushing the buffer.
        /// </summary>
        /// <param name="sender">
        /// The sender. 
        /// </param>
        private void _ScreenLogTimer_Elapsed(object sender)
        {
            this.CycleStatusView();
            this._ScreenLogTimer.Change(Convert.ToInt32(ScreenRefresh) * 1000, Timeout.Infinite);
        }

        /// <summary>
        /// Handles the Click event of the aboutToolStripMenuItem control.
        /// </summary>
        /// <param name="sender">
        /// The source of the event. 
        /// </param>
        /// <param name="e">
        /// The <see cref="System.EventArgs"/> instance containing the event data. 
        /// </param>
        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (this.frmAbout == null)
            {
                this.frmAbout = new AboutBox1();
            }

            this.frmAbout.ShowDialog();
        }

        /// <summary>
        /// Handles the DoWork event of the backgroundWorker1 control.
        /// </summary>
        /// <param name="sender">
        /// The source of the event. 
        /// </param>
        /// <param name="e">
        /// The <see cref="System.ComponentModel.DoWorkEventArgs"/> instance containing the event data. 
        /// </param>
        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            this.PerformSearch();
        }

        /// <summary>
        /// Handles the RunWorkerCompleted event of the backgroundWorker1 control.
        /// </summary>
        /// <param name="sender">
        /// The source of the event. 
        /// </param>
        /// <param name="e">
        /// The <see cref="System.ComponentModel.RunWorkerCompletedEventArgs"/> instance containing the event data. 
        /// </param>
        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            this.Cursor = Cursors.Default;
            this.progressBar1.Visible = false;
            this.SetButtons(true);
        }

        /// <summary>
        /// Handles the SelectedIndexChanged event of the cbAttributes control.
        /// </summary>
        /// <param name="sender">
        /// The source of the event. 
        /// </param>
        /// <param name="e">
        /// The <see cref="System.EventArgs"/> instance containing the event data. 
        /// </param>
        private void cbAttributes_SelectedIndexChanged(object sender, EventArgs e)
        {
            var _attribute = (Attribute)this.cbAttributes.SelectedItem;
            if (_attribute == null)
            {
                return;
            }

            this.txtNewValue.Text = _attribute.Value;
        }

        /// <summary>
        /// Handles the SelectedIndexChanged event of the cbMediaType control.
        /// </summary>
        /// <param name="sender">
        /// The source of the event. 
        /// </param>
        /// <param name="e">
        /// The <see cref="System.EventArgs"/> instance containing the event data. 
        /// </param>
        private void cbMediaType_SelectedIndexChanged(object sender, EventArgs e)
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
                    this.LoadTVAttributes();
                    this.gbTV.Visible = true;
                    break;
            }
        }

        /// <summary>
        /// Handles the Click event of the cmdAmazonSearch control.
        /// </summary>
        /// <param name="sender">
        /// The source of the event. 
        /// </param>
        /// <param name="e">
        /// The <see cref="System.EventArgs"/> instance containing the event data. 
        /// </param>
        private void cmdAmazonSearch_Click(object sender, EventArgs e)
        {
            this.pageNumber = 1;
            this.AmazonSearch();
        }

        /// <summary>
        /// Handles the Click event of the cmdBrowse control.
        /// </summary>
        /// <param name="sender">
        /// The source of the event. 
        /// </param>
        /// <param name="e">
        /// The <see cref="System.EventArgs"/> instance containing the event data. 
        /// </param>
        private void cmdBrowse_Click(object sender, EventArgs e)
        {
            // Get file
            var openFileDialog1 = new OpenFileDialog();
            openFileDialog1.InitialDirectory = Environment.CurrentDirectory;
            openFileDialog1.Filter = "Windows Media Video (*.wmv)|*.wmv";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.RestoreDirectory = false;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                this.lblMediaFile.Text = openFileDialog1.FileName;
                this.RegisterNewMediaFile();
            }
        }

        /// <summary>
        /// Handles the Click event of the cmdCopyAz control.
        /// </summary>
        /// <param name="sender">
        /// The source of the event. 
        /// </param>
        /// <param name="e">
        /// The <see cref="System.EventArgs"/> instance containing the event data. 
        /// </param>
        private void cmdCopyAz_Click(object sender, EventArgs e)
        {
            if (this.lbResults.SelectedItem == null)
            {
                return;
            }

            if (this.lblMediaFile.Text.Length == 0)
            {
                return;
            }

            var _entry = (AmazonEntry)this.lbResults.SelectedItem;
            if (_entry != null)
            {
                switch (this.cbMediaType.SelectedIndex)
                {
                    case 0:
                        break;

                    case 1:
                        this.AddLogEntry("Copying details to current media file");
                        this.txtVideoAuthor.Text = _entry.Director;
                        this.txtVideoDescription.Text = _entry.Description;
                        this.txtVideoGenre.Text = _entry.Genre;
                        this.txtVideoTitle.Text = _entry.Title;
                        this.txtVideoYear.Text = _entry.Year;
                        break;

                    case 2:
                        this.AddLogEntry("Copying details to current media file");
                        this.txtMovieAuthor.Text = _entry.Director;
                        this.txtMovieDate.Text = _entry.Date;
                        this.txtMovieDescription.Text = _entry.Description;
                        this.txtMovieGenre.Text = _entry.Genre;
                        this.txtMovieRating.Text = _entry.Rating;
                        this.txtMovieTitle.Text = _entry.Title;
                        this.txtMovieYear.Text = _entry.Year;
                        break;

                    case 3:
                        this.AddLogEntry("Copying details to current media file");
                        this.txtMusicAuthor.Text = _entry.Director;
                        this.txtMusicDate.Text = _entry.Date;
                        this.txtMusicDescription.Text = _entry.Description;
                        this.txtMusicGenre.Text = _entry.Genre;
                        this.txtMusicRating.Text = _entry.Rating;
                        this.txtMusicTitle.Text = _entry.Title;
                        this.txtMusicYear.Text = _entry.Year;
                        break;

                    case 4:
                        this.AddLogEntry("Copying details to current media file");
                        this.txtTVAuthor.Text = _entry.Director;
                        this.txtTVDate.Text = _entry.Date;
                        this.txtTVDescription.Text = _entry.Description;
                        this.txtTVGenre.Text = _entry.Genre;
                        this.txtTVRating.Text = _entry.Rating;
                        this.txtTVTitle.Text = _entry.Title;
                        this.txtTVYear.Text = _entry.Year;
                        break;
                }

                // Screenshot
                // pictureBox1.Image = pbCover.Image;

                // Switch back to the first tab to see the results
                this.tabControl1.SelectedIndex = 0;
            }
        }

        /// <summary>
        /// Handles the 1 event of the cmdModify_Click control.
        /// </summary>
        /// <param name="sender">
        /// The source of the event. 
        /// </param>
        /// <param name="e">
        /// The <see cref="System.EventArgs"/> instance containing the event data. 
        /// </param>
        private void cmdModify_Click_1(object sender, EventArgs e)
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

            var _attribute = (Attribute)this.cbAttributes.SelectedItem;
            if (_attribute == null)
            {
                return;
            }

            if (this.ModifyAttrib(
                this.lblMediaFile.Text, 
                Stream, 
                _attribute.Index, 
                Convert.ToUInt16(_attribute.Type), 
                this.txtNewValue.Text, 
                Language))
            {
                this.AddLogEntry(_attribute.Name + " successfully modified");
                this.InspectFile();
            }
        }

        /// <summary>
        /// Handles the Click event of the cmdMovieSave control.
        /// </summary>
        /// <param name="sender">
        /// The source of the event. 
        /// </param>
        /// <param name="e">
        /// The <see cref="System.EventArgs"/> instance containing the event data. 
        /// </param>
        private void cmdMovieSave_Click(object sender, EventArgs e)
        {
            this.EditAttribute("Title", this.txtMovieTitle.Text);
            this.EditAttribute("WM/SubTitleDescription", this.txtMovieDescription.Text);
            this.EditAttribute("Author", this.txtMovieAuthor.Text);
            this.EditAttribute("WM/Year", this.txtMovieYear.Text);
            this.EditAttribute("WM/OriginalBroadcastDateTime", this.txtMovieDate.Text);
            this.EditAttribute("WM/ParentalRating", this.txtMovieRating.Text);
            this.EditAttribute("WM/Genre", this.txtMovieGenre.Text);

            // EditPicture(pictureBox1);
            this.AddLogEntry(this.lblMediaFile.Text + " successfully modified", LogType.Success);
            this.InspectFile();
        }

        /// <summary>
        /// Handles the Click event of the cmdMusicSave control.
        /// </summary>
        /// <param name="sender">
        /// The source of the event. 
        /// </param>
        /// <param name="e">
        /// The <see cref="System.EventArgs"/> instance containing the event data. 
        /// </param>
        private void cmdMusicSave_Click(object sender, EventArgs e)
        {
            this.EditAttribute("Title", this.txtMusicTitle.Text);
            this.EditAttribute("WM/SubTitleDescription", this.txtMusicDescription.Text);
            this.EditAttribute("Author", this.txtMusicAuthor.Text);
            this.EditAttribute("WM/Year", this.txtMusicYear.Text);
            this.EditAttribute("WM/OriginalBroadcastDateTime", this.txtMusicDate.Text);
            this.EditAttribute("WM/ParentalRating", this.txtMusicRating.Text);
            this.EditAttribute("WM/Genre", this.txtMusicGenre.Text);

            // EditPicture(pictureBox1);
            this.AddLogEntry(this.lblMediaFile.Text + " successfully modified", LogType.Success);
            this.InspectFile();
        }

        /// <summary>
        /// Handles the Click event of the cmdNext control.
        /// </summary>
        /// <param name="sender">
        /// The source of the event. 
        /// </param>
        /// <param name="e">
        /// The <see cref="System.EventArgs"/> instance containing the event data. 
        /// </param>
        private void cmdNext_Click(object sender, EventArgs e)
        {
            this.pageNumber++;
            this.AmazonSearch();
        }

        /// <summary>
        /// Handles the Click event of the cmdPrev control.
        /// </summary>
        /// <param name="sender">
        /// The source of the event. 
        /// </param>
        /// <param name="e">
        /// The <see cref="System.EventArgs"/> instance containing the event data. 
        /// </param>
        private void cmdPrev_Click(object sender, EventArgs e)
        {
            this.pageNumber--;
            if (this.pageNumber <= 0)
            {
                this.pageNumber = 1;
            }
            else
            {
                this.AmazonSearch();
            }
        }

        /// <summary>
        /// Handles the Click event of the cmdReset control.
        /// </summary>
        /// <param name="sender">
        /// The source of the event. 
        /// </param>
        /// <param name="e">
        /// The <see cref="System.EventArgs"/> instance containing the event data. 
        /// </param>
        private void cmdReset_Click(object sender, EventArgs e)
        {
            // Delete the main keys
            foreach (Attribute _attribute in this._attributes)
            {
                if (_attribute.Name == "WM/MediaClassPrimaryID")
                {
                    this.DeleteAttrib(this.lblMediaFile.Text, Stream, _attribute.Index);
                    break;
                }

                if (_attribute.Name == "WM/MediaClassSecondaryID")
                {
                    this.DeleteAttrib(this.lblMediaFile.Text, Stream, _attribute.Index);
                    break;
                }
            }

            // Now re-add the keys
            this.AddAttrib(
                this.lblMediaFile.Text, 
                newStream, 
                "WM/MediaClassPrimaryID", 
                Convert.ToUInt16(WMT_ATTR_DATATYPE.WMT_TYPE_GUID), 
                TypeVideo, 
                Language);
            this.AddAttrib(
                this.lblMediaFile.Text, 
                newStream, 
                "WM/MediaClassSecondaryID", 
                Convert.ToUInt16(WMT_ATTR_DATATYPE.WMT_TYPE_GUID), 
                TypeVideo, 
                Language);
        }

        /// <summary>
        /// Handles the Click event of the cmdSave control.
        /// </summary>
        /// <param name="sender">
        /// The source of the event. 
        /// </param>
        /// <param name="e">
        /// The <see cref="System.EventArgs"/> instance containing the event data. 
        /// </param>
        private void cmdSave_Click(object sender, EventArgs e)
        {
            string newType = string.Empty;

            if (!File.Exists(this.lblMediaFile.Text))
            {
                return;
            }

            switch (this.cbMediaType.SelectedIndex)
            {
                case 1:
                    newType = "Generic Video";
                    this.EditAttribute("WM/MediaClassPrimaryID", TypeVideo);

                    // ModifyAttrib(lblMediaFile.Text, Stream, indexPrimaryVideo, Convert.ToUInt16(WMT_ATTR_DATATYPE.WMT_TYPE_GUID), TypeVideo, Language);
                    break;

                case 2:
                    newType = "Movie";
                    this.EditAttribute("WM/MediaClassPrimaryID", TypeVideo);
                    this.EditAttribute("WM/MediaClassSecondaryID", TypeMovie);

                    // ModifyAttrib(lblMediaFile.Text, Stream, indexPrimaryVideo, Convert.ToUInt16(WMT_ATTR_DATATYPE.WMT_TYPE_GUID), TypeVideo, Language);
                    // ModifyAttrib(lblMediaFile.Text, Stream, indexSecondaryVideo, Convert.ToUInt16(WMT_ATTR_DATATYPE.WMT_TYPE_GUID), TypeMovie, Language);
                    break;

                case 3:
                    newType = "Music Video";
                    this.EditAttribute("WM/MediaClassPrimaryID", TypeVideo);
                    this.EditAttribute("WM/MediaClassSecondaryID", TypeMusic);

                    // ModifyAttrib(lblMediaFile.Text, Stream, indexPrimaryVideo, Convert.ToUInt16(WMT_ATTR_DATATYPE.WMT_TYPE_GUID), TypeVideo, Language);
                    // ModifyAttrib(lblMediaFile.Text, Stream, indexSecondaryVideo, Convert.ToUInt16(WMT_ATTR_DATATYPE.WMT_TYPE_GUID), TypeMusic, Language);
                    break;

                case 4:
                    newType = "TV Show";
                    this.EditAttribute("WM/MediaClassPrimaryID", TypeVideo);
                    this.EditAttribute("WM/MediaClassSecondaryID", TypeTV);

                    // ModifyAttrib(lblMediaFile.Text, Stream, indexPrimaryVideo, Convert.ToUInt16(WMT_ATTR_DATATYPE.WMT_TYPE_GUID), TypeVideo, Language);
                    // ModifyAttrib(lblMediaFile.Text, Stream, indexSecondaryVideo, Convert.ToUInt16(WMT_ATTR_DATATYPE.WMT_TYPE_GUID), TypeTV, Language);
                    break;
            }

            // Refresh screen
            if (newType.Length > 0)
            {
                this.AddLogEntry("File modified to be a " + newType, LogType.Success);
                this.InspectFile();
            }
        }

        /// <summary>
        /// Handles the Click event of the cmdTVSave control.
        /// </summary>
        /// <param name="sender">
        /// The source of the event. 
        /// </param>
        /// <param name="e">
        /// The <see cref="System.EventArgs"/> instance containing the event data. 
        /// </param>
        private void cmdTVSave_Click(object sender, EventArgs e)
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

            // EditPicture(pictureBox1);
            this.AddLogEntry(this.lblMediaFile.Text + " successfully modified", LogType.Success);
            this.InspectFile();
        }

        /// <summary>
        /// Handles the Click event of the cmdVideoSave control.
        /// </summary>
        /// <param name="sender">
        /// The source of the event. 
        /// </param>
        /// <param name="e">
        /// The <see cref="System.EventArgs"/> instance containing the event data. 
        /// </param>
        private void cmdVideoSave_Click(object sender, EventArgs e)
        {
            this.EditAttribute("Title", this.txtVideoTitle.Text);
            this.EditAttribute("WM/SubTitleDescription", this.txtVideoDescription.Text);
            this.EditAttribute("Author", this.txtVideoAuthor.Text);
            this.EditAttribute("WM/Year", this.txtVideoYear.Text);
            this.EditAttribute("WM/Genre", this.txtVideoGenre.Text);

            // EditPicture(pictureBox1);
            this.AddLogEntry(this.lblMediaFile.Text + " successfully modified", LogType.Success);
            this.InspectFile();
        }

        /// <summary>
        /// Handles the Click event of the exitToolStripMenuItem1 control.
        /// </summary>
        /// <param name="sender">
        /// The source of the event. 
        /// </param>
        /// <param name="e">
        /// The <see cref="System.EventArgs"/> instance containing the event data. 
        /// </param>
        private void exitToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        /// <summary>
        /// Handles the DoubleClick event of the lbResults control.
        /// </summary>
        /// <param name="sender">
        /// The source of the event. 
        /// </param>
        /// <param name="e">
        /// The <see cref="System.EventArgs"/> instance containing the event data. 
        /// </param>
        private void lbResults_DoubleClick(object sender, EventArgs e)
        {
            if (this.lbResults.SelectedItem == null)
            {
                return;
            }

            var _entry = (AmazonEntry)this.lbResults.SelectedItem;
            if (_entry != null)
            {
                if (_entry.Url.Length > 0)
                {
                    Process.Start(_entry.Url);
                }
            }
        }

        /// <summary>
        /// Handles the SelectedIndexChanged event of the lbResults control.
        /// </summary>
        /// <param name="sender">
        /// The source of the event. 
        /// </param>
        /// <param name="e">
        /// The <see cref="System.EventArgs"/> instance containing the event data. 
        /// </param>
        private void lbResults_SelectedIndexChanged(object sender, EventArgs e)
        {
            var _entry = (AmazonEntry)this.lbResults.SelectedItem;
            if (_entry != null)
            {
                this.txtAzTitle.Text = _entry.Title;
                this.txtAzYear.Text = _entry.Year;
                this.txtAzDirector.Text = _entry.Director;
                this.txtAzDescription.Text = _entry.Description;
                this.ShowCoverArt(_entry.Cover);
            }
        }

        /// <summary>
        /// Handles the DragDrop event of the lblMediaFile control.
        /// </summary>
        /// <param name="sender">
        /// The source of the event. 
        /// </param>
        /// <param name="e">
        /// The <see cref="System.Windows.Forms.DragEventArgs"/> instance containing the event data. 
        /// </param>
        private void lblMediaFile_DragDrop(object sender, DragEventArgs e)
        {
            var s = (string[])e.Data.GetData(DataFormats.FileDrop, false);

            string ext = Path.GetExtension(s[0]).ToLower();
            if (ext == ".wmv")
            {
                this.lblMediaFile.Text = s[0];
                this.RegisterNewMediaFile();
            }
            else
            {
                this.AddLogEntry("Invalid media file format", LogType.Fail);
            }
        }

        /// <summary>
        /// Handles the DragEnter event of the lblMediaFile control.
        /// </summary>
        /// <param name="sender">
        /// The source of the event. 
        /// </param>
        /// <param name="e">
        /// The <see cref="System.Windows.Forms.DragEventArgs"/> instance containing the event data. 
        /// </param>
        private void lblMediaFile_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.All;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }

        /// <summary>
        /// Ors the null.
        /// </summary>
        /// <param name="inputString">
        /// The input string. 
        /// </param>
        /// <returns>
        /// The <see cref="string"/>.
        /// </returns>
        private string orNull(string inputString)
        {
            return string.IsNullOrEmpty(inputString) ? string.Empty : inputString;
        }

        /// <summary>
        /// Handles the Tick event of the timer1 control.
        /// </summary>
        /// <param name="sender">
        /// The source of the event. 
        /// </param>
        /// <param name="e">
        /// The <see cref="System.EventArgs"/> instance containing the event data. 
        /// </param>
        private void timer1_Tick(object sender, EventArgs e)
        {
            this.progressBar1.Value += 5;
            if (this.progressBar1.Value > 120)
            {
                this.progressBar1.Value = 0;
            }
        }

        /// <summary>
        /// Handles the KeyDown event of the txtSearchCriteria control.
        /// </summary>
        /// <param name="sender">
        /// The source of the event. 
        /// </param>
        /// <param name="e">
        /// The <see cref="System.Windows.Forms.KeyEventArgs"/> instance containing the event data. 
        /// </param>
        private void txtSearchCriteria_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                this.cmdAmazonSearch_Click(sender, e);
            }
        }

        /// <summary>
        /// Handles the TextChanged event of the txtSearchCriteria control.
        /// </summary>
        /// <param name="sender">
        /// The source of the event. 
        /// </param>
        /// <param name="e">
        /// The <see cref="System.EventArgs"/> instance containing the event data. 
        /// </param>
        private void txtSearchCriteria_TextChanged(object sender, EventArgs e)
        {
            this.pageNumber = 0;
        }

        #endregion
    }
}