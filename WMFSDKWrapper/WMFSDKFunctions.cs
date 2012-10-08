// ------------------------------------------------------------------
//  DrunkenBakery Zune Tag
//  ZuneTag.ManagedWMFSDKWrapper
// 
//  <copyright file="WMFSDKFunctions.cs" company="The Drunken Bakery">
//      Copyright (c) 2009-2012 The Drunken Bakery. All rights reserved.
//  </copyright>
// 
//  Author: IRS
// ------------------------------------------------------------------
namespace WMFSDKWrapper
{
    using System;
    using System.Runtime.InteropServices;

    /// <summary>
    /// The wmfsdk functions.
    /// </summary>
    public class WMFSDKFunctions
    {
        #region Public Methods and Operators

        /// <summary>
        /// The wm create editor.
        /// </summary>
        /// <param name="ppMetadataEditor">
        /// The pp metadata editor.
        /// </param>
        /// <returns>
        /// The <see cref="uint"/>.
        /// </returns>
        [DllImport("WMVCore.dll", EntryPoint = "WMCreateEditor", SetLastError = true, CharSet = CharSet.Unicode, 
            ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        public static extern uint WMCreateEditor(
            [Out] [MarshalAs(UnmanagedType.Interface)] out IWMMetadataEditor ppMetadataEditor);

        #endregion
    }

    /// <summary>
    /// The WMMetadataEditor interface.
    /// </summary>
    [Guid("96406BD9-2B2B-11d3-B36B-00C04F6108FF")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IWMMetadataEditor
    {
        /// <summary>
        /// The open.
        /// </summary>
        /// <param name="pwszFilename">
        /// The pwsz filename.
        /// </param>
        /// <returns>
        /// The <see cref="uint"/>.
        /// </returns>
        uint Open([In] [MarshalAs(UnmanagedType.LPWStr)] string pwszFilename);

        /// <summary>
        /// The close.
        /// </summary>
        /// <returns>
        /// The <see cref="uint"/>.
        /// </returns>
        uint Close();

        /// <summary>
        /// The flush.
        /// </summary>
        /// <returns>
        /// The <see cref="uint"/>.
        /// </returns>
        uint Flush();
    }

    /// <summary>
    /// The WMHeaderInfo3 interface.
    /// </summary>
    [Guid("15CC68E3-27CC-4ecd-B222-3F5D02D80BD5")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IWMHeaderInfo3
    {
        /// <summary>
        /// The get attribute count.
        /// </summary>
        /// <param name="wStreamNum">
        /// The w stream num.
        /// </param>
        /// <param name="pcAttributes">
        /// The pc attributes.
        /// </param>
        /// <returns>
        /// The <see cref="uint"/>.
        /// </returns>
        uint GetAttributeCount([In] ushort wStreamNum, [Out] out ushort pcAttributes);

        /// <summary>
        /// The get attribute by index.
        /// </summary>
        /// <param name="wIndex">
        /// The w index.
        /// </param>
        /// <param name="pwStreamNum">
        /// The pw stream num.
        /// </param>
        /// <param name="pwszName">
        /// The pwsz name.
        /// </param>
        /// <param name="pcchNameLen">
        /// The pcch name len.
        /// </param>
        /// <param name="pType">
        /// The p type.
        /// </param>
        /// <param name="pValue">
        /// The p value.
        /// </param>
        /// <param name="pcbLength">
        /// The pcb length.
        /// </param>
        /// <returns>
        /// The <see cref="uint"/>.
        /// </returns>
        uint GetAttributeByIndex(
            [In] ushort wIndex, 
            [Out] [In] ref ushort pwStreamNum, 
            [Out] [MarshalAs(UnmanagedType.LPWStr)] string pwszName, 
            [Out] [In] ref ushort pcchNameLen, 
            [Out] out WMT_ATTR_DATATYPE pType, 
            [Out] [MarshalAs(UnmanagedType.LPArray)] byte[] pValue, 
            [Out] [In] ref ushort pcbLength);

        /// <summary>
        /// The get attribute by name.
        /// </summary>
        /// <param name="pwStreamNum">
        /// The pw stream num.
        /// </param>
        /// <param name="pszName">
        /// The psz name.
        /// </param>
        /// <param name="pType">
        /// The p type.
        /// </param>
        /// <param name="pValue">
        /// The p value.
        /// </param>
        /// <param name="pcbLength">
        /// The pcb length.
        /// </param>
        /// <returns>
        /// The <see cref="uint"/>.
        /// </returns>
        uint GetAttributeByName(
            [Out] [In] ref ushort pwStreamNum, 
            [Out] [MarshalAs(UnmanagedType.LPWStr)] string pszName, 
            [Out] out WMT_ATTR_DATATYPE pType, 
            [Out] [MarshalAs(UnmanagedType.LPArray)] byte[] pValue, 
            [Out] [In] ref ushort pcbLength);

        /// <summary>
        /// The set attribute.
        /// </summary>
        /// <param name="wStreamNum">
        /// The w stream num.
        /// </param>
        /// <param name="pszName">
        /// The psz name.
        /// </param>
        /// <param name="Type">
        /// The type.
        /// </param>
        /// <param name="pValue">
        /// The p value.
        /// </param>
        /// <param name="cbLength">
        /// The cb length.
        /// </param>
        /// <returns>
        /// The <see cref="uint"/>.
        /// </returns>
        uint SetAttribute(
            [In] ushort wStreamNum, 
            [In] [MarshalAs(UnmanagedType.LPWStr)] string pszName, 
            [In] WMT_ATTR_DATATYPE Type, 
            [In] [MarshalAs(UnmanagedType.LPArray)] byte[] pValue, 
            [In] ushort cbLength);

        /// <summary>
        /// The get marker count.
        /// </summary>
        /// <param name="pcMarkers">
        /// The pc markers.
        /// </param>
        /// <returns>
        /// The <see cref="uint"/>.
        /// </returns>
        uint GetMarkerCount([Out] out ushort pcMarkers);

        /// <summary>
        /// The get marker.
        /// </summary>
        /// <param name="wIndex">
        /// The w index.
        /// </param>
        /// <param name="pwszMarkerName">
        /// The pwsz marker name.
        /// </param>
        /// <param name="pcchMarkerNameLen">
        /// The pcch marker name len.
        /// </param>
        /// <param name="pcnsMarkerTime">
        /// The pcns marker time.
        /// </param>
        /// <returns>
        /// The <see cref="uint"/>.
        /// </returns>
        uint GetMarker(
            [In] ushort wIndex, 
            [Out] [MarshalAs(UnmanagedType.LPWStr)] string pwszMarkerName, 
            [Out] [In] ref ushort pcchMarkerNameLen, 
            [Out] out ulong pcnsMarkerTime);

        /// <summary>
        /// The add marker.
        /// </summary>
        /// <param name="pwszMarkerName">
        /// The pwsz marker name.
        /// </param>
        /// <param name="cnsMarkerTime">
        /// The cns marker time.
        /// </param>
        /// <returns>
        /// The <see cref="uint"/>.
        /// </returns>
        uint AddMarker([In] [MarshalAs(UnmanagedType.LPWStr)] string pwszMarkerName, [In] ulong cnsMarkerTime);

        /// <summary>
        /// The remove marker.
        /// </summary>
        /// <param name="wIndex">
        /// The w index.
        /// </param>
        /// <returns>
        /// The <see cref="uint"/>.
        /// </returns>
        uint RemoveMarker([In] ushort wIndex);

        /// <summary>
        /// The get script count.
        /// </summary>
        /// <param name="pcScripts">
        /// The pc scripts.
        /// </param>
        /// <returns>
        /// The <see cref="uint"/>.
        /// </returns>
        uint GetScriptCount([Out] out ushort pcScripts);

        /// <summary>
        /// The get script.
        /// </summary>
        /// <param name="wIndex">
        /// The w index.
        /// </param>
        /// <param name="pwszType">
        /// The pwsz type.
        /// </param>
        /// <param name="pcchTypeLen">
        /// The pcch type len.
        /// </param>
        /// <param name="pwszCommand">
        /// The pwsz command.
        /// </param>
        /// <param name="pcchCommandLen">
        /// The pcch command len.
        /// </param>
        /// <param name="pcnsScriptTime">
        /// The pcns script time.
        /// </param>
        /// <returns>
        /// The <see cref="uint"/>.
        /// </returns>
        uint GetScript(
            [In] ushort wIndex, 
            [Out] [MarshalAs(UnmanagedType.LPWStr)] string pwszType, 
            [Out] [In] ref ushort pcchTypeLen, 
            [Out] [MarshalAs(UnmanagedType.LPWStr)] string pwszCommand, 
            [Out] [In] ref ushort pcchCommandLen, 
            [Out] out ulong pcnsScriptTime);

        /// <summary>
        /// The add script.
        /// </summary>
        /// <param name="pwszType">
        /// The pwsz type.
        /// </param>
        /// <param name="pwszCommand">
        /// The pwsz command.
        /// </param>
        /// <param name="cnsScriptTime">
        /// The cns script time.
        /// </param>
        /// <returns>
        /// The <see cref="uint"/>.
        /// </returns>
        uint AddScript(
            [In] [MarshalAs(UnmanagedType.LPWStr)] string pwszType, 
            [In] [MarshalAs(UnmanagedType.LPWStr)] string pwszCommand, 
            [In] ulong cnsScriptTime);

        /// <summary>
        /// The remove script.
        /// </summary>
        /// <param name="wIndex">
        /// The w index.
        /// </param>
        /// <returns>
        /// The <see cref="uint"/>.
        /// </returns>
        uint RemoveScript([In] ushort wIndex);

        /// <summary>
        /// The get codec info count.
        /// </summary>
        /// <param name="pcCodecInfos">
        /// The pc codec infos.
        /// </param>
        /// <returns>
        /// The <see cref="uint"/>.
        /// </returns>
        uint GetCodecInfoCount([Out] out uint pcCodecInfos);

        /// <summary>
        /// The get codec info.
        /// </summary>
        /// <param name="wIndex">
        /// The w index.
        /// </param>
        /// <param name="pcchName">
        /// The pcch name.
        /// </param>
        /// <param name="pwszName">
        /// The pwsz name.
        /// </param>
        /// <param name="pcchDescription">
        /// The pcch description.
        /// </param>
        /// <param name="pwszDescription">
        /// The pwsz description.
        /// </param>
        /// <param name="pCodecType">
        /// The p codec type.
        /// </param>
        /// <param name="pcbCodecInfo">
        /// The pcb codec info.
        /// </param>
        /// <param name="pbCodecInfo">
        /// The pb codec info.
        /// </param>
        /// <returns>
        /// The <see cref="uint"/>.
        /// </returns>
        uint GetCodecInfo(
            [In] uint wIndex, 
            [Out] [In] ref ushort pcchName, 
            [Out] [MarshalAs(UnmanagedType.LPWStr)] string pwszName, 
            [Out] [In] ref ushort pcchDescription, 
            [Out] [MarshalAs(UnmanagedType.LPWStr)] string pwszDescription, 
            [Out] out WMT_CODEC_INFO_TYPE pCodecType, 
            [Out] [In] ref ushort pcbCodecInfo, 
            [Out] [MarshalAs(UnmanagedType.LPArray)] byte[] pbCodecInfo);

        /// <summary>
        /// The get attribute count ex.
        /// </summary>
        /// <param name="wStreamNum">
        /// The w stream num.
        /// </param>
        /// <param name="pcAttributes">
        /// The pc attributes.
        /// </param>
        /// <returns>
        /// The <see cref="uint"/>.
        /// </returns>
        uint GetAttributeCountEx([In] ushort wStreamNum, [Out] out ushort pcAttributes);

        /// <summary>
        /// The get attribute indices.
        /// </summary>
        /// <param name="wStreamNum">
        /// The w stream num.
        /// </param>
        /// <param name="pwszName">
        /// The pwsz name.
        /// </param>
        /// <param name="pwLangIndex">
        /// The pw lang index.
        /// </param>
        /// <param name="pwIndices">
        /// The pw indices.
        /// </param>
        /// <param name="pwCount">
        /// The pw count.
        /// </param>
        /// <returns>
        /// The <see cref="uint"/>.
        /// </returns>
        uint GetAttributeIndices(
            [In] ushort wStreamNum, 
            [In] [MarshalAs(UnmanagedType.LPWStr)] string pwszName, 
            [In] ref ushort pwLangIndex, 
            [Out] [MarshalAs(UnmanagedType.LPArray)] ushort[] pwIndices, 
            [Out] [In] ref ushort pwCount);

        /// <summary>
        /// The get attribute by index ex.
        /// </summary>
        /// <param name="wStreamNum">
        /// The w stream num.
        /// </param>
        /// <param name="wIndex">
        /// The w index.
        /// </param>
        /// <param name="pwszName">
        /// The pwsz name.
        /// </param>
        /// <param name="pwNameLen">
        /// The pw name len.
        /// </param>
        /// <param name="pType">
        /// The p type.
        /// </param>
        /// <param name="pwLangIndex">
        /// The pw lang index.
        /// </param>
        /// <param name="pValue">
        /// The p value.
        /// </param>
        /// <param name="pdwDataLength">
        /// The pdw data length.
        /// </param>
        /// <returns>
        /// The <see cref="uint"/>.
        /// </returns>
        uint GetAttributeByIndexEx(
            [In] ushort wStreamNum, 
            [In] ushort wIndex, 
            [Out] [MarshalAs(UnmanagedType.LPWStr)] string pwszName, 
            [Out] [In] ref ushort pwNameLen, 
            [Out] out WMT_ATTR_DATATYPE pType, 
            [Out] out ushort pwLangIndex, 
            [Out] [MarshalAs(UnmanagedType.LPArray)] byte[] pValue, 
            [Out] [In] ref uint pdwDataLength);

        /// <summary>
        /// The modify attribute.
        /// </summary>
        /// <param name="wStreamNum">
        /// The w stream num.
        /// </param>
        /// <param name="wIndex">
        /// The w index.
        /// </param>
        /// <param name="Type">
        /// The type.
        /// </param>
        /// <param name="wLangIndex">
        /// The w lang index.
        /// </param>
        /// <param name="pValue">
        /// The p value.
        /// </param>
        /// <param name="dwLength">
        /// The dw length.
        /// </param>
        /// <returns>
        /// The <see cref="uint"/>.
        /// </returns>
        uint ModifyAttribute(
            [In] ushort wStreamNum, 
            [In] ushort wIndex, 
            [In] WMT_ATTR_DATATYPE Type, 
            [In] ushort wLangIndex, 
            [In] [MarshalAs(UnmanagedType.LPArray)] byte[] pValue, 
            [In] uint dwLength);

        /// <summary>
        /// The add attribute.
        /// </summary>
        /// <param name="wStreamNum">
        /// The w stream num.
        /// </param>
        /// <param name="pszName">
        /// The psz name.
        /// </param>
        /// <param name="pwIndex">
        /// The pw index.
        /// </param>
        /// <param name="Type">
        /// The type.
        /// </param>
        /// <param name="wLangIndex">
        /// The w lang index.
        /// </param>
        /// <param name="pValue">
        /// The p value.
        /// </param>
        /// <param name="dwLength">
        /// The dw length.
        /// </param>
        /// <returns>
        /// The <see cref="uint"/>.
        /// </returns>
        uint AddAttribute(
            [In] ushort wStreamNum, 
            [In] [MarshalAs(UnmanagedType.LPWStr)] string pszName, 
            [Out] out ushort pwIndex, 
            [In] WMT_ATTR_DATATYPE Type, 
            [In] ushort wLangIndex, 
            [In] [MarshalAs(UnmanagedType.LPArray)] byte[] pValue, 
            [In] uint dwLength);

        /// <summary>
        /// The delete attribute.
        /// </summary>
        /// <param name="wStreamNum">
        /// The w stream num.
        /// </param>
        /// <param name="wIndex">
        /// The w index.
        /// </param>
        /// <returns>
        /// The <see cref="uint"/>.
        /// </returns>
        uint DeleteAttribute([In] ushort wStreamNum, [In] ushort wIndex);

        /// <summary>
        /// The add codec info.
        /// </summary>
        /// <param name="pszName">
        /// The psz name.
        /// </param>
        /// <param name="pwszDescription">
        /// The pwsz description.
        /// </param>
        /// <param name="codecType">
        /// The codec type.
        /// </param>
        /// <param name="cbCodecInfo">
        /// The cb codec info.
        /// </param>
        /// <param name="pbCodecInfo">
        /// The pb codec info.
        /// </param>
        /// <returns>
        /// The <see cref="uint"/>.
        /// </returns>
        uint AddCodecInfo(
            [In] [MarshalAs(UnmanagedType.LPWStr)] string pszName, 
            [In] [MarshalAs(UnmanagedType.LPWStr)] string pwszDescription, 
            [In] WMT_CODEC_INFO_TYPE codecType, 
            [In] ushort cbCodecInfo, 
            [In] [MarshalAs(UnmanagedType.LPArray)] byte[] pbCodecInfo);

        /// <summary>
        /// The set pic attribute.
        /// </summary>
        /// <param name="wStreamNum">
        /// The w stream num.
        /// </param>
        /// <param name="pszName">
        /// The psz name.
        /// </param>
        /// <param name="Type">
        /// The type.
        /// </param>
        /// <param name="pValue">
        /// The p value.
        /// </param>
        /// <param name="cbLength">
        /// The cb length.
        /// </param>
        /// <returns>
        /// The <see cref="uint"/>.
        /// </returns>
        uint SetPicAttribute(
            [In] ushort wStreamNum, 
            [In] [MarshalAs(UnmanagedType.LPWStr)] string pszName, 
            [In] WMT_ATTR_DATATYPE Type, 
            [In] IntPtr pValue, 
            [In] ushort cbLength);
    }

    /// <summary>
    /// The wm t_ att r_ datatype.
    /// </summary>
    public enum WMT_ATTR_DATATYPE
    {
        /// <summary>
        /// The wm t_ typ e_ dword.
        /// </summary>
        WMT_TYPE_DWORD = 0, 

        /// <summary>
        /// The wm t_ typ e_ string.
        /// </summary>
        WMT_TYPE_STRING = 1, 

        /// <summary>
        /// The wm t_ typ e_ binary.
        /// </summary>
        WMT_TYPE_BINARY = 2, 

        /// <summary>
        /// The wm t_ typ e_ bool.
        /// </summary>
        WMT_TYPE_BOOL = 3, 

        /// <summary>
        /// The wm t_ typ e_ qword.
        /// </summary>
        WMT_TYPE_QWORD = 4, 

        /// <summary>
        /// The wm t_ typ e_ word.
        /// </summary>
        WMT_TYPE_WORD = 5, 

        /// <summary>
        /// The wm t_ typ e_ guid.
        /// </summary>
        WMT_TYPE_GUID = 6, 
    }

    /// <summary>
    /// The wm t_ code c_ inf o_ type.
    /// </summary>
    public enum WMT_CODEC_INFO_TYPE
    {
        /// <summary>
        /// The wm t_ codecinf o_ audio.
        /// </summary>
        WMT_CODECINFO_AUDIO = 0, 

        /// <summary>
        /// The wm t_ codecinf o_ video.
        /// </summary>
        WMT_CODECINFO_VIDEO = 1, 

        /// <summary>
        /// The wm t_ codecinf o_ unknown.
        /// </summary>
        WMT_CODECINFO_UNKNOWN = 0xffffff
    }

    /// <summary>
    /// The wm picture.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct WMPicture
    {
        /// <summary>
        /// The pwsz mime type.
        /// </summary>
        public IntPtr pwszMIMEType;

        /// <summary>
        /// The b picture type.
        /// </summary>
        public byte bPictureType;

        /// <summary>
        /// The pwsz description.
        /// </summary>
        public IntPtr pwszDescription;

        /// <summary>
        /// The dw data len.
        /// </summary>
        [MarshalAs(UnmanagedType.U4)]
        public int dwDataLen;

        /// <summary>
        /// The pb data.
        /// </summary>
        public IntPtr pbData;
    }
}