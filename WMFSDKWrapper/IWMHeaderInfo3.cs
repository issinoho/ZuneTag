// *****************************************************************************
//
// Microsoft Windows Media
// Copyright ( C) Microsoft Corporation. All rights reserved.
//
// FileName:            IWMHeaderInfo3.cs
//
// Abstract:            Wrapper used by managed-code samples.
//
// *****************************************************************************

namespace WMFSDKWrapper
{
    using System;
    using System.Runtime.InteropServices;

    [Guid("15CC68E3-27CC-4ecd-B222-3F5D02D80BD5")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IWMHeaderInfo3
    {
        uint GetAttributeCount(
            [In] ushort wStreamNum,
            [Out] out ushort pcAttributes);

        uint GetAttributeByIndex(
            [In] ushort wIndex,
            [Out, In] ref ushort pwStreamNum,
            [Out, MarshalAs(UnmanagedType.LPWStr)] string pwszName,
            [Out, In] ref ushort pcchNameLen,
            [Out] out WMT_ATTR_DATATYPE pType,
            [Out, MarshalAs(UnmanagedType.LPArray)] byte[] pValue,
            [Out, In] ref ushort pcbLength);

        uint GetAttributeByName(
            [Out, In] ref ushort pwStreamNum,
            [Out, MarshalAs(UnmanagedType.LPWStr)] string pszName,
            [Out] out WMT_ATTR_DATATYPE pType,
            [Out, MarshalAs(UnmanagedType.LPArray)] byte[] pValue,
            [Out, In] ref ushort pcbLength);

        uint SetAttribute(
            [In] ushort wStreamNum,
            [In, MarshalAs(UnmanagedType.LPWStr)] string pszName,
            [In] WMT_ATTR_DATATYPE type,
            [In, MarshalAs(UnmanagedType.LPArray)] byte[] pValue,
            [In] ushort cbLength);

        uint GetMarkerCount(
            [Out] out ushort pcMarkers);

        uint GetMarker(
            [In] ushort wIndex,
            [Out, MarshalAs(UnmanagedType.LPWStr)] string pwszMarkerName,
            [Out, In] ref ushort pcchMarkerNameLen,
            [Out] out ulong pcnsMarkerTime);

        uint AddMarker(
            [In, MarshalAs(UnmanagedType.LPWStr)] string pwszMarkerName,
            [In] ulong cnsMarkerTime);

        uint RemoveMarker(
            [In] ushort wIndex);

        uint GetScriptCount(
            [Out] out ushort pcScripts);

        uint GetScript(
            [In] ushort wIndex,
            [Out, MarshalAs(UnmanagedType.LPWStr)] string pwszType,
            [Out, In] ref ushort pcchTypeLen,
            [Out, MarshalAs(UnmanagedType.LPWStr)] string pwszCommand,
            [Out, In] ref ushort pcchCommandLen,
            [Out] out ulong pcnsScriptTime);

        uint AddScript(
            [In, MarshalAs(UnmanagedType.LPWStr)] string pwszType,
            [In, MarshalAs(UnmanagedType.LPWStr)] string pwszCommand,
            [In] ulong cnsScriptTime);

        uint RemoveScript(
            [In] ushort wIndex);

        uint GetCodecInfoCount(
            [Out] out uint pcCodecInfos);

        uint GetCodecInfo(
            [In] uint wIndex,
            [Out, In] ref ushort pcchName,
            [Out, MarshalAs(UnmanagedType.LPWStr)] string pwszName,
            [Out, In] ref ushort pcchDescription,
            [Out, MarshalAs(UnmanagedType.LPWStr)] string pwszDescription,
            [Out] out WMT_CODEC_INFO_TYPE pCodecType,
            [Out, In] ref ushort pcbCodecInfo,
            [Out, MarshalAs(UnmanagedType.LPArray)] byte[] pbCodecInfo);

        uint GetAttributeCountEx(
            [In] ushort wStreamNum,
            [Out] out ushort pcAttributes);

        uint GetAttributeIndices(
            [In] ushort wStreamNum,
            [In, MarshalAs(UnmanagedType.LPWStr)] string pwszName,
            [In] ref ushort pwLangIndex,
            [Out, MarshalAs(UnmanagedType.LPArray)] ushort[] pwIndices,
            [Out, In] ref ushort pwCount);

        uint GetAttributeByIndexEx(
            [In] ushort wStreamNum,
            [In] ushort wIndex,
            [Out, MarshalAs(UnmanagedType.LPWStr)] string pwszName,
            [Out, In] ref ushort pwNameLen,
            [Out] out WMT_ATTR_DATATYPE pType,
            [Out] out ushort pwLangIndex,
            [Out, MarshalAs(UnmanagedType.LPArray)] byte[] pValue,
            [Out, In] ref uint pdwDataLength);

        uint ModifyAttribute(
            [In] ushort wStreamNum,
            [In] ushort wIndex,
            [In] WMT_ATTR_DATATYPE type,
            [In] ushort wLangIndex,
            [In, MarshalAs(UnmanagedType.LPArray)] byte[] pValue,
            [In] uint dwLength);

        uint AddAttribute(
            [In] ushort wStreamNum,
            [In, MarshalAs(UnmanagedType.LPWStr)] string pszName,
            [Out] out ushort pwIndex,
            [In] WMT_ATTR_DATATYPE type,
            [In] ushort wLangIndex,
            [In, MarshalAs(UnmanagedType.LPArray)] byte[] pValue,
            [In] uint dwLength);

        uint DeleteAttribute(
            [In] ushort wStreamNum,
            [In] ushort wIndex);

        uint AddCodecInfo(
            [In, MarshalAs(UnmanagedType.LPWStr)] string pszName,
            [In, MarshalAs(UnmanagedType.LPWStr)] string pwszDescription,
            [In] WMT_CODEC_INFO_TYPE codecType,
            [In] ushort cbCodecInfo,
            [In, MarshalAs(UnmanagedType.LPArray)] byte[] pbCodecInfo);

        uint SetPicAttribute(
        [In] ushort wStreamNum,
        [In, MarshalAs(UnmanagedType.LPWStr)] string pszName,
        [In] WMT_ATTR_DATATYPE type,
        [In] IntPtr pValue,
        [In] ushort cbLength);
    }
}
