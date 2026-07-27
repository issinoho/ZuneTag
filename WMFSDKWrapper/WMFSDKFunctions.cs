// *****************************************************************************
//
// Microsoft Windows Media
// Copyright ( C) Microsoft Corporation. All rights reserved.
//
// FileName:            WMFSDKFunction.cs
//
// Abstract:            Wrapper used by managed-code samples.
//
// *****************************************************************************

namespace WMFSDKWrapper
{
    using System.Runtime.InteropServices;

    public class WMFSDKFunctions
    {
        public WMFSDKFunctions()
        {
            // TODO: Add constructor logic here
        }

        [DllImport(
             "WMVCore.dll",
             EntryPoint = "WMCreateEditor",
             SetLastError = true,
             CharSet = CharSet.Unicode,
             ExactSpelling = true,
             CallingConvention = CallingConvention.StdCall)]
        public static extern uint WMCreateEditor(
            [Out, MarshalAs(UnmanagedType.Interface)] out IWMMetadataEditor ppMetadataEditor);
    }
}
