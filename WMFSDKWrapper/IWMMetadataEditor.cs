// *****************************************************************************
//
// Microsoft Windows Media
// Copyright ( C) Microsoft Corporation. All rights reserved.
//
// FileName:            IWMMetadataEditor.cs
//
// Abstract:            Wrapper used by managed-code samples.
//
// *****************************************************************************

namespace WMFSDKWrapper
{
    using System.Runtime.InteropServices;

    [Guid("96406BD9-2B2B-11d3-B36B-00C04F6108FF")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IWMMetadataEditor
    {
        uint Open([In, MarshalAs(UnmanagedType.LPWStr)] string pwszFilename);

        uint Close();

        uint Flush();
    }
}
