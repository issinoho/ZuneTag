// *****************************************************************************
//
// Microsoft Windows Media
// Copyright ( C) Microsoft Corporation. All rights reserved.
//
// FileName:            WMPicture.cs
//
// Abstract:            Wrapper used by managed-code samples.
//
// *****************************************************************************

namespace WMFSDKWrapper
{
    using System;
    using System.Runtime.InteropServices;

    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct WMPicture
    {
        public IntPtr PwszMIMEType;
        public byte BPictureType;
        public IntPtr PwszDescription;
        [MarshalAs(UnmanagedType.U4)]
        public int DwDataLen;
        public IntPtr PbData;
    }
}
