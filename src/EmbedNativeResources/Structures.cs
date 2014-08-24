using System;
using System.Runtime.InteropServices;

namespace EmbedNativeResources
{
    [StructLayout(LayoutKind.Sequential, Pack = 2)]
    struct BITMAPFILEHEADER
    {
        public UInt16 bfType;
        public UInt32 bfSize;
        public UInt16 bfReserved1;
        public UInt16 bfReserved2;
        public UInt32 bfOffBits;
    }

    [Serializable]
    public struct NativeResource
    {
        public int Id;
        public string Path;
    }
}
