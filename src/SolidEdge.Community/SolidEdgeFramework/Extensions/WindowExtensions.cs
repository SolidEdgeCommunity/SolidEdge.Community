using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SolidEdgeFramework.Extensions
{
    public static class WindowExtensions
    {
        public static IntPtr GetDrawHandle(this SolidEdgeFramework.Window window)
        {
            return new IntPtr(window.DrawHwnd);
        }

        public static IntPtr GetHandle(this SolidEdgeFramework.Window window)
        {
            return new IntPtr(window.hWnd);
        }
    }
}
