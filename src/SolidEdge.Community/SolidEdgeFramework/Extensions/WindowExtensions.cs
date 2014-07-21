using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SolidEdgeFramework.Extensions
{
    /// <summary>
    /// SolidEdgeFramework.Window extension methods.
    /// </summary>
    public static class WindowExtensions
    {
        /// <summary>
        /// Returns an IntPtr representing the window handle.
        /// </summary>
        public static IntPtr GetDrawHandle(this SolidEdgeFramework.Window window)
        {
            return new IntPtr(window.DrawHwnd);
        }

        /// <summary>
        /// Returns an IntPtr representing the window handle.
        /// </summary>
        public static IntPtr GetHandle(this SolidEdgeFramework.Window window)
        {
            return new IntPtr(window.hWnd);
        }
    }
}
