using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace SolidEdgeFramework.Extensions
{
    public static class SheetExtensions
    {
        [DllImport("user32.dll")]
        static extern bool CloseClipboard();

        [DllImport("user32.dll", CharSet = CharSet.Unicode, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        static extern IntPtr GetClipboardData(uint format);

        [DllImport("user32.dll")]
        static extern IntPtr GetClipboardOwner();

        [DllImport("user32.dll")]
        static extern bool IsClipboardFormatAvailable(uint format);

        [DllImport("user32.dll")]
        static extern bool OpenClipboard(IntPtr hWndNewOwner);

        [DllImport("gdi32.dll")]
        static extern bool DeleteEnhMetaFile(IntPtr hemf);

        [DllImport("gdi32.dll")]
        static extern uint GetEnhMetaFileBits(IntPtr hemf, uint cbBuffer, [Out] byte[] lpbBuffer);

        const uint CF_ENHMETAFILE = 14;

        public static void SaveAsEMF(this SolidEdgeDraft.Sheet sheet, string filename)
        {
            try
            {
                // Copy the sheet as an EMF to the windows clipboard.
                sheet.CopyEMFToClipboard();

                if (OpenClipboard(IntPtr.Zero))
                {
                    if (IsClipboardFormatAvailable(CF_ENHMETAFILE))
                    {
                        // Get the handle to the EMF.
                        IntPtr hEMF = GetClipboardData(CF_ENHMETAFILE);

                        // Query the size of the EMF.
                        uint len = GetEnhMetaFileBits(hEMF, 0, null);
                        byte[] rawBytes = new byte[len];

                        // Get all of the bytes of the EMF.
                        GetEnhMetaFileBits(hEMF, len, rawBytes);

                        // Write all of the bytes to a file.
                        File.WriteAllBytes(filename, rawBytes);

                        // Delete the EMF handle.
                        DeleteEnhMetaFile(hEMF);
                    }
                    else
                    {
                        throw new System.Exception("CF_ENHMETAFILE is not available in clipboard.");
                    }
                }
                else
                {
                    throw new System.Exception("Error opening clipboard.");
                }
            }
            catch (System.Exception e)
            {
                throw e;
            }
            finally
            {
                CloseClipboard();
            }
        }
    }
}
