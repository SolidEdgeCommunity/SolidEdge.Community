using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace SolidEdgeCommunity.AddIn
{
    public class EdgeBarPage : NativeWindow, IDisposable
    {
        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr SetParent(IntPtr hWndChild, IntPtr hWndNewParent);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        private bool _disposed = false;
        private SolidEdgeFramework.SolidEdgeDocument _document;
        private EdgeBarControl _control;
        private bool _visible = true;

        internal EdgeBarPage(int hWnd, SolidEdgeFramework.SolidEdgeDocument document, EdgeBarControl control)
            : this(new IntPtr(hWnd), document, control)
        {
        }

        internal EdgeBarPage(IntPtr hWnd, SolidEdgeFramework.SolidEdgeDocument document, EdgeBarControl control)
        {
            _document = document;
            _control = control;
            _control.Initialize(this);

            this.AssignHandle(hWnd);

            // Reparent child control to this hWnd.
            SetParent(_control.Handle, this.Handle);

            // Show the child control and maximize it to fill the entire EdgeBarPage reigon.
            ShowWindow(_control.Handle, 3 /* SHOWMAXIMIZED */);
        }

        #region Properties

        public bool IsDisposed { get { return _disposed; } }

        public SolidEdgeFramework.SolidEdgeDocument Document
        {
            get { return _document; }
        }

        public EdgeBarControl Control
        {
            get { return _control; }
        }

        public virtual bool Visible
        {
            get { return _visible; }
            set { _visible = value; }
        }

        #endregion

        #region IDisposable implementation

        ~EdgeBarPage()
        {
            Dispose(false);
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    if ((_control != null) && (!_control.IsDisposed))
                    {
                        try
                        {
                            _control.Dispose();
                        }
                        catch
                        {
                        }
                    }
                }

                try
                {
                    ReleaseHandle();
                }
                catch
                {
                }

                _control = null;
                _document = null;
                _disposed = true;
            }
        }

        #endregion
    }
}
