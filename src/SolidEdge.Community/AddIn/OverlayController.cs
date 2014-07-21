using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SolidEdge.Community.AddIn
{
    public sealed class OverlayController : IDisposable
    {
        private List<Overlay> _overlays = new List<Overlay>();
        private bool _disposed = false;

        internal OverlayController()
        {
        }

        ~OverlayController()
        {
            Dispose(false);
        }

        #region Methods

        public void Add(SolidEdgeFramework.View view, Overlay overlay)
        {
            if (view == null) throw new ArgumentNullException("view");
            if (overlay == null) throw new ArgumentNullException("overlay");

            if (HasOverlay(view))
            {
                throw new System.Exception("Specified view already has an overlay.");
            }

            overlay.Controller = this;
            overlay.View = view;
            _overlays.Add(overlay);
        }

        public void Add(SolidEdgeFramework.Window window, Overlay overlay)
        {
            if (window == null) throw new ArgumentNullException("window");
            Add(window.View, overlay);
        }

        public void Add<TOverlay>(SolidEdgeFramework.View view) where TOverlay : Overlay
        {
            TOverlay overlay = Activator.CreateInstance<TOverlay>();
            Add(view, overlay);
        }

        public void Add<TOverlay>(SolidEdgeFramework.Window window) where TOverlay : Overlay
        {
            TOverlay overlay = Activator.CreateInstance<TOverlay>();
            Add(window, overlay);
        }

        public bool HasOverlay(SolidEdgeFramework.View view)
        {
            return _overlays.Where(x => x.View.Equals(view)).FirstOrDefault() != null;
        }

        public bool HasOverlay(SolidEdgeFramework.Window window)
        {
            if (window == null) throw new ArgumentNullException("window");
            return HasOverlay(window.View);
        }

        public void Remove(Overlay overlay)
        {
            if (overlay == null) throw new ArgumentNullException("overlay");

            if (_overlays.Contains(overlay))
            {
                _overlays.Remove(overlay);
            }

            overlay.Dispose();
        }

        public void RemoveAll()
        {
            foreach (var overlay in _overlays)
            {
                Remove(overlay);
            }
        }

        public void RemoveAll(SolidEdgeFramework.Window window)
        {
            if (window == null) throw new ArgumentNullException("window");

            RemoveAll(window.View);

        }

        public void RemoveAll(SolidEdgeFramework.View view)
        {
            if (view == null) throw new ArgumentNullException("view");

            //var viewOverlays = _overlays.Where(x => x.IsDisposed == false).Where(x => x.View.Equals(view)).ToArray();
            var overlays = _overlays.Where(x => x.View.Equals(view)).ToArray();

            foreach (var overlay in overlays)
            {
                Remove(overlay);
            }
        }

        #endregion

        #region Properties

        public IEnumerable<Overlay> Overlays { get { return _overlays.AsEnumerable(); } }

        #endregion

        #region IDisposable implementation

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        void Dispose(bool disposing)
        {
            if (_disposed) return;

            if (disposing)
            {
                // Free managed objects here. 
                foreach (var overlay in _overlays)
                {
                    try
                    {
                        Remove(overlay);
                    }
                    catch
                    {
                    }
                }

                _overlays.Clear();
            }

            // Free unmanaged objects here. 
            _disposed = true;
        }

        #endregion
    }
}
