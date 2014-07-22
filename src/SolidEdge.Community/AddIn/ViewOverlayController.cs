using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SolidEdgeCommunity.AddIn
{
    /// <summary>
    /// View overlay controller class.
    /// </summary>
    public sealed class ViewOverlayController : IDisposable
    {
        private List<ViewOverlay> _overlays = new List<ViewOverlay>();
        private bool _disposed = false;

        internal ViewOverlayController()
        {
        }

        /// <summary>
        /// Destructor.
        /// </summary>
        ~ViewOverlayController()
        {
            Dispose(false);
        }

        #region Methods

        /// <summary>
        /// Adds an overlay to the specified view.
        /// </summary>
        public void Add(SolidEdgeFramework.View view, ViewOverlay overlay)
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

        /// <summary>
        /// Adds an overlay to the specified window.
        /// </summary>
        public void Add(SolidEdgeFramework.Window window, ViewOverlay overlay)
        {
            if (window == null) throw new ArgumentNullException("window");
            Add(window.View, overlay);
        }

        /// <summary>
        /// Adds an overlay to the specified view.
        /// </summary>
        public TOverlay Add<TOverlay>(SolidEdgeFramework.View view) where TOverlay : ViewOverlay
        {
            TOverlay overlay = Activator.CreateInstance<TOverlay>();
            Add(view, overlay);
            return overlay;
        }

        /// <summary>
        /// Adds an overlay to the specified window.
        /// </summary>
        public TOverlay Add<TOverlay>(SolidEdgeFramework.Window window) where TOverlay : ViewOverlay
        {
            TOverlay overlay = Activator.CreateInstance<TOverlay>();
            Add(window, overlay);
            return overlay;
        }

        /// <summary>
        /// Gets the overlay for the specified view.
        /// </summary>
        public ViewOverlay GetOverlay(SolidEdgeFramework.View view)
        {
            if (view == null) throw new ArgumentNullException("view");
            return _overlays.Where(x => x.View.Equals(view)).FirstOrDefault();
        }

        /// <summary>
        /// Gets the overlay for the specified window.
        /// </summary>
        public ViewOverlay GetOverlay(SolidEdgeFramework.Window window)
        {
            if (window == null) throw new ArgumentNullException("window");
            return GetOverlay(window.View);
        }

        /// <summary>
        /// Determines if the specified view has an overlay.
        /// </summary>
        public bool HasOverlay(SolidEdgeFramework.View view)
        {
            return _overlays.Where(x => x.View.Equals(view)).FirstOrDefault() != null;
        }

        /// <summary>
        /// Determines if the specified window has an overlay.
        /// </summary>
        public bool HasOverlay(SolidEdgeFramework.Window window)
        {
            if (window == null) throw new ArgumentNullException("window");
            return HasOverlay(window.View);
        }

        /// <summary>
        /// Removes the specified overlay.
        /// </summary>
        public void Remove(ViewOverlay overlay)
        {
            if (overlay == null) throw new ArgumentNullException("overlay");

            if (_overlays.Contains(overlay))
            {
                _overlays.Remove(overlay);
            }

            overlay.Dispose();
        }

        /// <summary>
        /// Removes all overlays.
        /// </summary>
        public void RemoveAll()
        {
            foreach (var overlay in _overlays)
            {
                Remove(overlay);
            }
        }

        /// <summary>
        /// Removes all overlays for the specified window.
        /// </summary>
        public void RemoveAll(SolidEdgeFramework.Window window)
        {
            if (window == null) throw new ArgumentNullException("window");

            RemoveAll(window.View);

        }

        /// <summary>
        /// Removes all overlays for the specified view.
        /// </summary>
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

        /// <summary>
        /// Returns an IEnumerable of all overlays.
        /// </summary>
        public IEnumerable<ViewOverlay> Overlays { get { return _overlays.AsEnumerable(); } }

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
