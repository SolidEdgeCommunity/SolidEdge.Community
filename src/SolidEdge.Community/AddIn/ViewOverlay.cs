using SolidEdgeCommunity;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Threading;

namespace SolidEdgeCommunity.AddIn
{
    public abstract class ViewOverlay :
        SolidEdgeFramework.ISEViewEvents,
        SolidEdgeFramework.ISEIGLDisplayEvents,
        SolidEdgeFramework.ISEhDCDisplayEvents,
        IDisposable
    {
        SolidEdgeFramework.View _view;
        ViewOverlayController _controller;
        private Dictionary<IConnectionPoint, int> _connectionPointDictionary = new Dictionary<IConnectionPoint, int>();
        protected bool _disposed = false;

        public ViewOverlay()
        {
        }

        ~ViewOverlay()
        {
            Dispose(false);
        }

        #region SolidEdgeFramework.ISEViewEvents implementation

        void SolidEdgeFramework.ISEViewEvents.Changed()
        {
            Changed();
        }

        void SolidEdgeFramework.ISEViewEvents.Destroyed()
        {
            Destroyed();

            _controller.Remove(this);
        }

        void SolidEdgeFramework.ISEViewEvents.StyleChanged()
        {
            StyleChanged();
        }

        #endregion

        #region SolidEdgeFramework.ISEIGLDisplayEvents implementation

        void SolidEdgeFramework.ISEIGLDisplayEvents.BeginDisplay()
        {
            if (_disposed) return;

            BeginOpenGLDisplay();
        }

        void SolidEdgeFramework.ISEIGLDisplayEvents.BeginIGLMainDisplay(object pUnknownIGL)
        {
            if (_disposed) return;

            BeginOpenGLMainDisplay(pUnknownIGL as SolidEdgeSDK.IGL);
        }

        void SolidEdgeFramework.ISEIGLDisplayEvents.EndDisplay()
        {
            if (_disposed) return;

            EndOpenGLDisplay();
        }

        void SolidEdgeFramework.ISEIGLDisplayEvents.EndIGLMainDisplay(object pUnknownIGL)
        {
            if (_disposed) return;

            EndOpenGLMainDisplay(pUnknownIGL as SolidEdgeSDK.IGL);
        }

        #endregion

        #region SolidEdgeFramework.ISEhDCDisplayEvents implementation

        void SolidEdgeFramework.ISEhDCDisplayEvents.BeginDisplay()
        {
            BeginDeviceContextDisplay();
        }

        void SolidEdgeFramework.ISEhDCDisplayEvents.BeginhDCMainDisplay(int hDC, ref double ModelToDC, ref int Rect)
        {
            BeginDeviceContextMainDisplay(new IntPtr(hDC), ref ModelToDC, ref Rect);
        }

        void SolidEdgeFramework.ISEhDCDisplayEvents.EndDisplay()
        {
            EndDeviceContextDisplay();
        }

        void SolidEdgeFramework.ISEhDCDisplayEvents.EndhDCMainDisplay(int hDC, ref double ModelToDC, ref int Rect)
        {
            EndDeviceContextMainDisplay(new IntPtr(hDC), ref ModelToDC, ref Rect);
        }

        #endregion

        #region SolidEdgeFramework.ISEViewEvents virtual members

        /// <summary>
        /// Raised by the SolidEdgeFramework.ISEViewEvents.Changed event.
        /// </summary>
        public virtual void Changed()
        {
        }

        /// <summary>
        /// Raised by the SolidEdgeFramework.ISEViewEvents.Destroyed event.
        /// </summary>
        public virtual void Destroyed()
        {
        }

        /// <summary>
        /// Raised by the SolidEdgeFramework.ISEViewEvents.StyleChanged event.
        /// </summary>
        public virtual void StyleChanged()
        {
        }

        #endregion

        #region SolidEdgeFramework.ISEIGLDisplayEvents virtual members

        /// <summary>
        /// Raised by the SolidEdgeFramework.ISEIGLDisplayEvents.BeginDisplay event.
        /// </summary>
        public virtual void BeginOpenGLDisplay()
        {
        }

        /// <summary>
        /// Raised by the SolidEdgeFramework.ISEIGLDisplayEvents.BeginIGLMainDisplay event.
        /// </summary>
        public virtual void BeginOpenGLMainDisplay(SolidEdgeSDK.IGL gl)
        {
        }

        /// <summary>
        /// Raised by the SolidEdgeFramework.ISEIGLDisplayEvents.EndDisplay event.
        /// </summary>
        public virtual void EndOpenGLDisplay()
        {
        }

        /// <summary>
        /// Raised by the SolidEdgeFramework.ISEIGLDisplayEvents.EndIGLMainDisplay event.
        /// </summary>
        public virtual void EndOpenGLMainDisplay(object pUnknownIGL)
        {
        }

        #endregion

        #region SolidEdgeFramework.ISEhDCDisplayEvents virtual members

        /// <summary>
        /// Raised by the SolidEdgeFramework.ISEhDCDisplayEvents.BeginDisplay event.
        /// </summary>
        public virtual void BeginDeviceContextDisplay()
        {
        }

        /// <summary>
        /// Raised by the SolidEdgeFramework.ISEhDCDisplayEvents.BeginhDCMainDisplay event.
        /// </summary>
        public virtual void BeginDeviceContextMainDisplay(IntPtr hDC, ref double modelToDC, ref int rect)
        {
        }

        /// <summary>
        /// Raised by the SolidEdgeFramework.ISEhDCDisplayEvents.EndDisplay event.
        /// </summary>
        public virtual void EndDeviceContextDisplay()
        {
        }

        /// <summary>
        /// Raised by the SolidEdgeFramework.ISEhDCDisplayEvents.EndhDCMainDisplay event.
        /// </summary>
        public virtual void EndDeviceContextMainDisplay(IntPtr hDC, ref double modelToDC, ref int rect)
        {
        }

        #endregion

        #region Properties

        public bool IsDisposed { get { return _disposed; } }

        public ViewOverlayController Controller
        {
            get { return _controller; }
            internal set { _controller = value; }
        }

        /// <summary>
        /// The SolidEdgeFramework.View assigned to this overlay.
        /// </summary>
        public SolidEdgeFramework.View View
        {
            get { return _view; }
            internal set
            {
                if (value != null)
                {
                    AdviseSink<SolidEdgeFramework.ISEViewEvents>(value.ViewEvents);
                    AdviseSink<SolidEdgeFramework.ISEIGLDisplayEvents>(value.GLDisplayEvents);
                    AdviseSink<SolidEdgeFramework.ISEhDCDisplayEvents>(value.DisplayEvents);                    
                }
                else
                {
                    UnadviseAllSinks();
                }

                _view = value;
            }
        }
        public SolidEdgeFramework.Window Window { get { return _view.Window; } }

        #endregion

        #region IDisposable implementation

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
                    // Unhook all events.
                    UnadviseAllSinks();
                }

                _view = null;
                _controller = null;
                _disposed = true;
            }
        }

        #endregion

        #region IConnectionPoint implementation

        /// <summary>
        /// Establishes a connection between a connection point object and the client's sink.
        /// </summary>
        /// <typeparam name="TInterface">Interface type of the outgoing interface whose connection point object is being requested.</typeparam>
        /// <param name="container">An object that implements the IConnectionPointContainer inferface.</param>
        private void AdviseSink<TInterface>(object container) where TInterface : class
        {
            bool lockTaken = false;

            // Make sure instance implements specified TInterface.
            if (this is TInterface)
            {
                try
                {
                    System.Threading.Monitor.Enter(this, ref lockTaken);

                    // Prevent multiple event Advise() calls on same sink.
                    if (IsSinkAdvised<TInterface>(container))
                    {
                        return;
                    }

                    IConnectionPointContainer cpc = null;
                    IConnectionPoint cp = null;
                    int cookie = 0;

                    cpc = (IConnectionPointContainer)container;
                    cpc.FindConnectionPoint(typeof(TInterface).GUID, out cp);

                    if (cp != null)
                    {
                        cp.Advise(this, out cookie);
                        _connectionPointDictionary.Add(cp, cookie);
                    }
                }
                finally
                {
                    if (lockTaken)
                    {
                        Monitor.Exit(this);
                    }
                }
            }
            else
            {
                throw new NotImplementedException(String.Format("Instance does not implement {0}.", typeof(TInterface)));
            }
        }

        /// <summary>
        /// Determines if a connection between a connection point object and the client's sink is established.
        /// </summary>
        /// <param name="container">An object that implements the IConnectionPointContainer inferface.</param>
        private bool IsSinkAdvised<TInterface>(object container) where TInterface : class
        {
            bool lockTaken = false;

            // Make sure instance implements specified TInterface.
            if (this is TInterface)
            {
                try
                {
                    Monitor.Enter(this, ref lockTaken);

                    IConnectionPointContainer cpc = null;
                    IConnectionPoint cp = null;
                    int cookie = 0;

                    cpc = (IConnectionPointContainer)container;
                    cpc.FindConnectionPoint(typeof(TInterface).GUID, out cp);

                    if (cp != null)
                    {
                        if (_connectionPointDictionary.ContainsKey(cp))
                        {
                            cookie = _connectionPointDictionary[cp];
                            return true;
                        }
                    }
                }
                finally
                {
                    if (lockTaken)
                    {
                        Monitor.Exit(this);
                    }
                }
            }

            return false;
        }

        /// <summary>
        /// Terminates an advisory connection previously established between a connection point object and a client's sink.
        /// </summary>
        /// <typeparam name="TInterface">Interface type of the interface whose connection point object is being requested to be removed.</typeparam>
        /// <param name="container">An object that implements the IConnectionPointContainer inferface.</param>
        private void UnadviseSink<TInterface>(object container) where TInterface : class
        {
            bool lockTaken = false;

            // Make sure instance implements specified TInterface.
            if (this is TInterface)
            {
                try
                {
                    Monitor.Enter(this, ref lockTaken);

                    IConnectionPointContainer cpc = null;
                    IConnectionPoint cp = null;
                    int cookie = 0;

                    cpc = (IConnectionPointContainer)container;
                    cpc.FindConnectionPoint(typeof(TInterface).GUID, out cp);

                    if (cp != null)
                    {
                        if (_connectionPointDictionary.ContainsKey(cp))
                        {
                            cookie = _connectionPointDictionary[cp];
                            cp.Unadvise(cookie);
                            _connectionPointDictionary.Remove(cp);
                        }
                    }
                }
                finally
                {
                    if (lockTaken)
                    {
                        Monitor.Exit(this);
                    }
                }
            }
            else
            {
                // We could throw an exception here but does it really matter?
            }
        }

        /// <summary>
        /// Terminates all advisory connections previously established.
        /// </summary>
        private void UnadviseAllSinks()
        {
            bool lockTaken = false;

            try
            {
                Monitor.Enter(this, ref lockTaken);
                Dictionary<IConnectionPoint, int>.Enumerator enumerator = _connectionPointDictionary.GetEnumerator();
                while (enumerator.MoveNext())
                {
                    enumerator.Current.Key.Unadvise(enumerator.Current.Value);
                }
            }
            finally
            {
                _connectionPointDictionary.Clear();

                if (lockTaken)
                {
                    Monitor.Exit(this);
                }
            }
        }

        /// <summary>
        /// Establishes or terminates a connection between a connection point object and the client's sink.
        /// </summary>
        /// <typeparam name="TInterface">Interface type of the interface whose connection point object is being requested to be updated.</typeparam>
        /// <param name="container">An object that implements the IConnectionPointContainer inferface.</param>
        /// <param name="advise">Flag indicating whether to advise or unadvise.</param>
        private void UpdateSink<TInterface>(object container, bool advise) where TInterface : class
        {
            if (advise)
            {
                AdviseSink<TInterface>(container);
            }
            else
            {
                UnadviseSink<TInterface>(container);
            }
        }

        #endregion
    }
}
