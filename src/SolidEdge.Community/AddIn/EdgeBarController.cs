using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Threading;

namespace SolidEdgeCommunity.AddIn
{
    /// <summary>
    /// EdgeBar controller class.
    /// </summary>
    public sealed class EdgeBarController : IDisposable,
        SolidEdgeFramework.ISEAddInEdgeBarEvents,
        SolidEdgeFramework.ISEAddInEdgeBarEventsEx
    {
        private SolidEdgeCommunity.AddIn.SolidEdgeAddIn _addIn;
        private Dictionary<IConnectionPoint, int> _connectionPointDictionary = new Dictionary<IConnectionPoint, int>();
        private List<EdgeBarPage> _edgeBarPages = new List<EdgeBarPage>();
        private bool _disposed = false;

        internal EdgeBarController(SolidEdgeCommunity.AddIn.SolidEdgeAddIn addIn)
        {
            if (addIn == null) throw new ArgumentNullException("addIn");
            _addIn = addIn;

            AdviseSink<SolidEdgeFramework.ISEAddInEdgeBarEvents>(_addIn.AddInEx);
            AdviseSink<SolidEdgeFramework.ISEAddInEdgeBarEventsEx>(_addIn.AddInEx);
        }

        /// <summary>
        /// Destructor
        /// </summary>
        ~EdgeBarController()
        {
            Dispose(false);
        }

        #region SolidEdgeFramework.ISEAddInEdgeBarEvents implementation

        void SolidEdgeFramework.ISEAddInEdgeBarEvents.AddPage(object theDocument)
        {
            ((SolidEdgeFramework.ISEAddInEdgeBarEventsEx)this).AddPage(theDocument);
        }

        void SolidEdgeFramework.ISEAddInEdgeBarEvents.IsPageDisplayable(object theDocument, string EnvironmentCatID, out bool vbIsPageDisplayable)
        {
            ((SolidEdgeFramework.ISEAddInEdgeBarEventsEx)this).IsPageDisplayable(theDocument, EnvironmentCatID, out vbIsPageDisplayable);
        }

        void SolidEdgeFramework.ISEAddInEdgeBarEvents.RemovePage(object theDocument)
        {
            ((SolidEdgeFramework.ISEAddInEdgeBarEventsEx)this).RemovePage(theDocument);
        }

        #endregion

        #region SolidEdgeFramework.ISEAddInEdgeBarEventsEx implementation

        void SolidEdgeFramework.ISEAddInEdgeBarEventsEx.AddPage(object theDocument)
        {
            if (theDocument == null) return;

            var document = theDocument as SolidEdgeFramework.SolidEdgeDocument;

            if (document != null)
            {
                _addIn.OnCreateEdgeBarPage(this, document);
            }
        }

        void SolidEdgeFramework.ISEAddInEdgeBarEventsEx.IsPageDisplayable(object theDocument, string EnvironmentCatID, out bool vbIsPageDisplayable)
        {
            vbIsPageDisplayable = true;
        }

        void SolidEdgeFramework.ISEAddInEdgeBarEventsEx.RemovePage(object theDocument)
        {
            var edgeBarPages = _edgeBarPages.Where(x => x.Document.Equals(theDocument)).ToArray();

            foreach (var edgeBarPage in edgeBarPages)
            {
                _edgeBarPages.Remove(edgeBarPage);

                try
                {
                    int hWnd = edgeBarPage.Handle.ToInt32();

                    try
                    {
                        _addIn.EdgeBarEx.RemovePage(theDocument, hWnd, 0);
                    }
                    catch
                    {
                    }

                    edgeBarPage.Dispose();
                }
                catch
                {
                }
            }
        }

        #endregion

        #region Methods

        public void Add(SolidEdgeFramework.SolidEdgeDocument document, EdgeBarControl control, int imageId)
        {
            if (document == null) throw new ArgumentNullException("document");
            if (control == null) throw new ArgumentNullException("control");

            var options = (int)SolidEdgeConstants.EdgeBarConstant.DONOT_MAKE_ACTIVE;
            var hWnd = _addIn.EdgeBarEx2.AddPageEx(document, _addIn.NativeResourcesDllPath, imageId, control.ToolTip, options);

            if (hWnd.Equals(IntPtr.Zero) == false)
            {
                var edgeBarPage = new EdgeBarPage(hWnd, document, control);
                _edgeBarPages.Add(edgeBarPage);
            }
        }

        public void Add<TEdgeBarControl>(SolidEdgeFramework.SolidEdgeDocument document, int imageId) where TEdgeBarControl : EdgeBarControl
        {
            TEdgeBarControl control = Activator.CreateInstance<TEdgeBarControl>();
            Add(document, control, imageId);
        }

        #endregion

        #region Properties

        /// <summary>
        /// Enumerable collection of EdgeBarPage objects.
        /// </summary>
        public IEnumerable<EdgeBarPage> EdgeBarPages { get { return _edgeBarPages.AsEnumerable(); } }

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

                foreach (var edgeBarPage in _edgeBarPages)
                {
                    try
                    {
                        edgeBarPage.Dispose();
                    }
                    catch
                    {
                    }
                }

                _edgeBarPages.Clear();
            }

            // Free unmanaged objects here. 
            UnadviseAllSinks();

            _disposed = true;
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
