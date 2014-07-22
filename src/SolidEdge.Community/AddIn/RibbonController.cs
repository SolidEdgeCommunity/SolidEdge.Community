using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Threading;

namespace SolidEdgeCommunity.AddIn
{
    public sealed class RibbonController : IDisposable,
        SolidEdgeFramework.ISEAddInEvents,
        SolidEdgeFramework.ISEAddInEventsEx
    {
        private SolidEdgeCommunity.AddIn.SolidEdgeAddIn _addIn;
        private List<Ribbon> _ribbons = new List<Ribbon>();
        private Dictionary<IConnectionPoint, int> _connectionPointDictionary = new Dictionary<IConnectionPoint, int>();
        private bool _disposed = false;

        internal RibbonController(SolidEdgeCommunity.AddIn.SolidEdgeAddIn addIn)
        {
            if (addIn == null) throw new ArgumentNullException("addIn");
            _addIn = addIn;
        }

        ~RibbonController()
        {
            Dispose(false);
        }

        #region SolidEdgeFramework.ISEAddInEvents implentation

        void SolidEdgeFramework.ISEAddInEvents.OnCommand(int CommandID)
        {
            //((SolidEdgeFramework.ISEAddInEvents)this).OnCommand(CommandID);
        }

        void SolidEdgeFramework.ISEAddInEvents.OnCommandHelp(int hFrameWnd, int HelpCommandID, int CommandID)
        {
            //((SolidEdgeFramework.ISEAddInEvents)this).OnCommandHelp(hFrameWnd, HelpCommandID, CommandID);
        }

        void SolidEdgeFramework.ISEAddInEvents.OnCommandUpdateUI(int CommandID, ref int CommandFlags, out string MenuItemText, ref int BitmapID)
        {
            MenuItemText = null;
            //((SolidEdgeFramework.ISEAddInEvents)this).OnCommandUpdateUI(CommandID, ref CommandFlags, out MenuItemText, ref BitmapID);
        }

        #endregion

        #region SolidEdgeFramework.ISEAddInEventsEx implementation

        void SolidEdgeFramework.ISEAddInEventsEx.OnCommand(int CommandID)
        {
            var ribbon = ActiveRibbon;

            if (ribbon != null)
            {
                var control = ribbon.Controls.Where(x => x.CommandId == CommandID).FirstOrDefault();

                if (control != null)
                {
                    control.DoClick();
                    ribbon.OnControlClick(control);
                }
            }
        }

        void SolidEdgeFramework.ISEAddInEventsEx.OnCommandHelp(int hFrameWnd, int HelpCommandID, int CommandID)
        {
            var ribbon = ActiveRibbon;

            if (ribbon != null)
            {
                var control = ribbon.Controls.Where(x => x.CommandId == CommandID).FirstOrDefault();

                if (control != null)
                {
                    control.DoHelp(new IntPtr(hFrameWnd), HelpCommandID);
                }
            }
        }

        void SolidEdgeFramework.ISEAddInEventsEx.OnCommandOnLineHelp(int HelpCommandID, int CommandID, out string HelpURL)
        {
            HelpURL = null;
        }

        void SolidEdgeFramework.ISEAddInEventsEx.OnCommandUpdateUI(int CommandID, ref int CommandFlags, out string MenuItemText, ref int BitmapID)
        {
            MenuItemText = null;
            var ribbon = ActiveRibbon;
            var flags = default(SolidEdgeConstants.SECommandActivation); ;

            if (ribbon != null)
            {
                var control = ribbon.Controls.Where(x => x.CommandId == CommandID).FirstOrDefault();

                if (control != null)
                {
                    if (control.Enabled)
                    {
                        flags |= SolidEdgeConstants.SECommandActivation.seCmdActive_Enabled;
                    }

                    if (control.Checked)
                    {
                        flags |= SolidEdgeConstants.SECommandActivation.seCmdActive_Checked;
                    }

                    if (control.UseDotMark)
                    {
                        flags |= SolidEdgeConstants.SECommandActivation.seCmdActive_UseDotMark;
                    }

                    flags |= SolidEdgeConstants.SECommandActivation.seCmdActive_ChangeText;
                    MenuItemText = control.Label;

                    CommandFlags = (int)flags;
                }
            }
        }

        #endregion

        #region Methods

        public void Add<TRibbon>(Guid environmentCategory, bool firstTime) where TRibbon : Ribbon
        {
            TRibbon ribbon = Activator.CreateInstance<TRibbon>();
            Add(ribbon, environmentCategory, firstTime);
        }

        public void Add(Ribbon ribbon, Guid environmentCategory, bool firstTime)
        {
            if (ribbon == null) throw new ArgumentNullException("ribbon");

            var addInEx = _addIn.AddInEx;
            var EnvironmentCatID = environmentCategory.ToString("B");

            ribbon.EnvironmentCategory = environmentCategory;

            if (IsSinkAdvised<SolidEdgeFramework.ISEAddInEvents>(addInEx) == false)
            {
                AdviseSink<SolidEdgeFramework.ISEAddInEvents>(addInEx);
            }

            if (IsSinkAdvised<SolidEdgeFramework.ISEAddInEventsEx>(addInEx) == false)
            {
                AdviseSink<SolidEdgeFramework.ISEAddInEventsEx>(addInEx);
            }

            if (_ribbons.Exists(x => x.EnvironmentCategory.Equals(ribbon.EnvironmentCategory)))
            {
                throw new System.Exception(String.Format("A ribbon has already been added for environment category {0}.", ribbon.EnvironmentCategory));
            }

            if (ribbon.EnvironmentCategory.Equals(Guid.Empty))
            {
                throw new System.Exception(String.Format("{0} is not a valid environment category.", ribbon.EnvironmentCategory));
            }

            foreach (var tab in ribbon.Tabs)
            {
                foreach (var group in tab.Groups)
                {
                    foreach (var control in group.Controls)
                    {
                        // Allocate command arrays. Please see the addin.doc in the SDK folder for details.
                        Array commandNames = new string[] { control.ToCommandName() };
                        Array commandIDs = new int[] { control.CommandId };

                        addInEx.SetAddInInfoEx(
                            _addIn.NativeResourcesDllPath,
                            EnvironmentCatID,
                            tab.Name,
                            control.ImageId,
                            -1,
                            -1,
                            -1,
                            commandNames.Length,
                            ref commandNames,
                            ref commandIDs);

                        control.SolidEdgeCommandId = (int)commandIDs.GetValue(0);

                        if (firstTime)
                        {
                            // Properly format the command bar name string.
                            string commandBarName = String.Format("{0}\n{1}", tab.Name, group.Name);

                            // Add the command bar button.
                            SolidEdgeFramework.CommandBarButton pButton = addInEx.AddCommandBarButton(EnvironmentCatID, commandBarName, control.CommandId);

                            // Set the button style.
                            if (pButton != null)
                            {
                                pButton.Style = control.Style;
                            }
                        }
                    }
                }
            }

            _ribbons.Add(ribbon);
        }

        #endregion

        #region Properties

        public Ribbon ActiveRibbon
        {
            get
            {
                var application = _addIn.Application;
                var environments = application.Environments;
                var environment = environments.Item(application.ActiveEnvironment);
                var envCatId = Guid.Parse(environment.CATID);
                return _ribbons.Where(x => x.EnvironmentCategory.Equals(envCatId)).FirstOrDefault();
            }
        }

        public IEnumerable<Ribbon> Ribbons { get { return _ribbons.AsEnumerable(); } }

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
                foreach (var ribbon in _ribbons)
                {
                    try
                    {
                        ribbon.Dispose();
                    }
                    catch
                    {
                    }
                }

                _ribbons.Clear();
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
                    Monitor.Enter(this, ref lockTaken);

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
