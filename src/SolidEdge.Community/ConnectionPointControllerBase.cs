//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Runtime.InteropServices.ComTypes;
//using System.Text;
//using System.Threading;

//namespace SolidEdgeCommunity
//{
//    /// <summary>
//    /// Controller base class that handles connecting\disconnecting to COM events via IConnectionPointContainer and IConnectionPoint interfaces.
//    /// </summary>
//    public abstract class ConnectionPointControllerBase
//    {
//        private Dictionary<IConnectionPoint, int> _connectionPointDictionary = new Dictionary<IConnectionPoint, int>();

//        /// <summary>
//        /// Establishes a connection between a connection point object and the client's sink.
//        /// </summary>
//        /// <typeparam name="TInterface">Interface type of the outgoing interface whose connection point object is being requested.</typeparam>
//        /// <param name="container">An object that implements the IConnectionPointContainer inferface.</param>
//        protected void AdviseSink<TInterface>(object container) where TInterface : class
//        {
//            bool lockTaken = false;

//            try
//            {
//                Monitor.Enter(this, ref lockTaken);

//                // Prevent multiple event Advise() calls on same sink.
//                if (IsSinkAdvised<TInterface>(container))
//                {
//                    return;
//                }

//                IConnectionPointContainer cpc = null;
//                IConnectionPoint cp = null;
//                int cookie = 0;

//                cpc = (IConnectionPointContainer)container;
//                cpc.FindConnectionPoint(typeof(TInterface).GUID, out cp);

//                if (cp != null)
//                {
//                    cp.Advise(this, out cookie);
//                    _connectionPointDictionary.Add(cp, cookie);
//                }
//            }
//            finally
//            {
//                if (lockTaken)
//                {
//                    Monitor.Exit(this);
//                }
//            }
//        }

//        /// <summary>
//        /// Determines if a connection between a connection point object and the client's sink is established.
//        /// </summary>
//        /// <param name="container">An object that implements the IConnectionPointContainer inferface.</param>
//        protected bool IsSinkAdvised<TInterface>(object container) where TInterface : class
//        {
//            bool lockTaken = false;

//            try
//            {
//                Monitor.Enter(this, ref lockTaken);

//                IConnectionPointContainer cpc = null;
//                IConnectionPoint cp = null;
//                int cookie = 0;

//                cpc = (IConnectionPointContainer)container;
//                cpc.FindConnectionPoint(typeof(TInterface).GUID, out cp);

//                if (cp != null)
//                {
//                    if (_connectionPointDictionary.ContainsKey(cp))
//                    {
//                        cookie = _connectionPointDictionary[cp];
//                        return true;
//                    }
//                }
//            }
//            finally
//            {
//                if (lockTaken)
//                {
//                    Monitor.Exit(this);
//                }
//            }

//            return false;
//        }

//        /// <summary>
//        /// Terminates an advisory connection previously established between a connection point object and a client's sink.
//        /// </summary>
//        /// <typeparam name="TInterface">Interface type of the interface whose connection point object is being requested to be removed.</typeparam>
//        /// <param name="container">An object that implements the IConnectionPointContainer inferface.</param>
//        protected void UnadviseSink<TInterface>(object container) where TInterface : class
//        {
//            bool lockTaken = false;

//            try
//            {
//                Monitor.Enter(this, ref lockTaken);

//                IConnectionPointContainer cpc = null;
//                IConnectionPoint cp = null;
//                int cookie = 0;

//                cpc = (IConnectionPointContainer)container;
//                cpc.FindConnectionPoint(typeof(TInterface).GUID, out cp);

//                if (cp != null)
//                {
//                    if (_connectionPointDictionary.ContainsKey(cp))
//                    {
//                        cookie = _connectionPointDictionary[cp];
//                        cp.Unadvise(cookie);
//                        _connectionPointDictionary.Remove(cp);
//                    }
//                }
//            }
//            finally
//            {
//                if (lockTaken)
//                {
//                    Monitor.Exit(this);
//                }
//            }
//        }

//        /// <summary>
//        /// Terminates all advisory connections previously established.
//        /// </summary>
//        protected void UnadviseAllSinks()
//        {
//            bool lockTaken = false;

//            try
//            {
//                Monitor.Enter(this, ref lockTaken);
//                Dictionary<IConnectionPoint, int>.Enumerator enumerator = _connectionPointDictionary.GetEnumerator();
//                while (enumerator.MoveNext())
//                {
//                    enumerator.Current.Key.Unadvise(enumerator.Current.Value);
//                }
//            }
//            finally
//            {
//                _connectionPointDictionary.Clear();

//                if (lockTaken)
//                {
//                    Monitor.Exit(this);
//                }
//            }
//        }

//        /// <summary>
//        /// Establishes or terminates a connection between a connection point object and the client's sink.
//        /// </summary>
//        /// <typeparam name="TInterface">Interface type of the interface whose connection point object is being requested to be updated.</typeparam>
//        /// <param name="container">An object that implements the IConnectionPointContainer inferface.</param>
//        /// <param name="advise">Flag indicating whether to advise or unadvise.</param>
//        //protected void UpdateSink<TInterface>(object container, bool advise) where TInterface : class
//        //{
//        //    if (advise)
//        //    {
//        //        AdviseSink<TInterface>(container);
//        //    }
//        //    else
//        //    {
//        //        UnadviseSink<TInterface>(container);
//        //    }
//        //}
//    }
//}
