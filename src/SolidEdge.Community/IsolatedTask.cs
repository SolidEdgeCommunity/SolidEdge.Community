using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;

namespace SolidEdgeCommunity
{
    /// <summary>
    /// Generic class used to execute an IsolatedTaskProxy implementation.
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public sealed class IsolatedTask<T> : IDisposable where T : IsolatedTaskProxy
    {
        private Type _proxyType = null;
        private AppDomain _appDomain = null;
        private T _proxy = null;

        /// <summary>
        /// 
        /// </summary>
        public IsolatedTask()
        {
            // Get the proxy type.
            _proxyType = typeof(T);

            // Create a custom AppDomain to do COM Interop.
            _appDomain = AppDomain.CreateDomain(String.Format("{0} AppDomain", _proxyType.Name), null, AppDomain.CurrentDomain.SetupInformation);

            // Create a new instance of InteropProxy in the isolated application domain.
            _proxy = (T)_appDomain.CreateInstanceAndUnwrap(_proxyType.Assembly.FullName, _proxyType.FullName);
        }

        void IDisposable.Dispose()
        {
            if (_appDomain != null)
            {
                // Unload the Interop AppDomain. This will automatically free up any COM references.
                AppDomain.Unload(_appDomain);
            }

            _proxy = null;
            _appDomain = null;
            _proxyType = null;
        }

        public T Proxy { get { return _proxy; } }
    }
}
