using SolidEdgeCommunity;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace QA
{
    public class MyIsolatedTask : IsolatedTaskProxy
    {
        public string DoWork(string message)
        {            
            //return InvokeSTAThread<string>(DoWorkInternal);
            return InvokeSTAThread<string, string>(DoWorkInternal, message);
        }

        internal string DoWorkInternal()
        {
            return DateTime.Now.ToString();
        }

        internal string DoWorkInternal(string message)
        {
            return message;
        }
    }
}
