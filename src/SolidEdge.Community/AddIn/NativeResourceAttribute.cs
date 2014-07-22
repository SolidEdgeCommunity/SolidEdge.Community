using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SolidEdgeCommunity.AddIn
{
    /// <summary>
    /// Defines a resource file (.bmp | .png) to be embedded as a native Win32 resource at compile time.
    /// </summary>
    [AttributeUsage(AttributeTargets.Assembly, AllowMultiple=true)]
    public class NativeResourceAttribute : Attribute
    {
        private int _id;
        private string _path;

        /// <summary>
        /// Initializes a new instance of the class.
        /// </summary>
        /// <param name="id">Unique resource ID.</param>
        /// <param name="path">Path to resource file (.bmp | .png) relative to project.</param>
        public NativeResourceAttribute(int id, string path)
        {
            if (id < 0) throw new ArgumentException("id must be zero or greater.");
            if (String.IsNullOrWhiteSpace(path)) throw new ArgumentException("path must not be empty.");

            _id = id;
            _path = path;
        }

        public int Id { get { return _id; } }
        public string Path { get { return _path; } }

        public override string ToString()
        {
            return String.Format("{0} | {1}", Id, Path);
        }
    }
}
