using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SolidEdgeCommunity.AddIn
{
    static class GuidExtensions
    {
        public static string ToRegistryString(this Guid guid)
        {
            return guid.ToString("B");
        }
    }

    static class StringExtensions
    {
        public static bool Parse(this string value, bool defaultValue)
        {
            if (value == null) return defaultValue;

            bool result = defaultValue;

            if (bool.TryParse(value, out result))
            {
                return result;
            }

            return defaultValue;
        }

        public static Guid Parse(this string value, Guid defaultValue)
        {
            if (value == null) return defaultValue;

            Guid result = defaultValue;

            if (Guid.TryParse(value, out result))
            {
                return result;
            }

            return defaultValue;
        }

        public static int Parse(this string value, int defaultValue)
        {
            if (value == null) return defaultValue;

            int result = defaultValue;

            if (int.TryParse(value, out result))
            {
                return result;
            }

            return defaultValue;
        }

        public static T ParseEnum<T>(this string value, T defaultValue) where T : struct, IConvertible
        {
            if (value == null) return defaultValue;

            T result = defaultValue;

            if (Enum.TryParse<T>(value, out result))
            {
                return result;
            }

            return defaultValue;
        }
    }
}
