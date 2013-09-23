using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace HLU
{
    public class EnumCode : Attribute
    {
        public string Code;

        public EnumCode(string text)
        {
            Code = text;
        }
    }

    public static class Enum<T>
    {
        public static T Parse(string value)
        {
            return (T)Enum.Parse(typeof(T), value);
        }

        public static T Parse(string value, bool ignoreCase)
        {
            return (T)Enum.Parse(typeof(T), value, ignoreCase);
        }

        public static T[] GetValues()
        {
            return (T[])Enum.GetValues(typeof(T));
        }

        public static string[] GetNames()
        {
            return Enum.GetNames(typeof(T));
        }

        public static string[] GetCodes()
        {
            return Enum.GetValues(typeof(T)).Cast<T>().Select(e => GetCode(e)).ToArray();
        }

        public static string GetCode(T en)
        {
            MemberInfo[] memberInfo = typeof(T).GetMember(en.ToString());
            if ((memberInfo != null) && (memberInfo.Length > 0))
            {
                object[] attributes;
                attributes = memberInfo[0].GetCustomAttributes(typeof(EnumCode), false);
                if ((null != attributes) && (attributes.Length > 0))
                {
                    return ((EnumCode)attributes[0]).Code;
                }
            }
            return null;
        }

        public static Dictionary<T, string> ToValueCodeDictionary()
        {
            return Enum.GetValues(typeof(T)).Cast<T>().ToDictionary(k => k, v => GetCode(v));
        }

        public static Dictionary<string, T> ToCodeValueDictionary()
        {
            return Enum.GetValues(typeof(T)).Cast<T>().ToDictionary(k => GetCode(k), v => v);
        }

        public static Dictionary<string, string> ToCodeNameDictionary()
        {
            return Enum.GetValues(typeof(T)).Cast<T>().ToDictionary(k => GetCode(k), v => Enum.GetName(typeof(T), v));
        }

        public static Dictionary<string, string> ToNameCodeDictionary()
        {
            return Enum.GetValues(typeof(T)).Cast<T>().ToDictionary(k => Enum.GetName(typeof(T), k), v => GetCode(v));
        }
    }
}
