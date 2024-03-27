﻿// HLUTool is used to view and maintain habitat and land use GIS data.
// Copyright © 2011 Hampshire Biodiversity Information Centre
// 
// This file is part of HLUTool.
// 
// HLUTool is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
// 
// HLUTool is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
// 
// You should have received a copy of the GNU General Public License
// along with HLUTool.  If not, see <http://www.gnu.org/licenses/>.

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
