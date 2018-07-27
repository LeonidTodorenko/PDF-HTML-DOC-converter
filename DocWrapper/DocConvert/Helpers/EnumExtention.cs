using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Reflection;

namespace DocConvert.Helpers
{
    public static class EnumExtention
    {
        public static String Name(this Enum item)
        {
            return Enum.GetName(item.GetType(), item);
        }

        public static String Description(this Enum item)
        {
            String result = item.ToString();
            Type type = item.GetType();
            MemberInfo[] memInfo = type.GetMember(item.ToString());

            if (memInfo.Length > 0)
            {
                Object[] attrs = memInfo[0].GetCustomAttributes(typeof(DescriptionAttribute), false);
                if (attrs.Length > 0)
                {
                    result = ((DescriptionAttribute)attrs[0]).Description;
                }
            }

            return result;
        }

        public static T Attribute<T>(this Enum item) where T : Attribute
        {
            T result = null;

            Type type = item.GetType();
            MemberInfo[] memberInfo = type.GetMember(item.ToString());

            if (memberInfo.Length > 0)
            {
                Object[] attrs = memberInfo[0].GetCustomAttributes(typeof(T), false);
                if (attrs.Length > 0)
                {
                    result = (T)attrs[0];
                }
            }

            return result;
        }

        public static IEnumerable<T> GetItems<T>() where T : struct
        {
            List<T> result = new List<T>();

            foreach (T item in Enum.GetValues(typeof(T)))
            {
                result.Add(item);
            }

            return result;
        }
    }
}
