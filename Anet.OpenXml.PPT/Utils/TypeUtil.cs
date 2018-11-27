using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Reflection;

namespace Anet.OpenXml.PPT
{
    public static class TypeUtil
    {
        /// <summary>
        /// 获取按顺序排列的属性(没有Display的属性会被忽略)
        /// </summary>
        public static List<PropertyInfo> GetOrderedDisplayProps<T>()
        {
            var dic = new SortedDictionary<int, PropertyInfo>();
            var props = typeof(T).GetProperties();
            foreach (var prop in props)
            {
                var attr = prop.GetCustomAttribute<DisplayAttribute>();
                if (attr == null)
                    continue;
                dic.Add(attr.Order, prop);
            }
            return dic.Values.ToList();
        }
    }
}
