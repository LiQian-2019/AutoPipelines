using Autodesk.AutoCAD.DatabaseServices;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace AutoPipelines
{
    public static class  ListTools
    {
        public static IList<T> Init<T>(this IList<T> list, int number)
        {
            if (number<list.Count())
            {
                throw new IndexOutOfRangeException();
            }
            Type t = typeof(T);
            PropertyInfo[] propertyList = t.GetProperties();
            var count = list.Count();
            for (int i = 0; i < number-count; i++)
            {
                var newListItem = Activator.CreateInstance(t);

                foreach (var property in propertyList)
                {
                    if (property.PropertyType.Equals(typeof(int)))
                    {
                        property.SetValue(newListItem, 0);
                    }
                    else if (property.PropertyType.Equals(typeof(string)))
                    {
                        property.SetValue(newListItem, "");
                    }
                    else if (property.PropertyType.Equals(typeof(double)))
                    {
                        property.SetValue(newListItem, 0.0);
                    }
                    else if (property.PropertyType.Equals(typeof(ushort)))
                    {
                        property.SetValue(newListItem, (ushort)0);
                    }
                    else if (property.PropertyType.Equals(typeof(Enum)))
                    {
                        property.SetValue(newListItem, 0);
                    }
                }

                list.Add((T)newListItem);
            }
            return list;
        }

        public static IEnumerable<Entity> Filter1(this IEnumerable<Entity> oldEntities)
        {
            var newEntities = from e in oldEntities
                              where e is Polyline
                              where !e.Layer.EndsWith("FZL")
                              select e;
            return newEntities;
        }
    }
}
