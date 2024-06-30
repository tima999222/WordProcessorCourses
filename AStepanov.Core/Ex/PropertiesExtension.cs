using System.Reflection;

namespace AStepanov.Core.Ex
{
    public static class PropertiesExtension
    {
        public static Dictionary<string, object> GetProperties<T>(this T obj)
        {
            var properties = typeof(T).GetProperties(BindingFlags.Instance | BindingFlags.Public);
            var dict = new Dictionary<string, object>();

            foreach (var prop in properties)
            {
                dict[prop.Name] = prop.GetValue(obj);
            }

            return dict;
        }
    }
}
