using Microsoft.Extensions.Configuration;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.IO;

namespace AspNetCoreSSRS.Conversion
{
    public class JsonRead
    {
        public static T FromFile<T>(string path)
        {
            using (StreamReader file = new StreamReader(path))
            {
                T res = default;
                try
                {
                    string json = file.ReadToEnd();

                    res = JsonConvert.DeserializeObject<T>(json);

                }
                catch (Exception e)
                {
                    Console.WriteLine("Problem reading file" + e.ToString());
                }

                return res;
            }
        }      


        public static JToken Serialize(IConfiguration config)
        {
            JObject obj = new JObject();
            foreach (var child in config.GetChildren())
            {
                obj.Add(child.Key, Serialize(child));
            }

            if (!obj.HasValues && config is IConfigurationSection section)
                return new JValue(section.Value);

            return obj;
        }
    }
}
