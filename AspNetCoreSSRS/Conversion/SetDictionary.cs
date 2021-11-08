using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;
using System;
using System.Collections.Generic;
using System.IO;

namespace AspNetCoreSSRS.Conversion
{
    public class SetDictionary
    {
        public static Dictionary<string, object> Read(string path)
        {
            using (StreamReader file = new StreamReader(path))
            {
                Dictionary<string, object> res = default;
                try
                {
                    string json = file.ReadToEnd();

                    var serializerSettings = new JsonSerializerSettings
                    {
                        ContractResolver = new CamelCasePropertyNamesContractResolver()
                    };

                    res = JsonConvert.DeserializeObject<Dictionary<string, object>>(json, serializerSettings);                    

                }
                catch (Exception)
                {
                    Console.WriteLine("Problem reading file");
                }

                return res;
            }
        }
    }
}
