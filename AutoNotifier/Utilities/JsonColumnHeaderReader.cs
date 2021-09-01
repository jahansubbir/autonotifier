using System;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json;

namespace AutoNotifier.Utilities
{
    public class JsonColumnHeaderReader
    {
        public Dictionary<string, string> GetColumnDictionary()
        {
            //Dictionary<string,string> dictionary=new Dictionary<string, string>();

            string jsonFilePath = @"columns.json";
            try
            {
                string json = File.ReadAllText(jsonFilePath);
                var data = JsonConvert.DeserializeObject<Dictionary<string, string>>(json);

                return data;
            }
            catch (Exception e)
            {
                return null;
            }

        }
        /*public Dictionary<string, string> GetColumnDictionary()
        {
            Dictionary<string,string> dictionary=new Dictionary<string, string>();

            string jsonFilePath = @"columns.xml";
            try
            {
                string json = File.ReadAllText(jsonFilePath);


               return  dictionary;
            }
            catch (Exception e)
            {
                return null;
            }

        }*/
    }
}
