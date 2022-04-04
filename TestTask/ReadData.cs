using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace TestTask
{
    public static class ReadData
    {
        public static Channels ReadXMLDataBase()
        ///Read data from a file using a data model.
        {
            Channels channels;
            string path = @"data.xml";

        XmlSerializer serializer = new XmlSerializer(typeof(Channels));

        StreamReader reader = new StreamReader(path);
        channels = (Channels) serializer.Deserialize(reader);
        reader.Close();
            Console.WriteLine("Данные отлично считаны!");
            return channels;
        }
        public static Channels ReadXMLRegularExpressions()
        {
            Channels channels = null;
            return channels;
        }
    }
}
