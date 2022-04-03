using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace TestTask
{

    [Serializable()]
    public class Channel
    {
        [System.Xml.Serialization.XmlElement("title")]
        public string title { get; set; }

        [System.Xml.Serialization.XmlElement("link")]
        public string link { get; set; }

        [System.Xml.Serialization.XmlElement("description")]
        public string description { get; set; }
        [System.Xml.Serialization.XmlElement("category")]
        public string category { get; set; }
        [System.Xml.Serialization.XmlElement("pubDate")]
        public string pubDate { get; set; }
    }


    [XmlRootAttribute("channel")]
    public class Channels
    {
        [XmlElement("item")]
        public Channel[] ChannelList { get; set; }
    }

}