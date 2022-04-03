using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Xml;
using System.Xml.Serialization;
using Excel = Microsoft.Office.Interop.Excel;

namespace TestTask
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public Channel[] channelList;
        public MainWindow()
        {
            InitializeComponent();
        }

        private void XMLDataBase(object sender, RoutedEventArgs e)
        ///Read data from a file using a data model.
        {
            Channels channels;
            string path = @"data.xml";

            XmlSerializer serializer = new XmlSerializer(typeof(Channels));

            StreamReader reader = new StreamReader(path);
            channels = (Channels)serializer.Deserialize(reader);
            reader.Close();
            Console.WriteLine("Данные отлично считаны!");
            channelList = channels.ChannelList;
        }

        private void XMLRegularExpressions(object sender, RoutedEventArgs e)
        ///Read data from a file using regular expressions
        {

        }


        private void AddExel(object sender, RoutedEventArgs e)
        ///Write data to excel
        {
            /*Excel.Application xlApp = new Excel.Application();
            if(xlApp != null)
            {
                MessageBox.Show("Excel is not properly installed!!");
                return;
                Excel.Workbook xlWorkBook;
                Excel.Worksheet xlWorkSheet;
                object misValue = System.Reflection.Missing.Value;
            }*/

        }

        private void AddWord(object sender, RoutedEventArgs e)
        ///Write data to word
        {
        }

        private async void AddTxt(object sender, RoutedEventArgs e)
        ///Write data to txt
        {
            string path = "TxtAdd.txt";
            if (!File.Exists(path))
            {
                using (FileStream fs = File.Create(path));
            }
            if (channelList == null)
            {
                Console.WriteLine("Данные из файла не были взяты!");
                return;
            }
            using (StreamWriter writer = new StreamWriter(path, true))
            {   
                foreach (Channel channel in channelList)
                {
                    await writer.WriteLineAsync($"{channel.title}");
                    await writer.WriteLineAsync($"\t{channel.link}");
                    await writer.WriteLineAsync($"\t{channel.description}");
                    await writer.WriteLineAsync($"\t{channel.category}");
                    await writer.WriteLineAsync($"\t{channel.pubDate}");
                }
                Console.WriteLine("Данные успешно были записаны в TxtAdd.txt файл!");
                channelList = null;
                
            }
        }
    }
}
