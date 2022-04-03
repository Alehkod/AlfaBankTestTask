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

namespace TestTask
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }
        private void XMLDataBase(object sender, RoutedEventArgs e)
        {
            Channels channels;
            string path = @"data.xml";

            XmlSerializer serializer = new XmlSerializer(typeof(Channels));

            StreamReader reader = new StreamReader(path);
            channels = (Channels)serializer.Deserialize(reader);
            reader.Close();
            Console.WriteLine("Данные отлично считаны!");         

        }
        private void XMLRegularExpressions(object sender, RoutedEventArgs e)
        {
            
        }
        private void AddExel(object sender, RoutedEventArgs e)
        {

        }
        private void AddWord(object sender, RoutedEventArgs e)
        {
        }
        private void AddTxt(object sender, RoutedEventArgs e)
        {
            
        }
    }
}
