using System.Windows;





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
            channelList = ReadData.ReadXMLDataBase().ChannelList;
        }

        private void XMLRegularExpressions(object sender, RoutedEventArgs e)
        ///Read data from a file using regular expressions
        {
            channelList = ReadData.ReadXMLRegularExpressions().ChannelList;
        }

        private void AddExel(object sender, RoutedEventArgs e)
        ///Write data to excel
        {
            WriteData.WriteAddExel(channelList);
            channelList = null;
        }

        private async void AddWord(object sender, RoutedEventArgs e)
        ///Write data to word
        {
            WriteData.WriteAddWord(channelList);
            channelList = null;
        }
        private async void AddTxt(object sender, RoutedEventArgs e)
        ///Write data to txt
        {
            WriteData.WriteAddTxt(channelList);
            channelList = null;
        }
    }
}
