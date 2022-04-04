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
using Path = System.IO.Path;
using Word = Microsoft.Office.Interop.Word;

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

        private async void XMLDataBase(object sender, RoutedEventArgs e)
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

        private async void AddWord(object sender, RoutedEventArgs e)
        ///Write data to word
        {
            try
            {
                //Create an instance for word app  
                Word.Application winword = new Word.Application();

                //Set animation status for word application  
                 winword.ShowAnimation = false;

                    //Set status for word application is to be visible or not.  
                winword.Visible = true;

                    //Create a missing variable for missing value  
                object missing = System.Reflection.Missing.Value;

                //Create a new document  
                Word.Document document = winword.Documents.Add(ref missing, ref missing, ref missing, ref missing);
                if (channelList == null)
                {
                    Console.WriteLine("Данные из файла не были взяты!");
                    return;
                }

                Word.Paragraph para1 = document.Content.Paragraphs.Add(ref missing);

                foreach (Channel channel in channelList)
                {
                    para1.Range.Text = $"\t{channel.title} " + Environment.NewLine;
                    para1.Range.Text = $"\t{channel.link}" + Environment.NewLine;
                    para1.Range.Text = $"\t{channel.description} " + Environment.NewLine;
                    para1.Range.Text = $"\t{channel.category} " + Environment.NewLine;
                    para1.Range.Text = $"\t{channel.pubDate}\n\n\n " + Environment.NewLine;
                }
                Console.WriteLine("Данные успешно были записаны в WordAdd.docx файл!");
                channelList = null;


                //Save the document  
                object filename = @"D:\Projects\AlfaBank\TestTask\bin\Debug\WordApp.docx";
                document.SaveAs2(ref filename);
                document.Close(ref missing, ref missing, ref missing);
                document = null;
                winword.Quit(ref missing, ref missing, ref missing);
                winword = null;
            }
            catch (Exception ex)
            {
                 MessageBox.Show(ex.Message);
            }
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
