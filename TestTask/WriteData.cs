using System;
using System.IO;
using System.Windows;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;

namespace TestTask
{
    static public class WriteData
    {
        public static async void WriteAddTxt(Channel[] channelList)
        ///Write data to txt
        {
            string path = "TxtAdd.txt";
            if (!File.Exists(path))
            {
                using (FileStream fs = File.Create(path)) ;
            }
            if (channelList == null)
            {
                Console.WriteLine("Данные из файла не были взяты!");
                return;
            }
            using (FileStream fs = File.Create(path)) ;
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

            }
        }
        public static void WriteAddWord(Channel[] channelList)
        {
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
                    object filename = "WordApp.docx";
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
        }
        public static void WriteAddExel(Channel[] channelList)
        ///Write data to excel
        {
            Console.WriteLine("Данные успешно записаны в ExelAdd.xlsx");
        }
    }
}
