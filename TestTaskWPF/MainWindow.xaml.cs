using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Xml;
using System.Xml.Serialization;
using System.IO;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Win32;

namespace TestTaskWPF
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private string _xmlFilePath = @"data\data.xml";
        private Channel _channel = new Channel();

        private void ReadXMLButton_Click(object sender, EventArgs e)
        {
            XmlSerializer xmlSerializer = new XmlSerializer(typeof(Channel));

            try
            {
                using (FileStream fileStream = new FileStream(_xmlFilePath, FileMode.Open))
                {
                    _channel = xmlSerializer.Deserialize(fileStream) as Channel;

                    if (_channel != null)
                    {
                        textBox.Text = "";

                        foreach (Item item in _channel.Item)
                        {
                            textBox.Text += $"{item.GetNews()}\n\n";
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Не удалось загрузить файл с данными.");
                Console.WriteLine($"Ошибка: {ex}");
            }
        }

        private void SaveFileToExcelButton_Click(object sender, EventArgs e)
        {
            if (textBox.Text != "")
            {
                Excel.Application excelApplication = new Excel.Application();
                excelApplication.SheetsInNewWorkbook = 1;

                Excel.Workbook excelWorkbook = excelApplication.Workbooks.Add(Type.Missing);
                Excel.Worksheet excelWorksheet = excelApplication.Worksheets.get_Item(1);
                excelWorksheet.Name = "NewsDocument";

                int indexY = 1;
                foreach(Item item in _channel.Item)
                {
                    excelWorksheet.Cells[indexY, 1] = item.Title;
                    excelWorksheet.Cells[indexY, 2] = item.Link;
                    excelWorksheet.Cells[indexY, 3] = item.Description;
                    excelWorksheet.Cells[indexY, 4] = item.Category;
                    excelWorksheet.Cells[indexY, 5] = item.PubDate;

                    indexY++;
                }

                try
                {
                    excelWorkbook.Close();
                    excelApplication.Quit();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Ошибка: {ex}");
                }
            }
            else
            {
                MessageBox.Show("Нету данных для записи.");
            }
        }

        private void SaveFileToWordButton_Click(object sendet, EventArgs e)
        {
            if (textBox.Text != "")
            {
                Word.Application worldApplication = new Word.Application();
                Word.Document wordDocument = worldApplication.Documents.Add(Visible: true);

                Word.Range rangeText = wordDocument.Range();
                rangeText.Text = textBox.Text;

                try
                {
                    wordDocument.Close();
                    worldApplication.Quit();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Ошибка: {ex}");
                }
            }
            else
            {
                MessageBox.Show("Нету данных для записи.");
            }
        }

        private async void SaveFileToTXTButton_Click(object sender, EventArgs e)
        {
            if (textBox.Text != "")
            {
                SaveFileDialog saveDialog = new SaveFileDialog();
                saveDialog.Filter = "Text File |*.txt";
                saveDialog.FileName = "NewsDocument";
                saveDialog.ShowDialog();

                if (saveDialog.FileName != "")
                {
                    using (FileStream fileStream = (FileStream)saveDialog.OpenFile())
                    {
                        byte[] buffer = Encoding.Default.GetBytes(textBox.Text);
                        await fileStream.WriteAsync(buffer, 0, buffer.Length);
                    }
                }
            }
            else
            {
                MessageBox.Show("Нету данных для записи.");
            }
        }

        private void ExitBuuton_Click(object sender, EventArgs e)
        {
            Application.Current.Shutdown();
        }
    }

    [XmlRoot(ElementName = "channel")]
    public class Channel
    {
        [XmlElement(ElementName = "item")]
        public List<Item> Item { get; set; }
    }

    [XmlRoot(ElementName = "item")]
    public class Item
    {
        [XmlElement(ElementName = "title")]
        public string Title { get; set; }

        [XmlElement(ElementName = "link")]
        public string Link { get; set; }

        [XmlElement(ElementName = "description")]
        public string Description { get; set; }

        [XmlElement(ElementName = "category")]
        public string Category { get; set; }

        [XmlElement(ElementName = "pubDate")]
        public string PubDate { get; set; }

        public string GetNews()
        {
            return $"Заголовок: {Title}\nСсылка: {Link}\nОписание: {Description}\nКатегория: {Category}\nДата публикации {PubDate}";
        }
    }
}
