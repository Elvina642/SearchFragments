using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using Microsoft.Win32;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using System.Windows.Controls;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using Aspose.Words;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using iText.Kernel.Geom;
using iText.Kernel.Pdf.Canvas.Parser.Data;


public class SpacingAwareTextExtractionStrategy : ITextExtractionStrategy
{
    private readonly StringBuilder sb = new StringBuilder();
    private iText.Kernel.Geom.Vector lastEnd;
    private readonly float spaceThreshold;

    public SpacingAwareTextExtractionStrategy(float spaceThreshold = 3f)
    {
        this.spaceThreshold = spaceThreshold;
    }

    public void EventOccurred(IEventData data, EventType type)
    {
        if (type == EventType.RENDER_TEXT)
        {
            var renderInfo = (TextRenderInfo)data;
            iText.Kernel.Geom.LineSegment segment = renderInfo.GetBaseline();
            iText.Kernel.Geom.Vector start = segment.GetStartPoint();
            iText.Kernel.Geom.Vector end = segment.GetEndPoint();

            if (lastEnd != null)
            {
                float distance = start.Subtract(lastEnd).Length();
                if (distance > spaceThreshold)
                {
                    sb.Append(' ');
                }
            }

            sb.Append(renderInfo.GetText());
            lastEnd = end;
        }
    }

    public string GetResultantText() => sb.ToString();

    public ICollection<EventType> GetSupportedEvents() => null;

    public void BeginTextBlock() { }

    public void EndTextBlock() { }

    public void RenderText(TextRenderInfo renderInfo) { }

    public void RenderImage(ImageRenderInfo renderInfo) { }
}

namespace SearchFragments
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : IMainWindow 
    {
        private string loadedText; // Объявление переменной класса для хранения текста

        public MainWindow()
        {
            InitializeComponent();
        }

        private void TextBox_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        private void RadioButton_Checked(object sender, RoutedEventArgs e)
        {

        }

        private void RadioButton_Checked_1(object sender, RoutedEventArgs e)
        {

        }


        private void BtnLoadFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "All files (*.*)|*.*",
                FilterIndex = 1, // Устанавливаем начальный фильтр (все файлы)
                Title = "Выберите файл"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                string filePath = openFileDialog.FileName;

                try
                {
                    string fileExtension = System.IO.Path.GetExtension(filePath).ToLower();
                    if (fileExtension == ".txt" || fileExtension == ".docx" || fileExtension == ".doc" || fileExtension == ".pdf")
                    {
                        loadedText = LoadFile(filePath); // Загружаем содержимое файла
                        txtFileName.Text = System.IO.Path.GetFileName(filePath); // Выводим название файла в TextBox
                    }
                    else
                    {
                        MessageBox.Show("Формат файла не поддерживается.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ошибка загрузки файла: " + ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private string LoadFile(string filePath)
        {
            
            if (filePath.EndsWith(".txt"))
            {
                return File.ReadAllText(filePath); // Читает текст из .txt файла
            }
            else if (filePath.EndsWith(".docx"))
            {
                using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, false))
                {
                    return doc.MainDocumentPart.Document.Body.InnerText; // Получает текст из .docx файла
                }
            }
            else if (filePath.EndsWith(".pdf"))
            {
                using (PdfReader pdfReader = new PdfReader(filePath))
                using (PdfDocument pdf = new PdfDocument(pdfReader))
                {
                    StringBuilder sb = new StringBuilder();
                    for (int page = 1; page <= pdf.GetNumberOfPages(); page++)
                    {
                        var strategy = new SpacingAwareTextExtractionStrategy(2.5f); // указанный порог, можно изменять
                        var parser = new PdfCanvasProcessor(strategy);
                        parser.ProcessPageContent(pdf.GetPage(page));
                        sb.Append(strategy.GetResultantText());
                    }
                    return sb.ToString();
                }
            }
            else if (filePath.EndsWith(".doc"))
            {
                // Загружаем документ через Aspose.Words
                Document doc = new Document(filePath);
                return doc.GetText(); // Возвращаем текст документа
            }


            return ""; // Возвращает пустую строку, если файл не поддерживается
        }

        

        private void BtnSearch_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(loadedText))
            {
                MessageBox.Show("Сначала загрузите файл.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            string processedText = loadedText;


            if (!int.TryParse(txtFragmentLength.Text, out int fragmentLength) || fragmentLength < 1)
            {
                MessageBox.Show("Введите корректную длину фрагмента (больше 0).", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Словарь для хранения фрагментов и их количества
            Dictionary<string, int> fragmentCounts = new Dictionary<string, int>();

            // Определяем, какой вариант выбран
            bool checkUpperCaseCyrillicn = rbUpperCaseCyrillic.IsChecked == true;
            bool checkUpperCaseLatin = rbUpperCaseLatin.IsChecked == true;
            bool checkAnySymbol = rbAnySymbol.IsChecked == true;

            // Перебираем текст посимвольно для поиска фрагментов
            for (int i = 0; i <= loadedText.Length - fragmentLength; i++)
            {
                // Проверяем, что фрагмент находится в допустимых пределах
                if (i + fragmentLength > processedText.Length)
                {
                    break; // Прерываем цикл, если конец фрагмента выходит за пределы строки
                }

                // Извлекаем фрагмент текста указанной длины
                string fragment = loadedText.Substring(i, fragmentLength);

                // Проверяем, что фрагмент начинается с пробела или с начала строки
                if (i > 0 && char.IsLetterOrDigit(processedText[i - 1])) // Если перед фрагментом есть буква или цифра (это не начало слова)
                {
                    continue; // Пропускаем фрагмент, если он не начинается с пробела или начала строки
                }

                // Проверяем, что фрагмент начинается с нужной буквы
                if (checkUpperCaseCyrillicn && !(char.IsUpper(fragment[0]) && fragment[0] >= 'А' && fragment[0] <= 'Я'))
                {
                    continue;
                }
                else if (checkUpperCaseLatin && !(char.IsUpper(fragment[0]) && fragment[0] >= 'A' && fragment[0] <= 'Z'))
                {
                    continue;
                }
                else if (!checkAnySymbol && !char.IsLetter(fragment[0]))
                {
                    // Проверяем, что фрагмент начинается с буквы (русской или английской), без знаков препинания
                    if (!char.IsLetter(fragment[0]))
                    {
                        continue; // Пропускаем, если фрагмент не начинается с буквы
                    }

                    // Проверяем, что весь фрагмент состоит из букв
                    if (!fragment.All(c => char.IsLetter(c)))
                    {
                        continue; // Пропускаем фрагмент, если в нем есть символы, которые не являются буквами
                    }


                }

                // Подсчитываем фрагмент
                if (fragmentCounts.ContainsKey(fragment))
                {
                    fragmentCounts[fragment]++;
                }
                else
                {
                    fragmentCounts[fragment] = 1;
                }

            }

            var sortedFragments = fragmentCounts
         .Where(kv => kv.Value > 1)
         .OrderByDescending(kv => kv.Value) // Сортировка по количеству повторений (по убыванию)
         .ThenBy(kv => kv.Key, StringComparer.CurrentCulture) // Корректная сортировка по алфавиту для разных языков
         .ToList();



            StringBuilder resultBuilder = new StringBuilder();
            resultBuilder.AppendLine($"Найдено {sortedFragments.Count} фрагментов длиной {fragmentLength} символов:\n");

            foreach (var kv in sortedFragments)
            {
                resultBuilder.AppendLine($"Фрагмент: \"{kv.Key}\", повторов: {kv.Value}");
            }

            txtResult.Text = resultBuilder.ToString();
        }




        private void TxtFragmentLength_TextChanged(object sender, TextChangedEventArgs e)
        {

        }
        
        private void TxtFileName_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtResult.Text))
            {
                MessageBox.Show("Нет данных для сохранения.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "Text files (*.txt)|*.txt",
                Title = "Сохранить результаты",
                FileName = "SearchResults.txt"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                File.WriteAllText(saveFileDialog.FileName, txtResult.Text);
            }
        }
    }
}
