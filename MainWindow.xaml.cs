using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;

using Updater;

using Word_Generator.Resources;
using Word = Microsoft.Office.Interop.Word;

namespace Word_Generator
{
    public partial class MainWindow : Window
    {
        public List<string> ExcelColumns = new List<string>(); // список колонок из Excel
        public BindingList<PlaceholderMapping> Placeholders = new BindingList<PlaceholderMapping>(); // список плейсхолдеров из word

        public MainWindow()
        {
            InitializeComponent();
            PlaceholdersGrid.ItemsSource = Placeholders;
            Loaded += async (s, e) => await updater();
        }

        public async Task updater() {
            var updater = new AppUpdater("dan21321", "Word_Generator");
            var update = await updater.UpdateAsync();

            if (update != null)
            {
                MessageBoxResult result = MessageBox.Show($"{Strings.UpdaterText}", $"{Strings.UpdaterTitle}", MessageBoxButton.YesNo, MessageBoxImage.Warning, MessageBoxResult.Yes);
                if (result == MessageBoxResult.Yes) {
                    try {
                        var updaterPath = Path.Combine(AppContext.BaseDirectory, "UpdaterApp.exe");
                        var args = $"\"{update.DownloadUrl}\" \"{AppContext.BaseDirectory}\"";
                        var started = Process.Start(new ProcessStartInfo
                        {
                            FileName = updaterPath,
                            Arguments = $"\"{update.DownloadUrl}\" \"{AppContext.BaseDirectory.TrimEnd('\\')}\"",
                            CreateNoWindow = true,
                            UseShellExecute = false
                        });
                        Application.Current.Shutdown();
                    }
                    catch (Exception ex) {
                        MessageBox.Show($"{Strings.Error}{ex}");
                    }
                }
                return;
            }
        }

        // открываем выбор файла Excel
        private void BtnSelectExcel_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Excel Files|*.xlsx;*.xls";
            if (ofd.ShowDialog() == true)
            {
                TxtExcelPath.Text = ofd.FileName;
                LoadExcelColumns(ofd.FileName);
            }
        }

        // Загружаем колонки из Excel
        private void LoadExcelColumns(string path)
        {
            try
            {
                using var workbook = new XLWorkbook(path);
                var ws = workbook.Worksheet(1);
                ExcelColumns = ws.Row(1).Cells().Select(c => c.GetString().Trim()).Where(s => !string.IsNullOrEmpty(s)).ToList();

                // Обновляем ComboBox в DataGrid
                foreach (var col in PlaceholdersGrid.Columns.OfType<DataGridComboBoxColumn>())
                {
                    col.ItemsSource = ExcelColumns;
                }

                // Для выбора имени файла
                CmbFileNameColumn.ItemsSource = ExcelColumns;
            }
            catch (Exception ex)
            {
                MessageBox.Show(Strings.ErrorExcel + ex.Message);
            }
        }

        // Выбор Word шаблона
        private void BtnSelectWord_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Word Files|*.doc;*.docx";
            if (ofd.ShowDialog() == true)
                TxtWordPath.Text = ofd.FileName;
        }

        // Выбор папки для вывода
        private void BtnSelectOutput_Click(object sender, RoutedEventArgs e)
        {
            var fbd = new System.Windows.Forms.FolderBrowserDialog();
            var result = fbd.ShowDialog();
            if (result == System.Windows.Forms.DialogResult.OK)
                TxtOutputPath.Text = fbd.SelectedPath;
        }

        // Автонумерация включена
        private void ChkAutoNumber_Checked(object sender, RoutedEventArgs e)
        {
            CmbFileNameColumn.IsEnabled = false;
        }

        // Автонумерация выключена
        private void ChkAutoNumber_Unchecked(object sender, RoutedEventArgs e)
        {
            CmbFileNameColumn.IsEnabled = true;
        }

        // генерация документов
        private async void BtnGenerate_Click(object sender, RoutedEventArgs e)
        {
            // все файлы должеы быть выбраны
            if (!File.Exists(TxtExcelPath.Text) || !File.Exists(TxtWordPath.Text) || string.IsNullOrEmpty(TxtOutputPath.Text))
            {
                MessageBox.Show(Strings.ErrorSelectFiles);
                return;
            }

            ProgressBar.Value = 0; // обнуляем прогресс бар
            TxtLog.Clear(); // обнуляем логи

            // Копируем нужные данные в локальные переменные, чтобы не обращаться к UI из фонового потока
            string excelPath = TxtExcelPath.Text;
            string wordTemplatePath = TxtWordPath.Text;

            // ЕСЛИ ЭТО .DOC → КОНВЕРТИРУЕМ В .DOCX
            if (Path.GetExtension(wordTemplatePath).ToLower() == ".doc")
            {
                try
                {
                    wordTemplatePath = ConvertDocToDocx(wordTemplatePath);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(Strings.ErrorConvert + ex.Message);
                    return;
                }
            }

            string outputFolder = TxtOutputPath.Text;
            bool autoNumber = ChkAutoNumber.IsChecked == true;
            string fileNameColumn = null;

            if (!autoNumber && CmbFileNameColumn.SelectedItem != null)
                fileNameColumn = CmbFileNameColumn.SelectedItem.ToString();

            // Создаём копию плейсхолдеров
            List<PlaceholderMapping> placeholdersCopy = null;
            await Dispatcher.InvokeAsync(() =>
            {
                placeholdersCopy = Placeholders.Where(p => !string.IsNullOrEmpty(p.Placeholder) && !string.IsNullOrEmpty(p.Column))
                                               .Select(p => new PlaceholderMapping { Placeholder = p.Placeholder, Column = p.Column })
                                               .ToList();
            });

            // Всё в фоне чтобы работал прогрессбар
            await Task.Run(() =>
            {
                try
                {
                    // открываем Excel на первой странице первой строке
                    using var workbook = new XLWorkbook(excelPath);
                    var ws = workbook.Worksheet(1);
                    var headerRow = ws.Row(1);

                    // Создаём словарь placeholder → колонка
                    var phToCol = placeholdersCopy.ToDictionary(
                        p => p.Placeholder,
                        p => headerRow.Cells().First(c => c.GetString().Trim() == p.Column).Address.ColumnNumber
                    );

                    // генерация каждого документа
                    int totalRows = ws.RowsUsed().Count() - 1;
                    int fileNumber = 1;

                    foreach (var row in ws.RowsUsed().Skip(1))
                    {
                        // определяем имя
                        string outputName;
                        if (autoNumber)
                        {
                            outputName = fileNumber.ToString();
                        }
                        else if (fileNameColumn != null)
                        {
                            int colIndex = headerRow.Cells().First(c => c.GetString().Trim() == fileNameColumn).Address.ColumnNumber;
                            outputName = row.Cell(colIndex).GetString();
                        }
                        else
                        {
                            outputName = $"File{fileNumber}";
                        }

                        string safeFileName = string.Join("_", outputName.Split(Path.GetInvalidFileNameChars()));
                        string outputPath = Path.Combine(outputFolder, $"{safeFileName}.docx");

                        // копируем шаблон в отдельный файл
                        File.Copy(wordTemplatePath, outputPath, true);

                        using (var doc = WordprocessingDocument.Open(outputPath, true))
                        {
                            foreach (var ph in phToCol.Keys)
                            {
                                var cell = row.Cell(phToCol[ph]);
                                string val = cell.DataType == XLDataType.DateTime ? cell.GetDateTime().ToString("dd.MM.yyyy") : cell.GetString().Trim();
                                ReplacePlaceholder(doc, ph, val); // замена текста
                            }
                            doc.MainDocumentPart.Document.Save();
                        }

                        // Обновляем UI безопасно
                        Dispatcher.Invoke(() =>
                        {
                            TxtLog.AppendText($"{Strings.TxtLogCreateDoc}{outputPath}\n");
                            ProgressBar.Value = (double)fileNumber / totalRows * 100;
                        });

                        fileNumber++;
                    }

                    Dispatcher.Invoke(() => MessageBox.Show(Strings.TxtFinishGenerate));
                }
                catch (Exception ex)
                {
                    Dispatcher.Invoke(() => MessageBox.Show(Strings.Error + ex.Message));
                }
            });
        }
        // замена текста
        private void ReplacePlaceholder(WordprocessingDocument doc, string placeholder, string value)
        {
            // Проходим по параграфам
            foreach (var para in doc.MainDocumentPart.Document.Body.Descendants<Paragraph>())
                ReplaceTextInParagraph(para, placeholder, value);

            // Проходим по таблицам
            foreach (var tbl in doc.MainDocumentPart.Document.Body.Descendants<Table>())
                foreach (var cell in tbl.Descendants<TableCell>())
                    foreach (var para in cell.Descendants<Paragraph>())
                        ReplaceTextInParagraph(para, placeholder, value);
        }
        // замена текста в параграфе
        private void ReplaceTextInParagraph(Paragraph para, string placeholder, string replacement)
        {
            var runs = para.Elements<Run>().ToList();
            if (!runs.Any()) return;

            // Склеиваем весь текст параграфа
            string fullText = string.Concat(runs.Select(r => r.InnerText));

            if (!fullText.Contains(placeholder))
                return;

            string newText = fullText.Replace(placeholder, replacement);

            // Берём стиль ПЕРВОГО Run — это и есть формат шаблона
            RunProperties originalStyle = runs[0].RunProperties?.CloneNode(true) as RunProperties;

            // Удаляем старые Run
            foreach (var run in runs)
                run.Remove();

            // Создаём новый Run С ТЕМ ЖЕ СТИЛЕМ
            Run newRun = new Run();

            if (originalStyle != null)
                newRun.RunProperties = originalStyle;

            newRun.Append(new Text(newText) { Space = SpaceProcessingModeValues.Preserve });

            para.AppendChild(newRun);
        }

        // добавление плейсхолдера
        private void BtnAddPlaceholder_Click(object sender, RoutedEventArgs e)
        {
            Placeholders.Add(new PlaceholderMapping { Placeholder = "", Column = "" });
        }

        // удаление плейсхолдера
        private void BtnRemovePlaceholder_Click(object sender, RoutedEventArgs e)
        {
            if (PlaceholdersGrid.SelectedItem is PlaceholderMapping selected)
                Placeholders.Remove(selected);
        }
        // класс плейсхолдера
        public class PlaceholderMapping
        {
            public string Placeholder { get; set; }
            public string Column { get; set; }
        }

        // кнопка помощи
        private void HelpButton_Click(object sender, RoutedEventArgs e)
        {
            HelpWindow help = new HelpWindow();
            help.Owner = this;          // окно будет поверх главного
            help.ShowDialog();          // модально
        }
        private string ConvertDocToDocx(string docPath)
        {
            string newPath = Path.ChangeExtension(docPath, ".docx");

            Word.Application wordApp = null;
            Word.Document doc = null;

            try
            {
                wordApp = new Word.Application
                {
                    Visible = false,
                    DisplayAlerts = Word.WdAlertLevel.wdAlertsNone
                };

                doc = wordApp.Documents.Open(
                    docPath,
                    ReadOnly: true,
                    Visible: false
                );

                doc.SaveAs2(
                    newPath,
                    Word.WdSaveFormat.wdFormatXMLDocument
                );

                doc.Close(false);
                wordApp.Quit(false);
            }
            finally
            {
                if (doc != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(doc);

                if (wordApp != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp);
            }

            return newPath;
        }

    }
}

