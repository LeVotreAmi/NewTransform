using System;
using System.Threading.Tasks;
using System.Windows;
using WinForms = System.Windows.Forms;
using System.Windows.Controls;
using System.Windows.Input;
using System.Collections.Generic;
using System.IO;
using System.Windows.Media.Animation;
using Microsoft.Win32;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using OfficeConvert;
using System.Text.RegularExpressions;

namespace NewTransform
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        static string pathDoc = "";
        static string pathXls = "";
        static string pathSave = "";

        private void Properties(Button generatedButton)
        {
            generatedButton.IsEnabled = true;
        }

        public MainWindow()
        {
            InitializeComponent();
        }

        private void SaveFolderPanel_Drop(object sender, DragEventArgs e)
        {
            try
            {
                if (e.Data.GetDataPresent(DataFormats.FileDrop))
                {
                    string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                    string file = (string)files[0];

                    SaveFolderLabel.Content = file;
                    pathSave = file;
                }

                if (pathSave != "" && pathDoc != "" && pathXls != "")
                {
                    Properties(generatedButton);
                }
            } catch
            {
                MessageBox.Show("Что-то пошло не так", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            
        }

        private void XlsxDropPanel_Drop(object sender, DragEventArgs e)
        {
            try
            {
                if (e.Data.GetDataPresent(DataFormats.FileDrop))
                {
                    xlsxLabel.Visibility = Visibility.Hidden;
                    string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                    string file = (string)files[0];

                    xlsxLabelTitle.IsEnabled = true;
                    xlsxLabelTitle.Content = file;
                    pathXls = file;
                }

                if (pathSave != "" && pathDoc != "" && pathXls != "")
                {
                    Properties(generatedButton);
                }
            } catch
            {
                MessageBox.Show("Что-то пошло не так", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            
        }

        private void DocDropPanel_Drop(object sender, DragEventArgs e)
        {
            try
            {
                if (e.Data.GetDataPresent(DataFormats.FileDrop))
                {
                    docLabel.Visibility = Visibility.Hidden;
                    string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                    string file = (string)files[0];

                    docLabelTitle.IsEnabled = true;
                    docLabelTitle.Content = file;
                    pathDoc = file;
                }

                if (pathSave != "" && pathDoc != "" && pathXls != "")
                {
                    Properties(generatedButton);
                }
            } catch
            {
                MessageBox.Show("Что-то пошло не так", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            
        }

        private void Generate()
        {
            var WordApp = new Word.Application();
            WordApp.Visible = false;

            var WordDocument = WordApp.Documents.Open(@"" + pathDoc);
            Excel.Application ObjWorkExcel = new Excel.Application();
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(@"" + pathXls);
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];

            try
            {

                int nInLastRow = ObjWorkSheet.Cells.Find("*", System.Reflection.Missing.Value,
                    System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                    Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious,
                    false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;
                int nInLastCol = ObjWorkSheet.Cells.Find("*", System.Reflection.Missing.Value,
                    System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                    Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlPrevious,
                    false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Column;

                string[,] list = new string[nInLastRow, nInLastCol];

                for (int i = 0; i < nInLastRow; i++)
                {
                    for (int j = 0; j < nInLastCol; j++)
                    {
                        list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text.ToString();

                    }
                }

                for (int i = 0; i < nInLastRow; i++)
                {
                    for (int j = 0; j < nInLastCol; j++)
                    {

                        string tmp = list[i, j];
                        ReplaceWordStub("{" + j + "}", tmp, WordDocument);

                        if (j == nInLastCol - 1)
                        {
                            string fName = Regex.Replace(list[i, j], "[\"«»]", "");
                            WordDocument.SaveAs(@"" + pathSave + "\\" + fName + ".docx");
                            WordDocument.Close(false, Type.Missing, Type.Missing);
                            WordDocument = WordApp.Documents.Open(@"" + pathDoc);
                        }

                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                WordDocument.Close(false, Type.Missing, Type.Missing);
                WordApp.Quit();
                ObjWorkBook.Close(false, Type.Missing, Type.Missing);
                ObjWorkExcel.Quit();
                GC.Collect();
            }
        }

        private async void GeneratedButton_Click(object sender, RoutedEventArgs e)
        {
            generatedButton.IsEnabled = false;
            Panel.SetZIndex(imgGenerate, 10);

            await Task.Run(()=>Generate());

            generatedButton.IsEnabled = true;
            Panel.SetZIndex(imgGenerate, -1);
        }

        static void ReplaceWordStub(string StubToReplace, string Text, Word.Document WordDocument)
        {

            var Range = WordDocument.Content;
            Range.Find.Execute(FindText: StubToReplace, ReplaceWith: Text);

        }
        
        private void XlsxDropPanel_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            
            ofd.Filter = "Файлы xlsx|*.xlsx";
            ofd.ShowDialog();

            pathXls = ofd.FileName;
            if (pathXls != "" && pathDoc != "" && pathSave != "")
            {
                Properties(generatedButton);
            }
        }

        private void DocDropPanel_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Файлы docx|*.docx|Файлы doc|*.doc";
            ofd.ShowDialog();

            pathDoc = ofd.FileName;
            if (pathXls != "" && pathDoc != "" && pathSave != "")
            {
                Properties(generatedButton);
            }
        }

        private void SaveFolderPanel_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            WinForms.FolderBrowserDialog fbd = new WinForms.FolderBrowserDialog();
            fbd.ShowDialog();

            pathSave = fbd.SelectedPath;
            if (pathXls != "" && pathDoc != "" && pathSave != "")
            {
                Properties(generatedButton);
            }
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            GC.Collect();
        }

        private void tryConvert(Converter converter, String inputFile, String outputFile)
        {
            try
            {
                converter.Convert(inputFile, outputFile);
            }
            catch (ConvertException err)
            {
                MessageBox.Show(err.Message + "\n\n" + err.StackTrace);
            }
        }

        private void ConvertPanelFuncDrop(DragEventArgs e)
        {
            Dispatcher.Invoke(() => ConvertProgressBar.Visibility = Visibility.Visible);

            List<string> paths = new List<string>();

            foreach (string obj in (string[])e.Data.GetData(DataFormats.FileDrop))
                if (Directory.Exists(obj))
                    paths.AddRange(Directory.GetFiles(obj, "*.*", SearchOption.AllDirectories));
                else
                    paths.Add(obj);

            int count = paths.Count;
            Dispatcher.Invoke(() => ConvertProgressBar.Maximum = count);

            for (int i = 0; i < count; i++)
            {
                String inputFile = paths[i];
                int lngth = inputFile.Length;

                if (inputFile[lngth - 1] == 'c' || inputFile[lngth - 2] == 'c')
                {
                    String outputFile = String.Concat(inputFile, ".pdf");
                    Converter converter = new WordConverter();
                    tryConvert(converter, inputFile, outputFile.Replace(".docx", ""));
                }

                Dispatcher.Invoke(() => ConvertProgressBar.Value++);

            }

            if (count > 0)
                MessageBox.Show("Конвертация завершена", "Готово");

            Dispatcher.Invoke(() => ConvertProgressBar.Visibility = Visibility.Hidden);
            Dispatcher.Invoke(() => ConvertLabelText.Content = "Drop or click for convert file(s)");
        }

        private async void ConvertPanel_Drop(object sender, DragEventArgs e)
        {
            ConvertPanel.IsEnabled = false;
            ConvertLabelPanel.Visibility = Visibility.Hidden;
            imgConvert.Visibility = Visibility.Visible;

            await Task.Run(() => ConvertPanelFuncDrop(e));

            ConvertLabelPanel.Visibility = Visibility.Visible;
            imgConvert.Visibility = Visibility.Hidden;
            ConvertPanel.IsEnabled = true;
        }

        private void ConvertPanelFuncClick(MouseButtonEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Файлы docx|*.docx|Файлы doc|*.doc";
            ofd.ShowDialog();

            int count = ofd.FileNames.Length;
            if (count > 0)
            {
                Dispatcher.Invoke(() => ConvertProgressBar.Visibility = Visibility.Visible);
                Dispatcher.Invoke(() => ConvertLabelPanel.Visibility = Visibility.Hidden);
                Dispatcher.Invoke(() => imgConvert.Visibility = Visibility.Visible);
            }
            Dispatcher.Invoke(() => ConvertProgressBar.Maximum = count);
            for (int i = 0; i < count; i++)
            {
                String inputFile = ofd.FileNames[i];
                int lngth = inputFile.Length;
                if (inputFile[lngth - 1] == 'c' || inputFile[lngth - 2] == 'c')
                {
                    String outputFile = String.Concat(inputFile, ".pdf");
                    Converter converter = new WordConverter();
                    tryConvert(converter, inputFile, outputFile);
                }
                Dispatcher.Invoke(() => ConvertProgressBar.Value++);
            }
            if (count > 0) MessageBox.Show("Конвертация завершена", "Готово");

            Dispatcher.Invoke(() => ConvertProgressBar.Visibility = Visibility.Hidden);
            Dispatcher.Invoke(() => ConvertLabelPanel.Visibility = Visibility.Visible);
            Dispatcher.Invoke(() => imgConvert.Visibility = Visibility.Hidden);
            Dispatcher.Invoke(() => ConvertLabelText.Content = "Drop or click for convert file(s)");
        }

        private async void ConvertPanel_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            ConvertPanel.IsEnabled = false;

            await Task.Run(() => ConvertPanelFuncClick(e));
            
            ConvertPanel.IsEnabled = true;
        }

        private void ConvertPanel_DragEnter(object sender, DragEventArgs e)
        {
            Dispatcher.Invoke(() => ConvertLabelText.Content = "Are you sure?");
        }

        private void ConvertPanel_DragLeave(object sender, DragEventArgs e)
        {
            Dispatcher.Invoke(() => ConvertLabelText.Content = "Drop or click for convert file(s)");
        }
    }
}
