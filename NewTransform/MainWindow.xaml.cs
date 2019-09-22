using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
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
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace NewTransform
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        //Пути к файлам
        static string pathDoc = "";
        static string pathXls = "";
        static string pathSave = "";

        //Свойства кнопки "Сгенерировать"
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
        }

        private void XlsxDropPanel_Drop(object sender, DragEventArgs e)
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
        }

        private void DocDropPanel_Drop(object sender, DragEventArgs e)
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
        }

        private void Generate()
        {
            var WordApp = new Word.Application();
            WordApp.Visible = false;

            //Открыть ворд && открыть эксель, получить 1 лист
            var WordDocument = WordApp.Documents.Open(@"" + pathDoc);
            Excel.Application ObjWorkExcel = new Excel.Application();
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(@"" + pathXls);
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];

            //Получить кол-во заполненных ячеек эксель
            int nInLastRow = ObjWorkSheet.Cells.Find("*", System.Reflection.Missing.Value,
                System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious,
                false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;
            int nInLastCol = ObjWorkSheet.Cells.Find("*", System.Reflection.Missing.Value,
                System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlPrevious,
                false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Column;

            string[,] list = new string[nInLastRow, nInLastCol]; //Равен по размеру листу

            //Данные с листа в массив
            for (int i = 0; i < nInLastRow; i++)
            {
                for (int j = 0; j < nInLastCol; j++)
                {
                    list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text.ToString();

                }
            }

            //Данные из массива в ворд
            try
            {
                for (int i = 0; i < nInLastRow; i++)
                {
                    for (int j = 0; j < nInLastCol; j++)
                    {

                        string tmp = list[i, j]; //Значение текущей ячейки в переменную
                        ReplaceWordStub("{" + j + "}", tmp, WordDocument); //Меняем метку в шаблоне ворд на значение

                        //Сохранить, когда будет обрабатываться последняя ячейка в строке
                        if (j == nInLastCol - 1)
                        {
                            WordDocument.SaveAs(@"" + pathSave + "\\" + list[i, j] + ".docx");
                            //Закрыть, иначе будет множество экземпляров ворд в фоне, открыть шаблон снова для следующей строки
                            WordDocument.Close(false, Type.Missing, Type.Missing);
                            WordDocument = WordApp.Documents.Open(@"" + pathDoc);
                        }

                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Что-то пошло не так", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                //Закрыть после завершения работы
                WordDocument.Close(false, Type.Missing, Type.Missing);
                WordApp.Quit();
                ObjWorkBook.Close(false, Type.Missing, Type.Missing);
                ObjWorkExcel.Quit();
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

        private void XlsxPanelDropPanel_Click(object sender, RoutedEventArgs e)
        {
            xlsxLabelTitle.Content = "Test";
        }

        static void ReplaceWordStub(string StubToReplace, string Text, Word.Document WordDocument)
        {

            var Range = WordDocument.Content;
            Range.Find.Execute(FindText: StubToReplace, ReplaceWith: Text);

        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            GC.Collect();
        }
        
        private void XlsxDropPanel_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            /*
            OpenFileDialog ofd = new OpenFileDialog();
            
            ofd.Filter = "Файлы xlsx|*.xlsx";
            ofd.ShowDialog();

            pathXls = ofd.FileName;
            //Если все пути получены: включаем кнопку "Сгенерировать"
            if (pathXls != "" && pathDoc != "" && pathSave != "")
            {
                Properties(generatedButton);
            }
            */
        }

        private void MenuItem_Click(object sender, RoutedEventArgs e)
        {
            
        }
    }
}
