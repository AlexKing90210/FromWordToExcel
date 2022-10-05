using System;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;



namespace FromWordToExcel
{

    class WordReader
    {
        private string line;
        private string pathFile;
        private Errorhandler errorHandler;

        public string Line
        {
            get { return line; }
            set { line = value; }
        }

        public string PathFile
        {
            get { return pathFile; }
            set { pathFile = value; }
        }

        public void Read(string path)
        {
            // Переменная для чтения
            string parText = "";

            //Заполнения пути файла
            this.pathFile = Path.GetDirectoryName(path);

            // Подключение к Word
            Word.Application app = new Word.Application();
            Object fileName = path;
            app.Documents.Open(ref fileName, ReadOnly: true);
            Word.Document doc = app.ActiveDocument;

            // Считывание документа
            for (int i = 1; i <= doc.Paragraphs.Count; i++)
            {
                parText += doc.Paragraphs[i].Range.Text;
            }
            app.Quit();

            // Присвоение текста в поле класса
            this.line = parText;

            Console.WriteLine("Считывание завершено успешно!");
            
        }
    }

    class ExcelWriter
    {
        private WordReader word_text;

        public WordReader Word_Text
        {
            get { return word_text; }
            set { word_text = value; }
        }

        public void ExportToExcel()
        {
            // Загрузить Excel, затем создать новую пустую рабочую книгу
            Excel.Application excelApp = new Excel.Application();

            // Сделать приложение Excel видимым
            excelApp.Visible = true;
            excelApp.Workbooks.Add();
            Excel._Worksheet workSheet = excelApp.ActiveSheet;

            // Запись в эксель файл
            workSheet.Cells[2, "B"] = this.word_text.Line;

            excelApp.DisplayAlerts = false;

            // Сохранение файла
            string nameFile = DateTime.Now.ToString("dd-MMMM-yyyy-HH-mm-ss");

            workSheet.SaveAs(string.Format(@"{0}\{1}.xlsx", this.word_text.PathFile, nameFile));

            excelApp.Quit();

            Console.WriteLine("Запись занесена в excel файл {0}", nameFile);
        }
    }

    class Errorhandler
    {
        string[] extTrueMas = { ".doc", ".docx" };
        public int FileExists(string path)
        {
            if(File.Exists(path))
            {
                return 0;
            }
            
            Console.WriteLine("По указанному пути файла для считывания не обнаружено! Попробуйте еще раз!");
            return -1;
            
        }

        public int FileExtension(string path)
        {
            string ext = Path.GetExtension(path);
            if(extTrueMas.Contains(ext))
            {
                return 0;
            }
            Console.WriteLine("По указанному пути указан файл неверного расширения! Используйте расширение для файлов Microsoft Word!");
            return -1;
            
        }

    }
    

    class Program
    {
        static void Main()
        {
            

            while(true)
            {
                Console.WriteLine("Введите путь до файла:");

                string path = Console.ReadLine();

                Errorhandler errorhandler = new Errorhandler();
                

                if(errorhandler.FileExists(path) == 0 & errorhandler.FileExtension(path) == 0)
                {
                    WordReader reader = new WordReader();

                    reader.Read(path);
                    ExcelWriter writer = new ExcelWriter();
                    writer.Word_Text = reader;

                    writer.ExportToExcel();

                    Console.WriteLine();

                }
                else
                {
                    Console.WriteLine("Указаны некорректные данные! Попробуйте еще раз!");
                    
                }

                if (Console.ReadKey().Key == ConsoleKey.Escape)
                {
                    break;
                }

            }

        }

    }

}
