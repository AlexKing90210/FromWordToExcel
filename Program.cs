using System;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;



namespace FromWordToExcel
{
    
    class WordReader
    {
        private string line;

        public string Line
        {
            get { return line; }
            set { line = value; }
        }

        public void Read(string path)
        {

            // Переменная для чтения
            string parText = "";

            // Подключение к Word
            Word.Application app = new Word.Application();
            Object fileName = path;
            app.Documents.Open(ref fileName);
            Word.Document doc = app.ActiveDocument;
            
            // Считывание документа
            for (int i = 1; i < doc.Paragraphs.Count; i++)
            {
                parText += doc.Paragraphs[i].Range.Text;
            }
            app.Quit();


            this.line = parText;

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

            workSheet.Cells[2, "B"] = this.word_text.Line;

            excelApp.DisplayAlerts = false;

            // Сохранение файла
            string nameFile = DateTime.Now.ToString("dd-MMMM-yyyy-HH-mm-ss");

            workSheet.SaveAs(string.Format(@"{0}\{1}.xlsx", Environment.CurrentDirectory, nameFile));

            excelApp.Quit();

        }
        
    }

    class Program
    {
        static void Main()
        {
            Console.WriteLine("Введите путь до файла:");

            string path = Console.ReadLine();

            WordReader reader = new WordReader();

            reader.Read(path);

            ExcelWriter writer = new ExcelWriter();
            writer.Word_Text = reader;

            writer.ExportToExcel();


        }


    }




}
