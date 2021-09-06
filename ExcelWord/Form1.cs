using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using WordTool = Microsoft.Office.Tools.Word;

namespace ExcelWord
{
    public partial class Form1 : Form
    {
        
        Excel.Application appExcel = new Excel.Application();
        Word.Application appWord = new Word.Application();
        public Form1()
        {
            InitializeComponent();

        }
        List<Mark> marks = new List<Mark>();

        public class Mark // класс записи с таблицы
        {
            public string desRadComp;
            public string description;
            public string documentation;
            public string value;
            public string refDesignator;
            public int count;// количество изделий

            public Mark(string drc, string des, string doc, string rd, string val)
            {
                desRadComp = drc;
                description = des;
                value = val;
                refDesignator = rd;
                documentation = doc;
                count = 1;
            }
        }

        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {


            if (openFileDialog1.ShowDialog() == DialogResult.OK)// проверка на успешное открытие файла диалогом
            {
                Excel.Workbook book = appExcel.Workbooks.Open(openFileDialog1.FileName);
                openFileDialog2.ShowDialog();
                string templatePath = openFileDialog2.FileName;
                Word.Document doc = appWord.Documents.Add(templatePath);
                try
                {
                     // открытие книги экселя
                    
                    Excel.Worksheet tableExcel = book.Worksheets[1]; // выбор таблицы
                    int lastUsedRow = tableExcel.Cells.End[Excel.XlDirection.xlDown].Row;// последняя заполненная строка таблицы
                    Excel.Range range = tableExcel.Range[$"A2:T{lastUsedRow}"]; //вся таблица
                    Excel.Range refDes = tableExcel.Range[$"G2:G{lastUsedRow}"]; 
                    Excel.Range addRange = tableExcel.Range[$"S2:S{lastUsedRow}"]; //доп столбец для сортировки по первой букве
                    Excel.Range sortLevelSecond = tableExcel.Range[$"C2:C{lastUsedRow}"];
                    Excel.Range sortLevelThird = tableExcel.Range[$"H2:H{lastUsedRow}"];
                    addRange.FormulaLocal = $"=ЛЕВСИМВ({refDes.Address[Excel.XlReferenceStyle.xlA1]};1)"; //формула для доп столбца
                    tableExcel.Sort.SortFields.Add(addRange); //Уровни сортировки
                    tableExcel.Sort.SortFields.Add(sortLevelSecond);
                    tableExcel.Sort.SortFields.Add(sortLevelThird);
                    tableExcel.Sort.SetRange(range); //диапазон сортировки
                    tableExcel.Sort.Orientation = Excel.XlSortOrientation.xlSortColumns; //режим сортировки по колонкам
                    tableExcel.Sort.SortMethod = Excel.XlSortMethod.xlPinYin;
                    tableExcel.Sort.Apply();
                    for (int i = 2; i <= lastUsedRow; ++i)// заполнение списка записей из таблицы
                    {
                        marks.Add(new Mark(tableExcel.Range[$"B{i}"].Text, tableExcel.Range[$"C{i}"].Text, tableExcel.Range[$"F{i}"].Text, tableExcel.Range[$"G{i}"].Text, tableExcel.Range[$"H{i}"].Text));
                    }

                    for (int i = 0; i < marks.Count - 1; ++i)// подсчет одинаковых записей и удаление повторений
                    {
                        if (marks[i].description == marks[i + 1].description)
                        {
                            marks[i].count++;
                            marks[i].refDesignator += $", {marks[i + 1].refDesignator}";
                            marks.Remove(marks[i + 1]);
                            i--;
                        }
                    }
                }
                catch (Exception error)
                {
                    appExcel.Quit();
                    throw error;
                }
                try
                {
                    // открытие документа из шаблонов
                    int indexOfTable = 1;
                    int nRow = 5;
                    int indexOfMarks = 0;
                    int position = 8;
                    StringBuilder text = new StringBuilder();
                    while (indexOfMarks <= marks.Count-1)
                    {
                        
                        while (nRow < doc.Tables[indexOfTable].Rows.Count - 5 && indexOfMarks < marks.Count ) //цикл заполнения одной страницы документа
                        {
                            text.Append(doc.Tables[indexOfTable].Cell(nRow, 3).Range.Text = Convert.ToString(position));

                            text.Append(doc.Tables[indexOfTable].Cell(nRow, 5).Range.Text = marks[indexOfMarks].desRadComp);
                            text.Append(doc.Tables[indexOfTable].Cell(nRow + 1, 5).Range.Text = marks[indexOfMarks].description);
                            text.Append(doc.Tables[indexOfTable].Cell(nRow + 2, 5).Range.Text = marks[indexOfMarks].documentation);
                            text.Append(doc.Tables[indexOfTable].Cell(nRow, 6).Range.Text = Convert.ToString(marks[indexOfMarks].count));
                            text.Append(doc.Tables[indexOfTable].Cell(nRow, 7).Range.Text = marks[indexOfMarks].refDesignator);
                            text.Replace("\r", " ");
                            nRow += 4;
                            ++indexOfMarks;
                            position++;

                        }
                        nRow = 5;
                        text.Append(doc.Tables[indexOfTable].Cell(doc.Tables[indexOfTable].Rows.Count-1, doc.Tables[indexOfTable].Columns.Count).Range.Text = Convert.ToString(indexOfTable + 1));// проставление номера листа
                        doc.Range(doc.Content.End - 1).InsertFile(templatePath); // добавление нового листа из шаблона
                        indexOfTable++;// следующая таблица
                        
                        
                        

                    }
                    
                    doc.Save();
                    MessageBox.Show("Success");
                    appWord.Documents.Open(doc.FullName);
                }
                catch (Exception error)
                {
                    appWord.Quit();
                    throw error;
                }

            }
        }
    }
}
