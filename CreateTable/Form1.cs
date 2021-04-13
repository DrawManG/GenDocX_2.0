using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using Spire.Doc;

namespace CreateTable
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Переменные нужные для работы циклов
            int sar = 1;
            int cut = 8;
            int time = 0;
            int and2 = 2;
            //Оригинальное название файла на месяц
            DateTime NAMEFILE = new DateTime(DateTime.Today.Year, DateTime.Today.Month + 1, 1);
            // Создание DOCX отчёта (объявление)
            Document document = new Document();
            // Создание секции
            Section section = document.AddSection();
            // Создание таблицу 1 с линиями 
            Table table1 = section.AddTable(true);
            //Отделение первой от второй таблицы
            section.AddParagraph();
            // Создание таблицу 2 с линиями
            Table table2 = section.AddTable(true);
            // Статическая информация в таблице
            string[] TIME_BASE = new string[] { "10:00 - 11:00", "11:00 - 12:00", "12:00 - 13:00", "13:00 - 13:30", "13:30 - 14:30", "14:30 - 16:00", "16:00 - 17:00", "17:00 - 18:00" };
            string[] OVER_PLANE = new string[] { "Текущие объекты", "Изучение инструментария", "Разработка софта" };
            // Подсчёт дней в месяце 
            int Day_in_Mounth = System.DateTime.DaysInMonth(DateTime.Today.Year, DateTime.Today.Month+1);
            // Создание лист дней в месяце выше
            List<string> Days = new List<string>();
            // Добавление даты в лист дней (ДЕНЬ,МЕСЯЦ,ГОД)
            while (sar <= Day_in_Mounth + 1)
            {
                Days.Add(sar.ToString() + "." + Convert.ToInt16(DateTime.Today.Month + 1) + "." + DateTime.Today.Year);
                sar++;
            }
            // Создание первой информационной таблицы из 3ёх ячеек 
            table1.ResetCells(1, 2);
            table1[0, 0].AddParagraph().AppendText(Convert.ToString("Дата"));
            table1[0, 0].Width = 200;
            table1[0, 1].SplitCell(2, 0);
            table1[0, 1].AddParagraph().AppendText(Convert.ToString("Время"));
            table1[0, 1].Width = 230;
            table1[0, 2].AddParagraph().AppendText(Convert.ToString("Действие"));
            table1[0, 2].Width = 270;

            //Создание таблиц по количеству дней в месяце в два столбца
            table2.ResetCells(Day_in_Mounth, 2);
            // Создание уже нормального отчёта, разделение 2-ой ячейки ещё на 2 и их маштабирование
            for (int i = 0; i < Day_in_Mounth; i++)
            {
                table2[i * 9, 0].Width = 80;
                table2[i * 9, 1].Width = 250;
                table2[i * 9, 1].SplitCell(2, 9);
                table2[i * 9, 0].AddParagraph().AppendText(Convert.ToString(Days[i]));
                // Последняя ячейка будет с пробелом (отделять дни)
                if (i != 0)
                {
                    TableCell Flow = table2[i * 9 - 1, 0];
                    Flow.CellFormat.Borders.Left.Color = Color.White;
                }
                table2.ApplyHorizontalMerge(cut, 1, 2);
                cut = cut + 9;
            }
            sar = 0;

            // Суета с временем, вставка её на свои места
            for (int i = 0; i < Day_in_Mounth * 8; i++)
            {
                table2[sar, 1].AddParagraph().AppendText(Convert.ToString(TIME_BASE[time]));
                if (time == 7)
                {
                    time = 0;
                    sar = i + and2;
                    and2++;
                }
                else
                {
                    time++;
                    sar++;
                }
            }
            sar = 0;
            and2 = 8;
            //Вставка 3-и последних статических данных + ввод обед на каждую неделю
            while (and2 < Day_in_Mounth * 9)
            {
                table2[and2 + sar, 1].SplitCell(3, 2);
                table2[and2 + sar, 1].AddParagraph().AppendText(Convert.ToString(OVER_PLANE[0]));
                table2[and2 + sar, 2].AddParagraph().AppendText(Convert.ToString(OVER_PLANE[1]));
                table2[and2 + sar, 3].AddParagraph().AppendText(Convert.ToString(OVER_PLANE[2]));
                table2[and2 + sar - 5, 2].AddParagraph().AppendText("Обед");
                and2 = and2 + 9;
                sar = sar + 1;
            }
            //Поднастройка маштабов ячеек
            table2[0, 2].Width = 300;
            table1[0, 0].Width = table2[0, 0].Width - 15;
            // Сохранение на рабочий стол со статическим именем
            document.SaveToFile(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + "\\" + "Отчёт за " + NAMEFILE.ToString("MMMM") + ".docx", FileFormat.Docx);
            // Автозакрытие программы по завершении
            Application.Exit();
        }
    }
}
// Create by DrawManG.