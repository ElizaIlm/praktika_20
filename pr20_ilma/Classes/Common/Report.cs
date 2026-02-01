using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using pr20_ilma.Models;
using pr20_ilma.Pages;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
namespace pr20_ilma.Classes.Common
{
    public class Report
    {
        public static void Group(int IdGroup, Main Main)
        {
            // Создаём диалог для сохранения
            SaveFileDialog SFD = new SaveFileDialog()
            {
                // Указываем начальную директорию
                InitialDirectory = @"C:\",
                // Указываем формат сохранения файла
                Filter = "Excel (*.xlsx)|*.xlsx"
            };
            // Открываем диалоговое окно
            SFD.ShowDialog();
            // Проверяем, если прописано наименование файла
            if (SFD.FileName != "")
            {
                // Получаем группу, о которой сохраняем информацию
                GroupContext Group = Main.AllGroups.Find(x => x.Id == IdGroup);
                // Открываем Excel
                var ExcelApp = new Excel.Application();
                try
                {
                    // Скрываем его видимость
                    ExcelApp.Visible = false;
                    // Добавляем книгу
                    Excel.Workbook Workbook = ExcelApp.Workbooks.Add(Type.Missing);
                    // Получаем активный лист
                    Excel.Worksheet Worksheet = Workbook.ActiveSheet;

                    // Обращаемся к ячейке A1 и указываем текст
                    (Worksheet.Cells[1, 1] as Excel.Range).Value = $"Отчёт о группе {Group.Name}";
                    // Объединяем ячейки A1 и E1
                    Worksheet.Range[Worksheet.Cells[1, 1], Worksheet.Cells[1, 5]].Merge();
                    // Создаём стили для ячейки A1
                    Styles(Worksheet.Cells[1, 1], 18);

                    // Обращаемся к ячейке A3 и указываем текст
                    (Worksheet.Cells[3, 1] as Excel.Range).Value = $"Список группы:";
                    // Объединяем ячейки A3 и E3
                    Worksheet.Range[Worksheet.Cells[3, 1], Worksheet.Cells[3, 5]].Merge();
                    // Создаём стили для ячейки A3
                    Styles(Worksheet.Cells[3, 1], 12, Excel.XlHAlign.xlHAlignLeft);

                    // Обращаемся к ячейке, указываем текст
                    (Worksheet.Cells[4, 1] as Excel.Range).Value = $"ФИО";
                    // Создаём стили для ячейки
                    Styles(Worksheet.Cells[4, 1], 12, XlHAlign.xlHAlignCenter, true);
                    // Указываем ширину
                    (Worksheet.Cells[4, 1] as Excel.Range).ColumnWidth = 35.0f;
                    // Обращаемся к ячейке, указываем текст
                    (Worksheet.Cells[4, 2] as Excel.Range).Value = $"Кол-во не сданных практических";
                    // Создаём стили для ячейки
                    Styles(Worksheet.Cells[4, 2], 12, XlHAlign.xlHAlignCenter, true);

                    // Обращаемся к ячейке, указываем текст
                    (Worksheet.Cells[4, 3] as Excel.Range).Value = $"Кол-во не сданных теоретических";
                    // Создаём стили для ячейки
                    Styles(Worksheet.Cells[4, 3], 12, XlHAlign.xlHAlignCenter, true);

                    // Обращаемся к ячейке, указываем текст
                    (Worksheet.Cells[4, 4] as Excel.Range).Value = $"Отсутствовал на паре";
                    // Создаём стили для ячейки
                    Styles(Worksheet.Cells[4, 4], 12, XlHAlign.xlHAlignCenter, true);

                    // Обращаемся к ячейке, указываем текст
                    (Worksheet.Cells[4, 5] as Excel.Range).Value = $"Опоздал";
                    // Создаём стили для ячейки
                    Styles(Worksheet.Cells[4, 5], 12, XlHAlign.xlHAlignCenter, true);

                    // Указываем ячейку с которой начинаются студенты
                    int Height = 5;
                    // Получаем студентов учащихся в группе
                    List<StudentContext> Students = Main.AllStudents.FindAll(x => x.IdGroup == IdGroup);
                    // Перебираем студентов учащихся в группе
                    foreach (StudentContext Student in Students)
                    {
                        // Получаем дисциплины в которой учится студент
                        List<DisciplineContext> StudentDisciplines = Main.AllDisciplines.FindAll(
                        x => x.IdGroup == Student.IdGroup);

                        // Кол-во практик
                        int PracticeCount = 0;
                        // Кол-во теории
                        int TheoryCount = 0;
                        // Количество пропущенных занятий
                        int AbsenteeismCount = 0;
                        // Кол-во опозданий на занятия
                        int LateCount = 0;

                        // Перебираем дисциплины
                        foreach (DisciplineContext StudentDiscipline in StudentDisciplines)
                        {
                            // Получаем работы студента
                            List<WorkContext> StudentWorks = Main.AllWorks.FindAll(x => x.IdDiscipline == StudentDiscipline.Id);
                            // Перебираем работы студента
                            foreach (WorkContext StudentWork in StudentWorks)
                            {
                                // Получаем оценку за работу
                                EvaluationContext Evaluation = Main.AllEvaluations.Find(x =>
     x.IdWork == StudentWork.Id &&
     x.IdStudent == Student.Id);

                                // Если оценки нет, или она пустая, или равно 2
                                if ((Evaluation != null && (Evaluation.Value.Trim() == "" || Evaluation.Value.Trim() == "2"))
                                    || Evaluation == null ){
                                    // Если практика
                                    if (StudentWork.IdType == 1)
                                        // Считаем не сданную работу
                                        PracticeCount++;
                                    // Если теория
                                    else if (StudentWork.IdType == 2)
                                        // Считаем не сданную работу
                                        TheoryCount++;
                                }
                                // Проверяем что оценка не отсутствует и стоит пропуск
                                if (Evaluation != null && Evaluation.Lateness.Trim() != "")
                                {
                                    // Если пропуск 90 минут
                                    if (Convert.ToInt32(Evaluation.Lateness) == 90)
                                        // Считаем как пропущенную пару
                                        AbsenteeismCount++;
                                }
                                else
                                {
                                    // Считаем как опоздание
                                    LateCount++;
                                }
                            }

                           

                            // Обращаемся к ячейке, указываем текст
                            (Worksheet.Cells[Height, 1] as Excel.Range).Value = $"{Student.Lastname} {Student.Firstname}";
                            // Присваиваем стили
                            Styles(Worksheet.Cells[Height, 1], 12, XlHAlign.xlHAlignLeft, true);
                            // Обращаемся к ячейке, указываем текст
                            (Worksheet.Cells[Height, 2] as Excel.Range).Value = PracticeCount.ToString();
                            // Присваиваем стили
                            Styles(Worksheet.Cells[Height, 2], 12, XlHAlign.xlHAlignCenter, true);
                            // Обращаемся к ячейке, указываем текст
                            (Worksheet.Cells[Height, 3] as Excel.Range).Value = TheoryCount.ToString();
                            // Присваиваем стили
                            Styles(Worksheet.Cells[Height, 3], 12, XlHAlign.xlHAlignCenter, true);
                            // Обращаемся к ячейке, указываем текст
                            (Worksheet.Cells[Height, 4] as Excel.Range).Value = AbsenteeismCount.ToString();
                            // Присваиваем стили
                            Styles(Worksheet.Cells[Height, 4], 12, XlHAlign.xlHAlignCenter, true);
                            // Обращаемся к ячейке, указываем текст
                            (Worksheet.Cells[Height, 5] as Excel.Range).Value = LateCount.ToString();
                            // Присваиваем стили
                            Styles(Worksheet.Cells[Height, 5], 12, XlHAlign.xlHAlignCenter, true);
                            // Увеличиваем высоту
                            Height++;
                        }
                        // Сохраняем книгу
                        Workbook.SaveAs(SFD.FileName);
                        // Закрываем книгу
                        Workbook.Close();
                    } }
                catch (Exception exp) { };

                // Закрываем Excel
                ExcelApp.Quit();

            }
        }
        public static void Styles(Excel.Range Cell,
        int FontSize,
        Excel.XlHAlign Position = Excel.XlHAlign.xlHAlignCenter,
    bool Border = false)
        {
            // Присваиваем шрифт
            Cell.Font.Name = "Bahnschrift Light Condensed";
            // Присваиваем размер
            Cell.Font.Size = FontSize;
            // Указываем вертикальное центрирование
            Cell.HorizontalAlignment = Position;
            // Указываем горизонтальное центрирование
            Cell.VerticalAlignment = Excel.XlHAlign.xlHAlignCenter;
            // Если границы
            if (Border)
            {
                // Получаем границу ячейки
                Excel.Borders border = Cell.Borders;
                // Задаём стиль линии
                border.LineStyle = Excel.XlLineStyle.xlDouble;
                // Задаём ширину линии
                border.Weight = XlBorderWeight.xlThin;
                // Включаем перенос текста
                Cell.WrapText = true;
            }
        }
    }
}
