using pr20_ilma.Classes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
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

namespace pr20_ilma.Pages
{
    /// <summary>
    /// Логика взаимодействия для Main.xaml
    /// </summary>
    public partial class Main : Page
    {
        public List<GroupContext> AllGroups = GroupContext.AllGroups();
        public List<StudentContext> AllStudents = StudentContext.AllStudent();
        public List<WorkContext> AllWorks = WorkContext.AllWorks();
        public List<EvaluationContext> AllEvaluation = EvaluationContext.AllEvaluations();
        public List<DisciplineContext> AllDiscipline = DisciplineContext.AllDisciplines();
        public Main()
        {
            InitializeComponent();
        }
       
        public void CreateGroupUI()
        {
            foreach (GroupContext Group in AllGroups)
                CBGroups.Items.Add(Group.Name);
            CBGroups.Items.Add("Выберите");
            CBGroups.SelectedIndex = CBGroups.Items.Count - 1;
        }
        public void CreateStudents(List<StudentContext> AllStudents)
        {
            Parent.Children.Clear();
            // Перебираем студентов
            foreach (StudentContext Student in AllStudents)
            {
                // Добавляем студентов в список
                Parent.Children.Add(new Items.Student(Student, this));
            }
        }
        private void SelectGroup(object sender, SelectionChangedEventArgs e)
        {
            // Проверяем что в списке выбрана группа, а не элемент "Выберите"
            if (CBGroups.SelectedIndex != CBGroups.Items.Count - 1)
            {
                // Получаем группу
                int IdGroup = AllGroups.Find(x => x.Name == CBGroups.SelectedItem).Id;
                // Создаём студентов, из списка группы
                CreateStudents(AllStudents.FindAll(x => x.IdGroup == IdGroup));
            }
        }
        private void SelectStudents(object sender, System.Windows.Input.KeyEventArgs e)
        {
            // Получаем всех студентов
            List<StudentContext> SearchStudent = AllStudents;
            // Проверяем что в списке выбрана группа, а не элемент "Выберите"
            if (CBGroups.SelectedIndex != CBGroups.Items.Count - 1)
            {
                // Получаем группу
                int IdGroup = AllGroups.Find(x => x.Name == CBGroups.SelectedItem.Id);
                // Фильтруем студентов по группе
                SearchStudent = AllStudents.FindAll(x => x.IdGroup == IdGroup);
            }
            // Сортируем отсортированных студентов, по ФИО
            CreateStudents(SearchStudent.FindAll(x => $"{x.Lastname} {x.Firstname}".Contains(TBFIO.Text)));
        }

        private void ReportGeneration(object sender, RoutedEventArgs e)
        {

        }
    }
}
