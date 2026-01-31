using pr20_ilma.Classes;
using pr20_ilma.Pages;
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

namespace pr20_ilma.Items
{
    /// <summary>
    /// Логика взаимодействия для Student.xaml
    /// </summary>
    public partial class Student : UserControl
    {
        public StudentContext student;
        public MainWindow main;
       
        public Student(StudentContext student, MainWindow main)
        {
            InitializeComponent();
           
            this.main = main;
        }

        public Student(StudentContext student, Main main1)
        {
            this.student = student;
          
        }
    }
}
