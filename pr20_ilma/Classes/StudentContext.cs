using MySql.Data.MySqlClient;
using pr20_ilma.Classes.Common;
using pr20_ilma.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace pr20_ilma.Classes
{
    public class StudentContext : Student
    {
        public StudentContext(int Id, string Firstname, string Lastname, int IdGroup, bool Expelled, DateTime DateExpelled) :
        base(Id, Firstname, Lastname, IdGroup, Expelled, DateExpelled)  { }
        public static List<StudentContext> AllStudent()
        {
            List<StudentContext> allStudent = new List<StudentContext>();
            MySqlConnection connection = Connection.OpenConnection();
            MySqlDataReader BDStudents = Connection.Query("SELECT * FROM `student` ORDER BY `LastName`", connection);
            while (BDStudents.Read())
            {
                allStudent.Add(new StudentContext(
                BDStudents.GetInt32(0),
                BDStudents.GetString(1),
                BDStudents.GetString(2),
                BDStudents.GetInt32(3),
                BDStudents.GetBoolean(4),
                BDStudents.IsDBNull(5) ? DateTime.Now : BDStudents.GetDateTime(5)
                ));
            }
            Connection.CloseConnection(connection);
            return allStudent;
        }
    }
}
