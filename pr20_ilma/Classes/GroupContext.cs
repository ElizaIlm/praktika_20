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
    public class GroupContext : Group
    {
        // <summary> Конструктор для контекста rpynn
        // Ссылка: 1
        public GroupContext(int Id, string Name) : base(Id, Name) { }
        // <summary> Получение всех rpynn из БД
        // Ссылка: 1
        public static List<GroupContext> AllGroups()
        {
            // Коллекция rpynn
            List<GroupContext> allGroups = new List<GroupContext>();
            // Открываем соединение
            MySqlConnection connection = Connection.OpenConnection();
            // Выполняем запрос
            MySqlDataReader BDGroups = Connection.Query("SELECT * FROM `group` ORDER BY `Name`", connection);
            // Читаем данные из БД
            while (BDGroups.Read())
            {
                // Добавляем данные в коллекцию
                allGroups.Add(new GroupContext(
                BDGroups.GetInt32(0),
                BDGroups.GetString(1)));
            }
            // Закрываем подключение
            Connection.CloseConnection(connection);
            // Возвращаем группы
            return allGroups;
        }
    }
}
