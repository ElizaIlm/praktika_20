using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace pr20_ilma.Classes.Common
{
    public class Connection
    {
        public static string config = "server=127.0.0.1;uid=root;pwd=root;database=journal;";
        public static MySqlConnection OpenConnection()
        {
            // Создаём подключение
            MySqlConnection connection = new MySqlConnection(config);
            // Открываем подключение
            connection.Open();
            // Возвращаем открытое подключение
            return connection;
        }
    public static MySqlDataReader Query(string SQL, MySqlConnection connection)
        {
            return new MySqlCommand(SQL, connection).ExecuteReader();
        }

    public static void CloseConnection(MySqlConnection connection)
        {
            // Закрываем подключение
            connection.Close();
            // Очищаем пул подключений
            MySqlConnection.ClearPool(connection);
        }
    }
}
