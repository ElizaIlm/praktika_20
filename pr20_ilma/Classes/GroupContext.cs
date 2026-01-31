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
        public GroupContext(int Id, string Name) : base(Id, Name) { }
        public static List<GroupContext> AllGroups()
        {
            List<GroupContext> allGroups = new List<GroupContext>();
            MySqlConnection connection = Connection.OpenConnection();
            MySqlDataReader BDGroups = Connection.Query("SELECT * FROM `group` ORDER BY `Name`", connection);
            while (BDGroups.Read())
            {
                allGroups.Add(new GroupContext(
                BDGroups.GetInt32(0),
                BDGroups.GetString(1)));
            }
            Connection.CloseConnection(connection);
            return allGroups;
        }
    }
}
