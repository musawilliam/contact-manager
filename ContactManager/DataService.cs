using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.IO;
using System.Windows.Forms;

namespace ContactManager
{
    public class DataService
    {
        private readonly string _connectionString;
        private readonly string _dbPath;

        public DataService()
        {
            _dbPath = Path.Combine(Application.StartupPath, "contacts.db");
            _connectionString = $"Data Source={_dbPath};Version=3;";
            InitializeDatabase();
        }

        private void InitializeDatabase()
        {
            if (!File.Exists(_dbPath))
            {
                SQLiteConnection.CreateFile(_dbPath);
            }

            using (var connection = new SQLiteConnection(_connectionString))
            {
                connection.Open();
                string createTableQuery = @"
                    CREATE TABLE IF NOT EXISTS Contacts (
                        Id INTEGER PRIMARY KEY AUTOINCREMENT,
                        Name TEXT NOT NULL,
                        Surname TEXT,
                        Number TEXT,
                        Used BOOLEAN DEFAULT 0,
                        CreatedDate DATETIME DEFAULT CURRENT_TIMESTAMP
                    )";
                
                using (var command = new SQLiteCommand(createTableQuery, connection))
                {
                    command.ExecuteNonQuery();
                }
            }
        }

        public List<Contact> GetAllContacts()
        {
            var contacts = new List<Contact>();
            
            using (var connection = new SQLiteConnection(_connectionString))
            {
                connection.Open();
                string query = "SELECT * FROM Contacts ORDER BY Name";
                
                using (var command = new SQLiteCommand(query, connection))
                using (var reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        contacts.Add(new Contact
                        {
                            Id = Convert.ToInt32(reader["Id"]),
                            Name = reader["Name"].ToString() ?? "",
                            Surname = reader["Surname"].ToString() ?? "",
                            Number = reader["Number"].ToString() ?? "",
                            Used = Convert.ToBoolean(reader["Used"]),
                            CreatedDate = Convert.ToDateTime(reader["CreatedDate"])
                        });
                    }
                }
            }
            
            return contacts;
        }

        public void AddContact(Contact contact)
        {
            using (var connection = new SQLiteConnection(_connectionString))
            {
                connection.Open();
                string query = @"
                    INSERT INTO Contacts (Name, Surname, Number, Used, CreatedDate)
                    VALUES (@Name, @Surname, @Number, @Used, @CreatedDate)";
                
                using (var command = new SQLiteCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@Name", contact.Name);
                    command.Parameters.AddWithValue("@Surname", contact.Surname);
                    command.Parameters.AddWithValue("@Number", contact.Number);
                    command.Parameters.AddWithValue("@Used", contact.Used);
                    command.Parameters.AddWithValue("@CreatedDate", contact.CreatedDate);
                    command.ExecuteNonQuery();
                }
            }
        }

        public void UpdateContact(Contact contact)
        {
            using (var connection = new SQLiteConnection(_connectionString))
            {
                connection.Open();
                string query = @"
                    UPDATE Contacts 
                    SET Name = @Name, Surname = @Surname, Number = @Number, Used = @Used
                    WHERE Id = @Id";
                
                using (var command = new SQLiteCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@Id", contact.Id);
                    command.Parameters.AddWithValue("@Name", contact.Name);
                    command.Parameters.AddWithValue("@Surname", contact.Surname);
                    command.Parameters.AddWithValue("@Number", contact.Number);
                    command.Parameters.AddWithValue("@Used", contact.Used);
                    command.ExecuteNonQuery();
                }
            }
        }

        public void DeleteContact(int id)
        {
            using (var connection = new SQLiteConnection(_connectionString))
            {
                connection.Open();
                string query = "DELETE FROM Contacts WHERE Id = @Id";
                
                using (var command = new SQLiteCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@Id", id);
                    command.ExecuteNonQuery();
                }
            }
        }

        public void ClearAllContacts()
        {
            using (var connection = new SQLiteConnection(_connectionString))
            {
                connection.Open();
                string query = "DELETE FROM Contacts";
                
                using (var command = new SQLiteCommand(query, connection))
                {
                    command.ExecuteNonQuery();
                }
            }
        }
    }
}