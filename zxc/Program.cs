using System;
using System.Data;
using System.Data.OleDb;

class Program
{
    
    

        static string databasePath = @"C:\Users\ido-class2\Desktop\qwe\zxc.accdb";
        static string connectionString = $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={databasePath};";
        static void Main(string[] args)
    {
            // Пример добавления слова
            AddWord("Привет");

            // Пример изменения слова
            UpdateWord("Привет", "Здравствуйте");

            // Пример удаления слова
            RemoveWord("Здравствуйте");
        }

        static void AddWord(string word)
    {
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                try
                {
                    connection.Open();
                    Console.WriteLine("Соединение успешно установлено.");

                    string query = "INSERT INTO Таблица1(Word) VALUES (@word)";
                    using (OleDbCommand command = new OleDbCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@word", word);
                        int rowsAffected = command.ExecuteNonQuery();
                        Console.WriteLine($"{rowsAffected} слово(а) добавлено(ы).");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Ошибка: " + ex.Message);
                }
                finally
                {
                    connection.Close();
                    Console.WriteLine("Соединение закрыто.");
                }
            }
        }

        static void UpdateWord(string oldWord, string newWord)
        {
        using (OleDbConnection connection = new OleDbConnection(connectionString))
        {
            try
            {
                connection.Open();
                Console.WriteLine("Соединение успешно установлено.");

                string query = "UPDATE Таблица1 SET WordColumn = @newWord WHERE WordColumn = @oldWord";
                using (OleDbCommand command = new OleDbCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@newWord", newWord);
                    command.Parameters.AddWithValue("@oldWord", oldWord);
                    int rowsAffected = command.ExecuteNonQuery();
                    Console.WriteLine($"{rowsAffected} слово(а) изменено(ы).");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Ошибка: " + ex.Message);
            }
            finally
            {
                connection.Close();
                Console.WriteLine("Соединение закрыто.");
            }
        }
    }

    static void RemoveWord(string word)
    {
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                try
                {
                    connection.Open();
                    Console.WriteLine("Соединение успешно установлено.");

                    string query = "DELETE FROM Таблица1 WHERE WordColumn = @word";
                    using (OleDbCommand command = new OleDbCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@word", word);
                        int rowsAffected = command.ExecuteNonQuery();
                        Console.WriteLine($"{rowsAffected} слово(а) удалено(ы).");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Ошибка: " + ex.Message);
                }
                finally
                {
                    connection.Close();
                    Console.WriteLine("Соединение закрыто.");
                }
            }
        Console.ReadKey();
        }
    }
