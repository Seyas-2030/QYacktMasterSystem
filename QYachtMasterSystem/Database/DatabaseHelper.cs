using System;
using System.Collections.Generic;
using System.Data;
using System.Threading;
using Microsoft.Data.Sqlite;

namespace QYachtMaster.Database
{
    public static class DatabaseHelper
    {
        private const int MaxLockRetries = 5;
        private const int LockRetryDelayMs = 100;

        private static string GetConnectionString()
        {
            return DatabaseInitializer.GetConnectionString();
        }

        private static T ExecuteWithRetry<T>(Func<T> operation)
        {
            int attempts = 0;
            while (true)
            {
                try
                {
                    return operation();
                }
                catch (SqliteException ex) when (ex.SqliteErrorCode == 5 || ex.SqliteErrorCode == 6) // SQLITE_BUSY or SQLITE_LOCKED
                {
                    attempts++;
                    if (attempts >= MaxLockRetries) throw;
                    Thread.Sleep(LockRetryDelayMs * attempts);
                }
            }
        }

        public static int ExecuteNonQuery(string query, params SqliteParameter[] parameters)
        {
            return ExecuteWithRetry(() =>
            {
                using (var connection = new SqliteConnection(GetConnectionString()))
                {
                    connection.Open();
                    using (var command = new SqliteCommand(query, connection))
                    {
                        if (parameters != null)
                        {
                            command.Parameters.AddRange(parameters);
                        }
                        return command.ExecuteNonQuery();
                    }
                }
            });
        }

        public static object? ExecuteScalar(string query, params SqliteParameter[] parameters)
        {
            return ExecuteWithRetry(() =>
            {
                using (var connection = new SqliteConnection(GetConnectionString()))
                {
                    connection.Open();
                    using (var command = new SqliteCommand(query, connection))
                    {
                        if (parameters != null)
                        {
                            command.Parameters.AddRange(parameters);
                        }
                        return command.ExecuteScalar();
                    }
                }
            });
        }

        public static List<T> ExecuteQuery<T>(string query, Func<SqliteDataReader, T> mapFunction, params SqliteParameter[] parameters)
        {
            return ExecuteWithRetry(() =>
            {
                var results = new List<T>();
                using (var connection = new SqliteConnection(GetConnectionString()))
                {
                    connection.Open();
                    using (var command = new SqliteCommand(query, connection))
                    {
                        if (parameters != null)
                        {
                            command.Parameters.AddRange(parameters);
                        }
                        using (var reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                results.Add(mapFunction(reader));
                            }
                        }
                    }
                }
                return results;
            });
        }

        public static T? ExecuteSingleQuery<T>(string query, Func<SqliteDataReader, T> mapFunction, params SqliteParameter[] parameters)
        {
            return ExecuteWithRetry(() =>
            {
                using (var connection = new SqliteConnection(GetConnectionString()))
                {
                    connection.Open();
                    using (var command = new SqliteCommand(query, connection))
                    {
                        if (parameters != null)
                        {
                            command.Parameters.AddRange(parameters);
                        }
                        using (var reader = command.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                return mapFunction(reader);
                            }
                        }
                    }
                }
                return default;
            });
        }

        public static void ExecuteTransaction(Action<SqliteConnection, SqliteTransaction> action)
        {
            ExecuteWithRetry<object?>(() =>
            {
                using (var connection = new SqliteConnection(GetConnectionString()))
                {
                    connection.Open();
                    using (var transaction = connection.BeginTransaction())
                    {
                        try
                        {
                            action(connection, transaction);
                            transaction.Commit();
                        }
                        catch (Exception)
                        {
                            transaction.Rollback();
                            throw;
                        }
                    }
                }
                return null;
            });
        }
    }
}
