using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using System;
using System.IO;
using System.Text;
using System.Diagnostics;
using System.Threading.Tasks;
class Program
{
    static string connectionString;

    static async Task Main(string[] args)
    {
        // Завантаження конфігурації з appsettings.json
        var configuration = new ConfigurationBuilder()
            .SetBasePath(Directory.GetCurrentDirectory())
            .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
            .Build();

        connectionString = configuration.GetConnectionString("DefaultConnection");

        if (string.IsNullOrEmpty(connectionString))
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("Не вдалося знайти рядок підключення в конфігураційному файлі.");
            Console.ResetColor();
            return;
        }

        Console.OutputEncoding = Encoding.UTF8;

        while (true)
        {
            Console.Clear();
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine("Меню:");
            Console.ResetColor();
            Console.WriteLine("1 - Показати всі дані");
            Console.WriteLine("2 - Показати студентів з оцінкою вище заданої");
            Console.WriteLine("3 - Показати унікальні предмети з мінімальними оцінками");
            Console.WriteLine("4 - Показати мінімальну та максимальну середню оцінку");
            Console.WriteLine("5 - Показати кількість студентів з оцінками по математиці");
            Console.WriteLine("6 - Показати статистику по групах");
            Console.WriteLine("7 - Оновити оцінку студента");
            Console.WriteLine("8 - Видалити студента");
            Console.WriteLine("9 - Змінити СКБД");
            Console.WriteLine("10 - Експорт даних у CSV");
            Console.WriteLine("exit - Вихід з програми");
            Console.Write("\nВиберіть опцію: ");
            string choice = Console.ReadLine()?.ToLower();

            if (choice == "exit")
            {
                break;
            }

            try
            {
                if (choice == "9")
                {
                    await ChangeDatabaseConnectionStringAsync(configuration);
                    continue;
                }

                using (var connection = new SqlConnection(connectionString))
                {
                    await connection.OpenAsync();

                    switch (choice)
                    {
                        case "1":
                            await ShowAllDataAsync(connection);
                            break;
                        case "2":
                            Console.Write("\nВведіть мінімальну оцінку: ");
                            if (decimal.TryParse(Console.ReadLine(), out decimal minGrade))
                            {
                                await ShowStudentsWithMinGradeGreaterThanAsync(connection, minGrade);
                            }
                            else
                            {
                                Console.ForegroundColor = ConsoleColor.Red;
                                Console.WriteLine("Помилка: введено не число.");
                                Console.ResetColor();
                            }
                            break;
                        case "3":
                            await ShowSubjectsWithMinGradesAsync(connection);
                            break;
                        case "4":
                            await ShowMinMaxAverageGradeAsync(connection);
                            break;
                        case "5":
                            await ShowMathGradesCountAsync(connection);
                            break;
                        case "6":
                            await ShowGroupStatisticsAsync(connection);
                            break;
                        case "7":
                            await UpdateStudentGradeAsync(connection);
                            break;
                        case "8":
                            await DeleteStudentAsync(connection);
                            break;
                        case "10":
                            await ExportStudentsToCsvAsync(connection);
                            break;
                        default:
                            Console.ForegroundColor = ConsoleColor.Red;
                            Console.WriteLine("Некоректний вибір. Спробуйте ще раз.");
                            Console.ResetColor();
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"Помилка при підключенні до бази даних: {ex.Message}");
                Console.ResetColor();
            }

            Console.WriteLine("\nНатисніть будь-яку клавішу для продовження...");
            Console.ReadKey();
        }
    }

    static async Task ShowAllDataAsync(SqlConnection connection)
    {
        string query = "SELECT full_name, group_name, avg_grade, min_subject_name, max_subject_name FROM StudentsRating;";
        var command = new SqlCommand(query, connection);

        Stopwatch stopwatch = Stopwatch.StartNew();
        var reader = await command.ExecuteReaderAsync();
        stopwatch.Stop();

        Console.ForegroundColor = ConsoleColor.Yellow;
        Console.WriteLine("\n===== Всі дані =====");
        Console.ResetColor();
        while (await reader.ReadAsync())
        {
            Console.WriteLine($"Студент: {reader["full_name"]}, Група: {reader["group_name"]}, Середня оцінка: {reader["avg_grade"]}, Мін. предмет: {reader["min_subject_name"]}, Макс. предмет: {reader["max_subject_name"]}");
        }
        Console.WriteLine($"Час виконання запиту: {stopwatch.Elapsed.TotalSeconds} секунд.");
    }

    static async Task ShowStudentsWithMinGradeGreaterThanAsync(SqlConnection connection, decimal minGrade)
    {
        string query = "SELECT full_name, avg_grade FROM StudentsRating WHERE avg_grade > @minGrade;";
        var command = new SqlCommand(query, connection);
        command.Parameters.AddWithValue("@minGrade", minGrade);

        Stopwatch stopwatch = Stopwatch.StartNew();
        var reader = await command.ExecuteReaderAsync();
        stopwatch.Stop();

        Console.ForegroundColor = ConsoleColor.Yellow;
        Console.WriteLine($"\n===== Студенти з оцінкою вище {minGrade} =====");
        Console.ResetColor();
        while (await reader.ReadAsync())
        {
            Console.WriteLine($"Студент: {reader["full_name"]}, Середня оцінка: {reader["avg_grade"]}");
        }
        Console.WriteLine($"Час виконання запиту: {stopwatch.Elapsed.TotalSeconds} секунд.");
    }

    static async Task ShowSubjectsWithMinGradesAsync(SqlConnection connection)
    {
        string query = "SELECT DISTINCT subject_name, MIN(grade) AS min_grade FROM StudentGrades GROUP BY subject_name;";
        var command = new SqlCommand(query, connection);

        Stopwatch stopwatch = Stopwatch.StartNew();
        var reader = await command.ExecuteReaderAsync();
        stopwatch.Stop();

        Console.ForegroundColor = ConsoleColor.Yellow;
        Console.WriteLine("\n===== Мінімальні оцінки по предметах =====");
        Console.ResetColor();
        while (await reader.ReadAsync())
        {
            Console.WriteLine($"Предмет: {reader["subject_name"]}, Мінімальна оцінка: {reader["min_grade"]}");
        }
        Console.WriteLine($"Час виконання запиту: {stopwatch.Elapsed.TotalSeconds} секунд.");
    }

    static async Task ShowMinMaxAverageGradeAsync(SqlConnection connection)
    {
        string query = "SELECT MIN(avg_grade) AS min_avg_grade, MAX(avg_grade) AS max_avg_grade FROM StudentsRating;";
        var command = new SqlCommand(query, connection);

        Stopwatch stopwatch = Stopwatch.StartNew();
        var reader = await command.ExecuteReaderAsync();
        stopwatch.Stop();

        if (await reader.ReadAsync())
        {
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine("\n===== Мінімальна та максимальна середня оцінка =====");
            Console.ResetColor();
            Console.WriteLine($"Мінімальна середня оцінка: {reader["min_avg_grade"]}");
            Console.WriteLine($"Максимальна середня оцінка: {reader["max_avg_grade"]}");
        }
        Console.WriteLine($"Час виконання запиту: {stopwatch.Elapsed.TotalSeconds} секунд.");
    }

    static async Task ShowMathGradesCountAsync(SqlConnection connection)
    {
        string query = "SELECT COUNT(*) AS math_students_count FROM StudentGrades WHERE subject_name = 'Mathematics';";
        var command = new SqlCommand(query, connection);

        Stopwatch stopwatch = Stopwatch.StartNew();
        var reader = await command.ExecuteReaderAsync();
        stopwatch.Stop();

        if (await reader.ReadAsync())
        {
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine("\n===== Кількість студентів по математиці =====");
            Console.ResetColor();
            Console.WriteLine($"Студентів з оцінками по математиці: {reader["math_students_count"]}");
        }
        Console.WriteLine($"Час виконання запиту: {stopwatch.Elapsed.TotalSeconds} секунд.");
    }

    static async Task ShowGroupStatisticsAsync(SqlConnection connection)
    {
        string query = "SELECT group_name, AVG(avg_grade) AS avg_group_grade FROM StudentsRating GROUP BY group_name;";
        var command = new SqlCommand(query, connection);

        Stopwatch stopwatch = Stopwatch.StartNew();
        var reader = await command.ExecuteReaderAsync();
        stopwatch.Stop();

        Console.ForegroundColor = ConsoleColor.Yellow;
        Console.WriteLine("\n===== Статистика по групах =====");
        Console.ResetColor();
        while (await reader.ReadAsync())
        {
            Console.WriteLine($"Група: {reader["group_name"]}, Середня оцінка: {reader["avg_group_grade"]}");
        }
        Console.WriteLine($"Час виконання запиту: {stopwatch.Elapsed.TotalSeconds} секунд.");
    }

    static async Task UpdateStudentGradeAsync(SqlConnection connection)
    {
        Console.Write("\nВведіть ім'я студента: ");
        string studentName = Console.ReadLine();
        Console.Write("Введіть нову оцінку: ");
        if (decimal.TryParse(Console.ReadLine(), out decimal newGrade))
        {
            string query = "UPDATE StudentsRating SET avg_grade = @newGrade WHERE full_name = @studentName;";
            var command = new SqlCommand(query, connection);
            command.Parameters.AddWithValue("@newGrade", newGrade);
            command.Parameters.AddWithValue("@studentName", studentName);

            Stopwatch stopwatch = Stopwatch.StartNew();
            int rowsAffected = await command.ExecuteNonQueryAsync();
            stopwatch.Stop();

            if (rowsAffected > 0)
            {
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine($"\nОцінка оновлена для студента {studentName}.");
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"\nСтудента з ім'ям {studentName} не знайдено.");
            }
            Console.ResetColor();
            Console.WriteLine($"Час виконання запиту: {stopwatch.Elapsed.TotalSeconds} секунд.");
        }
        else
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("Помилка: введено не число.");
            Console.ResetColor();
        }
    }

    static async Task DeleteStudentAsync(SqlConnection connection)
    {
        Console.Write("\nВведіть ім'я студента для видалення: ");
        string studentName = Console.ReadLine();

        string query = "DELETE FROM StudentsRating WHERE full_name = @studentName;";
        var command = new SqlCommand(query, connection);
        command.Parameters.AddWithValue("@studentName", studentName);

        Stopwatch stopwatch = Stopwatch.StartNew();
        int rowsAffected = await command.ExecuteNonQueryAsync();
        stopwatch.Stop();

        if (rowsAffected > 0)
        {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine($"\nСтудент {studentName} видалений.");
        }
        else
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"\nСтудента з ім'ям {studentName} не знайдено.");
        }
        Console.ResetColor();
        Console.WriteLine($"Час виконання запиту: {stopwatch.Elapsed.TotalSeconds} секунд.");
    }

    static async Task ExportStudentsToCsvAsync(SqlConnection connection)
    {
        string query = "SELECT full_name, group_name, avg_grade, min_subject_name, max_subject_name FROM StudentsRating;";
        var command = new SqlCommand(query, connection);

        Stopwatch stopwatch = Stopwatch.StartNew();
        var reader = await command.ExecuteReaderAsync();
        stopwatch.Stop();

        string fileName = Path.Combine("C:\\Exports\\", "students_export.csv");
        Directory.CreateDirectory("Exports");

        using (var writer = new StreamWriter(fileName, false, Encoding.UTF8))
        {
            await writer.WriteLineAsync("Full Name,Group Name,Average Grade,Min Subject,Max Subject");
            while (await reader.ReadAsync())
            {
                string csvRow = $"{reader["full_name"]},{reader["group_name"]},{reader["avg_grade"]},{reader["min_subject_name"]},{reader["max_subject_name"]}";
                await writer.WriteLineAsync(csvRow);
            }
        }

        Console.ForegroundColor = ConsoleColor.Green;
        Console.WriteLine($"\nЕкспорт завершено. Файл збережено за адресою: {fileName}");
        Console.WriteLine($"Час виконання експорту: {stopwatch.Elapsed.TotalSeconds} секунд.");
        Console.ResetColor();
    }

    static async Task ChangeDatabaseConnectionStringAsync(IConfiguration configuration)
    {
        Console.WriteLine("Оберіть тип СКБД:");
        Console.WriteLine("1 - SQL Server");

        string choice = Console.ReadLine();

        if (choice == "1")
        {
            connectionString = configuration.GetConnectionString("DefaultConnection");
            Console.WriteLine("СКБД змінено на SQL Server.");
        }
        else
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("Некоректний вибір.");
            Console.ResetColor();
        }
    }
}
