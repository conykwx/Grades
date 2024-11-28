using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using Serilog;
using System;
using System.IO;
using System.Text;
using System.Diagnostics;
using System.Threading.Tasks;
using System.Globalization;

class Program
{
    static string connectionString;

    static async Task Main(string[] args)
    {
        // Ініціалізація Serilog для логування
        Log.Logger = new LoggerConfiguration()
            .WriteTo.Console()
            .WriteTo.File(@"C:\Exports\app.log", rollingInterval: RollingInterval.Day) // Логи записуються в папку C:\Exports\
            .CreateLogger();

        // Завантаження конфігурації з appsettings.json
        var configuration = new ConfigurationBuilder()
            .SetBasePath(Directory.GetCurrentDirectory())
            .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
            .Build();

        connectionString = configuration.GetConnectionString("DefaultConnection");

        if (string.IsNullOrEmpty(connectionString))
        {
            Log.Error("Не вдалося знайти рядок підключення в конфігураційному файлі.");
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
            Console.WriteLine("11 - Імпорт даних зі CSV");
            Console.WriteLine("12 - Очистити всіх студентів");
            Console.WriteLine("q - Вихід з програми");
            Console.Write("\nВиберіть опцію: ");
            string choice = Console.ReadLine()?.ToLower();

            if (choice == "q")
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
                        case "11":
                            await ImportStudentsFromCsvAsync(connection);
                            break;
                        case "12":
                            await ClearAllStudentsAsync(connection);
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
                Log.Error(ex, "Помилка при підключенні до бази даних.");
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"Помилка при підключенні до бази даних: {ex.Message}");
                Console.ResetColor();
            }

            Console.WriteLine("\nНатисніть будь-яку клавішу для продовження...");
            Console.ReadKey();
        }

        // Закриваємо логер в кінці програми
        Log.CloseAndFlush();
    }

    static async Task ShowAllDataAsync(SqlConnection connection)
    {
        string query = "SELECT full_name, group_name, avg_grade, min_subject_name, max_subject_name FROM StudentsRating;";
        var command = new SqlCommand(query, connection);

        Stopwatch stopwatch = Stopwatch.StartNew();
        var reader = await command.ExecuteReaderAsync();
        stopwatch.Stop();

        Log.Information("Показ всіх даних почався.");
        Console.ForegroundColor = ConsoleColor.Yellow;
        Console.WriteLine("\n===== Всі дані =====");
        Console.ResetColor();
        while (await reader.ReadAsync())
        {
            Console.WriteLine($"Студент: {reader["full_name"]}, Група: {reader["group_name"]}, Середня оцінка: {reader["avg_grade"]}, Мін. предмет: {reader["min_subject_name"]}, Макс. предмет: {reader["max_subject_name"]}");
        }
        Log.Information($"Час виконання запиту: {stopwatch.Elapsed.TotalSeconds} секунд.");
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

        Log.Information($"Показ студентів з оцінкою більше за {minGrade} почався.");
        Console.ForegroundColor = ConsoleColor.Yellow;
        Console.WriteLine($"\n===== Студенти з оцінкою вище {minGrade} =====");
        Console.ResetColor();
        while (await reader.ReadAsync())
        {
            Console.WriteLine($"Студент: {reader["full_name"]}, Середня оцінка: {reader["avg_grade"]}");
        }
        Log.Information($"Час виконання запиту: {stopwatch.Elapsed.TotalSeconds} секунд.");
        Console.WriteLine($"Час виконання запиту: {stopwatch.Elapsed.TotalSeconds} секунд.");
    }

    static async Task ShowSubjectsWithMinGradesAsync(SqlConnection connection)
    {
        string query = "SELECT DISTINCT subject_name, MIN(grade) AS min_grade FROM StudentGrades GROUP BY subject_name;";
        var command = new SqlCommand(query, connection);

        Stopwatch stopwatch = Stopwatch.StartNew();
        var reader = await command.ExecuteReaderAsync();
        stopwatch.Stop();

        Log.Information("Показ мінімальних оцінок по предметах почався.");
        Console.ForegroundColor = ConsoleColor.Yellow;
        Console.WriteLine("\n===== Мінімальні оцінки по предметах =====");
        Console.ResetColor();
        while (await reader.ReadAsync())
        {
            Console.WriteLine($"Предмет: {reader["subject_name"]}, Мінімальна оцінка: {reader["min_grade"]}");
        }
        Log.Information($"Час виконання запиту: {stopwatch.Elapsed.TotalSeconds} секунд.");
        Console.WriteLine($"Час виконання запиту: {stopwatch.Elapsed.TotalSeconds} секунд.");
    }

    static async Task ShowMinMaxAverageGradeAsync(SqlConnection connection)
    {
        string query = "SELECT MIN(avg_grade) AS min_avg_grade, MAX(avg_grade) AS max_avg_grade FROM StudentsRating;";
        var command = new SqlCommand(query, connection);

        Stopwatch stopwatch = Stopwatch.StartNew();
        var reader = await command.ExecuteReaderAsync();
        stopwatch.Stop();

        Log.Information("Показ мінімальної та максимальної середньої оцінки почався.");
        if (await reader.ReadAsync())
        {
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine("\n===== Мінімальна та максимальна середня оцінка =====");
            Console.ResetColor();
            Console.WriteLine($"Мінімальна середня оцінка: {reader["min_avg_grade"]}");
            Console.WriteLine($"Максимальна середня оцінка: {reader["max_avg_grade"]}");
        }
        Log.Information($"Час виконання запиту: {stopwatch.Elapsed.TotalSeconds} секунд.");
        Console.WriteLine($"Час виконання запиту: {stopwatch.Elapsed.TotalSeconds} секунд.");
    }

    static async Task ShowMathGradesCountAsync(SqlConnection connection)
    {
        string query = "SELECT COUNT(*) AS math_students_count FROM StudentGrades WHERE subject_name = 'Mathematics';";
        var command = new SqlCommand(query, connection);

        Stopwatch stopwatch = Stopwatch.StartNew();
        var reader = await command.ExecuteReaderAsync();
        stopwatch.Stop();

        Log.Information("Показ кількості студентів по математиці почався.");
        if (await reader.ReadAsync())
        {
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine($"\n===== Кількість студентів з оцінками по математиці =====");
            Console.ResetColor();
            Console.WriteLine($"Кількість студентів: {reader["math_students_count"]}");
        }
        Log.Information($"Час виконання запиту: {stopwatch.Elapsed.TotalSeconds} секунд.");
        Console.WriteLine($"Час виконання запиту: {stopwatch.Elapsed.TotalSeconds} секунд.");
    }

    static async Task ShowGroupStatisticsAsync(SqlConnection connection)
    {
        string query = "SELECT group_name, COUNT(*) AS student_count, AVG(avg_grade) AS avg_grade FROM StudentsRating GROUP BY group_name;";
        var command = new SqlCommand(query, connection);

        Stopwatch stopwatch = Stopwatch.StartNew();
        var reader = await command.ExecuteReaderAsync();
        stopwatch.Stop();

        Log.Information("Показ статистики по групах почався.");
        Console.ForegroundColor = ConsoleColor.Yellow;
        Console.WriteLine("\n===== Статистика по групах =====");
        Console.ResetColor();
        while (await reader.ReadAsync())
        {
            Console.WriteLine($"Група: {reader["group_name"]}, Кількість студентів: {reader["student_count"]}, Середня оцінка: {reader["avg_grade"]}");
        }
        Log.Information($"Час виконання запиту: {stopwatch.Elapsed.TotalSeconds} секунд.");
        Console.WriteLine($"Час виконання запиту: {stopwatch.Elapsed.TotalSeconds} секунд.");
    }

    static async Task UpdateStudentGradeAsync(SqlConnection connection)
    {
        Console.Write("\nВведіть ID студента для оновлення оцінки: ");
        if (int.TryParse(Console.ReadLine(), out int studentId))
        {
            Console.Write("Введіть нову оцінку: ");
            if (decimal.TryParse(Console.ReadLine(), out decimal newGrade))
            {
                string query = "UPDATE StudentGrades SET grade = @newGrade WHERE student_id = @studentId;";
                var command = new SqlCommand(query, connection);
                command.Parameters.AddWithValue("@newGrade", newGrade);
                command.Parameters.AddWithValue("@studentId", studentId);

                Stopwatch stopwatch = Stopwatch.StartNew();
                int rowsAffected = await command.ExecuteNonQueryAsync();
                stopwatch.Stop();

                if (rowsAffected > 0)
                {
                    Log.Information($"Оцінка студента {studentId} успішно оновлена.");
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine($"Оцінка студента {studentId} успішно оновлена.");
                }
                else
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("Студент не знайдений або оцінка не була оновлена.");
                }
                Log.Information($"Час виконання запиту: {stopwatch.Elapsed.TotalSeconds} секунд.");
                Console.WriteLine($"Час виконання запиту: {stopwatch.Elapsed.TotalSeconds} секунд.");
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Введена невірна оцінка.");
                Console.ResetColor();
            }
        }
        else
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("Введено невірне ID студента.");
            Console.ResetColor();
        }
    }

    static async Task DeleteStudentAsync(SqlConnection connection)
    {
        Console.Write("\nВведіть ID студента для видалення: ");
        if (int.TryParse(Console.ReadLine(), out int studentId))
        {
            string query = "DELETE FROM StudentsRating WHERE student_id = @studentId;";
            var command = new SqlCommand(query, connection);
            command.Parameters.AddWithValue("@studentId", studentId);

            Stopwatch stopwatch = Stopwatch.StartNew();
            int rowsAffected = await command.ExecuteNonQueryAsync();
            stopwatch.Stop();

            if (rowsAffected > 0)
            {
                Log.Information($"Студент з ID {studentId} успішно видалений.");
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine($"Студент з ID {studentId} успішно видалений.");
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Студент не знайдений.");
            }
            Log.Information($"Час виконання запиту: {stopwatch.Elapsed.TotalSeconds} секунд.");
            Console.WriteLine($"Час виконання запиту: {stopwatch.Elapsed.TotalSeconds} секунд.");
        }
        else
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("Введено невірне ID студента.");
            Console.ResetColor();
        }
    }

    static async Task ExportStudentsToCsvAsync(SqlConnection connection)
    {
        string query = "SELECT full_name, group_name, avg_grade, min_subject_name, max_subject_name FROM StudentsRating;";
        var command = new SqlCommand(query, connection);

        Stopwatch stopwatch = Stopwatch.StartNew();
        var reader = await command.ExecuteReaderAsync();
        stopwatch.Stop();

        Log.Information("Експорт даних до CSV почався.");
        using (var writer = new StreamWriter("students.csv"))
        {
            // Запис заголовка CSV
            await writer.WriteLineAsync("Full Name, Group Name, Average Grade, Min Subject, Max Subject");

            while (await reader.ReadAsync())
            {
                string line = $"{reader["full_name"]}, {reader["group_name"]}, {reader["avg_grade"]}, {reader["min_subject_name"]}, {reader["max_subject_name"]}";
                await writer.WriteLineAsync(line);
            }
        }

        Log.Information($"Експорт завершено. Час виконання запиту: {stopwatch.Elapsed.TotalSeconds} секунд.");
        Console.WriteLine("Дані успішно експортовані до файлу students.csv.");
        Console.WriteLine($"Час виконання запиту: {stopwatch.Elapsed.TotalSeconds} секунд.");
    }

    static async Task ImportStudentsFromCsvAsync(SqlConnection connection)
    {
        // Вкажіть шлях до вашого CSV файлу
        string filePath = "C:\\Exports\\students_import.csv";

        // Перевірка, чи існує файл
        if (!File.Exists(filePath))
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("Файл не знайдено. Перевірте шлях до файлу.");
            Console.ResetColor();
            return;
        }

        try
        {
            // Читання CSV файлу
            var lines = File.ReadAllLines(filePath);

            // Перевірка, чи є в файлі дані (перша лінія - заголовки)
            if (lines.Length <= 1)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("CSV файл порожній або не містить даних.");
                Console.ResetColor();
                return;
            }

            // Ініціалізація команди для вставки даних
            string query = "INSERT INTO StudentsRating (full_name, group_name, avg_grade, min_subject_name, max_subject_name) " +
                           "VALUES (@full_name, @group_name, @avg_grade, @min_subject_name, @max_subject_name)";

            using (var command = new SqlCommand(query, connection))
            {
                // Убедимся, что подключение открыто перед выполнением запроса
                if (connection.State != System.Data.ConnectionState.Open)
                {
                    await connection.OpenAsync();
                }

                // Пропускаємо перший рядок (заголовки)
                foreach (var line in lines.Skip(1))
                {
                    var columns = line.Split(',');

                    // Приведення значень до відповідних типів
                    string fullName = columns[0].Trim();
                    string groupName = columns[1].Trim();
                    decimal avgGrade = decimal.Parse(columns[2].Trim()); // Виправлено: не перевіряється на помилки
                    string minSubjectName = columns[3].Trim();
                    string maxSubjectName = columns[4].Trim();

                    // Додавання параметрів до запиту
                    command.Parameters.Clear();
                    command.Parameters.AddWithValue("@full_name", fullName);
                    command.Parameters.AddWithValue("@group_name", groupName);
                    command.Parameters.AddWithValue("@avg_grade", avgGrade);
                    command.Parameters.AddWithValue("@min_subject_name", minSubjectName);
                    command.Parameters.AddWithValue("@max_subject_name", maxSubjectName);

                    // Виконання запиту на вставку
                    await command.ExecuteNonQueryAsync();
                }
            }

            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("Дані успішно імпортовані з CSV.");
            Console.ResetColor();
        }
        catch (Exception ex)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"Сталася помилка при імпорті даних: {ex.Message}");
            Console.ResetColor();
        }
        finally
        {
            // Закриваємо підключення, якщо воно відкрите
            if (connection.State == System.Data.ConnectionState.Open)
            {
                await connection.CloseAsync();
            }
        }
    }

    static async Task ClearAllStudentsAsync(SqlConnection connection)
    {
        string query = "DELETE FROM StudentsRating;";
        var command = new SqlCommand(query, connection);

        Stopwatch stopwatch = Stopwatch.StartNew();
        int rowsAffected = await command.ExecuteNonQueryAsync();
        stopwatch.Stop();

        if (rowsAffected > 0)
        {
            Log.Information("Всі студенти успішно видалені.");
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("Всі студенти успішно видалені.");
        }
        else
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("Не вдалося очистити дані.");
        }
        Log.Information($"Час виконання запиту: {stopwatch.Elapsed.TotalSeconds} секунд.");
        Console.WriteLine($"Час виконання запиту: {stopwatch.Elapsed.TotalSeconds} секунд.");
    }

    static async Task ChangeDatabaseConnectionStringAsync(IConfiguration configuration)
    {
        Console.Write("\nВведіть новий рядок підключення: ");
        string newConnectionString = Console.ReadLine();

        if (!string.IsNullOrEmpty(newConnectionString))
        {
            connectionString = newConnectionString;

            // Оновлюємо connection string у конфігураційному файлі
            configuration["ConnectionStrings:DefaultConnection"] = newConnectionString;
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("Рядок підключення успішно оновлено.");
            Console.ResetColor();
        }
        else
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("Рядок підключення не може бути порожнім.");
            Console.ResetColor();
        }
    }
}
