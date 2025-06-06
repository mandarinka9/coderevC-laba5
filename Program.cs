using System;
using System.IO;

namespace RetailManagementSystem {
  /// <summary>
  /// Главный класс приложения для управления торговлей
  /// </summary>
  class Program {
    /// <summary>
    /// Точка входа в приложение
    /// </summary>
    static void Main(string[] args) {
      try {
        Console.WriteLine("=== СИСТЕМА УПРАВЛЕНИЯ РОЗНИЧНОЙ ТОРГОВЛЕЙ ===");

        string excelFilePath = GetExcelFilePath();
        string logFilePath = GetLogFilePath();

        using (var db = new Database(excelFilePath)) {
          Console.WriteLine("\nСистема готова к работе!");
          UserInterface.Run(db, logFilePath);
        }

        Console.WriteLine("\nРабота системы завершена. Нажмите любую клавишу...");
        Console.ReadKey();
      }
      catch (Exception ex) {
        Console.WriteLine($"\nКРИТИЧЕСКАЯ ОШИБКА: {ex.Message}");
        Console.WriteLine("Подробности смотрите в лог-файле");
        Console.WriteLine("Нажмите любую клавишу для выхода...");
        Console.ReadKey();
        Environment.Exit(1);
      }
    }

    /// <summary>
    /// Получает путь к файлу Excel с проверкой
    /// </summary>
    private static string GetExcelFilePath() {
      while (true) {
        Console.Write("Введите путь к файлу Excel с данными: ");
        string path = Console.ReadLine();

        if (string.IsNullOrWhiteSpace(path)) {
          Console.WriteLine("Ошибка: путь не может быть пустым");
          continue;
        }

        if (!File.Exists(path)) {
          Console.WriteLine("Ошибка: файл не найден");
          continue;
        }

        if (!path.EndsWith(".xls") && !path.EndsWith(".xlsx")) {
          Console.WriteLine("Ошибка: файл должен быть в формате .xls или .xlsx");
          continue;
        }

        return path;
      }
    }

    /// <summary>
    /// Получает путь к лог-файлу с созданием при необходимости
    /// </summary>
    private static string GetLogFilePath() {
      while (true) {
        Console.Write("Введите путь к файлу для логирования: ");
        string path = Console.ReadLine();

        if (string.IsNullOrWhiteSpace(path)) {
          Console.WriteLine("Ошибка: путь не может быть пустым");
          continue;
        }

        try {
          if (!File.Exists(path)) {
            File.Create(path).Close();
            Console.WriteLine($"Создан новый лог-файл: {path}");
          }
          return path;
        }
        catch (Exception ex) {
          Console.WriteLine($"Ошибка при работе с лог-файлом: {ex.Message}");
        }
      }
    }
  }
}
