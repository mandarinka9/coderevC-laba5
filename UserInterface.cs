using System;

namespace RetailManagementSystem {
    /// <summary>
    /// Класс для взаимодействия с пользователем через консольный интерфейс
    /// </summary>
    public static class UserInterface {
        /// <summary>
        /// Основной цикл работы с пользователем
        /// </summary>
        /// <param name="db">Экземпляр базы данных</param>
        /// <param name="logFilePath">Путь к файлу лога</param>
        public static void Run(Database db, string logFilePath) {
            try {
                bool isRunning = true;

                while (isRunning) {
                    try {
                        ShowMainMenu(db, logFilePath, ref isRunning);
                    }
                    catch (Exception ex) {
                        Console.WriteLine($"Ошибка: {ex.Message}");
                        db.LogAction(logFilePath, $"Ошибка в основном меню: {ex.Message}");
                    }
                }
            }
            catch (Exception ex) {
                Console.WriteLine($"Критическая ошибка: {ex.Message}");
                db.LogAction(logFilePath, $"Критическая ошибка: {ex.Message}");
            }
        }

        private static void ShowMainMenu(Database db, string logFilePath, ref bool isRunning) {
            Console.WriteLine("\n=== ГЛАВНОЕ МЕНЮ ===");
            Console.WriteLine("1. Движение товаров");
            Console.WriteLine("2. Товары");
            Console.WriteLine("3. Категории");
            Console.WriteLine("4. Магазины");
            Console.WriteLine("5. Выполнить запрос");
            Console.WriteLine("6. Выход");

            int choice = InputValidator.GetIntInput("Выберите пункт меню: ", 1, 6);

            switch (choice) {
                case 1: ShowTableMenu(db, logFilePath, 1, "Движение товаров"); break;
                case 2: ShowTableMenu(db, logFilePath, 2, "Товары"); break;
                case 3: ShowTableMenu(db, logFilePath, 3, "Категории"); break;
                case 4: ShowTableMenu(db, logFilePath, 4, "Магазины"); break;
                case 5: db.ExecuteQuery(logFilePath); break;
                case 6: isRunning = false; break;
            }
        }

        private static void ShowTableMenu(Database db, string logFilePath, int sheetNum, string tableName) {
            bool backToMain = false;

            while (!backToMain) {
                try {
                    Console.WriteLine($"\n=== МЕНЮ {tableName.ToUpper()} ===");
                    Console.WriteLine("1. Просмотреть данные");
                    Console.WriteLine("2. Добавить запись");
                    Console.WriteLine("3. Редактировать запись");
                    Console.WriteLine("4. Удалить запись");
                    Console.WriteLine("5. Вернуться в главное меню");

                    int choice = InputValidator.GetIntInput("Выберите действие: ", 1, 5);

                    switch (choice) {
                        case 1: db.ViewDatabase(logFilePath, sheetNum); break;
                        case 2: db.AddElement(logFilePath, sheetNum); break;
                        case 3: db.EditElement(logFilePath, sheetNum); break;
                        case 4: db.DeleteElement(logFilePath, sheetNum); break;
                        case 5: backToMain = true; break;
                    }
                }
                catch (Exception ex) {
                    Console.WriteLine($"Ошибка: {ex.Message}");
                    db.LogAction(logFilePath, $"Ошибка в меню {tableName}: {ex.Message}");
                }
            }
        }
    }
}
