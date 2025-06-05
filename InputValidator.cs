using System;
using System.Globalization;

namespace RetailManagementSystem {
    /// <summary>
    /// Класс для валидации и обработки пользовательского ввода
    /// </summary>
    public static class InputValidator {
        /// <summary>
        /// Получает целое число в заданном диапазоне
        /// </summary>
        /// <param name="prompt">Приглашение для ввода</param>
        /// <param name="min">Минимальное допустимое значение</param>
        /// <param name="max">Максимальное допустимое значение</param>
        /// <returns>Введенное целое число</returns>
        public static int GetIntInput(string prompt, int min, int max) {
            while (true) {
                Console.Write(prompt);
                string input = Console.ReadLine();

                if (int.TryParse(input, NumberStyles.Integer, CultureInfo.InvariantCulture, out int result)) {
                    if (result >= min && result <= max) {
                        return result;
                    }
                    Console.WriteLine($"Ошибка: число должно быть от {min} до {max}");
                }
                else {
                    Console.WriteLine("Ошибка: введите целое число");
                }
            }
        }

        /// <summary>
        /// Получает десятичное число
        /// </summary>
        /// <param name="prompt">Приглашение для ввода</param>
        /// <returns>Введенное десятичное число</returns>
        public static decimal GetDecimalInput(string prompt) {
            while (true) {
                Console.Write(prompt);
                string input = Console.ReadLine();

                if (decimal.TryParse(input, NumberStyles.Currency, CultureInfo.InvariantCulture, out decimal result)) {
                    if (result > 0) {
                        return result;
                    }
                    Console.WriteLine("Ошибка: число должно быть больше 0");
                }
                else {
                    Console.WriteLine("Ошибка: введите число");
                }
            }
        }

        /// <summary>
        /// Получает непустую строку
        /// </summary>
        /// <param name="prompt">Приглашение для ввода</param>
        /// <returns>Введенная строка</returns>
        public static string GetNonEmptyString(string prompt) {
            while (true) {
                Console.Write(prompt);
                string input = Console.ReadLine()?.Trim();

                if (!string.IsNullOrWhiteSpace(input)) {
                    return input;
                }
                Console.WriteLine("Ошибка: ввод не может быть пустым");
            }
        }

        /// <summary>
        /// Получает дату
        /// </summary>
        /// <param name="prompt">Приглашение для ввода</param>
        /// <returns>Введенная дата</returns>
        public static DateTime GetDateInput(string prompt) {
            while (true) {
                Console.Write(prompt);
                string input = Console.ReadLine();

                if (DateTime.TryParse(input, CultureInfo.CurrentCulture, DateTimeStyles.None, out DateTime result)) {
                    return result;
                }
                Console.WriteLine($"Ошибка: введите дату в формате {CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern}");
            }
        }

        /// <summary>
        /// Получает ответ да/нет
        /// </summary>
        /// <param name="prompt">Приглашение для ввода</param>
        /// <returns>true если 'да', false если 'нет'</returns>
        public static bool GetYesNoInput(string prompt) {
            while (true) {
                Console.Write(prompt);
                string input = Console.ReadLine()?.Trim().ToLower();

                if (input == "да" || input == "д" || input == "yes" || input == "y") {
                    return true;
                }
                if (input == "нет" || input == "н" || input == "no" || input == "n") {
                    return false;
                }
                Console.WriteLine("Ошибка: введите 'да' или 'нет'");
            }
        }

        /// <summary>
        /// Получает значение из перечисления
        /// </summary>
        /// <typeparam name="T">Тип перечисления</typeparam>
        /// <param name="prompt">Приглашение для ввода</param>
        /// <returns>Введенное значение перечисления</returns>
        public static T GetEnumInput<T>(string prompt) where T : struct, Enum {
            while (true) {
                Console.WriteLine(prompt);
                Console.WriteLine($"Допустимые значения: {string.Join(", ", Enum.GetNames(typeof(T))}");
                Console.Write("Ввод: ");
                string input = Console.ReadLine();

                if (Enum.TryParse<T>(input, true, out T result)) {
                    return result;
                }
                Console.WriteLine($"Ошибка: введите одно из допустимых значений");
            }
        }
    }
}
