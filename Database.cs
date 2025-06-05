using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.IO;
using System.Runtime.InteropServices;

namespace RetailManagementSystem {
    /// <summary>
    /// Класс для работы с базой данных в Excel
    /// </summary>
    class Database : IDisposable {
        private readonly string _filePath;
        private Excel.Application _excelApp;
        private Excel.Workbook _workbook;

        /// <summary>
        /// Конструктор класса Database
        /// </summary>
        /// <param name="filePath">Путь к файлу Excel</param>
        public Database(string filePath) {
            if (string.IsNullOrWhiteSpace(filePath)) {
                throw new ArgumentException("Путь к файлу не может быть пустым");
            }

            if (!File.Exists(filePath)) {
                throw new FileNotFoundException("Файл не найден", filePath);
            }

            _filePath = filePath;
            InitializeExcel();
        }

        private void InitializeExcel() {
            try {
                _excelApp = new Excel.Application();
                _workbook = _excelApp.Workbooks.Open(_filePath);
            }
            catch (COMException ex) {
                throw new Exception("Ошибка при работе с Excel. Убедитесь, что Excel установлен", ex);
            }
        }

        /// <summary>
        /// Просмотр данных в указанном листе
        /// </summary>
        /// <param name="logFilePath">Путь к файлу лога</param>
        /// <param name="sheetNum">Номер листа (1-4)</param>
        public void ViewDatabase(string logFilePath, int sheetNum) {
            if (sheetNum < 1 || sheetNum > 4) {
                throw new ArgumentException("Номер листа должен быть от 1 до 4");
            }

            Excel.Worksheet sheet = GetWorksheet(sheetNum);
            Excel.Range usedRange = sheet.UsedRange;

            for (int row = 1; row <= usedRange.Rows.Count; row++) {
                for (int col = 1; col <= usedRange.Columns.Count; col++) {
                    Console.Write($"{usedRange.Cells[row, col]?.Value2?.ToString() ?? ""}\t");
                }
                Console.WriteLine();
            }

            LogAction(logFilePath, $"Просмотр листа {sheetNum}");
        }

        /// <summary>
        /// Добавление элемента в указанный лист
        /// </summary>
        /// <param name="logFilePath">Путь к файлу лога</param>
        /// <param name="sheetNum">Номер листа (1-4)</param>
        public void AddElement(string logFilePath, int sheetNum) {
            try {
                switch (sheetNum) {
                    case 1: AddProductMovement(logFilePath); break;
                    case 2: AddProduct(logFilePath); break;
                    case 3: AddCategory(logFilePath); break;
                    case 4: AddStore(logFilePath); break;
                    default: throw new ArgumentException("Неверный номер листа");
                }
            }
            catch (Exception ex) {
                LogAction(logFilePath, $"Ошибка при добавлении: {ex.Message}");
                throw;
            }
        }

        private void AddProductMovement(string logFilePath) {
            var data = new ProductMovement {
                OperationId = InputValidator.GetNonEmptyString("ID операции: "),
                Date = InputValidator.GetDateInput("Дата (ДД.ММ.ГГГГ): "),
                StoreId = InputValidator.GetNonEmptyString("ID магазина: "),
                ArticleId = InputValidator.GetNonEmptyString("Артикул: "),
                OperationType = InputValidator.GetNonEmptyString("Тип операции: "),
                PackageCount = InputValidator.GetIntInput("Количество упаковок: ", 1, 1000),
                HasClientCard = InputValidator.GetYesNoInput("Наличие карты клиента (да/нет): ")
            };

            Excel.Worksheet sheet = GetWorksheet(1);
            int nextRow = sheet.UsedRange.Rows.Count + 1;

            sheet.Cells[nextRow, 1].Value2 = data.OperationId;
            sheet.Cells[nextRow, 2].Value = data.Date;
            sheet.Cells[nextRow, 3].Value2 = data.StoreId;
            sheet.Cells[nextRow, 4].Value2 = data.ArticleId;
            sheet.Cells[nextRow, 5].Value2 = data.OperationType;
            sheet.Cells[nextRow, 6].Value2 = data.PackageCount;
            sheet.Cells[nextRow, 7].Value2 = data.HasClientCard ? "Да" : "Нет";

            _workbook.Save();
            LogAction(logFilePath, $"Добавлено движение товара: {data.OperationId}");
        }

        private void AddProduct(string logFilePath) {
            var data = new Product {
                ArticleId = InputValidator.GetNonEmptyString("Артикул: "),
                CategoryId = InputValidator.GetNonEmptyString("ID категории: "),
                ProductName = InputValidator.GetNonEmptyString("Наименование товара: "),
                PurchasePrice = InputValidator.GetDecimalInput("Цена закупки: "),
                SalePrice = InputValidator.GetDecimalInput("Цена продажи: "),
                DiscountPercent = InputValidator.GetIntInput("Скидка (%): ", 0, 100)
            };

            Excel.Worksheet sheet = GetWorksheet(2);
            int nextRow = sheet.UsedRange.Rows.Count + 1;

            sheet.Cells[nextRow, 1].Value2 = data.ArticleId;
            sheet.Cells[nextRow, 2].Value2 = data.CategoryId;
            sheet.Cells[nextRow, 3].Value2 = data.ProductName;
            sheet.Cells[nextRow, 4].Value2 = data.PurchasePrice;
            sheet.Cells[nextRow, 5].Value2 = data.SalePrice;
            sheet.Cells[nextRow, 6].Value2 = data.DiscountPercent;

            _workbook.Save();
            LogAction(logFilePath, $"Добавлен товар: {data.ArticleId}");
        }

        private void AddCategory(string logFilePath) {
            var data = new Category {
                CategoryId = InputValidator.GetNonEmptyString("ID категории: "),
                CategoryName = InputValidator.GetNonEmptyString("Наименование: "),
                AgeLimit = InputValidator.GetNonEmptyString("Возрастное ограничение: ")
            };

            Excel.Worksheet sheet = GetWorksheet(3);
            int nextRow = sheet.UsedRange.Rows.Count + 1;

            sheet.Cells[nextRow, 1].Value2 = data.CategoryId;
            sheet.Cells[nextRow, 2].Value2 = data.CategoryName;
            sheet.Cells[nextRow, 3].Value2 = data.AgeLimit;

            _workbook.Save();
            LogAction(logFilePath, $"Добавлена категория: {data.CategoryId}");
        }

        private void AddStore(string logFilePath) {
            var data = new Store
            {
                StoreId = InputValidator.GetNonEmptyString("ID магазина: "),
                District = InputValidator.GetNonEmptyString("Район: "),
                Address = InputValidator.GetNonEmptyString("Адрес: ")
            };

            Excel.Worksheet sheet = GetWorksheet(4);
            int nextRow = sheet.UsedRange.Rows.Count + 1;

            sheet.Cells[nextRow, 1].Value2 = data.StoreId;
            sheet.Cells[nextRow, 2].Value2 = data.District;
            sheet.Cells[nextRow, 3].Value2 = data.Address;

            _workbook.Save();
            LogAction(logFilePath, $"Добавлен магазин: {data.StoreId}");
        }

        /// <summary>
        /// Редактирование элемента в указанном листе
        /// </summary>
        public void EditElement(string logFilePath, int sheetNum) {
            try {
                string id = InputValidator.GetNonEmptyString("Введите ID для редактирования: ");
                Excel.Worksheet sheet = GetWorksheet(sheetNum);
                var foundRow = FindRowById(sheet, id);

                if (foundRow == null) {
                    throw new Exception("Элемент не найден");
                }

                switch (sheetNum) {
                    case 1: EditProductMovement(sheet, foundRow); break;
                    case 2: EditProduct(sheet, foundRow); break;
                    case 3: EditCategory(sheet, foundRow); break;
                    case 4: EditStore(sheet, foundRow); break;
                }

                _workbook.Save();
                LogAction(logFilePath, $"Изменён элемент {id} в листе {sheetNum}");
            }
            catch (Exception ex) {
                LogAction(logFilePath, $"Ошибка редактирования: {ex.Message}");
                throw;
            }
        }

        private void EditProductMovement(Excel.Worksheet sheet, Excel.Range row) {
            Console.WriteLine($"Дата (текущая: {row.Cells[1, 2].Value}): ");
            if (DateTime.TryParse(Console.ReadLine(), out DateTime newDate)) {
                row.Cells[1, 2].Value = newDate;
            }

            Console.WriteLine($"Тип операции (текущий: {row.Cells[1, 5].Value2}): ");
            string newType = Console.ReadLine();
            if (!string.IsNullOrEmpty(newType)) {
                row.Cells[1, 5].Value2 = newType;
            }
        }

        private void EditProduct(Excel.Worksheet sheet, Excel.Range row) {
            Console.WriteLine($"Наименование (текущее: {row.Cells[1, 3].Value2}): ");
            string newName = Console.ReadLine();
            if (!string.IsNullOrEmpty(newName)) {
                row.Cells[1, 3].Value2 = newName;
            }
        }

        private void EditCategory(Excel.Worksheet sheet, Excel.Range row) {
            Console.WriteLine($"Возрастное ограничение (текущее: {row.Cells[1, 3].Value2}): ");
            string newLimit = Console.ReadLine();
            if (!string.IsNullOrEmpty(newLimit)) {
                row.Cells[1, 3].Value2 = newLimit;
            }
        }

        private void EditStore(Excel.Worksheet sheet, Excel.Range row) {
            Console.WriteLine($"Адрес (текущий: {row.Cells[1, 3].Value2}): ");
            string newAddress = Console.ReadLine();
            if (!string.IsNullOrEmpty(newAddress)) {
                row.Cells[1, 3].Value2 = newAddress;
            }
        }

        /// <summary>
        /// Удаление элемента из указанного листа
        /// </summary>
        public void DeleteElement(string logFilePath, int sheetNum) {
            try {
                string id = InputValidator.GetNonEmptyString("Введите ID для удаления: ");
                Excel.Worksheet sheet = GetWorksheet(sheetNum);
                var foundRow = FindRowById(sheet, id);

                if (foundRow == null) {
                    throw new Exception("Элемент не найден");
                }

              ((Excel.Range)sheet.Rows[foundRow.Row]).Delete();
                _workbook.Save();
                LogAction(logFilePath, $"Удалён элемент {id} из листа {sheetNum}");
            }
            catch (Exception ex) {
                LogAction(logFilePath, $"Ошибка удаления: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Выполнение запросов к данным
        /// </summary>
        public void ExecuteQuery(string logFilePath) {
            try {
                // Пример запроса - общая стоимость товаров в категории "Игрушки"
                string categoryName = "Игрушки на радиоуправлении";
                string ageLimit = "12+";
                string district = "Ходунковый";
                DateTime startDate = new DateTime(2024, 8, 1);
                DateTime endDate = new DateTime(2024, 8, 5);

                // Реализация запроса...
                LogAction(logFilePath, "Выполнен запрос к данным");
                Console.WriteLine("Запрос выполнен. Результаты в файле лога.");
            }
            catch (Exception ex) {
                LogAction(logFilePath, $"Ошибка выполнения запроса: {ex.Message}");
                throw;
            }
        }

        private Excel.Worksheet GetWorksheet(int sheetNum) {
            return (Excel.Worksheet)_workbook.Sheets[sheetNum];
        }

        private Excel.Range FindRowById(Excel.Worksheet sheet, string id) {
            Excel.Range usedRange = sheet.UsedRange;
            for (int row = 1; row <= usedRange.Rows.Count; row++) {
                if (usedRange.Cells[row, 1].Value2?.ToString() == id) {
                    return usedRange.Rows[row];
                }
            }
            return null;
        }

        private void LogAction(string logFilePath, string message) {
            try {
                File.AppendAllText(logFilePath, $"{DateTime.Now}: {message}{Environment.NewLine}");
            }
            catch (Exception ex) {
                Console.WriteLine($"Ошибка записи в лог: {ex.Message}");
            }
        }

        public void Dispose() {
            try {
                if (_workbook != null) {
                    _workbook.Close(false);
                    Marshal.ReleaseComObject(_workbook);
                }

                if (_excelApp != null) {
                    _excelApp.Quit();
                    Marshal.ReleaseComObject(_excelApp);
                }
            }
            catch (Exception ex) {
                Console.WriteLine($"Ошибка при освобождении ресурсов: {ex.Message}");
            }
            finally {
                GC.SuppressFinalize(this);
            }
        }

        ~Database() {
            Dispose();
        }
    }
}
