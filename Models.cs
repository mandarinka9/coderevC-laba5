using System;

namespace RetailManagementSystem {
  /// <summary>
  /// Модель данных для движения товаров
  /// </summary>
  public class ProductMovement {
    /// <summary>Уникальный идентификатор операции</summary>
    public string OperationId { get; set; }

    /// <summary>Дата выполнения операции</summary>
    public DateTime Date { get; set; }

    /// <summary>Идентификатор магазина</summary>
    public string StoreId { get; set; }

    /// <summary>Артикул товара</summary>
    public string ArticleId { get; set; }

    /// <summary>Тип операции (Поступление/Продажа/Возврат)</summary>
    public string OperationType { get; set; }

    /// <summary>Количество упаковок</summary>
    public int PackageCount { get; set; }

    /// <summary>Наличие карты клиента</summary>
    public bool HasClientCard { get; set; }
  }

  /// <summary>
  /// Модель данных для товара
  /// </summary>
  public class Product {
    /// <summary>Артикул товара</summary>
    public string ArticleId { get; set; }

    /// <summary>Идентификатор категории</summary>
    public string CategoryId { get; set; }

    /// <summary>Наименование товара</summary>
    public string ProductName { get; set; }

    /// <summary>Цена закупки</summary>
    public decimal PurchasePrice { get; set; }

    /// <summary>Цена продажи</summary>
    public decimal SalePrice { get; set; }

    /// <summary>Процент скидки</summary>
    public int DiscountPercent { get; set; }
  }

  /// <summary>
  /// Модель данных для категории
  /// </summary>
  public class Category {
    /// <summary>Идентификатор категории</summary>
    public string CategoryId { get; set; }

    /// <summary>Наименование категории</summary>
    public string CategoryName { get; set; }

    /// <summary>Возрастное ограничение</summary>
    public string AgeLimit { get; set; }
  }

  /// <summary>
  /// Модель данных для магазина
  /// </summary>
  public class Store {
    /// <summary>Идентификатор магазина</summary>
    public string StoreId { get; set; }

    /// <summary>Район расположения</summary>
    public string District { get; set; }

    /// <summary>Адрес магазина</summary>
    public string Address { get; set; }
  }
}
