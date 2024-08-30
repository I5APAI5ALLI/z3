using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel;

namespace OrderManagement
{
    class Program
    {
        static void Main(string[] args)
        {
            // Устанавливаем кодировку консоли на UTF-8
            Console.OutputEncoding = Encoding.UTF8;
            Console.InputEncoding = Encoding.UTF8;

            Console.WriteLine("Введите путь до файла с данными:");
            string filePath = Console.ReadLine();

            ExcelManager excelManager = new ExcelManager(filePath);

            while (true)
            {
                Console.WriteLine("\nВыберите команду:");
                Console.WriteLine("1. Поиск по наименованию товара");
                Console.WriteLine("2. Изменение контактного лица клиента");
                Console.WriteLine("3. Определение золотого клиента");
                Console.WriteLine("4. Выход");

                string command = Console.ReadLine();

                switch (command)
                {
                    case "1":
                        Console.WriteLine("Введите наименование товара:");
                        string productName = Console.ReadLine();
                        excelManager.GetClientsByProductName(productName);
                        break;

                    case "2":
                        Console.WriteLine("Введите название организации:");
                        string organizationName = Console.ReadLine();
                        Console.WriteLine("Введите новое ФИО контактного лица:");
                        string newContactPerson = Console.ReadLine();
                        excelManager.UpdateClientContactPerson(organizationName, newContactPerson);
                        break;

                    case "3":
                        Console.WriteLine("Введите год:");
                        int year = int.Parse(Console.ReadLine());
                        Console.WriteLine("Введите месяц:");
                        int month = int.Parse(Console.ReadLine());
                        excelManager.GetGoldenClient(year, month);
                        break;

                    case "4":
                        return;

                    default:
                        Console.WriteLine("Неверная команда, попробуйте еще раз.");
                        break;
                }
            }
        }
    }

    public class Order
    {
        public int ProductCode { get; set; }
        public int ClientCode { get; set; }
        public int Quantity { get; set; }
        public DateTime OrderDate { get; set; }
    }

    public class Product
    {
        public int ProductCode { get; set; }
        public string ProductName { get; set; }
        public string Unit { get; set; }
        public decimal Price { get; set; }
    }

    public class Client
    {
        public int ClientCode { get; set; }
        public string OrganizationName { get; set; }
        public string ContactPerson { get; set; }
        public string Address { get; set; }
        public List<Order> Orders { get; set; }

        public Client()
        {
            Orders = new List<Order>();
        }
    }

    public class ExcelManager
    {
        private string _filePath;
        private List<Client> _clients;
        private List<Product> _products;

        public ExcelManager(string filePath)
        {
            _filePath = filePath;
            _products = LoadProducts();
            _clients = LoadClientsAndOrders();
        }

        private List<Product> LoadProducts()
        {
            var products = new List<Product>();

            using (var workbook = new XLWorkbook(_filePath))
            {
                var worksheet = workbook.Worksheet("Товары");
                var rows = worksheet.RangeUsed().RowsUsed().Skip(1); // Пропускаем заголовок

                foreach (var row in rows)
                {
                    var product = new Product
                    {
                        ProductCode = row.Cell(1).GetValue<int>(),
                        ProductName = row.Cell(2).GetString(),
                        Unit = row.Cell(3).GetString(),
                        Price = row.Cell(4).GetValue<decimal>()
                    };
                    products.Add(product);
                }
            }

            return products;
        }

        private List<Client> LoadClientsAndOrders()
        {
            var clients = new List<Client>();

            using (var workbook = new XLWorkbook(_filePath))
            {
                var clientSheet = workbook.Worksheet("Клиенты");
                var orderSheet = workbook.Worksheet("Заявки");

                // Загрузка данных клиентов
                var clientRows = clientSheet.RangeUsed().RowsUsed().Skip(1); // Пропускаем заголовок

                foreach (var row in clientRows)
                {
                    var client = new Client
                    {
                        ClientCode = row.Cell(1).GetValue<int>(),
                        OrganizationName = row.Cell(2).GetString(),
                        Address = row.Cell(3).GetString(),
                        ContactPerson = row.Cell(4).GetString()
                    };
                    clients.Add(client);
                }

                // Загрузка данных заявок
                var orderRows = orderSheet.RangeUsed().RowsUsed().Skip(1); // Пропускаем заголовок

                foreach (var row in orderRows)
                {
                    var order = new Order
                    {
                        ProductCode = row.Cell(2).GetValue<int>(),
                        ClientCode = row.Cell(3).GetValue<int>(),
                        Quantity = row.Cell(5).GetValue<int>(),
                        OrderDate = row.Cell(6).GetDateTime()
                    };

                    var client = clients.FirstOrDefault(c => c.ClientCode == order.ClientCode);
                    client?.Orders.Add(order);
                }
            }

            return clients;
        }

        public void GetClientsByProductName(string productName)
        {
            var product = _products.FirstOrDefault(p => p.ProductName.Equals(productName, StringComparison.OrdinalIgnoreCase));

            if (product == null)
            {
                Console.WriteLine("Товар не найден.");
                return;
            }

            var clients = _clients.Where(c => c.Orders.Any(o => o.ProductCode == product.ProductCode));

            if (!clients.Any())
            {
                Console.WriteLine("Нет заказов на этот товар.");
                return;
            }

            foreach (var client in clients)
            {
                Console.WriteLine($"Клиент: {client.OrganizationName}, Контактное лицо: {client.ContactPerson}");
                foreach (var order in client.Orders.Where(o => o.ProductCode == product.ProductCode))
                {
                    Console.WriteLine($"Количество: {order.Quantity}, Цена: {product.Price}, Дата заказа: {order.OrderDate.ToShortDateString()}");
                }
            }
        }

        public void UpdateClientContactPerson(string organizationName, string newContactPerson)
        {
            var client = _clients.FirstOrDefault(c => c.OrganizationName.Equals(organizationName, StringComparison.OrdinalIgnoreCase));

            if (client == null)
            {
                Console.WriteLine("Организация не найдена.");
                return;
            }

            client.ContactPerson = newContactPerson;

            SaveDataToExcel();
            Console.WriteLine("Контактное лицо обновлено.");
        }

        public void GetGoldenClient(int year, int month)
        {
            var goldenClient = _clients
                .Select(c => new
                {
                    Client = c,
                    OrderCount = c.Orders.Count(o => o.OrderDate.Year == year && o.OrderDate.Month == month)
                })
                .OrderByDescending(c => c.OrderCount)
                .FirstOrDefault();

            if (goldenClient == null || goldenClient.OrderCount == 0)
            {
                Console.WriteLine("Нет заказов за указанный период.");
                return;
            }

            Console.WriteLine($"Золотой клиент: {goldenClient.Client.OrganizationName}, Количество заказов: {goldenClient.OrderCount}");
        }

        private void SaveDataToExcel()
        {
            using (var workbook = new XLWorkbook(_filePath))
            {
                var clientSheet = workbook.Worksheet("Клиенты");

                foreach (var client in _clients)
                {
                    var rows = clientSheet.RowsUsed().Skip(1); // Пропускаем заголовок
                    var row = rows.FirstOrDefault(r => r.Cell(1).GetValue<int>() == client.ClientCode);

                    if (row != null)
                    {
                        row.Cell(4).Value = client.ContactPerson;
                    }
                }

                workbook.Save();
            }
        }
    }
}
