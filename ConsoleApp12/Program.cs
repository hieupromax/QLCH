using StoreManagement.Models;
using System.Text.RegularExpressions;
using OfficeOpenXml;
class Program
{
    static List<Employee> employees = new();
    static List<Product> products = new();
    static List<Customer> customers = new();
    static List<Order> orders = new();

    static bool ValidatePhone(string phone)
    {
        return Regex.IsMatch(phone, @"^\d{10}$");
    }

    static bool ValidateAddress(string address)
    {
        return !string.IsNullOrWhiteSpace(address) && address.Length >= 5;
    }

    static bool ValidateEmail(string email)
    {
        return Regex.IsMatch(email, @"^[^@\s]+@[^@\s]+\.[^@\s]+$");
    }

    static void Main()
    {
        while (true)
        {
            Console.WriteLine("\n1. Manage Employees\n2. Manage Products\n3. Customer Purchase\n4. View Orders\n5. Export to Excel\n0. Exit");
            Console.Write("Select option: ");
            var choice = Console.ReadLine();

            switch (choice)
            {
                case "1":
                    ManageEmployees();
                    break;
                case "2":
                    ManageProducts();
                    break;
                case "3":
                    CustomerPurchase();
                    break;
                case "4":
                    ViewOrders();
                    break;
                case "5":
                    ExportData();
                    break;
                case "0":
                    return;
                default:
                    Console.WriteLine("Invalid option.");
                    break;
            }
        }
    }

    static void ManageEmployees()
    {
        Console.WriteLine("\n1. Add Employee\n2. Edit Employee\n3. Delete Employee\n4. List Employees");
        var option = Console.ReadLine();
        switch (option)
        {
            case "1":
                employees.Add(GetEmployee());
                break;
            case "2":
                EditEmployee();
                break;
            case "3":
                DeleteEmployee();
                break;
            case "4":
                foreach (var e in employees)
                {
                    Console.WriteLine($"{e.Name} | {e.Email} | {e.Phone}");
                }
                break;
        }
    }

    static void ManageProducts()
    {
        Console.WriteLine("\n1. Add Product\n2. Edit Product\n3. Delete Product\n4. List Products");
        var option = Console.ReadLine();
        switch (option)
        {
            case "1":
                products.Add(GetProduct());
                break;
            case "2":
                EditProduct();
                break;
            case "3":
                DeleteProduct();
                break;
            case "4":
                foreach (var p in products)
                {
                    Console.WriteLine($"{p.Name} | {p.Price:C} | Stock: {p.Stock}");
                }
                break;
        }
    }

    static void CustomerPurchase()
    {
        Console.Write("Enter Customer Name: ");
        string name = Console.ReadLine() ?? "";
        while (string.IsNullOrEmpty(name))
        {
            Console.Write("Invalid name. Re-enter: ");
            name = Console.ReadLine() ?? "";
        }

        var customer = customers.Find(c => c.Name == name);
        if (customer == null)
        {
            Console.Write("Customer not found. Enter Phone (10 digits): ");
            string phone = Console.ReadLine() ?? "";
            while (!ValidatePhone(phone))
            {
                Console.Write("Invalid phone. Re-enter: ");
                phone = Console.ReadLine() ?? "";
            }
            Console.Write("Shipping Address: ");
            string address = Console.ReadLine() ?? "";
            while (!ValidateAddress(address))
            {
                Console.Write("Invalid address. Re-enter: ");
                address = Console.ReadLine() ?? "";
            }
            Console.Write("Email: ");
            string email = Console.ReadLine() ?? "";
            while (!ValidateEmail(email))
            {
                Console.Write("Invalid email. Re-enter(xxx@xxx.xx): ");
                email = Console.ReadLine() ?? "";
            }
            customer = new Customer { Name = name, Phone = phone, Address = address, Email = email };
            customers.Add(customer);
        }

        Order order = new Order
        {
            CustomerName = customer.Name,
            ShippingAddress = customer.Address
        };

        Console.WriteLine("====Product List Stock====");
        foreach (var p in products)
        {
            Console.WriteLine($"{p.Name} | {p.Price:C} | Stock: {p.Stock}");
        }

        while (true)
        {
            Console.Write("Enter Product Name (or 'done'): ");
            string pname = Console.ReadLine() ?? "";
            if (pname.ToLower() == "done") break;

            var product = products.Find(p => p.Name == pname);
            if (product == null) { Console.WriteLine("Product not found."); continue; }

            Console.WriteLine($"Current stock: {product.Stock}");

            Console.Write("Quantity: ");
            if (!int.TryParse(Console.ReadLine(), out int qty) || qty <= 0 || qty > product.Stock)
            {
                Console.WriteLine("Invalid quantity.");
                continue;
            }

            order.Items.Add(new OrderDetail
            {
                ProductName = pname,
                Quantity = qty,
                UnitPrice = product.Price
            });
            product.Stock -= qty;
        }
        if (order.Items.Count > 0)
        {
            orders.Add(order);
            Console.WriteLine("\nOrder Summary:");
            foreach (var item in order.Items)
            {
                Console.WriteLine($" - {item.ProductName}, Qty: {item.Quantity}, Unit Price: {item.UnitPrice:C}, Subtotal: {item.UnitPrice * item.Quantity:C}");
            }
            Console.WriteLine($"Total Amount: {order.TotalAmount:C}");
        }
        else
        {
            Console.WriteLine("No items added. Order cancelled.");
        }
    }

    static void ViewOrders()
    {
        Console.Write("Enter Customer Name: ");
        string name = Console.ReadLine() ?? "";
        while (string.IsNullOrEmpty(name))
        {
            Console.Write("Invalid Name. Re-enter: ");
            name = Console.ReadLine() ?? "";
        }
        var filteredOrders = orders.Where(o => o.CustomerName.Equals(name, StringComparison.OrdinalIgnoreCase)).ToList();

        if (filteredOrders.Count == 0)
        {
            Console.WriteLine("No orders found.");
            return;
        }

        foreach (var order in filteredOrders)
        {
            Console.WriteLine($"\nCustomer: {order.CustomerName}, Address: {order.ShippingAddress}, Date: {order.OrderDate}");
            foreach (var item in order.Items)
            {
                Console.WriteLine($" - Product: {item.ProductName}, Qty: {item.Quantity}, Unit Price: {item.UnitPrice:C}, Subtotal: {item.UnitPrice * item.Quantity:C}");
            }
            Console.WriteLine($"Total Amount: {order.TotalAmount:C}");
        }
    }

    static void ExportData()
    {
        OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

        using (var package = new OfficeOpenXml.ExcelPackage())
        {
            var ws1 = package.Workbook.Worksheets.Add("Products");
            ws1.Cells[1, 1].Value = "Name";
            ws1.Cells[1, 2].Value = "Price";
            ws1.Cells[1, 3].Value = "Stock";
            for (int i = 0; i < products.Count; i++)
            {
                ws1.Cells[i + 2, 1].Value = products[i].Name;
                ws1.Cells[i + 2, 2].Value = products[i].Price;
                ws1.Cells[i + 2, 3].Value = products[i].Stock;
            }

            var ws2 = package.Workbook.Worksheets.Add("Customers");
            ws2.Cells[1, 1].Value = "Name";
            ws2.Cells[1, 2].Value = "Phone";
            ws2.Cells[1, 3].Value = "Address";
            ws2.Cells[1, 4].Value = "Email";
            for (int i = 0; i < customers.Count; i++)
            {
                ws2.Cells[i + 2, 1].Value = customers[i].Name;
                ws2.Cells[i + 2, 2].Value = customers[i].Phone;
                ws2.Cells[i + 2, 3].Value = customers[i].Address;
                ws2.Cells[i + 2, 4].Value = customers[i].Email;
            }
            string filePath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                "ExportedData.xlsx"
            );
            if (File.Exists(filePath))
            {
                File.Delete(filePath);
            }
            package.SaveAs(new FileInfo(filePath));
        }

        Console.WriteLine("Data exported to ExportedData.xlsx on DeskTop");
    }

    static Employee GetEmployee()
    {
        Console.Write("Name: ");
        var name = Console.ReadLine() ?? "";
        while (string.IsNullOrEmpty(name))
        {
            Console.Write("Invalid Name. Re-enter: ");
            name = Console.ReadLine() ?? "";
        }

        Console.Write("Email: ");
        var email = Console.ReadLine() ?? "";
        while (!ValidateEmail(email))
        {
            Console.Write("Invalid email. Re-enter(xxx@xxx.xx): ");
            email = Console.ReadLine() ?? "";
        }

        Console.Write("Phone (10 digits): ");
        var phone = Console.ReadLine() ?? "";
        while (!ValidatePhone(phone))
        {
            Console.Write("Invalid phone. Re-enter: ");
            phone = Console.ReadLine() ?? "";
        }
        return new Employee { Name = name, Email = email, Phone = phone };
    }

    static Product GetProduct()
    {
        Console.Write("Name: ");
        var name = Console.ReadLine();
        while (string.IsNullOrEmpty(name))
        {
            Console.Write("Invalid Name. Re-enter: ");
            name = Console.ReadLine();
        }
        Console.Write("Price: ");
        decimal price = 0;
        while (!decimal.TryParse(Console.ReadLine(), out price))
        {
            Console.Write("Invalid Price. Re-enter: ");
        }
        Console.Write("Stock: ");
        int stock = 0;
        while (!int.TryParse(Console.ReadLine(), out stock))
        {
            Console.Write("Invalid Stock. Re-enter: ");
        }
        return new Product { Name = name, Price = price, Stock = stock };
    }

    static void EditEmployee()
    {
        Console.Write("Enter Name: ");
        var name = Console.ReadLine();
        while (string.IsNullOrEmpty(name))
        {
            Console.Write("Invalid Name. Re-enter: ");
            name = Console.ReadLine();
        }

        var emp = employees.Find(e => e.Name == name);

        if (emp != null)
        {
            Console.Write("New Email: ");
            var Email = Console.ReadLine() ?? "";
            while (!ValidateEmail(Email))
            {
                Console.Write("Invalid email. Re-enter(xxx@xxx.xx): ");
                Email = Console.ReadLine() ?? "";
            }
            emp.Email = Email;

            Console.Write("New Phone(10 digits): ");
            var phone = Console.ReadLine() ?? "";
            while (!ValidatePhone(phone))
            {
                Console.Write("Invalid phone. Re-enter: ");
                phone = Console.ReadLine() ?? "";
            }
            emp.Phone = phone;
        }
        else
        {
            Console.Write("Cant find Employee");
        }
    }

    static void EditProduct()
    {
        Console.Write("Enter Name: ");
        var name = Console.ReadLine();
        while (string.IsNullOrEmpty(name))
        {
            Console.Write("Invalid Name. Re-enter: ");
            name = Console.ReadLine() ?? "";
        }

        var prod = products.Find(p => p.Name == name);
        if (prod != null)
        {
            Console.Write("New Price: ");
            if (decimal.TryParse(Console.ReadLine(), out decimal price))
            {
                prod.Price = price;
            }
            Console.Write("New Stock: ");
            if (int.TryParse(Console.ReadLine(), out int stock))
            {
                prod.Stock = stock;
            }
        }
    }

    static void DeleteEmployee()
    {
        Console.Write("Enter Employee Name to delete: ");
        string name = Console.ReadLine() ?? "";
        while (string.IsNullOrEmpty(name))
        {
            Console.Write("Invalid Name. Re-enter: ");
            name = Console.ReadLine() ?? "";
        }
        var emp = employees.Find(e => e.Name == name);
        if (emp != null) employees.Remove(emp);
    }

    static void DeleteProduct()
    {
        Console.Write("Enter Product Name to delete: ");
        string name = Console.ReadLine() ?? "";
        while (string.IsNullOrEmpty(name))
        {
            Console.Write("Invalid Name. Re-enter: ");
            name = Console.ReadLine() ?? "";
        }
        var prod = products.Find(p => p.Name == name);
        if (prod != null) products.Remove(prod);
    }
}