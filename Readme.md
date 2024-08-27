Here's the content in markdown format:

# Excel Function Examples Using Northwind Dataset

## 1. Logical Operations

### IF()

1. **Problem**: Categorize products as "Expensive" if their unit price is over $50, otherwise "Affordable".
   **Solution**: 
   ```excel
   =IF(Products[UnitPrice] > 50, "Expensive", "Affordable")
   ```

2. **Problem**: Check if an order's freight cost is high (over $100) or low.
   **Solution**: 
   ```excel
   =IF(Orders[Freight] > 100, "High freight", "Low freight")
   ```

3. **Problem**: Determine if an employee is "Senior" (hired before 1994) or "Junior".
   **Solution**: 
   ```excel
   =IF(YEAR(Employees[HireDate]) < 1994, "Senior", "Junior")
   ```

### IFS()

1. **Problem**: Categorize products based on units in stock.
   **Solution**: 
   ```excel
   =IFS(Products[UnitsInStock]=0, "Out of stock", 
        Products[UnitsInStock]<10, "Low stock", 
        Products[UnitsInStock]<50, "Medium stock", 
        TRUE, "High stock")
   ```

2. **Problem**: Classify orders based on total value.
   **Solution**: 
   ```excel
   =IFS(Orders[TotalValue]<100, "Small order", 
        Orders[TotalValue]<1000, "Medium order", 
        Orders[TotalValue]<10000, "Large order", 
        TRUE, "Huge order")
   ```

3. **Problem**: Categorize employees by their title.
   **Solution**: 
   ```excel
   =IFS(Employees[Title]="Sales Representative", "Sales", 
        Employees[Title]="Sales Manager", "Management", 
        Employees[Title]="Vice President, Sales", "Executive", 
        TRUE, "Other")
   ```

### SWITCH()

1. **Problem**: Assign a region based on the country of a customer.
   **Solution**: 
   ```excel
   =SWITCH(Customers[Country], 
           "USA", "North America", 
           "Canada", "North America", 
           "Mexico", "North America", 
           "Brazil", "South America", 
           "Argentina", "South America", 
           "UK", "Europe", 
           "France", "Europe", 
           "Germany", "Europe", 
           "Other")
   ```

2. **Problem**: Determine shipping method based on order ID's last digit.
   **Solution**: 
   ```excel
   =SWITCH(RIGHT(Orders[OrderID], 1), 
           "0", "Air", 
           "1", "Sea", 
           "2", "Ground", 
           "3", "Express", 
           "Other")
   ```

3. **Problem**: Assign a discount tier based on the customer's company name first letter.
   **Solution**: 
   ```excel
   =SWITCH(LEFT(Customers[CompanyName], 1), 
           "A", "Tier 1", 
           "B", "Tier 1", 
           "C", "Tier 2", 
           "D", "Tier 2", 
           "Tier 3")
   ```

### SUMIF()

1. **Problem**: Calculate total freight for orders shipped to the USA.
   **Solution**: 
   ```excel
   =SUMIF(Orders[ShipCountry], "USA", Orders[Freight])
   ```

2. **Problem**: Find the total units in stock for all products in category 1.
   **Solution**: 
   ```excel
   =SUMIF(Products[CategoryID], 1, Products[UnitsInStock])
   ```

3. **Problem**: Compute the total value of orders placed by customer "ALFKI".
   **Solution**: 
   ```excel
   =SUMIF(Orders[CustomerID], "ALFKI", Orders[TotalValue])
   ```

### AVERAGEIF()

1. **Problem**: Calculate the average unit price of products in category 3.
   **Solution**: 
   ```excel
   =AVERAGEIF(Products[CategoryID], 3, Products[UnitPrice])
   ```

2. **Problem**: Find the average freight cost for orders shipped to Germany.
   **Solution**: 
   ```excel
   =AVERAGEIF(Orders[ShipCountry], "Germany", Orders[Freight])
   ```

3. **Problem**: Determine the average quantity per order for product ID 11.
   **Solution**: 
   ```excel
   =AVERAGEIF(OrderDetails[ProductID], 11, OrderDetails[Quantity])
   ```

### COUNTIF()

1. **Problem**: Count how many products are in category 2.
   **Solution**: 
   ```excel
   =COUNTIF(Products[CategoryID], 2)
   ```

2. **Problem**: Determine how many orders were shipped to France.
   **Solution**: 
   ```excel
   =COUNTIF(Orders[ShipCountry], "France")
   ```

3. **Problem**: Find out how many employees have the title "Sales Representative".
   **Solution**: 
   ```excel
   =COUNTIF(Employees[Title], "Sales Representative")
   ```

### SUMIFS()

1. **Problem**: Calculate total freight for orders shipped to the USA in 1997.
   **Solution**: 
   ```excel
   =SUMIFS(Orders[Freight], 
           Orders[ShipCountry], "USA", 
           Orders[OrderDate], ">=1/1/1997", 
           Orders[OrderDate], "<1/1/1998")
   ```

2. **Problem**: Find the total value of orders for products in category 1 with unit price > $20.
   **Solution**: 
   ```excel
   =SUMIFS(OrderDetails[TotalValue], 
           Products[CategoryID], 1, 
           Products[UnitPrice], ">20")
   ```

3. **Problem**: Compute the total quantity ordered for products supplied by supplier ID 1 and not discontinued.
   **Solution**: 
   ```excel
   =SUMIFS(OrderDetails[Quantity], 
           Products[SupplierID], 1, 
           Products[Discontinued], 0)
   ```

### AVERAGEIFS()

1. **Problem**: Calculate the average freight cost for orders shipped to Germany in 1998.
   **Solution**: 
   ```excel
   =AVERAGEIFS(Orders[Freight], 
               Orders[ShipCountry], "Germany", 
               Orders[OrderDate], ">=1/1/1998", 
               Orders[OrderDate], "<1/1/1999")
   ```

2. **Problem**: Find the average unit price of products in category 3 with more than 20 units in stock.
   **Solution**: 
   ```excel
   =AVERAGEIFS(Products[UnitPrice], 
               Products[CategoryID], 3, 
               Products[UnitsInStock], ">20")
   ```

3. **Problem**: Determine the average quantity per order for product ID 11 in orders placed by customers in the USA.
   **Solution**: 
   ```excel
   =AVERAGEIFS(OrderDetails[Quantity], 
               OrderDetails[ProductID], 11, 
               Orders[ShipCountry], "USA")
   ```

### COUNTIFS()

1. **Problem**: Count how many orders were shipped to Germany in 1997.
   **Solution**: 
   ```excel
   =COUNTIFS(Orders[ShipCountry], "Germany", 
             Orders[OrderDate], ">=1/1/1997", 
             Orders[OrderDate], "<1/1/1998")
   ```

2. **Problem**: Determine how many products in category 1 have a unit price over $20 and are not discontinued.
   **Solution**: 
   ```excel
   =COUNTIFS(Products[CategoryID], 1, 
             Products[UnitPrice], ">20", 
             Products[Discontinued], 0)
   ```

3. **Problem**: Find out how many employees were hired before 1994 and have the title "Sales Representative".
   **Solution**: 
   ```excel
   =COUNTIFS(Employees[HireDate], "<1/1/1994", 
             Employees[Title], "Sales Representative")
   ```

### MAXIFS()

1. **Problem**: Find the highest freight cost for orders shipped to France.
   **Solution**: 
   ```excel
   =MAXIFS(Orders[Freight], Orders[ShipCountry], "France")
   ```

2. **Problem**: Determine the highest unit price among products in category 2 that are not discontinued.
   **Solution**: 
   ```excel
   =MAXIFS(Products[UnitPrice], 
           Products[CategoryID], 2, 
           Products[Discontinued], 0)
   ```

3. **Problem**: Find the maximum quantity ordered for any product supplied by supplier ID 1.
   **Solution**: 
   ```excel
   =MAXIFS(OrderDetails[Quantity], Products[SupplierID], 1)
   ```

### MINIFS()

1. **Problem**: Find the lowest freight cost for orders shipped to the USA in 1998.
   **Solution**: 
   ```excel
   =MINIFS(Orders[Freight], 
           Orders[ShipCountry], "USA", 
           Orders[OrderDate], ">=1/1/1998", 
           Orders[OrderDate], "<1/1/1999")
   ```

2. **Problem**: Determine the lowest unit price among products in category 3 with more than 10 units in stock.
   **Solution**: 
   ```excel
   =MINIFS(Products[UnitPrice], 
           Products[CategoryID], 3, 
           Products[UnitsInStock], ">10")
   ```

3. **Problem**: Find the minimum quantity ordered for any product in orders placed by customers in Germany.
   **Solution**: 
   ```excel
   =MINIFS(OrderDetails[Quantity], Orders[ShipCountry], "Germany")
   ```

### AND()

1. **Problem**: Check if a product is both expensive (>$50) and low in stock (<10 units).
   **Solution**: 
   ```excel
   =AND(Products[UnitPrice] > 50, Products[UnitsInStock] < 10)
   ```

2. **Problem**: Verify if an order was both shipped to the USA and had a high freight cost (>$100).
   **Solution**: 
   ```excel
   =AND(Orders[ShipCountry] = "USA", Orders[Freight] > 100)
   ```

3. **Problem**: Determine if an employee is both a senior staff member (hired before 1994) and a sales representative.
   **Solution**: 
   ```excel
   =AND(YEAR(Employees[HireDate]) < 1994, Employees[Title] = "Sales Representative")
   ```

### OR()

1. **Problem**: Check if a product is either out of stock or discontinued.
   **Solution**: 
   ```excel
   =OR(Products[UnitsInStock] = 0, Products[Discontinued] = 1)
   ```

2. **Problem**: Verify if an order was shipped to either the USA or Canada.
   **Solution**: 
   ```excel
   =OR(Orders[ShipCountry] = "USA", Orders[ShipCountry] = "Canada")
   ```

3. **Problem**: Determine if an employee is either in sales or management.
   **Solution**: 
   ```excel
   =OR(Employees[Title] = "Sales Representative", Employees[Title] = "Sales Manager")
   ```

### NOT()

1. **Problem**: Find products that are not in category 1.
   **Solution**: 
   ```excel
   =NOT(Products[CategoryID] = 1)
   ```

2. **Problem**: Identify orders that were not shipped to the USA.
   **Solution**: 
   ```excel
   =NOT(Orders[ShipCountry] = "USA")
   ```

3. **Problem**: Determine which employees are not sales representatives.
   **Solution**: 
   ```excel
   =NOT(Employees[Title] = "Sales Representative")
   ```

### LET()

1. **Problem**: Calculate the total value of orders for a specific customer, including a volume discount.
   **Solution**: 
   ```excel
   =LET(CustomerOrders, SUMIF(Orders[CustomerID], A1, Orders[TotalValue]), 
        Discount, IF(CustomerOrders > 10000, 0.05, 0), 
        CustomerOrders * (1 - Discount))
   ```

2. **Problem**: Find the average order value for products in a specific category, excluding discontinued products.
   **Solution**: 
   ```excel
   =LET(CategoryProducts, FILTER(Products, Products[CategoryID] = A1), 
        ActiveProducts, FILTER(CategoryProducts, CategoryProducts[Discontinued] = 0), 
        AVERAGE(ActiveProducts[UnitPrice]))
   ```

3. **Problem**: Calculate the percentage of total sales contributed by the top 5 customers.
   **Solution**: 
   ```excel
   =LET(TotalSales, SUM(Orders[TotalValue]), 
        Top5Sales, SUMPRODUCT(--(RANK(Orders[TotalValue], Orders[TotalValue]) <= 5), Orders[TotalValue]), 
        Top5Sales / TotalSales)
   ```

## 2. Looking Up Data

### XLOOKUP()

1. **Problem**: Find the company name for a given customer ID.
   **Solution**: 
   ```excel
   =XLOOKUP(A1, Customers[CustomerID], Customers[CompanyName], "Customer not found")
   ```

2. **Problem**: Retrieve the category name for a given product ID.
   **Solution**: 
   ```excel
   =XLOOKUP(XLOOKUP(A1, Products[ProductID], Products[CategoryID]), Categories[CategoryID], Categories[CategoryName], "Category not found")
   ```

3. **Problem**: Get the employee's last name for a given order ID.
   **Solution**: 
   ```excel
   =XLOOKUP(XLOOKUP(A1, Orders[OrderID], Orders[EmployeeID]), Employees[EmployeeID], Employees[LastName], "Employee not found")
   ```

### VLOOKUP()

1. **Problem**: Find the unit price of a product given its product ID.
   **Solution**: 
   ```excel
   =VLOOKUP(A1, Products, MATCH("UnitPrice", Products[#Headers], 0), FALSE)
   ```

2. **Problem**: Retrieve the ship country for a given order ID.
   **Solution**: 
   ```excel
   =VLOOKUP(A1, Orders, MATCH("ShipCountry", Orders[#Headers], 0), FALSE)
   ```

3. **Problem**: Get the category description for a given category name.
   **Solution**: 
   ```excel
   =VLOOKUP(A1, Categories, MATCH("Description", Categories[#Headers], 0), FALSE)
   ```

### HLOOKUP()

1. **Problem**: Find the reorder level for a product given its product name.
   **Solution**: 
   ```excel
   =HLOOKUP(A1, Products, MATCH("ReorderLevel", Products[#Headers], 0), FALSE)
   ```

2. **Problem**: Get the birth date of an employee given their last name.
   **Solution**: 
   ```excel
   =HLOOKUP(A1, Employees, MATCH("BirthDate", Employees[#Headers], 0), FALSE)
   ```

3. **Problem**: Retrieve the fax number for a supplier given their company name.
   **Solution**: 
   ```excel
   =HLOOKUP(A1, Suppliers, MATCH("Fax", Suppliers[#Headers], 0), FALSE)
   ```

### MATCH()

1. **Problem**: Find the position of a specific product in the product list sorted by unit price.
   **Solution**: 
   ```excel
   =MATCH(A1, SORT(Products[ProductName], SORTBY(Products[UnitPrice], Products[UnitPrice], -1)), 0)
   ```

2. **Problem**: Determine the rank of an employee's hire date among all employees.
   **Solution**: 
   ```excel
   =MATCH(XLOOKUP(A1, Employees[EmployeeID], Employees[HireDate]), SORT(Employees[HireDate]), 0)
   ```

Certainly! I'll continue from the MATCH() function examples:

3. **Problem**: Find the position of a specific country in the list of unique ship countries.
   **Solution**: 
   ```excel
   =MATCH(A1, UNIQUE(Orders[ShipCountry]), 0)
   ```

### INDEX()

1. **Problem**: Get the nth most expensive product name.
   **Solution**: 
   ```excel
   =INDEX(Products[ProductName], MATCH(n, RANK(Products[UnitPrice], Products[UnitPrice]), 0))
   ```

2. **Problem**: Retrieve the company name and contact name for the customer with the most orders.
   **Solution**: 
   ```excel
   =INDEX(Customers, MATCH(MAX(COUNTIF(Orders[CustomerID], Customers[CustomerID])), COUNTIF(Orders[CustomerID], Customers[CustomerID]), 0), {2,3})
   ```

3. **Problem**: Find the product name and unit price of the most ordered product.
   **Solution**: 
   ```excel
   =INDEX(Products, MATCH(MAX(SUMIF(OrderDetails[ProductID], Products[ProductID], OrderDetails[Quantity])), SUMIF(OrderDetails[ProductID], Products[ProductID], OrderDetails[Quantity]), 0), {2,5})
   ```

## 3. Date and Time Functions

### NOW()

1. **Problem**: Calculate how many days have passed since the last order.
   **Solution**: 
   ```excel
   =NOW() - MAX(Orders[OrderDate])
   ```

2. **Problem**: Determine if the current time is within business hours (9 AM to 5 PM).
   **Solution**: 
   ```excel
   =AND(HOUR(NOW()) >= 9, HOUR(NOW()) < 17)
   ```

3. **Problem**: Calculate the age of each employee in years and days.
   **Solution**: 
   ```excel
   =DATEDIF(Employees[BirthDate], NOW(), "Y") & " years, " & MOD(DATEDIF(Employees[BirthDate], NOW(), "MD"), 30) & " days"
   ```

### TODAY()

1. **Problem**: Find orders that are due to be shipped today.
   **Solution**: 
   ```excel
   =IF(Orders[RequiredDate] = TODAY(), "Ship today", "")
   ```

2. **Problem**: Calculate how many days are left in the current month.
   **Solution**: 
   ```excel
   =EOMONTH(TODAY(), 0) - TODAY()
   ```

3. **Problem**: Determine if it's the last day of the quarter.
   **Solution**: 
   ```excel
   =IF(EOMONTH(TODAY(), 0) = TODAY() AND MONTH(TODAY()) IN {3,6,9,12}, "Last day of quarter", "Not last day of quarter")
   ```

### WEEKDAY()

1. **Problem**: Categorize orders by the day of the week they were placed.
   **Solution**: 
   ```excel
   =CHOOSE(WEEKDAY(Orders[OrderDate]), "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday")
   ```

2. **Problem**: Count how many orders were placed on weekends.
   **Solution**: 
   ```excel
   =COUNTIF(WEEKDAY(Orders[OrderDate]), ">5")
   ```

3. **Problem**: Calculate the average order value for each day of the week.
   **Solution**: 
   ```excel
   =AVERAGEIFS(Orders[TotalValue], WEEKDAY(Orders[OrderDate]), ROW(1:7))
   ```

### WORKDAY()

1. **Problem**: Calculate the expected ship date for an order (3 business days after the order date).
   **Solution**: 
   ```excel
   =WORKDAY(Orders[OrderDate], 3)
   ```

2. **Problem**: Determine how many business days an order was delayed.
   **Solution**: 
   ```excel
   =NETWORKDAYS(Orders[RequiredDate], Orders[ShippedDate]) - 1
   ```

3. **Problem**: Find the date of the next performance review (every 6 months from hire date, on a business day).
   **Solution**: 
   ```excel
   =WORKDAY(EDATE(Employees[HireDate], ROUNDUP((DATEDIF(Employees[HireDate], TODAY(), "M") / 6), 0) * 6), 0)
   ```

This completes the list of examples for all the functions you requested, presented in markdown format and using the Northwind dataset context.