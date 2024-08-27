# Excel Function Examples with Sample Northwind Data

## Sample Dataset

Here's a sample of the first few rows of our hypothetical Northwind products dataset:

| A    | B    | C        | D      | E              | F            | G     | H    | I     | J     | K      |
|------|------|----------|--------|----------------|--------------|-------|------|-------|-------|--------|
| ID   | Price| Category | Value  | Supplier       | Discontinued | Stock | Order| Reorder | Country | Ship To |
| 1    | 18   | Beverages| 270    | Exotic Liquids | FALSE        | 15    | 10   | 10    | USA   | France |
| 2    | 19   | Condiments| 190   | New Orleans    | FALSE        | 10    | 5    | 5     | USA   | Germany|
| 3    | 10   | Confections| 100  | Tokyo Traders  | FALSE        | 20    | 15   | 15    | Japan | Brazil |
| 4    | 22   | Dairy Products| 440 | Cooperativa  | FALSE        | 20    | 10   | 30    | Italy | USA    |
| 5    | 21.35| Grains/Cereals| 213.5 | Exotic Liquids| TRUE     | 0     | 0    | 0     | UK    | France |

Now, let's go through each function with examples based on this dataset:

## 1. IF Function

**Sample Data:**
| ID   | Price| Category | Discontinued | Stock |
|------|------|----------|--------------|-------|
| 1    | 18   | Beverages| FALSE        | 15    |
| 2    | 19   | Condiments| FALSE       | 10    |
| 3    | 10   | Confections| FALSE      | 20    |
| 4    | 22   | Dairy Products| FALSE   | 20    |
| 5    | 21.35| Grains/Cereals| TRUE    | 0     |

### Example 1.1
**Exercise:** Create a column that categorizes products as "Expensive" if the unit price is over $20, and "Affordable" otherwise.

**Solution:**
```
=IF(B2>20, "Expensive", "Affordable")
```

### Example 1.2
**Exercise:** Display "Reorder" if the units in stock are less than 15, otherwise display "Sufficient".

**Solution:**
```
=IF(E2<15, "Reorder", "Sufficient")
```

### Example 1.3
**Exercise:** Show "Discontinued" if the product is discontinued, otherwise "Active".

**Solution:**
```
=IF(D2=TRUE, "Discontinued", "Active")
```

## 2. IFS Function

**Sample Data:**
| ID   | Price| Category | Stock | Country |
|------|------|----------|-------|---------|
| 1    | 18   | Beverages| 15    | USA     |
| 2    | 19   | Condiments| 10   | USA     |
| 3    | 10   | Confections| 20  | Japan   |
| 4    | 22   | Dairy Products| 20| Italy   |
| 5    | 21.35| Grains/Cereals| 0 | UK      |

### Example 2.1
**Exercise:** Categorize products based on their unit price: "Budget" if under $15, "Mid-range" if between $15 and $20, and "Premium" if over $20.

**Solution:**
```
=IFS(B2<15, "Budget", B2<=20, "Mid-range", B2>20, "Premium")
```

### Example 2.2
**Exercise:** Classify products based on units in stock: "Out of Stock" if 0, "Low Stock" if 1-10, "Medium Stock" if 11-20, "High Stock" if over 20.

**Solution:**
```
=IFS(D2=0, "Out of Stock", D2<=10, "Low Stock", D2<=20, "Medium Stock", D2>20, "High Stock")
```

### Example 2.3
**Exercise:** Categorize suppliers based on their country: "North America" for USA, "Europe" for UK and Italy, "Asia" for Japan, "Other" for any other country.

**Solution:**
```
=IFS(E2="USA", "North America", OR(E2="UK", E2="Italy"), "Europe", E2="Japan", "Asia", TRUE, "Other")
```

## 3. SWITCH Function

**Sample Data:**
| ID   | Category ID | Ship To | Supplier ID |
|------|-------------|---------|-------------|
| 1    | 1           | France  | 1           |
| 2    | 2           | Germany | 2           |
| 3    | 3           | Brazil  | 3           |
| 4    | 4           | USA     | 4           |
| 5    | 5           | France  | 1           |

### Example 3.1
**Exercise:** Display the full category name based on the category ID (1: Beverages, 2: Condiments, 3: Confections, 4: Dairy Products, 5: Grains/Cereals).

**Solution:**
```
=SWITCH(B2, 1, "Beverages", 2, "Condiments", 3, "Confections", 4, "Dairy Products", 5, "Grains/Cereals", "Other")
```

### Example 3.2
**Exercise:** Assign a shipping priority based on the ship country (France: High, Germany: Medium, Brazil: Medium, USA: Standard, Others: Low).

**Solution:**
```
=SWITCH(C2, "France", "High", "Germany", "Medium", "Brazil", "Medium", "USA", "Standard", "Low")
```

### Example 3.3
**Exercise:** Determine the supplier name based on the supplier ID (1: Exotic Liquids, 2: New Orleans, 3: Tokyo Traders, 4: Cooperativa, Others: Unknown).

**Solution:**
```
=SWITCH(D2, 1, "Exotic Liquids", 2, "New Orleans", 3, "Tokyo Traders", 4, "Cooperativa", "Unknown")
```

## 4. SUMIF Function

**Sample Data:**
| ID   | Price| Category | Value  | Discontinued |
|------|------|----------|--------|--------------|
| 1    | 18   | Beverages| 270    | FALSE        |
| 2    | 19   | Condiments| 190   | FALSE        |
| 3    | 10   | Confections| 100  | FALSE        |
| 4    | 22   | Dairy Products| 440| FALSE        |
| 5    | 21.35| Grains/Cereals| 213.5 | TRUE     |

### Example 4.1
**Exercise:** Calculate the total value of products in the "Beverages" category.

**Solution:**
```
=SUMIF(C2:C6, "Beverages", D2:D6)
```

### Example 4.2
**Exercise:** Find the total value of products with a unit price over $20.

**Solution:**
```
=SUMIF(B2:B6, ">20", D2:D6)
```

### Example 4.3
**Exercise:** Calculate the total value of non-discontinued products.

**Solution:**
```
=SUMIF(E2:E6, FALSE, D2:D6)
```

## 5. AVERAGEIF Function

**Sample Data:**
| ID   | Price| Category | Stock | Supplier    |
|------|------|----------|-------|-------------|
| 1    | 18   | Beverages| 15    | Exotic Liquids |
| 2    | 19   | Condiments| 10   | New Orleans |
| 3    | 10   | Confections| 20  | Tokyo Traders |
| 4    | 22   | Dairy Products| 20| Cooperativa |
| 5    | 21.35| Grains/Cereals| 0 | Exotic Liquids |

### Example 5.1
**Exercise:** Find the average unit price of products supplied by "Exotic Liquids".

**Solution:**
```
=AVERAGEIF(E2:E6, "Exotic Liquids", B2:B6)
```

### Example 5.2
**Exercise:** Calculate the average units in stock for products in the "Confections" category.

**Solution:**
```
=AVERAGEIF(C2:C6, "Confections", D2:D6)
```

### Example 5.3
**Exercise:** Determine the average unit price of products that have stock (more than 0 units).

**Solution:**
```
=AVERAGEIF(D2:D6, ">0", B2:B6)
```

(The examples continue in this format for the remaining functions...)