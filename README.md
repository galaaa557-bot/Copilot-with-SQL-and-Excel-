# Copilot-with-SQL-and-Excel-
Copilot with SQL and Excel   

Question 1: Use GitHub Copilot in SSMS to create a table SalesData whose columns are CustomerID, Name, Age, City, PurchaseAmount, PurchaseDate and then ask to insert 10,000 rows of random data.
solution:- 

-- Step 1: Create the SalesData table
CREATE TABLE SalesData (
    CustomerID INT PRIMARY KEY IDENTITY(1,1),
    Name VARCHAR(100),
    Age INT,
    City VARCHAR(100),
    PurchaseAmount DECIMAL(10, 2),
    PurchaseDate DATE
);

-- Step 2: Insert 10,000 rows of random data
WITH CTE AS (
    SELECT 1 AS RowNum
    UNION ALL
    SELECT RowNum + 1 FROM CTE WHERE RowNum < 10000
)
INSERT INTO SalesData (Name, Age, City, PurchaseAmount, PurchaseDate)
SELECT 
    'Customer_' + CAST(ABS(CHECKSUM(NEWID())) % 10000 AS VARCHAR),
    20 + (ABS(CHECKSUM(NEWID())) % 45), -- Age between 20 and 64
    CHOOSE((ABS(CHECKSUM(NEWID())) % 5) + 1, 'New York', 'Los Angeles', 'Chicago', 'Houston', 'Phoenix'),
    ROUND((RAND(CHECKSUM(NEWID())) * 990) + 10, 2), -- Purchase between 10.00 and 1000.00

Question 2: Use Copilot to find total sales per city and top 5 cities by revenue.

Sol. 
-- Find total sales per city
SELECT 
    City, 
    SUM(PurchaseAmount) AS TotalSales
FROM SalesData
GROUP BY City;

-- Top 5 cities by revenue
SELECT TOP 5 
    City, 
    SUM(PurchaseAmount) AS TotalRevenue
FROM SalesData
GROUP BY City
ORDER BY TotalRevenue DESC;

Question 3: Ask Copilot to find customers with purchases above average.
 
Sol. 
-- Find customers whose purchase amount is strictly above the overall average
SELECT 
    CustomerID, 
    Name, 
    PurchaseAmount, 
    PurchaseDate
FROM SalesData
WHERE PurchaseAmount > (SELECT AVG(PurchaseAmount) FROM SalesData);

Copilot in Excel

Question 4: Identify and handle missing values in the dataset using Copilot.
Answer:
Identification:
"Highlight all rows or cells containing missing or blank values across the table."
Findings: The dataset contains missing values across multiple columns: Customer_Name (68 missing), Region (192 missing), Product (105 missing), Sales (167 missing), and Quantity (90 missing).
Handling: 
"Fill blank cells in 'Sales' and 'Quantity' with 0, and fill missing text values in 'Region' and 'Product' with 'Unknown'."

Question 5: Identify duplicate rows and remove duplicates using Copilot suggestion.
Answer:
Prompt to Copilot: 
"Identify duplicate rows in this dataset and remove them."
Execution: Copilot identifies 51 duplicate rows in the dataset. Using the Remove Duplicates suggestion, Excel removes the 51 redundant rows, leaving 499 unique records.

Question 6: Generate a Column chart OR bar chart and Interpret which region has highest sales.
Answer:
Prompt to Copilot: 
"Generate a Column Chart showing total Sales by Region."
Interpretation:
West region generated the highest total sales ($16,700).
North region follows with $16,500.
South region recorded $16,000.
East region recorded $15,600.

Question 7: Create a pivot table showing total sales by region and product.
Answer:
Prompt to Copilot: 
"Create a Pivot Table showing total Sales with Region as rows and Product as columns."
Resulting Pivot Table Matrix:

Region
Laptop
Monitor
Phone
Tablet
Total
East
$1,900
$4,900
$3,100
$2,700
$12,600
North
$2,400
$4,600
$3,400
$2,900
$13,300
South
$4,100
$2,900
$2,900
$3,500
$13,400
West
$4,300
$3,600
$5,600
$2,200
$15,700


Question 8: Use Copilot to generate a new calculated column Total Sales.
Answer:
Prompt to Copilot: 
"Add a new column named 'Total Sales' that multiplies 'Sales' by 'Quantity'."
Formula Applied: = [@Sales] * VALUE([@Quantity])

Question 9: Summarize the dataset and find the key insights.
Answer:
Total Records & Completeness: The original dataset contains 550 entries with 51 full duplicate rows and missing data across several key columns (Region, Sales, Quantity).
Revenue Breakdown: Total recorded baseline sales sum to $97,700.
Top Regional Performer: West leads overall regional sales at $16,700, driven heavily by Phone sales ($5,600).
Product Insights: Phones represent the top revenue generator in the West region, whereas Laptops perform best in the South region ($4,100). Monitors represent the highest sales share in both East ($4,900) and North ($4,600) regions.

Question 10: Generate VBA code using Copilot for cleaning data (removing duplicates, filling null values, format table). 
Sol. 
Sub CleanData()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' 1. Remove duplicate rows
    ws.UsedRange.RemoveDuplicates Columns:=Array(1, 2, 3, 4, 5, 6), Header:=xlYes
    
    ' 2. Fill blank values in numeric and text columns
    Dim lastRow As Long, lastCol As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    Dim cell As Range
    For Each cell In ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol))
        If IsEmpty(cell) Then
            If IsNumeric(cell.Offset(0, -1).Value) Then
                cell.Value = 0
            Else
                cell.Value = "N/A"
            End If
        End If
    Next cell
    
    ' 3. Format as standard table
    Dim rng As Range
    Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
    ws.ListObjects.Add(xlSrcRange, rng, , xlYes).Name = "CleanedDataSet"
    
    ' 4. Auto-fit columns
    ws.Columns.AutoFit
    
    MsgBox "Data cleaning complete!", vbInformation
End Sub


    
    DATEADD(DAY, - ABS(CHECKSUM(NEWID())) % 365, GETDATE()) -- Date within the past year
FROM CTE
OPTION (MAXRECURSION 10000);
