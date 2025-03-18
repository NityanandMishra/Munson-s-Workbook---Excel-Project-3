# Munson-s-Workbook-Excel-Project-3
Project Overview

This project involves working with multiple worksheets to perform various Excel tasks such as importing data, formatting, creating tables, using formulas, and visualizing data with charts. The goal is to organize, calculate, and present the shipping costs effectively while applying advanced Excel functionalities.

Project Structure

The workbook contains the following worksheets:

Outsourcing: Contains imported data from an external text file.

Flowers: Holds shipping costs and calculated discounts.

Shipping Cost: Displays shipping costs across different regions and shipping lines.

Cost Chart: A chart visualizing the shipping cost comparison.

Tasks Completed

1. Excel Customization

Added Spelling and Insert Function commands to the Quick Access toolbar.

Moved the Quick Access toolbar below the ribbon.

2. Data Import and Management

Imported Outsourcing_text.txt as a table in a new worksheet and renamed it to Outsourcing.

Converted the Outsourcing table into a regular cell range.

Deleted the following from the Outsourcing worksheet:

Blank cells: E4 and F9

Rows: 1 and 10

3. Table Creation and Modification

Created a table from the data in the Flowers worksheet.

Deleted conditional formatting from the Miles to Munson's column (D2:D8).

4. Data Formatting and Cell Manipulation

Region Workbook:

Used Format Painter to copy column heading formatting to the Flowers worksheet.

Closed the Region workbook.

Shipping Cost Worksheet:

Resized column widths:

Column A: 8pt

Columns D-G: 14pt (Aligned vertically and horizontally)

Applied General Number Format to Miles to Munson's (D2:D8).

Used Conditional Formatting:

Highlighted cells over 1,700 in red.

Highlighted cells under 1,000 in green.

Merged cells A1:H1, titled "Shipping Cost" in 16pt font.

Vertically aligned the "Zone" header in A2.

Wrapped text in D2:G2.

5. Formula Application

Flowers Worksheet:

Added a "Discount" column to calculate a 10% discount on the shipping cost.

Defined named ranges:

"cost": J2:J28 (Shipping Cost)

"sale": K2:K28 (Discount)

Used these ranges to calculate the final shipping cost in a new column labeled "Discounted Shipping Cost".

Freeze Panes at Column C.

Displayed all formulas.

Shipping Cost Worksheet:

Used COUNT functions to:

Count flower shipping costs less than 200 in cell C11.

Count flower shipping costs greater than 500 in cell C12.

Applied absolute cell references to calculate service fees for Edison, Washington (E14:G16).

6. Data Sorting and Filtering

Shipping Cost Worksheet:

Sorted the Red Line Shipping Cost (Red values at the top).

Sorted Miles to Munson's from lowest to highest.

Filtered records to display only Green Redline Shipping Cost.

7. Data Visualization

Created a 3-D Clustered Bar Chart for Red Line and Blue Line shipping costs.

Moved the chart to a new worksheet named "Cost Chart".

Added Green Line Shipping Cost to the chart.

Changed the chart to a Stacked Bar Chart with Chart Layout 2.

Applied Chart Style 8.

Updated chart elements:

Changed the Chart Title to "Shipping Cost".

Removed the Legend.

Added Alternative Text to the Shipping Cost Chart.

8. Sparklines

Inserted a Sparkline in H3 (Data Range: E3:G3).

Auto-filled Sparklines for H4:H9.

9. Navigation and Workbook Properties

Used Go To to navigate to the "flower" range.

To be completed:

Add workbook properties:

Title: "Munson's Workbook"

Keywords: "Flowers, Shipping Cost"

10. Printing and Page Setup

Configured each worksheet for printing:

Print with Gridlines.

Custom Header:

Left: Date

Right: Sheet Name

Centered on page Horizontally and Vertically.

Fit to 1 page.

11. Document Inspection

To be completed:

Inspect the document for comments and hidden worksheets.

Save and close the workbook.

Instructions for Use

Open Project3_datafile.xlsx.

Navigate to respective worksheets to view data, formulas, and charts.

Ensure the workbook properties and document inspection tasks are completed before final submission.

Requirements

Microsoft Excel (2019 or later)

Basic understanding of Excel formulas and functions

Future Improvements

Automate repetitive tasks using Excel macros.

Enhance data visualization with dynamic charts.

Improve workbook accessibility by adding detailed alt-text for all visual elements.

