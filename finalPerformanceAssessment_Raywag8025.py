# ECPI University - Python Final Performance Assessment
# Student ID: Raywag8025
# Place this script beside the ORIGINAL final.csv in C:\PythonFiles.

import csv
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import PieChart, Reference
import matplotlib.pyplot as plt

STUDENT_ID = "Raywag8025"
DATA_DIR = Path(__file__).resolve().parent
CSV_PATH = DATA_DIR / "final.csv"
EXCEL_PATH = DATA_DIR / "final.xlsx"


# Read the names and incomes from final.csv without changing the existing file.
def readEntries():
    entries = []
    with open(CSV_PATH, "r", newline="", encoding="utf-8-sig") as file:
        reader = csv.reader(file)
        for row in reader:
            if not row or all(not value.strip() for value in row):
                continue
            if len(row) < 2:
                raise ValueError("Each CSV row must contain a name and an income.")
            name = row[0].strip()
            incomeText = row[1].strip().replace(",", "").replace("$", "")
            if name.lower() == "name" and incomeText.lower() in ("income", "annual income"):
                continue
            income = int(float(incomeText))
            if not name:
                raise ValueError("Names cannot be blank in final.csv.")
            entries.append((name, income))
    return entries


# Ask for five numbers and display their accumulated total.
def askUser():
    total = 0
    # This loop collects five numeric inputs, adds each to total, and then shows the sum.
    for numberIndex in range(5):
        while True:
            try:
                number = int(input("Please enter a number: "))
                total += number
                break
            except ValueError:
                print("Please enter a whole number.")
    print(f"The sum of the numbers entered is {total}")
    return total


# Ask for five names and annual incomes and append them to final.csv.
def askIncome():
    newEntries = []
    # This loop gathers five people and their incomes, retaining each entry for the CSV.
    for personIndex in range(5):
        name = input("Please enter a name: ").strip()
        while not name:
            name = input("Name cannot be blank. Please enter a name: ").strip()
        while True:
            try:
                income = int(input("Please enter their Income: "))
                if income < 0:
                    print("Please enter a nonnegative income.")
                    continue
                break
            except ValueError:
                print("Please enter a whole number for income.")
        newEntries.append((name, income))

    # The file already has four rows; appending adds five more without overwriting them.
    with open(CSV_PATH, "a", newline="", encoding="utf-8") as file:
        writer = csv.writer(file)
        writer.writerows(newEntries)
    print(f"Saved {len(newEntries)} new entries to {CSV_PATH}.")
    return newEntries


# Create an Excel workbook with all CSV records and an embedded pie chart.
def excelPie(entries):
    workbook = Workbook()  # Create an Excel workbook for the nine records.
    sheet = workbook.active  # Choose the worksheet that will contain the data and chart.
    sheet.title = "Income Data"  # Give the worksheet a meaningful name.
    sheet.append(["Name", "Income"])  # Add headers above the data for Excel.
    for name, income in entries:
        sheet.append([name, int(income)])  # Save each name and numeric income in separate cells.
    sheet.column_dimensions["A"].width = 22  # Make names easier to read.
    sheet.column_dimensions["B"].width = 17  # Make income values easier to read.
    for cells in sheet.iter_rows(min_row=2, min_col=2, max_col=2):
        cells[0].number_format = '#,##0'  # Display incomes as readable whole numbers.

    pieChart = PieChart()  # Create the pie-chart object.
    pieChart.title = f"{STUDENT_ID} {datetime.now():%B} {datetime.now().day}, {datetime.now().year}"  # Title with ID and today's date.
    pieChart.style = 10  # Apply a built-in chart style.
    pieChart.height = 10  # Ensure the chart is tall enough to see the slices.
    pieChart.width = 16  # Ensure the chart is wide enough for labels and legend.
    incomeValues = Reference(sheet, min_col=2, min_row=1, max_row=len(entries) + 1)  # Select incomes including the header.
    nameLabels = Reference(sheet, min_col=1, min_row=2, max_row=len(entries) + 1)  # Select names without the header.
    pieChart.add_data(incomeValues, titles_from_data=True)  # Add the numeric income series.
    pieChart.set_categories(nameLabels)  # Label slices with the corresponding names.
    sheet.add_chart(pieChart, "A12")  # Position the Excel pie chart below the nine rows.
    workbook.save(EXCEL_PATH)  # Save chart and data together as final.xlsx.
    print(f"Excel pie chart saved to {EXCEL_PATH}.")


# Display a vertical bar graph of the same name-and-income data using Matplotlib.
def verticalBar(entries):
    names = [name for name, income in entries]
    incomes = [income for name, income in entries]
    plt.figure(figsize=(10, 6))
    plt.bar(names, incomes, color="green", label="Income")
    plt.title(f"{STUDENT_ID} {datetime.now():%B} {datetime.now().day}, {datetime.now().year}")
    plt.xlabel("Names")
    plt.ylabel("Income")
    plt.xticks(rotation=0, fontsize=9)
    plt.legend()
    plt.tight_layout()
    print("Matplotlib vertical bar graph is now displayed.")
    plt.show()


# Run the six assessment tasks in the required order.
def main():
    print(STUDENT_ID)
    if not CSV_PATH.exists():
        print(f"ERROR: Missing {CSV_PATH}")
        print("Download the ORIGINAL final.csv from Canvas and put it beside this script in C:\\PythonFiles.")
        return

    try:
        originalEntries = readEntries()
    except (OSError, ValueError) as error:
        print(f"Unable to read final.csv: {error}")
        return

    if len(originalEntries) != 4:
        print(f"SAFETY CHECK: final.csv contains {len(originalEntries)} records; expected four original records.")
        print("Restore the original Canvas final.csv before running this assessment again.")
        print("Do not delete the original four records.")
        return

    askUser()
    try:
        askIncome()
        allEntries = readEntries()
        if len(allEntries) != 9:
            raise ValueError(f"Expected nine records after appending; found {len(allEntries)}.")
        excelPie(allEntries)
        verticalBar(allEntries)
    except (OSError, ValueError) as error:
        print(f"ERROR: {error}")


if __name__ == "__main__":
    main()
