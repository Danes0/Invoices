from docx import Document
from openpyxl import Workbook
#import libraries

#function to read the invoices
def readInvoices(numFiles):
    invoices_data = []
    for i in range(numFiles):
        invoiceNum = "100" + str(i).zfill(4)
        doc_file = "INV" + invoiceNum + ".docx"
        print(f"Reading file: {doc_file}")
        doc = Document(doc_file)
        # Initialize invoice data dictionary with fields
        invoice_data = {
            "Invoice Number": "INV" + invoiceNum + ".docx",
            "Total Quantity": 0,
            "Subtotal": 0,
            "Tax": 0,
            "Total": 0,
        }
        # Iterar sobre los párrafos del documento
        for paragraph in doc.paragraphs:
            # Iterate the first paragraph "products"
            if "PRODUCTS" in paragraph.text:
                for line in paragraph.text.split("\n"):
                    if ":" in line:
                        product, quantity = line.split(":")
                        invoice_data["Total Quantity"] += int(quantity)
        else:
            # Iterate the second paragraph line to line to find subtotal, tax, and total
            # Extracts the total value from total_line and splits it using ':' as a delimiter, then selects the second part
            for line in paragraph.text.split("\n"):
                if "SUBTOTAL" in line:
                    subtotal_line = line
                    subtotal_value = subtotal_line.split(":")[1]
                    invoice_data["Subtotal"] = subtotal_value
                if "TAX" in line:
                    tax_line = line
                    tax_value = tax_line.split(":")[1]
                    invoice_data["Tax"] = tax_value
                if "TOTAL" in line:
                    total_line = line
                    total_value = total_line.split(":")[1]
                    invoice_data["Total"] = total_value
            invoices_data.append(invoice_data)
    return invoices_data

# Function to write invoice data to an Excel file
def writeExcel(invoices_data, excel_file):
    # Create a new Excel
    wb = Workbook()
    # Select the active sheet
    ws = wb.active
    # Append a row with the column headers to the active sheet
    ws.append(["Invoice Number", "Total Quantity", "Subtotal", "Tax", "Total"])
    for invoice_data in invoices_data:
        ws.append(
            [
                invoice_data["Invoice Number"],
                invoice_data["Total Quantity"],
                invoice_data["Subtotal"],
                invoice_data["Tax"],
                invoice_data["Total"],
            ]
        )
    # Save the Excel file
    wb.save(excel_file)

# Read invoice data
invoices_data = readInvoices(2)
# Write the data to an Excel file
writeExcel(invoices_data, "invoices.xlsx")
