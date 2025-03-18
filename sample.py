
import pandas as pd
from azure.ai.formrecognizer import DocumentAnalysisClient
from azure.core.credentials import AzureKeyCredential

# Azure Form Recognizer credentials
endpoint = ""
api_key = ""

# Initialize Form Recognizer client
client = DocumentAnalysisClient(endpoint=endpoint, credential=AzureKeyCredential(api_key))

# Path to your PDF invoice
invoice_path ="C:/Users/Omkar/PycharmProjects/langchainuitlity/3402024.pdf"

# Read the PDF file
with open(invoice_path, "rb") as f:
    poller = client.begin_analyze_document("prebuilt-invoice", document=f)
    result = poller.result()

# List to store final structured data
structured_data = []

# Loop through each invoice detected in the PDF
for document in result.documents:
    invoice_dict = {}

    # Extract invoice-level fields
    for field_name, field_value in document.fields.items():
        if field_name in ["BillingAddress", "ShippingAddress", "VendorAddress"]:
            if field_value and field_value.value:
                address = field_value.value
                invoice_dict[field_name] = f"{address.street_address}, {address.city}, {address.state} {address.postal_code}"
        else:
            invoice_dict[field_name] = field_value.value

    # Extracting line items and appending them to the final structured data
    if "Items" in document.fields and document.fields["Items"].value:
        for idx, item in enumerate(document.fields["Items"].value, start=1):
            line_item_dict = invoice_dict.copy()  # Copy invoice details for each line item
            line_item_dict["Item Number"] = idx  # Add item number
            for item_field_name, item_field_value in item.value.items():
                line_item_dict[item_field_name] = item_field_value.value
            structured_data.append(line_item_dict)
    else:
        # If no line items exist, add just the invoice details
        structured_data.append(invoice_dict)

# Convert structured data into a Pandas DataFrame
df = pd.DataFrame(structured_data)

# File path for the output
excel_path = "extracted_invoice_data11.xlsx"

# Save to Excel (single sheet)
df.to_excel(excel_path, index=False)

print(f"Data successfully extracted and saved to {excel_path}")