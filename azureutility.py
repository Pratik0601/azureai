import os
import pandas as pd
import concurrent.futures
from azure.ai.formrecognizer import DocumentAnalysisClient
from azure.core.credentials import AzureKeyCredential

# Azure Form Recognizer credentials
ENDPOINT = ""
API_KEY = ""

# Initialize Form Recognizer client
client = DocumentAnalysisClient(endpoint=ENDPOINT, credential=AzureKeyCredential(API_KEY))

# Folder containing PDFs
input_folder = "C:/Users/Omkar/azureinvoice/pythonProject/invoice folder"  # Update with your folder path
output_folder = "C:/Users/Omkar/azureinvoice/pythonProject/output folder"  # Update with your output folder path

# Ensure output folder exists
os.makedirs(output_folder, exist_ok=True)


def process_invoice(pdf_path):
    """Extracts structured invoice data from a given PDF."""
    with open(pdf_path, "rb") as f:
        poller = client.begin_analyze_document("prebuilt-invoice", document=f)
        result = poller.result()

    structured_data = []

    for document in result.documents:
        invoice_dict = {}

        for field_name, field_value in document.fields.items():
            if "Address" in field_name:
                if field_value and field_value.value:
                    address = field_value.value
                    if isinstance(address, str):
                            invoice_dict[field_name] = address
                    else:
                        # Otherwise, assume it has the desired attributes.
                        invoice_dict[field_name] = (
                            f"{address.street_address}, {address.city}, {address.state} {address.postal_code}"
                        )
            else:
                invoice_dict[field_name] = field_value.value

        if "Items" in document.fields and document.fields["Items"].value:
            for idx, item in enumerate(document.fields["Items"].value, start=1):
                line_item_dict = invoice_dict.copy()
                line_item_dict["Item Number"] = idx
                for item_field_name, item_field_value in item.value.items():
                    line_item_dict[item_field_name] = item_field_value.value
                structured_data.append(line_item_dict)
        else:
            structured_data.append(invoice_dict)

    return pd.DataFrame(structured_data)


def process_and_save(pdf_path):
    """Processes a single invoice and saves its result as an Excel file."""
    file_name = os.path.basename(pdf_path)
    sheet_name = os.path.splitext(file_name)[0]

    print(f"Processing: {file_name}")
    df = process_invoice(pdf_path)

    excel_path = os.path.join(output_folder, f"{sheet_name}.xlsx")
    df.to_excel(excel_path, index=False)
    print(f"Saved: {excel_path}")


def main():
    """Processes all PDFs in the folder concurrently."""
    pdf_files = [os.path.join(input_folder, file) for file in os.listdir(input_folder) if file.lower().endswith(".pdf")]

    with concurrent.futures.ThreadPoolExecutor(max_workers=5) as executor:  # Adjust max_workers based on CPU cores
        executor.map(process_and_save, pdf_files)

    print("Processing completed for all PDFs.")


if __name__ == "__main__":
    main()
