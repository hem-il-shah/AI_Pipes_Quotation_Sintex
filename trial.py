import json
from azure.core.credentials import AzureKeyCredential
from azure.ai.formrecognizer import DocumentAnalysisClient

# Configuration
ENDPOINT = "https://doc-intel-ai-pipes-quotation.cognitiveservices.azure.com/"
KEY = "2xpainCqgQ9BxPcpMhEkAVBblsBXzMsbMHjZZVo7lmxftzqzUlj0JQQJ99CBACGhslBXJ3w3AAALACOGbGIY"
IMAGE_PATH = r"C:\Users\Shreyas Shah\Desktop\Pipes_Quotation_08042026\Image (2).jpg"

def analyze_quotation_table(image_path):
    client = DocumentAnalysisClient(ENDPOINT, AzureKeyCredential(KEY))

    with open(image_path, "rb") as f:
        poller = client.begin_analyze_document("prebuilt-layout", document=f)
    
    result = poller.result()
    detections = []

    for table in result.tables:
        # 1. Map column headers (OD Sizes)
        # Headers are usually in row 0 and 1. We need 15MM, 20MM, etc.
        header_map = {}
        for cell in table.cells:
            if cell.row_index == 1 and cell.column_index >= 2:
                header_map[cell.column_index] = cell.content.replace('\n', ' ')

        # 2. Iterate through rows to find quantities
        for cell in table.cells:
            # Skip headers and SKU/Description columns
            if cell.row_index > 1 and cell.column_index >= 2:
                content = cell.content.strip()
                
                # If we found a handwritten number
                if content and content.isdigit():
                    # Find the Description and SKU for this row
                    description = ""
                    sku = ""
                    for c in table.cells:
                        if c.row_index == cell.row_index:
                            if c.column_index == 0: description = c.content.replace('\n', ' ')
                            if c.column_index == 1: sku = c.content.replace('\n', ' ')
                    
                    size = header_map.get(cell.column_index, "Unknown Size")
                    
                    # Format the string as requested
                    detection_str = f"{description}: {sku} {size}: {content}"
                    detections.append({
                        "description": description,
                        "sku": sku,
                        "size": size,
                        "quantity": content,
                        "formatted": detection_str
                    })

    # Print the formatted list
    print("--- OCR Detections ---")
    for d in detections:
        print(d["formatted"])

    # Save to JSON
    output_file = "ocr_results.json"
    with open(output_file, "w") as jf:
        json.dump(detections, jf, indent=4)
    
    print(f"\nResults stored in {output_file}")

if __name__ == "__main__":
    analyze_quotation_table(IMAGE_PATH)