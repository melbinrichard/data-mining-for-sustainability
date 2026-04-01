import fitz  # PyMuPDF
import re
import pandas as pd
import os

# Loop through years from 2020 to 2023
for year in range(2020, 2024):
    # Open the PDF file
    file_path = f"Annual Report Extract/BP/{year}.pdf"

    # Check if the file exists
    if not os.path.exists(file_path):
        print(f"File not found: {file_path}. Skipping...")
        continue  # Skip this iteration if the file doesn't exist

    doc = fitz.open(file_path)

    # Variables to store the relevant text for each aim
    aims_text = {f"aim_{i}_texts": [] for i in range(1, 6)}
    capture_flags = {f"capture_aim_{i}": False for i in range(1, 6)}  # Flags to capture text for each aim

    # Function to clean illegal characters from text
    def clean_text(text):
        # Remove all non-printable characters
        return re.sub(r'[^\x20-\x7E]+', '', text)

    # Function to extract text related to specific aims
    def extract_text_related_to_aim(page):
        text_info = page.get_text("dict")  # Get detailed text information

        # Define regular expression patterns for each aim and the next section headers
        aim_patterns = {f"aim_{i}_pattern": re.compile(rf'\bAim\s?{i}\b', re.IGNORECASE) for i in range(1, 6)}
        next_section_pattern = re.compile(r'\bAim\s?\d+\b|\bAppendix\b|\bGlossary\b', re.IGNORECASE)

        relevant_texts = {f"aim_{i}_texts": [] for i in range(1, 6)}

        for block in text_info['blocks']:
            if 'lines' in block:  # Check if the block contains lines
                for line in block['lines']:
                    line_text = ""
                    line_font_size = 0  # Initialize font size
                    for span in line['spans']:
                        line_text += span['text']  # Combine text spans
                        line_font_size = max(line_font_size, span['size'])  # Get the maximum font size in the line

                    # Check font size to skip small footnotes or headers
                    if line_font_size < 7.5:
                        continue  # Skip lines with font size less than 7

                    # Update capture flags based on detected patterns
                    for i in range(1, 6):
                        if aim_patterns[f"aim_{i}_pattern"].search(line_text):  # Start capturing for the current aim
                            for j in range(1, 6):
                                capture_flags[f"capture_aim_{j}"] = (i == j)  # Only activate the current aim's capture flag
                            break
                    else:
                        if next_section_pattern.search(line_text):  # If another section is found, stop capturing
                            for j in range(1, 6):
                                capture_flags[f"capture_aim_{j}"] = False

                    # Append text to the respective list based on active capture flags
                    for i in range(1, 6):
                        if capture_flags[f"capture_aim_{i}"]:
                            relevant_texts[f"aim_{i}_texts"].append(line_text.strip())

        return relevant_texts

    # Extract text across all pages for all aims
    for page_num in range(doc.page_count):
        page = doc[page_num]
        page_texts = extract_text_related_to_aim(page)
        for i in range(1, 6):
            aims_text[f"aim_{i}_texts"].extend(page_texts[f"aim_{i}_texts"])

    # Combine the extracted text into a single string for easier reading
    aims_content = {f"aim_{i}_content": " ".join(aims_text[f"aim_{i}_texts"]) for i in range(1, 6)}

    # Clean the content for each aim
    aims_content = {key: clean_text(content) for key, content in aims_content.items()}

    # Prepare data for saving
    data_to_save = {
        'Aim': [f"Aim {i}" for i in range(1, 6)],
        'Year': [year] * 5,
        'Content': [aims_content[f"aim_{i}_content"] for i in range(1, 6)]
    }
    df = pd.DataFrame(data_to_save)

    # Define the output Excel file path
    output_path = "../Paragraph Extract/Aims_BP.xlsx"

    # Append the new data to the Excel file
    if os.path.exists(output_path):
        # Load existing data
        existing_df = pd.read_excel(output_path)
        # Concatenate the new data
        updated_df = pd.concat([existing_df, df], ignore_index=True)
    else:
        # If the file doesn't exist, use the new DataFrame
        updated_df = df

    # Save the updated DataFrame to an Excel file
    updated_df.to_excel(output_path, index=False)

    print(f"Extracted content related to Aims 1 through 5 for {year} saved to {output_path}.")
