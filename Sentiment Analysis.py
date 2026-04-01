import pandas as pd
import openai
import re
from datetime import datetime

# Configure OpenAI API
openai.api_key = ""

def load_excel_file(file_path):
    """Load data from Excel file."""
    print("Loading Excel file...")
    try:
        df = pd.read_excel(file_path)
        print(f"Successfully loaded {len(df)} rows from Excel.")
        print("\nColumns in Excel file:")
        print(df.columns.tolist())
        return df
    except Exception as e:
        print(f"Error loading Excel file: {str(e)}")
        return None

def get_gpt_analysis(text_list):
    """Send text to GPT for analysis in batches."""
    print("Sending requests to OpenAI API...")
    all_responses = []
    
    # Process in batches of 5 sentences
    batch_size = 5
    for i in range(0, len(text_list), batch_size):
        batch = text_list[i:i + batch_size]
        batch_text = "\n".join(batch)
        
        try:
            response = openai.ChatCompletion.create(
                model="gpt-3.5-turbo",
                messages=[
                    {"role": "system", "content": """You are an ESG report analyst. Analyze the following sentences from a company's annual ESG report. 
                    For each sentence, you MUST provide ALL of the following information, inferring from context if necessary:

                    1. Sentiment (Positive/Negative/Neutral)
                    2. Brief explanation
                    3. Percentage change comparison with previous year (ALWAYS include + or - sign)
                    4. Percentage change comparison with 2019 baseline (ALWAYS include + or - sign)
                    
                    If a specific percentage is not mentioned:
                    - For no change, use "+0%"
                    - For increases without specific numbers, use "+1%"
                    - For decreases without specific numbers, use "-1%"
                    
                    Format your response EXACTLY as follows for each sentence:
                    Sentence 1:
                    Sentiment: [sentiment]
                    Brief explanation: [explanation]
                    Previous year change: [MUST include sign, e.g., +0%, +1%, -1%]
                    Baseline change: [MUST include sign, e.g., +0%, +1%, -1%]"""},
                    {"role": "user", "content": batch_text}
                ]
            )
            print(f"Received response for batch {i//batch_size + 1}")
            all_responses.append(response.choices[0].message.content)
            
        except Exception as e:
            print(f"Error with batch {i//batch_size + 1}: {str(e)}")
            all_responses.append("")
            
    return "\n".join(all_responses)

def extract_percentage(text, pattern_list, default="+0%"):
    """Helper function to extract percentage with multiple patterns."""
    for pattern in pattern_list:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            # Handle patterns with capture groups for increase/decrease
            if len(match.groups()) > 1 and match.group(2):
                value = match.group(2)
                sign = '+' if match.group(1).lower() == 'increased' else '-'
                return f"{sign}{value}%"
            # Ensure the percentage has a sign
            percentage = match.group(1)
            if not percentage.startswith('+') and not percentage.startswith('-'):
                return f"+{percentage}"
            return percentage
    return default

def parse_gpt_response(response_text):
    """Parse the GPT response with improved pattern matching."""
    print("Parsing the GPT response...")
    
    sentiments = []
    explanations = []
    prev_year_changes = []
    baseline_changes = []
    
    # Split into sentence sections
    sections = response_text.split("Sentence")[1:]
    
    for section in sections:
        if not section.strip():
            continue
            
        # Extract sentiment
        sentiment_patterns = [
            r'Sentiment:\s*(\w+)',
            r'sentiment is\s*(\w+)',
            r'sentiment:\s*(\w+)'
        ]
        sentiment = extract_percentage(section, sentiment_patterns, default="Neutral")
        sentiments.append(sentiment if sentiment != "+0%" else "Neutral")
        
        # Extract explanation
        explanation_match = re.search(r'Brief explanation:\s*(.*?)(?=Previous year change|Baseline change|Sentiment|\Z)', 
                                    section, 
                                    re.DOTALL | re.IGNORECASE)
        explanation = explanation_match.group(1).strip() if explanation_match else "Analysis of performance"
        explanations.append(explanation)
        
        # Extract previous year change
        prev_year_patterns = [
            r'Previous year change.*?([+-]?\d+\.?\d*%)',
            r'compared to (?:previous year|last year).*?([+-]?\d+\.?\d*%)',
            r'(increased|decreased) by (\d+\.?\d*)%.*?from.*?previous',
            r'change of ([+-]?\d+\.?\d*%)',
            r'([+-]?\d+\.?\d*%).*?compared to.*?previous'
        ]
        prev_year = extract_percentage(section, prev_year_patterns)
        prev_year_changes.append(prev_year)
        
        # Extract baseline change
        baseline_patterns = [
            r'Baseline change.*?([+-]?\d+\.?\d*%)',
            r'compared to 2019.*?([+-]?\d+\.?\d*%)',
            r'against.*?2019 baseline.*?([+-]?\d+\.?\d*%)',
            r'since 2019.*?([+-]?\d+\.?\d*%)',
            r'from.*?2019 baseline.*?([+-]?\d+\.?\d*%)'
        ]
        baseline = extract_percentage(section, baseline_patterns)
        baseline_changes.append(baseline)
    
    print("\nParsed Results:")
    print(f"Sentiments: {sentiments}")
    print(f"Explanations: {explanations}")
    print(f"Previous Year Changes: {prev_year_changes}")
    print(f"Baseline Changes: {baseline_changes}")
    
    return sentiments, explanations, prev_year_changes, baseline_changes

def prepare_text_for_analysis(df):
    """Prepare text from all columns for analysis."""
    all_texts = []
    
    # Sort DataFrame by Year to ensure chronological order
    df_sorted = df.sort_values('Year')
    
    for i, (_, row) in enumerate(df_sorted.iterrows(), 1):
        row_text = []
        for column in ['Aim', 'Year', 'Content']:
            value = str(row[column])
            if value != 'nan' and value.strip():
                row_text.append(f"{column}: {value}")
        
        if row_text:
            combined_text = " | ".join(row_text)
            all_texts.append(f"Sentence {i}: {combined_text}")
    
    return all_texts  # Return list instead of joined string

def update_excel(df, sentiments, explanations, prev_year_changes, baseline_changes, excel_file):
    """Update Excel file with analysis results."""
    print("Updating Excel file...")
    try:
        # Sort DataFrame by Year to match the analysis order
        df_sorted = df.sort_values('Year')
        
        # Ensure we have enough values for all rows
        num_rows = len(df_sorted)
        sentiments = (sentiments + ["Neutral"] * num_rows)[:num_rows]
        explanations = (explanations + ["Analysis of performance"] * num_rows)[:num_rows]
        prev_year_changes = (prev_year_changes + ["+0%"] * num_rows)[:num_rows]
        baseline_changes = (baseline_changes + ["+0%"] * num_rows)[:num_rows]
        
        # Add new columns with analysis results
        df_sorted['Sentiment'] = sentiments
        df_sorted['Explanation'] = explanations
        df_sorted['vs previous year'] = prev_year_changes
        df_sorted['vs 2019'] = baseline_changes
        df_sorted['Analysis_Date'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        
        # Save updated DataFrame to the same Excel file
        with pd.ExcelWriter(excel_file, mode='w', engine='openpyxl') as writer:
            df_sorted.to_excel(writer, index=False)
        print(f"Successfully updated {excel_file}")
    except Exception as e:
        print(f"Error updating Excel file: {str(e)}")

def main():
    # File path
    excel_file = "c:/Users/Sidarth/Desktop/Combo/TCD/Course/Data Mining/Group Project/Extra companies/TotalEnergies_Sustainability_Aims_Detailed_2015_2023  - SA.xlsx"  # Replace with your Excel file path
    
    # Load Excel file
    df = load_excel_file(excel_file)
    if df is None:
        return
    
    # Prepare text from all columns
    all_texts = prepare_text_for_analysis(df)
    print(f"\nPrepared {len(all_texts)} sentences for analysis")
    
    # Get GPT analysis
    gpt_response = get_gpt_analysis(all_texts)
    if gpt_response is None:
        return
    
    # Parse GPT response
    sentiments, explanations, prev_year_changes, baseline_changes = parse_gpt_response(gpt_response)
    
    # Update Excel file with results
    update_excel(df, sentiments, explanations, prev_year_changes, baseline_changes, excel_file)

if __name__ == "__main__":
    main()
