import pandas as pd
import numpy as np
import re
import matplotlib.pyplot as plt
import seaborn as sns
from colorama import Fore

# LOAD DATASET
file_path = "Process Notes Responses (1).xlsx"

df = pd.read_excel(file_path)

# Explore the dataset
print(Fore.GREEN + "Data loaded successfully. Shape:", df.shape)
print(Fore.BLUE + "Columns:", list(df.columns))
print(Fore.CYAN + "First 5 rows:")
print(df.head())

# Data Cleansing
print("\n--- Data Cleansing ---")

# Handle missing values
text_cols = ['Process', 'Evaluation', 'Planning']
for col in text_cols:
    if col in df.columns:
        df[col].fillna('Not provided', inplace=True)
        print(f"Filled missing values in '{col}' with 'Not provided'")

# Handle missing Session Date
if 'Session Date' in df.columns:
    df['Session Date'].fillna(method='ffill', inplace=True)  # Forward fill for dates
    print("Filled missing values in 'Session Date' using forward fill")

# Check for duplicates
duplicates = df.duplicated().sum()
if duplicates > 0:
    df.drop_duplicates(inplace=True)
    print(f"Removed {duplicates} duplicate rows")
else:
    print("No duplicate rows found")

# Clean text in text columns - remove extra whitespaces, standardize
def clean_text(text):
    if pd.isna(text):
        return text
    # Remove extra whitespaces
    text = re.sub(r'\s+', ' ', str(text).strip())
    return text

for col in text_cols:
    if col in df.columns:
        df[col] = df[col].apply(clean_text)
        print(f"Cleaned text in '{col}'")

# Ensure data types
if 'Session Date' in df.columns:
    df['Session Date'] = pd.to_datetime(df['Session Date'], errors='coerce')
    print("Ensured 'Session Date' is datetime")

# Convert Session Duration to numeric if possible
if 'Session Duration' in df.columns:
    # Extract numeric part (assuming format like "60 minutes" or just "60")
    df['Session Duration Numeric'] = df['Session Duration'].str.extract(r'(\d+)').astype(float)
    print("Extracted numeric session duration")

print("Cleaned data shape:", df.shape)

# Further Analysis
print("\n--- Further Analysis ---")

# Summary statistics
print("Data types:")
print(df.dtypes)

print("\nMissing values after cleansing:")
print(df.isnull().sum())

# Value counts for categorical columns
categorical_cols = ['Counsellor', 'Phase']

for col in categorical_cols:
    if col in df.columns:
        print(f"\nValue counts for {col}:")
        print(df[col].value_counts())

# Session Duration statistics
if 'Session Duration Numeric' in df.columns:
    print("\nSession Duration statistics:")
    print(df['Session Duration Numeric'].describe())

# Text analysis - word count for text columns
for col in text_cols:
    word_count_col = f'{col} Word Count'
    df[word_count_col] = df[col].apply(lambda x: len(str(x).split()) if x != 'Not provided' else 0)
    print(f"\n{col} word count statistics:")
    print(df[word_count_col].describe())

# Save cleaned data
df.to_excel('cleaned_process_notes.xlsx', index=False)
print("\nCleaned data saved to 'cleaned_process_notes.xlsx'")

# Save analysis results to text file
with open('process_notes_analysis_summary.txt', 'w') as f:
    f.write("Data Shape: " + str(df.shape) + "\n")
    f.write("Columns: " + str(list(df.columns)) + "\n\n")
    f.write("Data Types:\n" + str(df.dtypes) + "\n\n")
    f.write("Missing Values:\n" + str(df.isnull().sum()) + "\n\n")

    for col in categorical_cols:
        if col in df.columns:
            f.write(f"Value counts for {col}:\n")
            f.write(str(df[col].value_counts()) + "\n\n")

    if 'Age' in df.columns:
        f.write("Age Statistics:\n" + str(df['Age'].describe()) + "\n\n")

    if 'Process Notes Word Count' in df.columns:
        f.write("Process Notes Word Count Statistics:\n" + str(df['Process Notes Word Count'].describe()) + "\n\n")

print("Analysis summary saved to 'process_notes_analysis_summary.txt'")

# Data Visualization
print("\n--- Data Visualization ---")

# Set up matplotlib for saving plots
plt.style.use('seaborn-v0_8')
sns.set_palette('husl')

# Counsellor distribution bar chart
if 'Counsellor' in df.columns:
    plt.figure(figsize=(10, 6))
    df['Counsellor'].value_counts().plot(kind='bar')
    plt.title('Counsellor Distribution')
    plt.xlabel('Counsellor')
    plt.ylabel('Number of Sessions')
    plt.xticks(rotation=45)
    plt.savefig('counsellor_distribution_bar.png')
    plt.close()
    print("Saved counsellor distribution bar chart as 'counsellor_distribution_bar.png'")

# Phase distribution pie chart
if 'Phase' in df.columns:
    plt.figure(figsize=(8, 8))
    df['Phase'].value_counts().plot(kind='pie', autopct='%1.1f%%', startangle=140)
    plt.title('Phase Distribution')
    plt.ylabel('')
    plt.savefig('phase_distribution_pie.png')
    plt.close()
    print("Saved phase distribution pie chart as 'phase_distribution_pie.png'")

# Session Duration histogram
if 'Session Duration Numeric' in df.columns:
    plt.figure(figsize=(8, 6))
    sns.histplot(df['Session Duration Numeric'].dropna(), kde=True, bins=20)
    plt.title('Session Duration Distribution')
    plt.xlabel('Session Duration (minutes)')
    plt.ylabel('Frequency')
    plt.savefig('session_duration_histogram.png')
    plt.close()
    print("Saved session duration histogram as 'session_duration_histogram.png'")

# Sessions over time
if 'Session Date' in df.columns:
    df['Session Month'] = df['Session Date'].dt.to_period('M')
    plt.figure(figsize=(12, 6))
    df['Session Month'].value_counts().sort_index().plot(kind='line', marker='o')
    plt.title('Sessions Over Time')
    plt.xlabel('Month')
    plt.ylabel('Number of Sessions')
    plt.xticks(rotation=45)
    plt.savefig('sessions_over_time_line.png')
    plt.close()
    print("Saved sessions over time line chart as 'sessions_over_time_line.png'")

# Word count histograms for text columns
text_cols = ['Process', 'Evaluation', 'Planning']
for col in text_cols:
    word_count_col = f'{col} Word Count'
    if word_count_col in df.columns:
        plt.figure(figsize=(8, 6))
        sns.histplot(df[df[word_count_col] > 0][word_count_col], kde=True, bins=20)
        plt.title(f'{col} Word Count Distribution')
        plt.xlabel('Word Count')
        plt.ylabel('Frequency')
        plt.savefig(f'{col.lower()}_word_count_histogram.png')
        plt.close()
        print(f"Saved {col.lower()} word count histogram as '{col.lower()}_word_count_histogram.png'")

print("All visualizations saved as PNG files.")

