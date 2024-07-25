import os
import pandas as pd

class LengthCalculator:
    def __init__(self, folder_path, existing_file):
        self.folder_path = folder_path
        self.existing_file = existing_file
        self.results = {}

    def process_files(self):
        for filename in os.listdir(self.folder_path):
            if filename.endswith('.csv'):
                self.process_file(filename)

    def process_file(self, filename):
        file_path = os.path.join(self.folder_path, filename)
        print(f"Processing file: {file_path}")
        
        try:
            df = pd.read_csv(file_path)
            df['Length'] = pd.to_numeric(df['Length'], errors='coerce')
            df = df.dropna(subset=['Length'])
            total_length = df['Length'].sum()
            clean_filename = filename.replace('_WELD REPORT.csv', '')
            self.results[clean_filename] = total_length
            print(f"Total Length for {clean_filename}: {total_length}")
        except Exception as e:
            print(f"Error processing file {filename}: {e}")

    def update_existing_file(self):
        try:
            existing_df = pd.read_excel(self.existing_file)
            print(existing_df.head())
            
            for index, row in existing_df.iterrows():
                piece_mark = row['Piece Mark']
                if piece_mark in self.results:
                    existing_df.at[index, 'Total Length'] = self.results[piece_mark]
            
            existing_df.to_excel(self.existing_file, index=False)
            print(f"Results have been written to {self.existing_file}")
        except Exception as e:
            print(f"Error updating the Excel file {self.existing_file}: {e}")

    def run(self):
        self.process_files()
        self.update_existing_file()

if __name__ == "__main__":
    folder_path = input("Enter the folder path: ")
    existing_file = input("Enter the full path to the existing Excel file: ").strip('"')
    calculator = LengthCalculator(folder_path, existing_file)
    calculator.run()
