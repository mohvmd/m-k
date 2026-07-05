import pandas as pd
import numpy as np

class ExcelHandler:
    def __init__(self, file_path):
        self.file_path = file_path
        self.df = None
        self.load_data()

    def load_data(self):
        """Loads all sheets of the Excel file, keeping track of sheet name and row number."""
        try:
            xls = pd.ExcelFile(self.file_path)
            sheets_to_combine = []
            
            for sheet_name in xls.sheet_names:
                df_sheet = pd.read_excel(xls, sheet_name=sheet_name, dtype=str)
                # Clean column names (strip whitespace)
                df_sheet.columns = df_sheet.columns.astype(str).str.strip()
                
                # Check if this sheet has potential site database columns
                has_site_cols = any(col in df_sheet.columns for col in ['SiteCode', 'Site Code', 'Site ID', 'sitecode'])
                if has_site_cols:
                    # Add metadata columns
                    df_sheet['_sheet_name'] = sheet_name
                    # Excel row number is index + 2 (1-based index + 1 for header)
                    df_sheet['_row_number'] = df_sheet.index + 2
                    sheets_to_combine.append(df_sheet)
            
            if sheets_to_combine:
                # Combine all sheets
                self.df = pd.concat(sheets_to_combine, ignore_index=True)
            else:
                # Fallback to loading first sheet if no sheets match standard column headers
                self.df = pd.read_excel(self.file_path, dtype=str)
                self.df.columns = self.df.columns.astype(str).str.strip()
                self.df['_sheet_name'] = xls.sheet_names[0]
                self.df['_row_number'] = self.df.index + 2
                
        except Exception as e:
            raise Exception(f"خطأ أثناء تحميل ملف Excel: {str(e)}")

    def search_sites(self, site_codes):
        """
        Searches for a list of site codes in the combined DataFrame.
        Returns a list of dictionaries containing the details of the matching sites with sheet metadata.
        """
        if self.df is None:
            return []

        # Find target column for SiteCode matching
        target_col = None
        for col in ['SiteCode', 'Site Code', 'sitecode', 'site code']:
            if col in self.df.columns:
                target_col = col
                break
        
        if not target_col:
            for col in self.df.columns:
                if 'site' in col.lower():
                    target_col = col
                    break
            if not target_col:
                target_col = self.df.columns[0]

        results = []
        cleaned_series = self.df[target_col].astype(str).str.strip().str.upper()

        for code in site_codes:
            code_clean = str(code).strip().upper()
            if not code_clean:
                continue

            # 1. Direct search in SiteCode
            matches = self.df[cleaned_series == code_clean]
            
            # 2. Fallback search by Site ID (if direct search fails and Site ID column exists)
            if matches.empty and 'Site ID' in self.df.columns:
                digits_only = ''.join(filter(str.isdigit, code_clean))
                if digits_only:
                    cleaned_site_id = self.df['Site ID'].astype(str).str.strip().str.upper()
                    matches = self.df[cleaned_site_id == digits_only]
                    if matches.empty:
                        padded = digits_only.zfill(4)
                        matches = self.df[cleaned_site_id == padded]

            for _, row in matches.iterrows():
                row_dict = row.replace({np.nan: ''}).to_dict()
                row_dict['_matched_code'] = code
                results.append(row_dict)

        # De-duplicate results strictly by SiteCode value
        unique_results = []
        seen_sites = set()
        for res in results:
            site_val = str(res.get(target_col, res.get('SiteCode', ''))).strip().upper()
            if site_val and site_val not in seen_sites:
                seen_sites.add(site_val)
                unique_results.append(res)

        return unique_results
