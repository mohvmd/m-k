import os
import json
import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk
from PIL import Image

# Import our modules
from excel_handler import ExcelHandler
from photo_processor import PhotoProcessor

# Set theme and appearance
ctk.set_appearance_mode("Dark")  # Modes: "System", "Dark", "Light"
ctk.set_default_color_theme("blue")  # Themes: "blue", "green", "dark-blue"

CONFIG_FILE = "config.json"

class SitePhotoMatcherApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Configure window
        self.title("Faqous Communication")
        self.geometry("1200x750")
        self.minsize(1000, 650)

        # Initialize core logic variables
        self.excel_handler = None
        self.photo_processor = PhotoProcessor()
        self.excel_path = ""
        self.processed_images = []  # list of dicts: {'path': str, 'codes': list}
        self.matched_sites = []     # list of dicts of matching site records
        self.unmatched_codes = []   # list of strings of codes not found in Excel
        self.selected_site = None   # dict of the currently selected site

        # Load saved configuration (like Excel path)
        self.load_config()

        # Build GUI Layout
        self.create_sidebar()
        self.create_main_content()

        # Load Excel if path was saved
        if self.excel_path and os.path.exists(self.excel_path):
            self.load_excel_database(self.excel_path)

    def load_config(self):
        """Loads configuration file containing previous paths."""
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    self.excel_path = config.get("excel_path", "")
            except Exception as e:
                print(f"Error loading config: {e}")

    def save_config(self):
        """Saves current configuration to file."""
        try:
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump({"excel_path": self.excel_path}, f, ensure_ascii=False, indent=4)
        except Exception as e:
            print(f"Error saving config: {e}")

    def create_sidebar(self):
        """Creates the sidebar containing database load buttons, processed images list, and stats."""
        # Sidebar Frame
        self.sidebar_frame = ctk.CTkFrame(self, width=280, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, rowspan=2, sticky="nsew")
        self.grid_columnconfigure(0, weight=0)
        self.grid_rowconfigure(0, weight=1)

        # Title / Logo
        self.logo_label = ctk.CTkLabel(self.sidebar_frame, text="قاهر المنشن", 
                                       font=ctk.CTkFont(size=20, weight="bold", family="Helvetica"))
        self.logo_label.pack(pady=(20, 5), padx=20)
        
        self.subtitle_label = ctk.CTkLabel(self.sidebar_frame, text="Prod by Mohamed Ayman Samir\n" \
        " Eng/ Mohand Ahmed", 
                                          font=ctk.CTkFont(size=15, family="Helvetica"), text_color="gray")
        self.subtitle_label.pack(pady=(0, 20), padx=20)

        # Excel Database Section
        self.db_label = ctk.CTkLabel(self.sidebar_frame, text="قاعدة البيانات (Excel):", 
                                     font=ctk.CTkFont(size=14, weight="bold"))
        self.db_label.pack(anchor="w", padx=20, pady=(10, 5))

        self.db_status_label = ctk.CTkLabel(self.sidebar_frame, text="غير محملة ", 
                                            text_color="#ff5555", font=ctk.CTkFont(size=12))
        self.db_status_label.pack(anchor="w", padx=20, pady=2)

        self.btn_load_db = ctk.CTkButton(self.sidebar_frame, text="تحميل ملف Excel", 
                                         command=self.select_excel_file, 
                                         fg_color="#1f538d", hover_color="#14375e")
        self.btn_load_db.pack(fill="x", padx=20, pady=10)

        # Divider
        self.divider = ctk.CTkFrame(self.sidebar_frame, height=2, fg_color="#3a3a3a")
        self.divider.pack(fill="x", padx=20, pady=15)

        # Stats section
        self.stats_label = ctk.CTkLabel(self.sidebar_frame, text="الإحصائيات الحالية:", 
                                        font=ctk.CTkFont(size=14, weight="bold"))
        self.stats_label.pack(anchor="w", padx=20, pady=(5, 5))

        self.stat_images_label = ctk.CTkLabel(self.sidebar_frame, text="الصور المعالجة: 0", font=ctk.CTkFont(size=13))
        self.stat_images_label.pack(anchor="w", padx=20, pady=2)

        self.stat_matched_label = ctk.CTkLabel(self.sidebar_frame, text="مواقع متطابقة: 0", font=ctk.CTkFont(size=13), text_color="#55ff55")
        self.stat_matched_label.pack(anchor="w", padx=20, pady=2)

        self.stat_unmatched_label = ctk.CTkLabel(self.sidebar_frame, text="مواقع غير موجودة: 0", font=ctk.CTkFont(size=13), text_color="#ffbb55")
        self.stat_unmatched_label.pack(anchor="w", padx=20, pady=2)

        # Divider 2
        self.divider2 = ctk.CTkFrame(self.sidebar_frame, height=2, fg_color="#3a3a3a")
        self.divider2.pack(fill="x", padx=20, pady=15)

        # List of processed images
        self.images_list_label = ctk.CTkLabel(self.sidebar_frame, text="الصور المرفوعة:", 
                                             font=ctk.CTkFont(size=14, weight="bold"))
        self.images_list_label.pack(anchor="w", padx=20, pady=(5, 5))

        self.images_scroll_frame = ctk.CTkScrollableFrame(self.sidebar_frame, height=180, label_text="")
        self.images_scroll_frame.pack(fill="both", expand=True, padx=20, pady=5)

        # Permanent Dark Mode footer spacer
        self.footer_spacer = ctk.CTkLabel(self.sidebar_frame, text="", font=ctk.CTkFont(size=12))
        self.footer_spacer.pack(side="bottom", pady=10)

    def create_main_content(self):
        """Creates the main area: Actions bar, match results list, and site details view."""
        self.main_container = ctk.CTkFrame(self, fg_color="transparent")
        self.main_container.grid(row=0, column=1, sticky="nsew", padx=20, pady=20)
        self.grid_columnconfigure(1, weight=1)

        # 1. Action Bar (Upload Photos, Search, Clear)
        self.action_bar = ctk.CTkFrame(self.main_container, height=60, fg_color="#2b2b2b", corner_radius=10)
        self.action_bar.pack(fill="x", pady=(0, 15))

        self.btn_upload_photos = ctk.CTkButton(self.action_bar, text=" رفع صور", 
                                               command=self.upload_photos,
                                               fg_color="#2da44e", hover_color="#22863a",
                                               font=ctk.CTkFont(size=14, weight="bold"))
        self.btn_upload_photos.pack(side="right", padx=15, pady=15)

        self.btn_clear = ctk.CTkButton(self.action_bar, text="مسح النتائج", 
                                       command=self.clear_results,
                                       fg_color="#d73a49", hover_color="#cb2431",
                                       font=ctk.CTkFont(size=13, weight="bold"))
        self.btn_clear.pack(side="left", padx=15, pady=15)

        # Search Frame in the center of Action Bar
        self.search_frame = ctk.CTkFrame(self.action_bar, fg_color="transparent")
        self.search_frame.pack(side="right", padx=(0, 40), pady=15)

        self.search_entry = ctk.CTkEntry(self.search_frame, placeholder_text="ابحث عن موقع", width=220, justify="right")
        self.search_entry.pack(side="right", padx=5)
        self.search_entry.bind("<Return>", lambda event: self.manual_search())

        self.btn_search = ctk.CTkButton(self.search_frame, text="بحث ", command=self.manual_search, width=70, fg_color="#1f538d", hover_color="#14375e")
        self.btn_search.pack(side="right", padx=5)

        # 2. Main Workspace (Cards list on left, details pane on right)
        self.workspace_frame = ctk.CTkFrame(self.main_container, fg_color="transparent")
        self.workspace_frame.pack(fill="both", expand=True)

        # Left Column: Matched & Unmatched Sites list
        self.results_column = ctk.CTkFrame(self.workspace_frame, fg_color="transparent")
        self.results_column.pack(side="right", fill="both", expand=True, padx=(10, 0))

        # Title for results
        self.results_title = ctk.CTkLabel(self.results_column, text="المواقع :", 
                                          font=ctk.CTkFont(size=16, weight="bold"))
        self.results_title.pack(anchor="e", pady=(0, 5))

        self.results_scroll_frame = ctk.CTkScrollableFrame(self.results_column, label_text="")
        self.results_scroll_frame.pack(fill="both", expand=True)

        # Right Column: Details dashboard (hidden/empty state initially)
        self.details_column = ctk.CTkFrame(self.workspace_frame, width=420, fg_color="#202020", corner_radius=12)
        self.details_column.pack(side="left", fill="both", padx=(0, 10))
        self.details_column.pack_propagate(False) # Prevent shrinking

        # Initial placeholder in details pane
        self.show_details_placeholder()

    def show_details_placeholder(self):
        """Displays a friendly message when no site card is clicked."""
        # Clear details pane
        for widget in self.details_column.winfo_children():
            widget.destroy()

        placeholder_label = ctk.CTkLabel(self.details_column, 
                                         text="التفاصيل هنا",
                                         font=ctk.CTkFont(size=14, family="Helvetica"),
                                         text_color="gray",
                                         wraplength=350)
        placeholder_label.pack(expand=True, padx=20, pady=20)

    def select_excel_file(self):
        """Opens file dialog to choose an Excel file."""
        file_path = filedialog.askopenfilename(
            title="اختر ملف قاعدة البيانات (Excel)",
            filetypes=[
                ("Excel Files", (".xlsx", ".xls", ".xlsb")),
                ("All Files", "*.*")
            ]
        )
        if file_path:
            self.load_excel_database(file_path)

    def load_excel_database(self, file_path):
        """Loads and processes the Excel database file."""
        try:
            self.excel_handler = ExcelHandler(file_path)
            self.excel_path = file_path
            self.save_config()

            # Update sidebar UI
            self.db_status_label.configure(
                text=f"محملة: {os.path.basename(file_path)} ",
                text_color="#55ff55"
            )
            # Re-process images if any are already uploaded and database changed
            if self.processed_images:
                self.reprocess_matches()
        except Exception as e:
            messagebox.showerror("خطأ", f"فشل تحميل ملف الـ Excel:\n{str(e)}")

    def upload_photos(self):
        """Opens file dialog to choose one or more photo files."""
        if not self.excel_handler:
            messagebox.showwarning("تنبيه", "حمل قاعد البيانات الاول")
            return

        file_paths = filedialog.askopenfilenames(
            title="اختر صور الجداول ",
            filetypes=[
                ("Image Files", (".png", ".jpg", ".jpeg", ".bmp", ".webp")),
                ("All Files", "*.*")
            ]
        )
        if not file_paths:
            return

        # Show loading indicator (using standard modal since Tkinter doesn't have async easily,
        # but we will process them with updates to the UI)
        loading_dialog = ctk.CTkToplevel(self)
        loading_dialog.title("جاري المعالجة...")
        loading_dialog.geometry("400x150")
        loading_dialog.transient(self)
        loading_dialog.grab_set()
        
        # Center loading dialog on parent
        x = self.winfo_x() + (self.winfo_width() // 2) - 200
        y = self.winfo_y() + (self.winfo_height() // 2) - 75
        loading_dialog.geometry(f"+{x}+{y}")

        loading_label = ctk.CTkLabel(loading_dialog, text="اتقل",
                                     font=ctk.CTkFont(size=14))
        loading_label.pack(pady=30)
        
        progress_bar = ctk.CTkProgressBar(loading_dialog, width=300)
        progress_bar.pack(pady=10)
        progress_bar.set(0)
        
        self.update()

        total_files = len(file_paths)
        new_codes_found = []

        for idx, file_path in enumerate(file_paths):
            # Check if already processed
            if any(p['path'] == file_path for p in self.processed_images):
                continue

            try:
                # Extract site codes via OCR
                site_codes = self.photo_processor.find_site_codes(file_path)
                
                # Add to processed images
                self.processed_images.append({
                    'path': file_path,
                    'codes': site_codes
                })

                new_codes_found.extend(site_codes)
                
                # Add file label to sidebar
                img_name = os.path.basename(file_path)
                img_btn = ctk.CTkButton(self.images_scroll_frame, text=img_name, 
                                        fg_color="#333333", hover_color="#444444", 
                                        anchor="w", font=ctk.CTkFont(size=11),
                                        command=lambda p=file_path: self.show_image_codes_info(p))
                img_btn.pack(fill="x", pady=2)
                
            except Exception as e:
                messagebox.showerror("في مشكلة", f"فشل معالجة الصورة: {os.path.basename(file_path)}\n{str(e)}")
            
            # Update progress bar
            progress_bar.set((idx + 1) / total_files)
            self.update()

        loading_dialog.destroy()

        if new_codes_found:
            self.reprocess_matches()

    def show_image_codes_info(self, image_path):
        """Shows details about codes found in a specific clicked image."""
        img_info = next((p for p in self.processed_images if p['path'] == image_path), None)
        if img_info:
            codes_str = ", ".join(img_info['codes']) if img_info['codes'] else "لا يوجد"
            messagebox.showinfo("معلومات الصورة", 
                                f"الصورة: {os.path.basename(image_path)}\n"
                                f"المسار: {image_path}\n\n"
                                f"أكواد المواقع المستخرجة منها ({len(img_info['codes'])}):\n{codes_str}")

    def reprocess_matches(self):
        """Collects all codes from processed images and queries the Excel DB."""
        if not self.excel_handler:
            return

        # Aggregate unique codes from all images
        all_codes = set()
        for img in self.processed_images:
            all_codes.update(img['codes'])
        
        all_codes = list(all_codes)

        # Search Excel
        try:
            # Query matching sites
            db_matches = self.excel_handler.search_sites(all_codes)
            
            # Identify matched codes
            matched_codes_set = set()
            self.matched_sites = []
            
            for site in db_matches:
                # Store match
                self.matched_sites.append(site)
                # Keep track of which search code it matched
                matched_codes_set.add(site.get('_matched_code', '').upper())
                
                # Also check column SiteCode/Site ID matching to map properly
                for col in ['SiteCode', 'Site Code', 'Site ID']:
                    val = site.get(col, '')
                    if val:
                        matched_codes_set.add(str(val).strip().upper())
            
            # Find unmatched codes (codes we searched but got no matches)
            self.unmatched_codes = []
            for code in all_codes:
                code_upper = code.upper()
                # Check if it was matched directly or by digits
                digits_only = ''.join(filter(str.isdigit, code_upper))
                
                matched = False
                if code_upper in matched_codes_set:
                    matched = True
                else:
                    # check if digits matched any site ID
                    for site in self.matched_sites:
                        sid = str(site.get('Site ID', '')).strip().upper()
                        sc = str(site.get('SiteCode', '')).strip().upper()
                        if (digits_only and digits_only == sid) or digits_only in sc or code_upper in sc:
                            matched = True
                            break
                            
                if not matched:
                    self.unmatched_codes.append(code)

            # Update stats
            self.stat_images_label.configure(text=f"الصور المعالجة: {len(self.processed_images)}")
            self.stat_matched_label.configure(text=f"مواقع متطابقة: {len(self.matched_sites)}")
            self.stat_unmatched_label.configure(text=f"مواقع غير موجودة: {len(self.unmatched_codes)}")

            # Stats and matching complete
            pass

            # Refresh results UI list
            self.refresh_results_list()
        except Exception as e:
            messagebox.showerror("خطأ أثناء البحث", f"حدث خطأ أثناء مطابقة الأكواد بالـ Excel:\n{str(e)}")

    def refresh_results_list(self):
        """Re-draws the list of matched and unmatched site cards in the center panel."""
        # Clear frame
        for widget in self.results_scroll_frame.winfo_children():
            widget.destroy()

        if not self.matched_sites and not self.unmatched_codes:
            no_results_lbl = ctk.CTkLabel(self.results_scroll_frame, text="Eng/Mohand Ahmed", font=ctk.CTkFont(size=15), text_color="gray")
            no_results_lbl.pack(pady=40)
            return

        # 1. Matched Sites Sub-Section
        if self.matched_sites:
            section_lbl = ctk.CTkLabel(self.results_scroll_frame, text="تبعنا", font=ctk.CTkFont(size=14, weight="bold"), text_color="#55ff55")
            section_lbl.pack(anchor="e", padx=10, pady=(10, 5))

            for site in self.matched_sites:
                site_code = site.get('SiteCode', site.get('Site Code', 'بدون كود'))
                site_id = site.get('Site ID', '')
                office = site.get('Office', '')
                address = site.get('Address (Arabic)', site.get('Address', ''))
                
                # Card frame
                card = ctk.CTkFrame(self.results_scroll_frame, fg_color="#2b2b2b", height=80, corner_radius=8)
                card.pack(fill="x", padx=5, pady=4)
                
                # Click event (lambda binding)
                card.bind("<Button-1>", lambda event, s=site: self.show_site_details(s))
                
                # Text labels on card (bind click to children too)
                lbl_code = ctk.CTkLabel(card, text=f"{site_code} ({site_id})", font=ctk.CTkFont(size=14, weight="bold"))
                lbl_code.pack(side="right", padx=15, pady=10)
                lbl_code.bind("<Button-1>", lambda event, s=site: self.show_site_details(s))
                
                lbl_details = ctk.CTkLabel(card, text=f"المكتب: {office} | العنوان: {address[:30]}...", font=ctk.CTkFont(size=12), text_color="#aaaaaa")
                lbl_details.pack(side="left", padx=15, pady=10)
                lbl_details.bind("<Button-1>", lambda event, s=site: self.show_site_details(s))

        # 2. Unmatched Codes Sub-Section
        if self.unmatched_codes:
            section2_lbl = ctk.CTkLabel(self.results_scroll_frame, text="مش تبعنا ", font=ctk.CTkFont(size=14, weight="bold"), text_color="#ffbb55")
            section2_lbl.pack(anchor="e", padx=10, pady=(15, 5))

            for code in self.unmatched_codes:
                card = ctk.CTkFrame(self.results_scroll_frame, fg_color="#3a2e22", height=50, corner_radius=8)
                card.pack(fill="x", padx=5, pady=4)
                
                lbl_code = ctk.CTkLabel(card, text=code, font=ctk.CTkFont(size=14, weight="bold"), text_color="#ffbb55")
                lbl_code.pack(side="right", padx=15, pady=10)
                
                lbl_msg = ctk.CTkLabel(card, text="غير مطابق لقاعدة البيانات", font=ctk.CTkFont(size=11, slant="italic"), text_color="#cc9966")
                lbl_msg.pack(side="left", padx=15, pady=10)

    def show_site_details(self, site_dict):
        """Displays the full key-value details of a matched site in the right dashboard panel."""
        self.selected_site = site_dict
        
        # Clear details panel
        for widget in self.details_column.winfo_children():
            widget.destroy()

        # Site Title Header inside panel
        site_code = site_dict.get('SiteCode', site_dict.get('Site Code', 'UNKNOWN'))
        site_id = site_dict.get('Site ID', '')
        
        title_frame = ctk.CTkFrame(self.details_column, fg_color="#2b2b2b", corner_radius=0, height=80)
        title_frame.pack(fill="x")
        
        lbl_title = ctk.CTkLabel(title_frame, text=f"تفاصيل الموقع: {site_code}", 
                                 font=ctk.CTkFont(size=18, weight="bold"), text_color="#55ff55")
        lbl_title.pack(side="right", padx=20, pady=15)

        lbl_id = ctk.CTkLabel(title_frame, text=f"ID: {site_id}", font=ctk.CTkFont(size=12), text_color="#aaaaaa")
        lbl_id.pack(side="left", padx=20, pady=15)

        # Scrollable container for the key-value attributes
        attrs_scroll = ctk.CTkScrollableFrame(self.details_column, label_text="")
        attrs_scroll.pack(fill="both", expand=True, padx=10, pady=10)

        # Display important fields beautifully as structured cards
        # We define Arabic titles for common columns
        headers_translation = {
            'SiteCode': 'كود الموقع (Site Code)',
            'Site ID': 'معرف الموقع (Site ID)',
            '_sheet_name': 'موقع البيانات (اسم الشيت / Sheet Name)',
            '_row_number': 'موقع البيانات (رقم الصف / Excel Row)',
            'Region': 'المنطقة (Region)',
            'Sub. Reg': 'المنطقة الفرعية (Sub Reg)',
            'Office': 'المكتب المسؤول (Office)',
            'Priority': 'الأولية (Priority)',
            'Accessibility': 'إمكانية الوصول (Accessibility)',
            'Address (Arabic)': 'العنوان بالعربي (Address)',
            'Status': 'حالة الموقع (Status)',
            'Comment': 'التعليق (Comment)',
            'Vendor': 'المورد (Vendor)',
            '#Site Cells': 'عدد الخلايا (#Site Cells)',
            'HUB Type': 'نوع الهب (HUB Type)',
            'Power Source': 'مصدر الطاقة (Power Source)',
            'MTTR': 'وقت الإصلاح (MTTR)'
        }

        # Filter and sort columns: show translated ones first, then others
        important_keys = [k for k in headers_translation.keys() if k in site_dict]
        remaining_keys = [k for k in site_dict.keys() if k not in important_keys and not k.startswith('_')]

        # Draw important attributes
        for key in important_keys + remaining_keys:
            val = site_dict[key]
            if val is None or val == 'nan' or val == '':
                val = "غير متوفر"

            label_name = headers_translation.get(key, key)
            
            # Card for this attribute
            attr_card = ctk.CTkFrame(attrs_scroll, fg_color="#2a2a2a", corner_radius=6)
            attr_card.pack(fill="x", pady=4, padx=5)
            
            lbl_key = ctk.CTkLabel(attr_card, text=label_name, font=ctk.CTkFont(size=11, weight="bold"), text_color="#888888")
            lbl_key.pack(anchor="e", padx=10, pady=(4, 1))

            lbl_val = ctk.CTkLabel(attr_card, text=str(val), font=ctk.CTkFont(size=13), justify="right", wraplength=350)
            lbl_val.pack(anchor="e", padx=10, pady=(1, 5))

    def export_results(self):
        """Exports the matching results list to a new Excel file."""
        if not self.matched_sites and not self.unmatched_codes:
            messagebox.showwarning("تنبيه", "لا توجد نتائج لتصديرها!")
            return

        file_path = filedialog.asksaveasfilename(
            title="حفظ تقرير المطابقة كـ Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")]
        )
        if not file_path:
            return

        try:
            # 1. Sheet for Matched Sites
            matched_df = pd.DataFrame(self.matched_sites)
            # Remove helper keys
            if '_matched_code' in matched_df.columns:
                matched_df = matched_df.drop(columns=['_matched_code'])
                
            # 2. Sheet for Unmatched Codes
            unmatched_df = pd.DataFrame(self.unmatched_codes, columns=['Unmatched Site Codes'])

            # Write to multiple sheets
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                if not matched_df.empty:
                    matched_df.to_excel(writer, sheet_name='المواقع المطابقة', index=False)
                if not unmatched_df.empty:
                    unmatched_df.to_excel(writer, sheet_name='أكواد غير مطابقة', index=False)

            messagebox.showinfo("تم التصدير بنجاح", f"تم حفظ تقرير المطابقة بنجاح في:\n{file_path}")
        except Exception as e:
            messagebox.showerror("خطأ أثناء التصدير", f"فشل حفظ ملف التصدير:\n{str(e)}")

    def clear_results(self):
        """Clears all uploaded images and matching results."""
        if messagebox.askyesno("تأكيد", "هتمسح النتائج "):
            self.processed_images = []
            self.matched_sites = []
            self.unmatched_codes = []
            self.selected_site = None
            
            # Clear UI list of processed images in sidebar
            for widget in self.images_scroll_frame.winfo_children():
                widget.destroy()

            # Update stats
            self.stat_images_label.configure(text="الصور المعالجة: 0")
            self.stat_matched_label.configure(text="مواقع متطابقة: 0")
            self.stat_unmatched_label.configure(text="مواقع غير موجودة: 0")

            # Clear complete
            pass

            # Refresh lists
            self.refresh_results_list()
            self.show_details_placeholder()

    def manual_search(self):
        """Manually search for a site code in the Excel database."""
        if not self.excel_handler:
            messagebox.showwarning("تنبيه", "حمل قاعده البيانات")
            return

        code = self.search_entry.get().strip()
        if not code:
            return

        try:
            # Query Excel database for this code
            matches = self.excel_handler.search_sites([code])
            
            if matches:
                # Add to matched list if not already present
                for site in matches:
                    site_code = site.get('SiteCode', '').upper()
                    # Check if already in self.matched_sites
                    if not any(s.get('SiteCode', '').upper() == site_code for s in self.matched_sites):
                        self.matched_sites.append(site)
                
                # Refresh list UI
                self.refresh_results_list()
                
                # Automatically select and show details of the matched site
                matched_site = next(s for s in self.matched_sites if s.get('SiteCode', '').upper() == code.upper() or s.get('Site ID', '').upper() == code.upper())
                self.show_site_details(matched_site)
                
                # Clear search box
                self.search_entry.delete(0, tk.END)
            else:
                # If not found, add to unmatched list if not already there
                code_upper = code.upper()
                if code_upper not in self.unmatched_codes:
                    # check it's not in matched either
                    if not any(s.get('SiteCode', '').upper() == code_upper for s in self.matched_sites):
                        self.unmatched_codes.append(code)
                
                self.refresh_results_list()
                messagebox.showinfo("نتائج البحث", f"الموقع ({code}) غير موجود في قاعدة البيانات!")
                
        except Exception as e:
            # If next() failed, just show the first one found or first in list
            if matches:
                self.show_site_details(matches[0])
                self.search_entry.delete(0, tk.END)
            else:
                messagebox.showerror("خطأ أثناء البحث", f"حدث خطأ أثناء البحث عن الموقع:\n{str(e)}")

if __name__ == "__main__":
    app = SitePhotoMatcherApp()
    app.mainloop()
