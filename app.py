import os
import threading
import urllib.parse
import pandas as pd
import gspread
import customtkinter as ctk
from tkinter import filedialog, messagebox
from datetime import datetime

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("лень двигатель прогреси")
        self.geometry("700x500")

        self.telega_path = ctk.StringVar()
        self.crm_path = ctk.StringVar()
        
        self.sheet_url = "" 
        self.sheet_name = ""

        self.setup_ui()

    def setup_ui(self):
        ctk.CTkLabel(self, text="Выгрузка Telega.in (.xlsx):").pack(pady=(20, 0), padx=20, anchor="w")
        frame_telega = ctk.CTkFrame(self, fg_color="transparent")
        frame_telega.pack(fill="x", padx=20, pady=5)
        ctk.CTkEntry(frame_telega, textvariable=self.telega_path, state="disabled", width=450).pack(side="left")
        ctk.CTkButton(frame_telega, text="Обзор", command=lambda: self.select_file(self.telega_path)).pack(side="right", padx=10)

        ctk.CTkLabel(self, text="Выгрузка CRM (.xlsx):").pack(pady=(10, 0), padx=20, anchor="w")
        frame_crm = ctk.CTkFrame(self, fg_color="transparent")
        frame_crm.pack(fill="x", padx=20, pady=5)
        ctk.CTkEntry(frame_crm, textvariable=self.crm_path, state="disabled", width=450).pack(side="left")
        ctk.CTkButton(frame_crm, text="Обзор", command=lambda: self.select_file(self.crm_path)).pack(side="right", padx=10)

        ctk.CTkLabel(self, text="Кто закупал:").pack(pady=(10, 0), padx=20, anchor="w")
        self.buyer_var = ctk.StringVar(value="Все")
        buyers = ["Все", "Ваня", "Глеб", "Юра"]
        ctk.CTkOptionMenu(self, variable=self.buyer_var, values=buyers).pack(padx=20, pady=5, anchor="w")

        self.start_btn = ctk.CTkButton(self, text="работать быстр..", command=self.start_process, height=40)
        self.start_btn.pack(pady=20)

        self.log_box = ctk.CTkTextbox(self, width=660, height=150, state="disabled")
        self.log_box.pack(padx=20, pady=10)

    def select_file(self, string_var):
        filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if filepath:
            string_var.set(filepath)

    def log(self, text):
        self.log_box.configure(state="normal")
        self.log_box.insert("end", f"{datetime.now().strftime('%H:%M:%S')} - {text}\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

    def get_buyer_from_project_name(self, project_name):
        project_name = str(project_name).lower()
        if project_name.startswith("gleb"): return "Глеб"
        if project_name.startswith("ivan"): return "Ваня"
        if project_name.startswith("yuriy"): return "Юра"
        return "Не указан"

    def extract_utm_campaign(self, url_string):
        if pd.isna(url_string): return None
        first_url = str(url_string).split(',')[0].strip()
        parsed_url = urllib.parse.urlparse(first_url)
        params = urllib.parse.parse_qs(parsed_url.query)
        return params.get('utm_campaign', [None])[0]

    def format_date(self, date_string):
        if pd.isna(date_string): return ""
        try:
            dt = pd.to_datetime(str(date_string), dayfirst=True)
            return dt.strftime('%d.%m.%Y')
        except:
            return str(date_string).split(' ')[0].replace('-', '.')

    def start_process(self):
        if not self.telega_path.get() or not self.crm_path.get():
            messagebox.showwarning("Ошибка", "Выберите оба файла!")
            return
        self.start_btn.configure(state="disabled")
        self.log("Начинаем обработку...")
        threading.Thread(target=self.run_logic, daemon=True).start()

    def run_logic(self):
        try:
            self.log("Чтение файла CRM...")
            df_crm = pd.read_excel(self.crm_path.get())
            leads_dict = df_crm.groupby('Utm Campaign')['Лиды'].sum().to_dict()
            
            self.log("Чтение выгрузки Telega.in...")
            df_tel = pd.read_excel(self.telega_path.get())
            df_tel.columns = df_tel.columns.str.strip()
            if 'Название проекта' in df_tel.columns:
                df_tel['Название проекта'] = df_tel['Название проекта'].ffill()

            selected_buyer = self.buyer_var.get()
            if selected_buyer != "Все":
                buyer_map = {"Ваня": "ivan", "Глеб": "gleb", "Юра": "yuriy"}
                prefix = buyer_map.get(selected_buyer, "")
                df_tel = df_tel[df_tel['Название проекта'].astype(str).str.startswith(prefix, na=False)]

            to_process = []
            for _, row in df_tel.iterrows():
                if pd.isna(row.get('Название канала')) or str(row.get('Название канала')).strip() == "": continue 
                
                price = row.get('Цена, руб', 0)
                if isinstance(price, str): price = float(price.replace(' ', '').replace(',', '.'))
                
                utm_camp = self.extract_utm_campaign(row.get('Оригинальная ссылка'))
                
                to_process.append({
                    "Название": str(row['Название канала']),
                    "Ссылка": str(row.get('Ссылка на канал', '')),
                    "Цена": int(price),
                    "Дата": self.format_date(row.get('Размещен')),
                    "Лиды": leads_dict.get(utm_camp, 0) if utm_camp else 0,
                    "Кто закупил": self.get_buyer_from_project_name(row.get('Название проекта', ""))
                })

            self.log("Подключение к Google Sheets...")
            gc = gspread.service_account(filename='credentials.json')
            sh = gc.open_by_url(self.sheet_url)
            worksheet = sh.worksheet(self.sheet_name)

            all_values = worksheet.get_all_values()
            def clean_url(url): return str(url).strip().lower().rstrip('/')

            existing_urls = {}
            max_row = 1 
            for i, row_data in enumerate(all_values):
                url = clean_url(row_data[1]) if len(row_data) > 1 else ""
                name = str(row_data[0]).strip() if len(row_data) > 0 else ""
                if url: existing_urls[url] = i + 1 
                if name or url: max_row = i + 1

            new_rows_AE, new_rows_G, update_cells = [], [], []

            for item in to_process:
                item_url = clean_url(item['Ссылка'])
                if item_url in existing_urls:
                    row_idx = existing_urls[item_url]
                    update_cells.append(gspread.Cell(row=row_idx, col=3, value=item['Цена']))
                    update_cells.append(gspread.Cell(row=row_idx, col=4, value=item['Дата']))
                    update_cells.append(gspread.Cell(row=row_idx, col=5, value=item['Лиды']))
                    update_cells.append(gspread.Cell(row=row_idx, col=7, value=item['Кто закупил']))
                else:
                    new_rows_AE.append([item['Название'], item['Ссылка'], item['Цена'], item['Дата'], item['Лиды']])
                    new_rows_G.append([item['Кто закупил']])

            if update_cells: worksheet.update_cells(update_cells, value_input_option='USER_ENTERED')
            if new_rows_AE:
                start_row = max_row + 1
                worksheet.update(range_name=f"A{start_row}:E{start_row + len(new_rows_AE) - 1}", values=new_rows_AE, value_input_option='USER_ENTERED')
                worksheet.update(range_name=f"G{start_row}:G{start_row + len(new_rows_G) - 1}", values=new_rows_G, value_input_option='USER_ENTERED')

            self.log("✅ УСПЕШНО ЗАВЕРШЕНО!")
            messagebox.showinfo("Успех", "Данные синхронизированы!")
        except Exception as e:
            self.log(f"❌ ОШИБКА: {str(e)}")
            messagebox.showerror("Ошибка", str(e))
        finally:
            self.start_btn.configure(state="normal")

if __name__ == "__main__":
    app = App()
    app.mainloop()