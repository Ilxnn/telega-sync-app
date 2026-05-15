import os
import json
import tkinter
import threading
import urllib.parse
import pandas as pd
import gspread
import customtkinter as ctk
from tkinter import filedialog, messagebox
from datetime import datetime

ctk.set_appearance_mode("Light")
ctk.set_default_color_theme("blue")
CONFIG_FILE = "config.json"

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Auto-Buyer Sync")
        self.geometry("750x650")
        self.configure(fg_color="#F3F4F6")

        self.telega_path = ctk.StringVar()
        self.crm_path = ctk.StringVar()
        self.buyer_var = ctk.StringVar(value="Все")
        
        self.config_data = self.load_config()

        self.main_container = ctk.CTkFrame(self, fg_color="transparent")
        self.main_container.pack(fill="both", expand=True, padx=20, pady=20)

        if not self.config_data.get("sheet_url") or not self.config_data.get("sheet_name"):
            self.show_setup_ui()
        else:
            self.show_main_ui()

    def add_right_click_menu(self, widget):
        menu = tkinter.Menu(widget, tearoff=0)

        def paste_text():
            try:
                clipboard_content = self.clipboard_get()
                try:
                    widget.delete("sel.first", "sel.last")
                except tkinter.TclError:
                    pass
                widget.insert("insert", clipboard_content)
            except tkinter.TclError:
                pass
            return "break"

        menu.add_command(label="Вырезать", command=lambda: widget.event_generate("<<Cut>>"))
        menu.add_command(label="Копировать", command=lambda: widget.event_generate("<<Copy>>"))
        menu.add_command(label="Вставить", command=paste_text)
        menu.add_separator()
        menu.add_command(label="Выделить все", command=lambda: widget.event_generate("<<SelectAll>>"))

        def show_menu(event):
            menu.post(event.x_root, event.y_root)

        widget.bind("<Button-3>", show_menu)
        widget.bind_class("Entry", "<Control-a>", lambda event: event.widget.select_range(0, 'end'))
        widget.bind_class("Entry", "<Control-v>", lambda event: paste_text())


    def load_config(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    return json.load(f)
            except:
                return {}
        return {}

    def save_config(self, url, name):
        self.config_data = {"sheet_url": url, "sheet_name": name}
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(self.config_data, f, indent=4)

    def clear_container(self):
        for widget in self.main_container.winfo_children():
            widget.destroy()

    def show_setup_ui(self):
        self.clear_container()
        card = ctk.CTkFrame(self.main_container, fg_color="#FFFFFF", corner_radius=15)
        card.pack(fill="x", pady=50, padx=50)

        ctk.CTkLabel(card, text="Первичная настройка", font=("Helvetica", 20, "bold"), text_color="#1F2937").pack(pady=(30, 20))

        ctk.CTkLabel(card, text="Ссылка на Google Таблицу:", text_color="#4B5563", font=("Helvetica", 13)).pack(anchor="w", padx=30)
        url_entry = ctk.CTkEntry(card, width=400, height=35, fg_color="#F9FAFB", border_color="#D1D5DB", text_color="#111827")
        url_entry.pack(pady=(5, 15), padx=30)
        self.add_right_click_menu(url_entry)
        if self.config_data.get("sheet_url"): url_entry.insert(0, self.config_data.get("sheet_url"))

        ctk.CTkLabel(card, text="Название листа (вкладки внизу):", text_color="#4B5563", font=("Helvetica", 13)).pack(anchor="w", padx=30)
        name_entry = ctk.CTkEntry(card, width=400, height=35, fg_color="#F9FAFB", border_color="#D1D5DB", text_color="#111827")
        name_entry.pack(pady=(5, 30), padx=30)
        self.add_right_click_menu(name_entry)
        if self.config_data.get("sheet_name"): name_entry.insert(0, self.config_data.get("sheet_name"))

        def save_and_continue():
            url = url_entry.get().strip()
            name = name_entry.get().strip()
            if not url or not name:
                messagebox.showwarning("Ошибка", "Заполните оба поля!")
                return
            self.save_config(url, name)
            self.show_main_ui()

        ctk.CTkButton(card, text="Сохранить и продолжить", font=("Helvetica", 14, "bold"), height=40, command=save_and_continue).pack(pady=(0, 30), padx=30, fill="x")

    def show_main_ui(self):
        self.clear_container()
        header_frame = ctk.CTkFrame(self.main_container, fg_color="transparent")
        header_frame.pack(fill="x", pady=(0, 15))
        ctk.CTkLabel(header_frame, text="Синхронизация закупок", font=("Helvetica", 24, "bold"), text_color="#111827").pack(side="left")
        ctk.CTkButton(header_frame, text="Настройки", width=100, fg_color="#E5E7EB", text_color="#374151", hover_color="#D1D5DB", command=self.show_setup_ui).pack(side="right")

        card_files = ctk.CTkFrame(self.main_container, fg_color="#FFFFFF", corner_radius=15)
        card_files.pack(fill="x", pady=(0, 15))

        ctk.CTkLabel(card_files, text="Выгрузка Telega.in (.xlsx):", text_color="#4B5563", font=("Helvetica", 13, "bold")).pack(pady=(15, 0), padx=20, anchor="w")
        row1 = ctk.CTkFrame(card_files, fg_color="transparent")
        row1.pack(fill="x", padx=20, pady=(5, 10))
        ctk.CTkEntry(row1, textvariable=self.telega_path, state="disabled", fg_color="#F3F4F6", text_color="#111827", border_color="#D1D5DB").pack(side="left", fill="x", expand=True, padx=(0, 10))
        ctk.CTkButton(row1, text="Выбрать файл", width=120, command=lambda: self.select_file(self.telega_path)).pack(side="right")

        ctk.CTkLabel(card_files, text="Выгрузка CRM (.xlsx):", text_color="#4B5563", font=("Helvetica", 13, "bold")).pack(pady=(5, 0), padx=20, anchor="w")
        row2 = ctk.CTkFrame(card_files, fg_color="transparent")
        row2.pack(fill="x", padx=20, pady=(5, 20))
        ctk.CTkEntry(row2, textvariable=self.crm_path, state="disabled", fg_color="#F3F4F6", text_color="#111827", border_color="#D1D5DB").pack(side="left", fill="x", expand=True, padx=(0, 10))
        ctk.CTkButton(row2, text="Выбрать файл", width=120, command=lambda: self.select_file(self.crm_path)).pack(side="right")

        card_action = ctk.CTkFrame(self.main_container, fg_color="#FFFFFF", corner_radius=15)
        card_action.pack(fill="x", pady=(0, 15))

        row_action = ctk.CTkFrame(card_action, fg_color="transparent")
        row_action.pack(fill="x", padx=20, pady=20)
        
        ctk.CTkLabel(row_action, text="Кто закупал:", text_color="#4B5563", font=("Helvetica", 13, "bold")).pack(side="left")
        buyers = ["Все", "Ваня", "Глеб", "Юра"]
        ctk.CTkOptionMenu(row_action, variable=self.buyer_var, values=buyers, fg_color="#FFFFFF", button_color="#D1D5DB", button_hover_color="#9CA3AF", text_color="#111827").pack(side="left", padx=(10, 0))

        self.start_btn = ctk.CTkButton(row_action, text="Cинхронизировать", font=("Helvetica", 14, "bold"), height=40, fg_color="#10B981", hover_color="#059669",text_color="#FFFFFF", command=self.start_process)
        self.start_btn.pack(side="right")

        self.log_box = ctk.CTkTextbox(self.main_container, height=150, state="disabled", fg_color="#111827", text_color="#10B981", font=("Consolas", 12))
        self.log_box.pack(fill="x", expand=True, pady=(0, 5))

        ctk.CTkLabel(self.main_container, text="доля воркера 15 рублей", text_color="#B0B0B0", font=("Helvetica", 10)).pack(side="bottom", anchor="se")

    def select_file(self, string_var):
        filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if filepath:
            string_var.set(filepath)

    def log(self, text):
        self.log_box.configure(state="normal")
        self.log_box.insert("end", f"[{datetime.now().strftime('%H:%M:%S')}] {text}\n")
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
            messagebox.showwarning("Внимание", "Пожалуйста, выберите оба файла выгрузки.")
            return
        self.start_btn.configure(state="disabled")
        self.log("Процесс запущен...")
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
            sh = gc.open_by_url(self.config_data["sheet_url"])
            worksheet = sh.worksheet(self.config_data["sheet_name"])

            all_values = worksheet.get_all_values()
            def clean_url(url): return str(url).strip().lower().rstrip('/')

            existing_data = {}
            max_row = 1 
            for i, row_data in enumerate(all_values):
                url = clean_url(row_data[1]) if len(row_data) > 1 else ""
                name = str(row_data[0]).strip() if len(row_data) > 0 else ""
                date_str = str(row_data[3]).strip() if len(row_data) > 3 else ""
                
                if url: existing_data[url] = {"row": i + 1, "date": date_str}
                if name or url: max_row = i + 1

            new_rows_AE, new_rows_G, update_cells = [], [], []
            today_date = datetime.now().date()

            for item in to_process:
                item_url = clean_url(item['Ссылка'])
                
                if item_url in existing_data:
                    row_idx = existing_data[item_url]["row"]
                    sheet_date_str = existing_data[item_url]["date"]
                    
                    is_future = False
                    if sheet_date_str:
                        try:
                            sheet_dt = pd.to_datetime(sheet_date_str, dayfirst=True).date()
                            if sheet_dt >= today_date: is_future = True
                        except:
                            pass
                    
                    if is_future: continue

                    update_cells.append(gspread.Cell(row=row_idx, col=3, value=item['Цена']))
                    update_cells.append(gspread.Cell(row=row_idx, col=4, value=item['Дата']))
                    update_cells.append(gspread.Cell(row=row_idx, col=7, value=item['Кто закупил']))
                    
                    if item['Кто закупил'] != 'Юра':
                        update_cells.append(gspread.Cell(row=row_idx, col=5, value=item['Лиды']))
                else:
                    leads_for_new_row = 0 if item['Кто закупил'] == 'Юра' else item['Лиды']
                    new_rows_AE.append([item['Название'], item['Ссылка'], item['Цена'], item['Дата'], leads_for_new_row])
                    new_rows_G.append([item['Кто закупил']])

            if update_cells: worksheet.update_cells(update_cells, value_input_option='USER_ENTERED')
            if new_rows_AE:
                start_row = max_row + 1
                worksheet.update(range_name=f"A{start_row}:E{start_row + len(new_rows_AE) - 1}", values=new_rows_AE, value_input_option='USER_ENTERED')
                worksheet.update(range_name=f"G{start_row}:G{start_row + len(new_rows_G) - 1}", values=new_rows_G, value_input_option='USER_ENTERED')

            self.log("УСПЕШНО ЗАВЕРШЕНО!")
            messagebox.showinfo("Успех", "Данные успешно синхронизированы!")
        except Exception as e:
            self.log(f"ОШИБКА: {str(e)}")
            messagebox.showerror("Ошибка", f"Произошла ошибка:\n{str(e)}")
        finally:
            self.start_btn.configure(state="normal")

if __name__ == "__main__":
    app = App()
    app.mainloop()
