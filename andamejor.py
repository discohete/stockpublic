import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import json
import os

class StockComparisonApp:
    def __init__(self, master):
        self.master = master
        self.master.title("Comparador de Stocks")
        self.master.geometry("800x600")

        self.df_real = pd.DataFrame()
        self.df_meli = pd.DataFrame()
        self.df_tidi = pd.DataFrame()
        self.df_summed = pd.DataFrame()  # Nueva variable para almacenar los datos sumados

        self.config_file = 'config.json'
        self.load_config()

        self.create_widgets()

    def create_widgets(self):
        # Frame para botones de carga
        load_frame = ttk.Frame(self.master, padding="10")
        load_frame.pack(fill=tk.X)

        self.real_button = ttk.Button(load_frame, text="Cargar Stock Real", command=lambda: self.load_file("real"))
        self.real_button.grid(row=0, column=0, padx=5, pady=5)

        self.meli_button = ttk.Button(load_frame, text="Cargar Stock MELI", command=lambda: self.load_file("meli"))
        self.meli_button.grid(row=0, column=1, padx=5, pady=5)

        self.tidi_button = ttk.Button(load_frame, text="Cargar Stock TIDI", command=lambda: self.load_file("tidi"))
        self.tidi_button.grid(row=0, column=2, padx=5, pady=5)

        self.status_label = ttk.Label(load_frame, text="")
        self.status_label.grid(row=1, column=0, columnspan=3, pady=5)

        ttk.Button(self.master, text="Procesar y Comparar Stocks", command=self.process_and_compare_stocks).pack(pady=10)

        # Treeview para resultados
        columns = ("SKU", "Stock Real", "Stock Sumado", "Exceso")
        self.tree = ttk.Treeview(self.master, columns=columns, show="headings")
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100)
        self.tree.pack(expand=True, fill=tk.BOTH, padx=10, pady=10)

        # Scrollbar para Treeview
        scrollbar = ttk.Scrollbar(self.master, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.update_button_states()

    def load_config(self):
        if os.path.exists(self.config_file):
            with open(self.config_file, 'r') as f:
                self.config = json.load(f)
        else:
            self.config = {'real': '', 'meli': '', 'tidi': ''}

    def save_config(self):
        with open(self.config_file, 'w') as f:
            json.dump(self.config, f)

    def load_file(self, file_type):
        initial_dir = os.path.dirname(self.config[file_type]) if self.config[file_type] else '/'
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")], initialdir=initial_dir)
        if not file_path:
            return

        try:
            if file_type == "real":
                self.df_real = pd.read_excel(file_path)
                self.status_label.config(text="Stock Real cargado")
            elif file_type == "meli":
                self.df_meli = pd.read_excel(file_path, sheet_name="Publicaciones", skiprows=1)
                self.status_label.config(text="Stock MELI cargado")
            elif file_type == "tidi":
                self.df_tidi = pd.read_excel(file_path)
                self.status_label.config(text="Stock TIDI cargado")
            
            self.config[file_type] = file_path
            self.save_config()
            self.update_button_states()
        except Exception as e:
            messagebox.showerror("Error", f"Error al cargar el archivo: {str(e)}")

    def update_button_states(self):
        self.real_button.config(text=f"Cargar Stock Real: {'✓' if self.config['real'] else '❌'}")
        self.meli_button.config(text=f"Cargar Stock MELI: {'✓' if self.config['meli'] else '❌'}")
        self.tidi_button.config(text=f"Cargar Stock TIDI: {'✓' if self.config['tidi'] else '❌'}")

    def sum_meli_tidi_stocks(self):
        # Sumar stocks de MELI
        meli_sum = self.df_meli.groupby('SKU')['En mi depósito'].sum().reset_index()
        meli_sum = meli_sum.rename(columns={'En mi depósito': 'Stock_MELI'})

        # Sumar stocks de TIDI
        tidi_sum = self.df_tidi.groupby('RefId')['AvailableQuantity'].sum().reset_index()
        tidi_sum = tidi_sum.rename(columns={'RefId': 'SKU', 'AvailableQuantity': 'Stock_TIDI'})

        # Combinar los datos sumados de MELI y TIDI
        self.df_summed = pd.merge(meli_sum, tidi_sum, on='SKU', how='outer').fillna(0)
        self.df_summed['Stock_Total'] = self.df_summed['Stock_MELI'] + self.df_summed['Stock_TIDI']

        return self.df_summed

    def process_and_compare_stocks(self):
        if self.df_real.empty or self.df_meli.empty or self.df_tidi.empty:
            messagebox.showerror("Error", "Por favor, cargue todas las planillas antes de procesar y comparar.")
            return

        try:
            # Sumar stocks de MELI y TIDI
            self.sum_meli_tidi_stocks()

            # Merge con el stock real
            df_comparison = pd.merge(self.df_real, self.df_summed, left_on='Articulo', right_on='SKU', how='left')

            # Calcular exceso
            df_comparison['Exceso'] = df_comparison['Stock_Total'] - df_comparison['Stock disponible']

            # Filtrar resultados donde hay exceso
            df_excess = df_comparison[df_comparison['Exceso'] > 0]

            # Limpiar y poblar el Treeview
            for item in self.tree.get_children():
                self.tree.delete(item)

            for _, row in df_excess.iterrows():
                self.tree.insert("", "end", values=(
                    row['Articulo'],
                    row['Stock disponible'],
                    row['Stock_Total'],
                    row['Exceso']
                ))

            messagebox.showinfo("Comparación completada", f"Se encontraron {len(df_excess)} SKUs con exceso de stock.")

        except Exception as e:
            messagebox.showerror("Error", f"Error al procesar y comparar stocks: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = StockComparisonApp(root)
    root.mainloop()