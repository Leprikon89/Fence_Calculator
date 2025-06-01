import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, IntVar
import os

class FenceCalculator:
    def __init__(self, root):
        self.root = root
        self.root.title("Калькулятор металлических заборов")
        self.root.geometry("550x700")
        
        # Параметры по умолчанию
        self.prices_file = "prices.xlsx"
        self.output_file = "result.xlsx"
        self.price_data = None
        self.metal_types = []
        self.profile_heights = []
        self.thicknesses = []
        
        # Создаем интерфейс
        self.create_widgets()
        self.load_prices()

    def create_widgets(self):
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Заголовок
        ttk.Label(main_frame, text="Калькулятор металлических заборов", 
                  font=("Arial", 14, "bold")).grid(row=0, columnspan=3, pady=10)
        
        # Поля для ввода
        fields = [
            ("Длина забора (м):", "length", "entry"),
            ("Высота забора (м):", "height", "entry"),
            ("Количество столбов:", "posts", "entry"),
            ("Заглубление столбов (м):", "post_depth", "entry"),
            ("Количество ворот:", "gates", "entry"),
            ("Количество калиток:", "doors", "entry"),
            ("Расстояние доставки (км):", "delivery_distance", "entry"),
            ("Тип металла:", "metal_type", "combobox"),
            ("Высота профиля (мм):", "profile_height", "combobox"),
            ("Толщина металла (мм):", "thickness", "combobox")
        ]
        
        self.entries = {}
        for i, (label, name, field_type) in enumerate(fields, start=1):
            ttk.Label(main_frame, text=label).grid(row=i, column=0, padx=5, pady=5, sticky="e")
            if field_type == "combobox":
                cb = ttk.Combobox(main_frame, width=22)
                cb.grid(row=i, column=1, padx=5, pady=5, sticky="w")
                cb.set("")
                self.entries[name] = cb
                # Привязка событий для обновления зависимых списков
                if name == "metal_type":
                    cb.bind("<<ComboboxSelected>>", self.update_metal_params)
            else:
                entry = ttk.Entry(main_frame, width=25)
                entry.grid(row=i, column=1, padx=5, pady=5, sticky="w")
                self.entries[name] = entry
                # Установка значений по умолчанию
                if name == "length":
                    entry.insert(0, "10")
                elif name == "height":
                    entry.insert(0, "1.8")
                elif name == "posts":
                    entry.insert(0, "10")
                elif name == "post_depth":
                    entry.insert(0, "1.2")

        # Забутовка
        ttk.Label(main_frame, text="Забутовка:").grid(row=len(fields)+1, column=0, padx=5, pady=5, sticky="e")
        self.foundation_var = IntVar(value=1)
        chk = ttk.Checkbutton(main_frame, variable=self.foundation_var)
        chk.grid(row=len(fields)+1, column=1, padx=5, pady=5, sticky="w")
        
        # Покрытие
        ttk.Label(main_frame, text="Полимерное покрытие:").grid(row=len(fields)+2, column=0, padx=5, pady=5, sticky="e")
        self.coating_var = IntVar(value=1)
        chk_coating = ttk.Checkbutton(main_frame, variable=self.coating_var)
        chk_coating.grid(row=len(fields)+2, column=1, padx=5, pady=5, sticky="w")
        
        # Кнопки
        btn_frame = ttk.Frame(main_frame)
        btn_frame.grid(row=len(fields)+3, columnspan=2, pady=15)
        ttk.Button(btn_frame, text="Загрузить цены", command=self.load_prices).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Рассчитать", command=self.calculate).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Сохранить результат", command=self.save_results).pack(side=tk.LEFT, padx=5)
        
        # Результаты
        result_frame = ttk.LabelFrame(main_frame, text="Результаты расчета", padding=10)
        result_frame.grid(row=len(fields)+4, columnspan=2, sticky="we", pady=10)
        self.results_var = tk.StringVar()
        self.results_var.set("Итоговая стоимость: 0.00 руб.")
        ttk.Label(result_frame, textvariable=self.results_var, font=("Arial", 12, "bold")).pack()
        self.details_var = tk.StringVar()
        ttk.Label(result_frame, textvariable=self.details_var, justify=tk.LEFT).pack(anchor="w", pady=5)
        
        # Статус
        self.status_var = tk.StringVar()
        ttk.Label(main_frame, textvariable=self.status_var, foreground="gray").grid(row=len(fields)+5, columnspan=2, pady=5)

    def update_metal_params(self, event=None):
        """Обновляет доступные параметры профиля при выборе типа металла"""
        metal_type = self.entries["metal_type"].get()
        if not metal_type or self.price_data is None:
            return
        
        # Фильтруем данные по выбранному типу металла
        filtered = self.price_data[self.price_data["metal_type"] == metal_type]
        
        # Обновляем доступные высоты профиля
        heights = filtered["profile_height"].unique().tolist()
        heights = sorted(heights)
        self.entries["profile_height"]["values"] = heights
        if heights:
            self.entries["profile_height"].current(0)
        
        # Обновляем доступные толщины
        thicknesses = filtered["thickness"].unique().tolist()
        thicknesses = sorted(thicknesses)
        self.entries["thickness"]["values"] = thicknesses
        if thicknesses:
            self.entries["thickness"].current(0)

    def load_prices(self):
        try:
            file_path = filedialog.askopenfilename(
                filetypes=[("Excel files", "*.xlsx *.xls")],
                title="Выберите файл с ценами",
                initialfile=self.prices_file
            )
            if not file_path:
                return
            
            self.prices_file = file_path
            self.price_data = pd.read_excel(file_path)
            
            # Обновляем список типов металла
            if "metal_type" in self.entries:
                metal_types = self.price_data["metal_type"].unique().tolist()
                self.entries["metal_type"]["values"] = metal_types
                if metal_types:
                    self.entries["metal_type"].current(0)
                    # Обновляем параметры профиля
                    self.update_metal_params()
            
            self.status_var.set(f"Загружены цены из: {os.path.basename(file_path)}")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить цены:\n{str(e)}")

    def get_float(self, entry, name):
        """Безопасное получение числовых значений"""
        try:
            return float(entry.get().replace(',', '.'))
        except ValueError:
            raise ValueError(f"Некорректное числовое значение в поле {name}")

    def calculate(self):
        try:
            # Получаем значения из полей ввода
            length = self.get_float(self.entries["length"], "Длина забора")
            height = self.get_float(self.entries["height"], "Высота забора")
            posts = int(self.get_float(self.entries["posts"], "Количество столбов"))
            post_depth = self.get_float(self.entries["post_depth"], "Заглубление столбов")
            gates = int(self.get_float(self.entries["gates"], "Количество ворот"))
            doors = int(self.get_float(self.entries["doors"], "Количество калиток"))
            delivery_distance = self.get_float(self.entries["delivery_distance"], "Расстояние доставки")
            metal_type = self.entries["metal_type"].get()
            profile_height = self.entries["profile_height"].get()
            thickness = self.entries["thickness"].get()
            foundation = bool(self.foundation_var.get())
            coating = bool(self.coating_var.get())
            
            # Ищем цены для выбранной конфигурации
            mask = (
                (self.price_data["metal_type"] == metal_type) &
                (self.price_data["profile_height"] == float(profile_height)) &
                (self.price_data["thickness"] == float(thickness))
            )
            material_prices = self.price_data[mask]
            if material_prices.empty:
                raise ValueError("Цены для выбранной конфигурации не найдены")
            material_prices = material_prices.iloc[0]
            
            # Расчет стоимости
            total = 0
            cost_breakdown = {}
            details = []
            
            # Основные материалы
            base_cost = length * height * material_prices.get("base_price", 0)
            cost_breakdown["Материал"] = base_cost
            details.append(f"Материал ({metal_type}, выс.{profile_height}мм, толщ.{thickness}мм): {base_cost:,.2f} руб.")
            total += base_cost
            
            # Полимерное покрытие
            if coating and "coating_price" in material_prices:
                coating_cost = length * height * material_prices["coating_price"]
                cost_breakdown["Покрытие"] = coating_cost
                details.append(f"Полимерное покрытие: {coating_cost:,.2f} руб.")
                total += coating_cost
            
            # Столбы
            if "post_price" in material_prices:
                posts_cost = posts * material_prices["post_price"]
                cost_breakdown["Столбы"] = posts_cost
                details.append(f"Столбы ({posts} шт.): {posts_cost:,.2f} руб.")
                total += posts_cost
            
            # Заглубление столбов
            if "post_depth_price" in material_prices:
                depth_cost = posts * post_depth * material_prices["post_depth_price"]
                cost_breakdown["Заглубление"] = depth_cost
                details.append(f"Заглубление столбов ({post_depth} м): {depth_cost:,.2f} руб.")
                total += depth_cost
            
            # Ворота и калитки
            if gates > 0:
                gates_cost = gates * material_prices["gate_price"]
                cost_breakdown["Ворота"] = gates_cost
                details.append(f"Ворота ({gates} шт.): {gates_cost:,.2f} руб.")
                total += gates_cost
            
            if doors > 0:
                doors_cost = doors * material_prices["door_price"]
                cost_breakdown["Калитки"] = doors_cost
                details.append(f"Калитки ({doors} шт.): {doors_cost:,.2f} руб.")
                total += doors_cost
            
            # Забутовка
            if foundation and "foundation_price" in material_prices:
                foundation_cost = length * material_prices["foundation_price"]
                cost_breakdown["Забутовка"] = foundation_cost
                details.append(f"Забутовка: {foundation_cost:,.2f} руб.")
                total += foundation_cost
            
            # Доставка
            if delivery_distance > 0 and "delivery_price_per_km" in material_prices:
                delivery_cost = delivery_distance * material_prices["delivery_price_per_km"]
                cost_breakdown["Доставка"] = delivery_cost
                details.append(f"Доставка ({delivery_distance} км): {delivery_cost:,.2f} руб.")
                total += delivery_cost
            
            # Сохраняем результаты для экспорта
            self.last_calculation = {
                "parameters": {
                    "Длина": f"{length} м",
                    "Высота": f"{height} м",
                    "Количество столбов": posts,
                    "Заглубление столбов": f"{post_depth} м",
                    "Тип металла": metal_type,
                    "Высота профиля": f"{profile_height} мм",
                    "Толщина металла": f"{thickness} мм",
                    "Покрытие": "Да" if coating else "Нет",
                    "Ворота": gates,
                    "Калитки": doors,
                    "Забутовка": "Да" if foundation else "Нет",
                    "Расстояние доставки": f"{delivery_distance} км"
                },
                "cost_breakdown": cost_breakdown,
                "details": details,
                "total": total
            }
            
            # Обновляем интерфейс
            self.results_var.set(f"Итоговая стоимость: {total:,.2f} руб.")
            self.details_var.set("\n".join(details))
            self.status_var.set("Расчет завершен успешно")
        except ValueError as ve:
            messagebox.showerror("Ошибка ввода", str(ve))
            self.status_var.set("Ошибка расчета")
        except KeyError as ke:
            messagebox.showerror("Ошибка данных", f"Отсутствует необходимый параметр: {str(ke)}")
            self.status_var.set("Ошибка расчета")
        except Exception as e:
            messagebox.showerror("Неизвестная ошибка", str(e))
            self.status_var.set("Ошибка расчета")

    def save_results(self):
        try:
            if not hasattr(self, 'last_calculation'):
                raise ValueError("Сначала выполните расчет!")
            
            # Создаем DataFrame для экспорта
            result_data = []
            for component, cost in self.last_calculation["cost_breakdown"].items():
                result_data.append({
                    "Компонент": component,
                    "Стоимость, руб.": cost,
                    "Примечание": ""
                })
            
            # Добавляем итоговую строку
            result_data.append({
                "Компонент": "ИТОГО",
                "Стоимость, руб.": self.last_calculation["total"],
                "Примечание": ""
            })
            
            df = pd.DataFrame(result_data)
            
            # Запрашиваем место сохранения
            file_path = filedialog.asksaveasfilename(
                filetypes=[("Excel files", "*.xlsx")],
                title="Сохранить результаты",
                initialfile=self.output_file,
                defaultextension=".xlsx"
            )
            if not file_path:
                return
            
            # Сохраняем в Excel с несколькими листами
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name="Детализация", index=False)
                
                # Добавляем лист с параметрами
                params_df = pd.DataFrame(
                    list(self.last_calculation["parameters"].items()),
                    columns=["Параметр", "Значение"]
                )
                params_df.to_excel(writer, sheet_name="Параметры", index=False)
                
                # Форматируем итоговую стоимость
                total_df = pd.DataFrame({
                    "Итоговая стоимость": [f"{self.last_calculation['total']:,.2f} руб."]
                })
                total_df.to_excel(writer, sheet_name="Итог", index=False)
            
            self.status_var.set(f"Файл сохранен: {os.path.basename(file_path)}")
        except Exception as e:
            messagebox.showerror("Ошибка сохранения", str(e))

if __name__ == "__main__":
    root = tk.Tk()
    app = FenceCalculator(root)
    root.mainloop()