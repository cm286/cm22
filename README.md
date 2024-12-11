import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import csv
import datetime
from openpyxl import Workbook

class EmployeeManager:
    def __init__(self, root):
        self.root = root
        self.root.title("Quản lý nhân viên")
        self.root.geometry("600x400")
        self.create_widgets()
        self.employees = []

    def create_widgets(self):
        frame = ttk.Frame(self.root, padding=10)
        frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frame, text="Mã:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.id_entry = ttk.Entry(frame)
        self.id_entry.grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(frame, text="Tên:").grid(row=0, column=2, padx=5, pady=5, sticky=tk.W)
        self.name_entry = ttk.Entry(frame)
        self.name_entry.grid(row=0, column=3, padx=5, pady=5)

        ttk.Label(frame, text="Đơn vị:").grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        self.department_entry = ttk.Entry(frame)
        self.department_entry.grid(row=1, column=1, padx=5, pady=5)

        ttk.Label(frame, text="Chức danh:").grid(row=1, column=2, padx=5, pady=5, sticky=tk.W)
        self.position_entry = ttk.Entry(frame)
        self.position_entry.grid(row=1, column=3, padx=5, pady=5)

        ttk.Label(frame, text="Ngày sinh (dd/mm/yyyy):").grid(row=2, column=0, padx=5, pady=5, sticky=tk.W)
        self.birth_entry = ttk.Entry(frame)
        self.birth_entry.grid(row=2, column=1, padx=5, pady=5)

        ttk.Label(frame, text="Giới tính:").grid(row=2, column=2, padx=5, pady=5, sticky=tk.W)
        self.gender_var = tk.StringVar(value="Nam")
        ttk.Radiobutton(frame, text="Nam", variable=self.gender_var, value="Nam").grid(row=2, column=3, sticky=tk.W)
        ttk.Radiobutton(frame, text="Nữ", variable=self.gender_var, value="Nữ").grid(row=2, column=3, sticky=tk.E)

        ttk.Label(frame, text="Số CMND:").grid(row=3, column=0, padx=5, pady=5, sticky=tk.W)
        self.id_number_entry = ttk.Entry(frame)
        self.id_number_entry.grid(row=3, column=1, padx=5, pady=5)

        ttk.Label(frame, text="Nơi cấp:").grid(row=3, column=2, padx=5, pady=5, sticky=tk.W)
        self.place_of_issue_entry = ttk.Entry(frame)
        self.place_of_issue_entry.grid(row=3, column=3, padx=5, pady=5)

        ttk.Button(frame, text="Lưu thông tin", command=self.save_employee).grid(row=4, column=0, padx=5, pady=10)
        ttk.Button(frame, text="Sinh nhật hôm nay", command=self.show_today_birthdays).grid(row=4, column=1, padx=5, pady=10)
        ttk.Button(frame, text="Xuất danh sách", command=self.export_to_excel).grid(row=4, column=2, padx=5, pady=10)

    def save_employee(self):
        employee = {
            "id": self.id_entry.get(),
            "name": self.name_entry.get(),
            "department": self.department_entry.get(),
            "position": self.position_entry.get(),
            "birth_date": self.birth_entry.get(),
            "gender": self.gender_var.get(),
            "id_number": self.id_number_entry.get(),
            "place_of_issue": self.place_of_issue_entry.get(),
        }

        self.employees.append(employee)
        with open("employees.csv", "a", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=employee.keys())
            if f.tell() == 0:
                writer.writeheader()
            writer.writerow(employee)

        messagebox.showinfo("Thông báo", "Đã lưu thông tin nhân viên!")
        self.clear_entries()

    def clear_entries(self):
        self.id_entry.delete(0, tk.END)
        self.name_entry.delete(0, tk.END)
        self.department_entry.delete(0, tk.END)
        self.position_entry.delete(0, tk.END)
        self.birth_entry.delete(0, tk.END)
        self.id_number_entry.delete(0, tk.END)
        self.place_of_issue_entry.delete(0, tk.END)

    def show_today_birthdays(self):
        today = datetime.datetime.now().strftime("%d/%m")
        today_birthdays = [e for e in self.employees if e["birth_date"][:5] == today]

        if not today_birthdays:
            messagebox.showinfo("Thông báo", "Không có nhân viên nào sinh nhật hôm nay!")
        else:
            msg = "Nhân viên sinh nhật hôm nay:\n"
            msg += "\n".join([f"{e['name']} ({e['birth_date']})" for e in today_birthdays])
            messagebox.showinfo("Thông báo", msg)

    def export_to_excel(self):
        self.employees.sort(key=lambda x: datetime.datetime.strptime(x["birth_date"], "%d/%m/%Y"), reverse=True)
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Danh sách nhân viên"
        headers = ["Mã", "Tên", "Đơn vị", "Chức danh", "Ngày sinh", "Giới tính", "Số CMND", "Nơi cấp"]
        sheet.append(headers)
        for emp in self.employees:
            sheet.append([emp["id"], emp["name"], emp["department"], emp["position"], emp["birth_date"], emp["gender"], emp["id_number"], emp["place_of_issue"]])
        filepath = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if filepath:
            workbook.save(filepath)
            messagebox.showinfo("Thông báo", "Đã xuất danh sách nhân viên ra Excel!")

if __name__ == "__main__":
    root = tk.Tk()
    app = EmployeeManager(root)
    root.mainloop()
