import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import mysql.connector
from mysql.connector import Error
from openpyxl import Workbook
# 1. KẾT NỐI MYSQL
def connect_db():
    try:
        conn = mysql.connector.connect(
            host="localhost",
            user="root",
            password="123456",
            database="qlnhanvien",
            port=3305
        )
        return conn
    except Error as e:
        messagebox.showerror("Lỗi CSDL", f"Không thể kết nối MySQL:\n{e}")
        return None
# 2. CANH GIỮA CỬA SỔ
def center_window(win, w=900, h=600):
    ws = win.winfo_screenwidth()
    hs = win.winfo_screenheight()
    x = (ws - w) // 2
    y = (hs - h) // 2
    win.geometry(f"{w}x{h}+{x}+{y}")
# 3. XÓA Ô NHẬP
def clear_input():
    entry_matuyen.config(state="normal")
    entry_matuyen.delete(0, tk.END)
    entry_tentuyen.delete(0, tk.END)
    entry_diemdi.delete(0, tk.END)
    entry_thoigian.delete(0, tk.END)
    entry_giatien.delete(0, tk.END)
    cbb_phuongtien.set("")
    entry_timkiem.delete(0, tk.END)
# 4. TẢI DỮ LIỆU TỪ SQL
def load_data():
    for i in tree.get_children():
        tree.delete(i)
    conn = connect_db()
    if not conn:
        return
    try:
        cur = conn.cursor()
        cur.execute("SELECT * FROM tuyendulich ORDER BY matuyen")
        rows = cur.fetchall()
        if not rows:
            messagebox.showinfo("Thông tin", "Chưa có tuyến du lịch nào.")
        for row in rows:
            row_display = list(row)
            # Format giá tiền: 200 -> 200,000
            row_display[4] = f"{row_display[4]:,.0f}"
            tree.insert("", tk.END, values=row_display)
        conn.close()
    except Error as e:
        messagebox.showerror("Lỗi CSDL", str(e))

# 5. THÊM TUYẾN DU LỊCH
def them_tuyendulich():
    matuyen = entry_matuyen.get().strip()
    tentuyen = entry_tentuyen.get().strip()
    diemdi = entry_diemdi.get().strip()
    phuongtien = cbb_phuongtien.get().strip()

    try:
        thoigian = int(entry_thoigian.get().strip())
    except ValueError:
        messagebox.showerror("Lỗi", "Thời gian phải là số nguyên!")
        return
    try:
        giatien = float(entry_giatien.get().strip())
    except ValueError:
        messagebox.showerror("Lỗi", "Giá tiền phải là số!")
        return

    if not all([matuyen, tentuyen, diemdi, thoigian, giatien, phuongtien]):
        messagebox.showwarning("Thiếu dữ liệu", "Vui lòng nhập đủ thông tin!")
        return

    conn = connect_db()
    if not conn:
        return
    try:
        cur = conn.cursor()
        cur.execute("SELECT matuyen FROM tuyendulich WHERE matuyen=%s", (matuyen,))
        if cur.fetchone():
            messagebox.showerror("Lỗi", "Mã tuyến đã tồn tại!")
            return
        sql = "INSERT INTO tuyendulich VALUES (%s, %s, %s, %s, %s, %s)"
        cur.execute(sql, (matuyen, tentuyen, diemdi, thoigian, giatien, phuongtien))
        conn.commit()
        messagebox.showinfo("Thành công", "Đã thêm tuyến du lịch!")

        # Hiển thị Treeview với giá tiền định dạng
        row_display = [matuyen, tentuyen, diemdi, thoigian, f"{giatien:,.0f}", phuongtien]
        tree.insert("", tk.END, values=row_display)

        clear_input()
    except Error as e:
        messagebox.showerror("Lỗi Thêm", str(e))
    finally:
        conn.close()

# =============================
# 6. CHỌN DÒNG SỬA
# =============================
def sua_tuyendulich():
    selected = tree.selection()
    if not selected:
        messagebox.showwarning("Chưa chọn", "Chọn 1 tuyến để sửa!")
        return
    values = tree.item(selected[0])["values"]
    clear_input()
    entry_matuyen.insert(0, values[0])
    entry_matuyen.config(state="disabled")
    entry_tentuyen.insert(0, values[1])
    entry_diemdi.insert(0, values[2])
    entry_thoigian.insert(0, values[3])
    # Xóa dấu phẩy trước khi hiển thị trong Entry
    entry_giatien.insert(0, str(values[4]).replace(',', ''))
    cbb_phuongtien.set(values[5])

# =============================
# 7. LƯU SỬA
# =============================
def luu_tuyendulich():
    entry_matuyen.config(state="normal")
    matuyen = entry_matuyen.get().strip()
    tentuyen = entry_tentuyen.get().strip()
    diemdi = entry_diemdi.get().strip()
    phuongtien = cbb_phuongtien.get().strip()

    try:
        thoigian = int(entry_thoigian.get().strip())
    except ValueError:
        messagebox.showerror("Lỗi", "Thời gian phải là số nguyên!")
        return
    try:
        giatien = float(entry_giatien.get().strip())
    except ValueError:
        messagebox.showerror("Lỗi", "Giá tiền phải là số!")
        return

    if not all([matuyen, tentuyen, diemdi, thoigian, giatien, phuongtien]):
        messagebox.showwarning("Thiếu dữ liệu", "Vui lòng nhập đủ thông tin!")
        return

    conn = connect_db()
    if not conn:
        return
    try:
        cur = conn.cursor()
        sql = """
            UPDATE tuyendulich
            SET tentuyen=%s, diemdi=%s, thoigian=%s, giatien=%s, phuongtien=%s
            WHERE matuyen=%s
        """
        cur.execute(sql, (tentuyen, diemdi, thoigian, giatien, phuongtien, matuyen))
        conn.commit()
        messagebox.showinfo("Thành công", "Đã lưu thay đổi!")

        # Cập nhật Treeview
        load_data()
        clear_input()
    except Error as e:
        messagebox.showerror("Lỗi Lưu", str(e))
    finally:
        conn.close()

# =============================
# 8. XÓA TUYẾN
# =============================
def xoa_tuyendulich():
    selected = tree.selection()
    if not selected:
        messagebox.showwarning("Chưa chọn", "Chọn 1 tuyến để xóa!")
        return
    matuyen = tree.item(selected[0])["values"][0]
    if not messagebox.askyesno("Xóa", f"Xóa tuyến {matuyen}?"):
        return
    conn = connect_db()
    if not conn:
        return
    try:
        cur = conn.cursor()
        cur.execute("DELETE FROM tuyendulich WHERE matuyen=%s", (matuyen,))
        conn.commit()
        messagebox.showinfo("Đã xóa", "Xóa thành công!")
        load_data()
    except Error as e:
        messagebox.showerror("Lỗi Xóa", str(e))
    finally:
        conn.close()

# =============================
# 9. TÌM KIẾM
# =============================
def tim_kiem():
    keyword = entry_timkiem.get().strip()
    for i in tree.get_children():
        tree.delete(i)
    conn = connect_db()
    if not conn:
        return
    try:
        cur = conn.cursor()
        cur.execute("SELECT * FROM tuyendulich WHERE tentuyen LIKE %s", (f"%{keyword}%",))
        rows = cur.fetchall()
        if not rows:
            messagebox.showinfo("Thông tin", "Không tìm thấy tuyến nào.")
        for row in rows:
            row_display = list(row)
            row_display[4] = f"{row_display[4]:,.0f}"
            tree.insert("", tk.END, values=row_display)
    except Error as e:
        messagebox.showerror("Lỗi", str(e))
    finally:
        conn.close()

# =============================
# 10. XUẤT EXCEL
# =============================
def xuat_excel():
    if not tree.get_children():
        messagebox.showwarning("Thông báo", "Không có dữ liệu để xuất!")
        return

    file_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")],
        title="Lưu file Excel"
    )
    if not file_path:
        return

    wb = Workbook()
    ws = wb.active
    ws.title = "Tuyến Du Lịch"

    headers = ["Mã tuyến", "Tên tuyến", "Điểm đi", "Thời gian (ngày)", "Giá tiền (VNĐ)", "Phương tiện"]
    ws.append(headers)

    for row_id in tree.get_children():
        row = tree.item(row_id)["values"]
        # Chuyển giá tiền thành số thực trước khi xuất
        price = int(str(row[4]).replace(',', ''))
        row[4] = f"{price:,} VNĐ"
        ws.append(row)

    try:
        wb.save(file_path)
        messagebox.showinfo("Thành công", f"Đã xuất dữ liệu ra Excel:\n{file_path}")
    except Exception as e:
        messagebox.showerror("Lỗi", f"Không thể lưu file:\n{e}")

# =============================
# 11. GIAO DIỆN
# =============================
root = tk.Tk()
root.title("Quản lý Tuyến Du lịch")
center_window(root, 900, 600)
root.resizable(False, False)

lbl_title = tk.Label(root, text="QUẢN LÝ TUYẾN DU LỊCH", font=("Arial", 18, "bold"), fg="blue")
lbl_title.pack(pady=10)

# -------- Frame nhập liệu --------
frame_info = tk.Frame(root)
frame_info.pack(padx=10, pady=5, fill="x")

tk.Label(frame_info, text="Mã tuyến").grid(row=0, column=0, padx=5, pady=5, sticky="w")
entry_matuyen = tk.Entry(frame_info, width=20)
entry_matuyen.grid(row=0, column=1, padx=5)

tk.Label(frame_info, text="Phương tiện").grid(row=0, column=2, padx=5, pady=5, sticky="w")
cbb_phuongtien = ttk.Combobox(frame_info, width=18, values=["Máy bay", "Tàu hỏa", "Ô tô", "Xe khách"])
cbb_phuongtien.grid(row=0, column=3, padx=5)

tk.Label(frame_info, text="Tên tuyến").grid(row=1, column=0, padx=5,  sticky="w")
entry_tentuyen = tk.Entry(frame_info, width=20)
entry_tentuyen.grid(row=1, column=1, padx=5)

tk.Label(frame_info, text="Điểm đi").grid(row=1, column=2, padx=5,  sticky="w")
entry_diemdi = tk.Entry(frame_info, width=20)
entry_diemdi.grid(row=1, column=3, padx=5)

tk.Label(frame_info, text="Thời gian (ngày)").grid(row=2, column=0, padx=5, sticky="w")
entry_thoigian = tk.Entry(frame_info, width=20)
entry_thoigian.grid(row=2, column=1, padx=5)

tk.Label(frame_info, text="Giá tiền (VNĐ)").grid(row=2, column=2, padx=5)
entry_giatien = tk.Entry(frame_info, width=20)
entry_giatien.grid(row=2, column=3, padx=5)

# Tìm kiếm
tk.Label(frame_info, text="Tìm theo tên tuyến").grid(row=3, column=0, padx=5, pady=5)
entry_timkiem = tk.Entry(frame_info, width=20)
entry_timkiem.grid(row=3, column=1, padx=5)
tk.Button(frame_info, text="Tìm kiếm", command=tim_kiem).grid(row=3, column=2, padx=5)

# -------- Treeview --------
columns = ("matuyen", "tentuyen", "diemdi", "thoigian", "giatien", "phuongtien")
tree = ttk.Treeview(root, columns=columns, show="headings", height=15)
tree.pack(padx=10, pady=10, fill="both", expand=True)

scrollbar = ttk.Scrollbar(root, orient="vertical", command=tree.yview)
tree.configure(yscroll=scrollbar.set)
scrollbar.pack(side="right", fill="y")

tree.heading("matuyen", text="Mã tuyến")
tree.heading("tentuyen", text="Tên tuyến")
tree.heading("diemdi", text="Điểm đi")
tree.heading("thoigian", text="Thời gian (ngày)")
tree.heading("giatien", text="Giá tiền (VNĐ)")
tree.heading("phuongtien", text="Phương tiện")

for col in columns:
    tree.column(col, width=130, anchor="center")

# -------- Nút chức năng --------
frame_btn = tk.Frame(root)
frame_btn.pack(pady=5)
tk.Button(frame_btn, text="Thêm", width=12, command=them_tuyendulich).grid(row=0, column=0, padx=5)
tk.Button(frame_btn, text="Sửa", width=12, command=sua_tuyendulich).grid(row=0, column=1, padx=5)
tk.Button(frame_btn, text="Lưu", width=12, command=luu_tuyendulich).grid(row=0, column=2, padx=5)
tk.Button(frame_btn, text="Hủy", width=12, command=clear_input).grid(row=0, column=3, padx=5)
tk.Button(frame_btn, text="Xóa", width=12, command=xoa_tuyendulich).grid(row=0, column=4, padx=5)
tk.Button(frame_btn, text="Thoát", width=12, command=root.quit).grid(row=0, column=5, padx=5)
tk.Button(frame_btn, text="Xuất Excel", width=12, command=xuat_excel).grid(row=0, column=6, padx=5)

# Load dữ liệu ban đầu
load_data()

root.mainloop()
