import os
import win32com.client as win32

def convert_all_excel_to_pdf(folder_path, output_folder):
    # Kiểm tra thư mục đầu vào và đầu ra
    if not os.path.exists(folder_path):
        print("Thư mục không tồn tại!")
        return
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # Mở ứng dụng Excel
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False

    # Lặp qua tất cả các file trong thư mục
    for filename in os.listdir(folder_path):
        if filename.endswith(".xlsx"):
            input_file = os.path.join(folder_path, filename)
            output_file = os.path.join(output_folder, filename.replace(".xlsx", ".pdf"))
            print(f"Đang xử lý: {input_file}")

            # Mở file Excel và xuất PDF
            try:
                workbook = excel.Workbooks.Open(input_file)
                workbook.ExportAsFixedFormat(0, output_file)
                workbook.Close(False)
                print(f"Đã chuyển đổi: {output_file}")
            except Exception as e:
                print(f"Lỗi khi chuyển đổi file {filename}: {e}")

    # Đóng Excel
    excel.Quit()
    print("Hoàn tất quá trình chuyển đổi.")