import os
from PyPDF2 import PdfMerger

def merge_pdfs_in_folder(folder_path_in, folder_path_out, output_file_name):
    # Kiểm tra nếu output_file_name chưa có đuôi .pdf thì thêm vào
    if not output_file_name.endswith(".pdf"):
        output_file_name += ".pdf"

    merger = PdfMerger()

    # Duyệt qua tất cả file PDF trong folder, sắp xếp theo tên
    for file_name in sorted(os.listdir(folder_path_in)):
        if file_name.endswith(".pdf"):
            file_path = os.path.join(folder_path_in, file_name)
            merger.append(file_path)

    # Lưu file PDF đã gộp với tên được cung cấp
    output_path = os.path.join(folder_path_out, output_file_name)
    merger.write(output_path)
    merger.close()

    return output_path
