import os
import pandas as pd

# Đọc file Excel
input_file = "product.xlsx"
output_folder = "images"  # Thư mục gốc để tạo thư mục sản phẩm

data = pd.read_excel(input_file)

# Kiểm tra nếu thư mục gốc tồn tại, nếu không thì tạo mới
if not os.path.exists(output_folder):
    os.makedirs(output_folder)
    print(f"Đã tạo thư mục gốc: {output_folder}")

# Duyệt qua từng dòng trong file Excel để tạo thư mục tương ứng với tên sản phẩm
for index, row in data.iterrows():
    product_name = str(row['Tên hàng']).strip()  # Lấy tên sản phẩm, loại bỏ khoảng trắng thừa

    if product_name:  # Kiểm tra nếu tên sản phẩm không rỗng
        product_folder = os.path.join(output_folder, product_name)

        # Tạo thư mục nếu chưa tồn tại
        if not os.path.exists(product_folder):
            os.makedirs(product_folder)
            print(f"Đã tạo thư mục: {product_folder}")
        else:
            print(f"Thư mục đã tồn tại: {product_folder}")
    else:
        print(f"Bỏ qua dòng {index} vì không có tên sản phẩm.")

print("\nHoàn tất tạo thư mục cho tất cả sản phẩm trong file Excel.")
