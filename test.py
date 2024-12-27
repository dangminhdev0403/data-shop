import os
import pandas as pd
import cloudinary
import cloudinary.uploader

# Cấu hình Cloudinary
cloudinary.config(
    cloud_name="dje0fenqy",
    api_key="666449536579451",
    api_secret="9RKnJZ9fdwhozAPwrtnUKrR0Hsk",
    secure=True,
)

# Đọc file Excel
input_file = "product.xlsx"
image_folder = "images"  # Thư mục chứa các thư mục ảnh

data = pd.read_excel(input_file)

# Kiểm tra nếu thư mục ảnh tồn tại
if not os.path.exists(image_folder):
    print(f"Thư mục '{image_folder}' không tồn tại. Vui lòng kiểm tra lại.")
    exit()

# Hàm tải ảnh lên Cloudinary
def upload_to_cloudinary(image_path):
    try:
        response = cloudinary.uploader.upload(image_path)
        return response['url']
    except Exception as e:
        print(f"Không thể tải lên Cloudinary: {e}")
        return None

# Lưu log các sản phẩm đã được xử lý và các lỗi
log_entries = []
updated_count = 0  # Số sản phẩm đã được cập nhật

# Duyệt qua tất cả các sản phẩm trong file Excel
for index, row in data.iterrows():
    product_name = row['Tên hàng'].strip()  # Tên sản phẩm, giữ nguyên chữ hoa chữ thường và dấu cách

    # Kiểm tra nếu cột "Hình ảnh" đã có dữ liệu, thì bỏ qua
    if pd.notna(row['Hình ảnh']) and row['Hình ảnh'].strip() != "":
        print(f"Đã có hình ảnh cho sản phẩm: {product_name}. Bỏ qua.")
        continue  # Bỏ qua sản phẩm này nếu đã có hình ảnh

    print(f"\nĐang xử lý sản phẩm: {product_name} (Hàng {index})")

    # Kiểm tra thư mục ảnh tương ứng với tên sản phẩm (Không thay đổi tên sản phẩm)
    product_folder = os.path.join(image_folder, product_name)  # Tên thư mục là tên sản phẩm

    if os.path.exists(product_folder) and os.path.isdir(product_folder):
        print(f"Tìm thấy thư mục ảnh: {product_folder}")
        # Lấy tất cả các ảnh trong thư mục sản phẩm
        image_paths = []
        for filename in os.listdir(product_folder):
            if filename.lower().endswith(('.jpg', '.jpeg', '.png', '.gif', '.webp')):
                image_paths.append(os.path.join(product_folder, filename))

        # In ra các ảnh tìm thấy trong thư mục
        print(f"Các ảnh tìm thấy trong thư mục {product_folder}: {image_paths}")

        # Nếu tìm thấy ảnh
        if image_paths:
            cloudinary_urls = []
            for image_path in image_paths:
                cloud_url = upload_to_cloudinary(image_path)
                if cloud_url:
                    cloudinary_urls.append(cloud_url)
                    print(f"Đã tải lên Cloudinary: {cloud_url}")

            # Cập nhật URL vào cột "Hình ảnh" của sản phẩm
            if cloudinary_urls:
                data.at[index, 'Hình ảnh'] = ",".join(cloudinary_urls)
                log_entries.append(f"Sản phẩm: {product_name}\nLink hình ảnh: {', '.join(cloudinary_urls)}")
                updated_count += 1
                print(f"\nĐã cập nhật cột 'Hình ảnh' cho {product_name}: {data.at[index, 'Hình ảnh']}")
        else:
            print(f"Không tìm thấy ảnh trong thư mục của sản phẩm: {product_name}")
    else:
        print(f"Không tìm thấy thư mục ảnh cho sản phẩm: {product_name}")

# Sau khi hoàn tất xử lý tất cả sản phẩm, lưu lại file Excel với cột "Hình ảnh" đã cập nhật
data.to_excel(input_file, index=False)
print("\nHoàn tất tải ảnh và cập nhật file Excel!")

# In ra log cho tất cả sản phẩm đã xử lý
print("\nTất cả sản phẩm và link hình ảnh đã upload:")
for log in log_entries:
    print(log)

# Thông báo kết quả tổng quan
print(f"\nTổng số sản phẩm đã được cập nhật: {updated_count}")
