import os
import pandas as pd
import requests
from bs4 import BeautifulSoup
import cloudinary
import cloudinary.uploader
from urllib.parse import urlparse, urljoin
import sys

# Đặt encoding mặc định cho sys.stdout là utf-8

sys.stdout.reconfigure(encoding='utf-8')

# Cấu hình Cloudinary
cloudinary.config(
    cloud_name="dje0fenqy",
    api_key="666449536579451",
    api_secret="9RKnJZ9fdwhozAPwrtnUKrR0Hsk",
    secure=True,
)

# Đường dẫn đến file Excel
input_file = "product.xlsx"

# Kiểm tra file Excel tồn tại
if not os.path.exists(input_file):
    print(f"File '{input_file}' không tồn tại. Vui lòng kiểm tra lại đường dẫn.")
    exit()

# Đọc file Excel
try:
    data = pd.read_excel(input_file)
    print(f"Đã đọc thành công file '{input_file}' với {len(data)} dòng dữ liệu.")
except Exception as e:
    print(f"Lỗi khi đọc file Excel: {e}")
    exit()

# Hàm lấy URL cơ sở
def get_base_url(url):
    parsed_url = urlparse(url)
    if not parsed_url.scheme:  # Nếu không có scheme (http/https)
        return f"http://{parsed_url.netloc}"
    return f"{parsed_url.scheme}://{parsed_url.netloc}"

# Hàm tải ảnh từ URL lên Cloudinary
def upload_to_cloudinary(image_path):
    try:
        response = cloudinary.uploader.upload(image_path)
        return response['url']
    except Exception as e:
        print(f"Failed to upload {image_path} to Cloudinary: {e}")
        return None

# Hàm tải ảnh tạm thời về máy
def download_temp_image(url):
    try:
        response = requests.get(url, stream=True, timeout=10)
        if response.status_code == 200:
            temp_folder = "temp"
            os.makedirs(temp_folder, exist_ok=True)
            temp_file = os.path.join(temp_folder, "temp_image.jpg")
            with open(temp_file, 'wb') as f:
                for chunk in response.iter_content(1024):
                    f.write(chunk)
            return temp_file
    except Exception as e:
        print(f"Failed to download {url}: {e}")
    return None

# Nhập số thứ tự hàng bắt đầu
try:
    start_index = int(input("Nhập số thứ tự hàng bắt đầu : "))-2
    if start_index < 0 or start_index >= len(data):
        raise ValueError("Số thứ tự không hợp lệ.")
except ValueError as ve:
    print(f"Lỗi: {ve}")
    exit()

# Kiểm tra số lượng sản phẩm còn lại
remaining_products = len(data) - start_index
if remaining_products <= 0:
    print(f"Không có sản phẩm nào để xử lý từ hàng {start_index}. Vui lòng nhập số thứ tự hợp lệ.")
    exit()

# Nhập giới hạn số lượng sản phẩm xử lý
try:
    limit = int(input(f"Nhập số lượng sản phẩm cần xử lý (tối đa {remaining_products}): "))
    if limit <= 0 or limit > remaining_products:
        raise ValueError(f"Giới hạn không hợp lệ. Bạn chỉ có thể xử lý tối đa {remaining_products} sản phẩm.")
except ValueError as ve:
    print(f"Lỗi: {ve}")
    exit()

# Tính toán lại số lượng sản phẩm thực tế sẽ được xử lý
end_index = start_index + limit
print(f"\nĐang xử lý từ hàng {start_index} đến {end_index - 1} (tổng cộng {limit} sản phẩm).")

# Danh sách lưu trữ log sản phẩm và hình ảnh đã tải lên
log_entries = []

# Duyệt qua từng sản phẩm trong phạm vi được chỉ định
for index in range(start_index, end_index):
    row = data.iloc[index]
    product_name = row['Tên hàng']
    search_url = row['URL tìm kiếm hình ảnh']

    # Kiểm tra nếu cột "Hình ảnh" đã có dữ liệu, thì bỏ qua
    if pd.notna(row['Hình ảnh']) and row['Hình ảnh'].strip() != "":
        print(f"Đã có hình ảnh cho sản phẩm: {product_name}. Bỏ qua.")
        continue  # Bỏ qua sản phẩm này nếu đã có hình ảnh

    print(f"\nĐang xử lý sản phẩm: {product_name} (Hàng {index})")
    print(f"URL tìm kiếm: {search_url}")

    try:
        # Lấy URL cơ sở
        base_url = get_base_url(search_url)

        # Gửi yêu cầu đến URL tìm kiếm hình ảnh
        response = requests.get(search_url, timeout=10)
        soup = BeautifulSoup(response.text, 'html.parser')

        # Tìm tất cả các thẻ <img>
        image_tags = soup.find_all('img')

        # Kiểm tra xem có ít nhất 5 hình ảnh không (để lấy 3 từ thẻ thứ 3 trở đi)
        if len(image_tags) >= 5:
            # Chọn các thẻ img từ thứ 3 đến thứ 5 (tính từ 0)
            img_urls = [urljoin(base_url, image_tags[i]['src']) for i in range(2, 5)]

            print(f"\nĐang xử lý các hình ảnh thứ 3 đến 5:")
            for img_url in img_urls:
                print(f"URL gốc: {img_url}")

            # Tải xuống và upload các hình ảnh lên Cloudinary
            cloudinary_urls = []
            for img_url in img_urls:
                # Tải xuống hình ảnh
                temp_file = download_temp_image(img_url)
                if temp_file:
                    cloud_url = upload_to_cloudinary(temp_file)
                    if cloud_url:
                        cloudinary_urls.append(cloud_url)
                        print(f"Đã tải lên Cloudinary: {cloud_url}")
                    else:
                        print("Không thể tải lên Cloudinary")
                    os.remove(temp_file)  # Xóa file tạm sau khi upload
                else:
                    print(f"Không thể tải xuống hình ảnh: {img_url}")

            # Ghi các liên kết ảnh vào cột "Hình ảnh"
            if cloudinary_urls:
                data.at[index, 'Hình ảnh'] = ",".join(cloudinary_urls)
                # Thêm sản phẩm và URL vào log
                log_entries.append(f"Sản phẩm: {product_name}\nLink hình ảnh: {', '.join(cloudinary_urls)}")
                print(f"\nĐã cập nhật cột 'Hình ảnh' cho {product_name}: {data.at[index, 'Hình ảnh']}")
            else:
                print(f"Không tìm thấy hình ảnh hợp lệ cho {product_name}")

        else:
            print(f"Không đủ hình ảnh (cần ít nhất 5) cho {product_name}")

    except Exception as e:
        print(f"Lỗi khi xử lý {product_name}: {e}")

# Sau khi hoàn tất xử lý tất cả sản phẩm, lưu lại file Excel với cột "Hình ảnh" đã cập nhật
data.to_excel(input_file, index=False)
print("\nHoàn tất tải ảnh và cập nhật file Excel!")

# In ra log cho tất cả sản phẩm
print("\nTất cả sản phẩm và link hình ảnh đã upload:")
for log in log_entries:
    print(log)
