import os
import pandas as pd
import requests
from bs4 import BeautifulSoup
import unicodedata
import re

# Đường dẫn đến file Excel (cùng thư mục với file Python)
input_file = "test.xlsx"

# Đọc file Excel
data = pd.read_excel(input_file)

# Hàm loại bỏ dấu tiếng Việt
def remove_vietnamese_accents(text):
    text = unicodedata.normalize('NFD', text)
    text = re.sub(r'[\u0300-\u036f]', '', text)  # Loại bỏ dấu
    text = re.sub(r'[\u02c6\u02dc]', '', text)  # Loại bỏ ký tự đặc biệt
    text = re.sub(r'[\u2018\u2019\u201c\u201d]', '', text)  # Loại bỏ ký tự quote
    text = re.sub(r'\s+', '_', text)  # Thay dấu cách bằng dấu gạch dưới
    return text

# Hàm tải ảnh
def download_image(url, folder, filename):
    try:
        response = requests.get(url, stream=True, timeout=10)
        if response.status_code == 200:
            with open(os.path.join(folder, filename), 'wb') as f:
                for chunk in response.iter_content(1024):
                    f.write(chunk)
            print(f"Downloaded: {filename}")
    except Exception as e:
        print(f"Failed to download {url}: {e}")

# Duyệt qua từng sản phẩm
for index, row in data.iterrows():
    product_name = row['Tên hàng']
    search_url = row['URL tìm kiếm hình ảnh']

    # Loại bỏ dấu và thay dấu cách trong tên sản phẩm
    clean_product_name = remove_vietnamese_accents(product_name)

    # Tạo thư mục riêng cho mỗi sản phẩm (dùng tên sản phẩm đã loại bỏ dấu làm tên thư mục)
    product_folder = f"downloaded_images/{clean_product_name}"
    os.makedirs(product_folder, exist_ok=True)

    try:
        # Gửi yêu cầu đến URL tìm kiếm hình ảnh
        response = requests.get(search_url, timeout=10)
        soup = BeautifulSoup(response.text, 'html.parser')

        # Tìm các thẻ hình ảnh
        image_tags = soup.find_all('img', limit=4)
        image_urls = [img['src'] for img in image_tags if 'src' in img.attrs]

        # Tải tối đa 3 ảnh vào thư mục riêng của sản phẩm
        for i, img_url in enumerate(image_urls):
            filename = f"{clean_product_name}_{i + 1}.jpg"
            download_image(img_url, product_folder, filename)

    except Exception as e:
        print(f"Error processing {product_name}: {e}")

print("Hoàn tất tải ảnh!")
