import os
import pandas as pd
import requests
from bs4 import BeautifulSoup
import unicodedata
import re

# Đường dẫn đến file Excel (cùng thư mục với file Python)
input_file = "shoppe.xlsx"

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

# Hàm lấy các sản phẩm từ URL tìm kiếm
def get_product_links(search_url):
    try:
        response = requests.get(search_url, timeout=10)
        soup = BeautifulSoup(response.text, 'html.parser')
        
        # Tìm tất cả các link sản phẩm trong kết quả tìm kiếm
        product_links = []
        product_tags = soup.find_all('a', {'class': 'link'}, href=True)  # Lọc các link sản phẩm
        
        for tag in product_tags:
            product_links.append(f"https://shopee.vn{tag['href']}")  # Xây dựng URL đầy đủ
        
        return product_links
    except Exception as e:
        print(f"Error getting product links from {search_url}: {e}")
        return []

# Duyệt qua từng sản phẩm
for index, row in data.iterrows():
    product_name = row['Tên hàng']
    search_url = row['URL tìm kiếm hình ảnh']

    # Loại bỏ dấu và thay dấu cách trong tên sản phẩm
    clean_product_name = remove_vietnamese_accents(product_name)

    # Tạo thư mục riêng cho mỗi sản phẩm (dùng tên sản phẩm đã loại bỏ dấu làm tên thư mục)
    product_folder = f"downloaded_images/{clean_product_name}"
    os.makedirs(product_folder, exist_ok=True)

    # Lấy danh sách các sản phẩm từ URL tìm kiếm
    product_links = get_product_links(search_url)

    for product_url in product_links:
        try:
            # Truy cập vào trang chi tiết sản phẩm
            response = requests.get(product_url, timeout=10)
            soup = BeautifulSoup(response.text, 'html.parser')
            
            # Tìm các thẻ hình ảnh trong trang sản phẩm
            image_tags = soup.find_all('img', limit=4)
            image_urls = [img['src'] for img in image_tags if 'src' in img.attrs]
            
            # Tải tối đa 3 ảnh vào thư mục riêng của sản phẩm
            for i, img_url in enumerate(image_urls):
                filename = f"{clean_product_name}_{i + 1}.jpg"
                download_image(img_url, product_folder, filename)

        except Exception as e:
            print(f"Error processing {product_url}: {e}")

print("Hoàn tất tải ảnh!")
