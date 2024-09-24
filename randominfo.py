import subprocess
import sys
import random
import pandas as pd
import time
from unidecode import unidecode
from faker import Faker

# Khởi tạo Faker để tạo địa chỉ email ngẫu nhiên
fake = Faker()

def install_package(package):
    """Cài đặt thư viện nếu chưa được cài đặt."""
    subprocess.check_call([sys.executable, "-m", "pip", "install", package])

def check_and_install_packages():
    """Kiểm tra và cài đặt các thư viện cần thiết."""
    packages = ['pandas', 'openpyxl', 'unidecode']
    for package in packages:
        try:
            __import__(package)
            print(f"{package} đã được cài đặt.")
        except ImportError:
            print(f"{package} chưa được cài đặt. Đang cài đặt...")
            install_package(package)
            print(f"{package} đã được cài đặt.")

# Kiểm tra và cài đặt các thư viện cần thiết
check_and_install_packages()

# Load dữ liệu về tỉnh, huyện, xã từ file Excel
file_path = 'data.xlsx'
df = pd.read_excel(file_path, sheet_name='Sheet1')

# Danh sách họ phổ biến
ho = [
    "Nguyễn", "Trần", "Lê", "Phạm", "Hoàng", "Huỳnh", "Phan", "Vũ", "Võ", "Đặng",
    "Bùi", "Đỗ", "Hồ", "Ngô", "Dương", "Lý", "Tô", "Đinh", "Trương", "Chu",
    "Tôn", "Vương", "Lâm", "Thái", "Diệp", "Hà", "Kiều", "Lương", "Mã", "Quách",
    "Đoàn", "Thạch", "Phùng", "Tạ", "Cao", "Văn", "Hứa", "Trịnh", "Triệu", "Đinh"
]

# Danh sách tên đệm phổ biến cho nam
ten_dem_nam = [
    "Văn", "Hữu", "Đức", "Minh", "Quang", "Hải", "Tuấn", "Anh", "Trọng", "Xuân",
    "Tiến", "Phúc", "Bảo", "Khánh", "Thành", "Chí", "Nhật", "Thế", "Công", "Việt",
    "Thịnh", "Phước", "Trí", "Đăng", "Ngọc", "Hoàng", "Gia", "Thanh", "Kiến", "Đình",
    "Quốc", "Thái", "Tấn", "Duy", "Nhân", "Kỳ", "Tài", "Sỹ", "Khắc", "Thắng"
]

# Danh sách tên đệm phổ biến cho nữ
ten_dem_nu = [
    "Thị", "Ngọc", "Diệu", "Thanh", "Bích", "Lan", "Mai", "Thu", "Ánh", "Hoài",
    "Kim", "Trúc", "Hồng", "Như", "Thùy", "Phương", "Tuyết", "Hương", "Yến", "My",
    "Oanh", "Lệ", "Tâm", "Quỳnh", "Vy", "Tiểu", "Như", "Tường", "Khuê", "Thảo",
    "Liên", "Giang", "Linh", "Châu", "Xuân", "Khánh", "Minh", "Nhã", "Phúc", "Thi"
]

# Danh sách tên riêng phổ biến cho nam
ten_nam = [
    "Dương", "Tùng", "Phong", "Nam", "Linh", "Huy", "Sơn", "Đạt", "Thắng", "Khoa",
    "Bình", "Hiếu", "Vinh", "Kiên", "Tài", "Khôi", "Phát", "Trí", "Vũ", "Long",
    "Hưng", "Hào", "Khải", "Toàn", "Nhật", "Nguyên", "Phước", "Tuấn", "Phúc", "Đông",
    "Tiến", "Khang", "Thịnh", "Đăng", "Dũng", "Quân", "Thái", "Lợi", "Bảo", "Thiện"
]

# Danh sách tên riêng phổ biến cho nữ
ten_nu = [
    "Tuyết", "Mi", "Trang", "Lan", "Hương", "Anh", "Yến", "Hà", "Nhi", "Vân",
    "Ly", "Thảo", "Phượng", "Hoa", "Linh", "Vy", "Dung", "Mai", "Quỳnh", "My",
    "Giang", "Ngân", "Khánh", "Diễm", "Hạnh", "Nhung", "Tâm", "Lệ", "Thúy", "Oanh",
    "Cúc", "Hiền", "Châu", "Tuyền", "Loan", "Kim", "Trâm", "Thy", "Xuân", "Ngọc"
]

# Hàm tạo tên ngẫu nhiên
def generate_random_name():
    gioi_tinh = random.choice(["Nam", "Nữ"])
    ho_chon = random.choice(ho)
    if gioi_tinh == "Nam":
        ten_dem = random.choice(ten_dem_nam)
        ten = random.choice(ten_nam)
    else:
        ten_dem = random.choice(ten_dem_nu)
        ten = random.choice(ten_nu)
    return f"{ho_chon} {ten_dem} {ten}", gioi_tinh

# Hàm tạo ngày tháng năm sinh ngẫu nhiên
def generate_random_dob():
    year = random.randint(1970, 2005)
    month = random.randint(1, 12)
    day = random.randint(1, 28)  # Để đảm bảo ngày hợp lệ cho mọi tháng
    return f"{day:02d}/{month:02d}/{year}"

# Hàm tạo Gmail ngẫu nhiên, loại bỏ dấu và ký tự không hợp lệ
def generate_random_email(name):
    # Chuyển tên thành chữ không dấu và loại bỏ ký tự không hợp lệ
    name = unidecode(name)
    name = ''.join(c for c in name if c.isalnum())
    if len(name) > 15:  # Đảm bảo tên người dùng không quá dài
        name = name[:15]
    username = name + str(random.randint(1, 99))  # Thêm số ngẫu nhiên để đảm bảo tính duy nhất
    email = f"{username}@gmail.com"
    # Chỉ viết hoa chữ cái đầu tiên của tên người dùng
    return email.capitalize()

# Hàm tạo số điện thoại, tên, và địa chỉ gồm xã, huyện, tỉnh
def generate_phone_numbers_with_names(num_records):
    data = []

    for _ in range(num_records):
        # Random đầu số và số điện thoại
        head = random.choice(["03", "09"])
        phone_number = head + "".join([str(random.randint(0, 9)) for _ in range(8)])
        
        # Random tên và giới tính
        name, gioi_tinh = generate_random_name()

        # Random địa chỉ từ dữ liệu: Tỉnh, Huyện, Xã
        row = df.sample(1).iloc[0]
        tinh_chon = row["Tỉnh Thành Phố"]
        huyen_chon = row["Quận Huyện"]
        xa_chon = row["Phường Xã"]

        # Random ngày tháng năm sinh và Gmail
        dob = generate_random_dob()
        email = generate_random_email(name)

        # Thêm vào danh sách dữ liệu
        data.append({
            "Tên": name,
            "Giới tính": gioi_tinh,
            "Địa chỉ": f"{xa_chon}, {huyen_chon}, {tinh_chon}",
            "Số điện thoại": phone_number,
            "Ngày sinh": dob,
            "Email": email
        })

    return data

# Nhập số lượng bản ghi từ người dùng
try:
    num_records = int(input("Nhập số lượng bản ghi cần tạo: "))
    if num_records <= 0:
        raise ValueError("Số lượng bản ghi phải lớn hơn 0.")
except ValueError as e:
    print(f"Lỗi: {e}. Sử dụng số lượng mặc định 100.")
    num_records = 100

# Gọi hàm và lưu kết quả vào file Excel
results = generate_phone_numbers_with_names(num_records)

# Chuyển đổi kết quả thành DataFrame
df_results = pd.DataFrame(results)

# Xuất DataFrame ra file Excel
output_file_path = 'output.xlsx'
df_results.to_excel(output_file_path, index=False)

print(f"Dữ liệu đã được xuất ra file {output_file_path}.")
