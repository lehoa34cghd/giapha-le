# Gia pha dong ho Le - Ban V8

## Chay project

```powershell
pip install -r requirements.txt
python app.py
```

Mo trinh duyet:

```text
http://127.0.0.1:5000
```

## Tai khoan mau

```text
admin / admin123
khach / xem123
```

## Diem moi V8

- Nen trang chinh dung anh `static/images/bg.jpg`.
- Khi giu chuot trai keo so do sang trai/phai, nen van giu nguyen, chi so do di chuyen.
- Da copy san anh nen mau vao `static/images/bg.jpg`.
- Neu muon doi anh nen, chi can thay file `static/images/bg.jpg` bang anh khac cung ten.
- Export anh PNG giu nen nay.

## Luu y

Neu may bao thieu Flask/openpyxl, chay:

```powershell
python -m pip install flask openpyxl
```


## Bản V9
- Trong chế độ Sửa đổi, mỗi ô người có nút **Sửa** nằm ngay cạnh nút **Xóa**.
- Chế độ Công khai vẫn ẩn toàn bộ dấu +, Sửa, Xóa.


## V10
- Trong chế độ Sửa đổi, mỗi ô người có nút Sửa nằm ngay cạnh nút Xóa.
- Chế độ Công khai vẫn ẩn toàn bộ nút Sửa/Xóa/+.

## Ghi chú bản V11
- Database luôn được lưu trong chính thư mục project: `giapha_le.db`.
- Khi tắt Python rồi bật lại, dữ liệu vẫn giữ nguyên nếu bạn không xóa file `giapha_le.db`.
- Nếu trước đây bạn chạy từ thư mục khác và có dữ liệu cũ, hãy tìm file `giapha_le.db` cũ rồi copy vào thư mục project này.
- Tài khoản xem: `khach / xem123`.
