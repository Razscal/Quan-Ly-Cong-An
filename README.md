# Ứng dụng Quản Lý Nhiệm Vụ - Công An

Ứng dụng quản lý nhiệm vụ cho lực lượng công an, cho phép tạo và theo dõi nhiệm vụ, trộn file Excel và quản lý danh hiệu khen thưởng.

## Tính năng chính

1. **Tạo nhiệm vụ**: Tạo nhiệm vụ mới với file Excel tùy chỉnh các cột và vị trí lưu file.
2. **Trộn file**: Trộn nhiều file Excel sau khi đồng đội cập nhật và import dữ liệu vào hệ thống.
3. **Danh sách nhiệm vụ**: Xem danh sách nhiệm vụ theo năm, đơn vị và tên, hiển thị thông tin người và danh hiệu đã được khen thưởng.

## Cài đặt

1. Cài đặt Python (phiên bản 3.8 trở lên)
2. Cài đặt các thư viện cần thiết:

```
pip install -r requirements.txt
```

## Chạy ứng dụng

```
python main.py
```

## Cấu trúc dự án

```
quan_ly_nhiem_vu/
├── database/           # Module quản lý cơ sở dữ liệu
│   ├── __init__.py
│   └── db_manager.py
├── models/             # Các model dữ liệu
│   ├── __init__.py
│   ├── task.py
│   ├── person.py
│   └── award.py
├── ui/                 # Giao diện người dùng
│   ├── __init__.py
│   ├── main_window.py
│   ├── task_creation.py
│   ├── task_merge.py
│   └── task_list.py
├── utils/              # Tiện ích
│   ├── __init__.py
│   └── excel_manager.py
├── main.py             # Điểm khởi đầu ứng dụng
└── requirements.txt    # Các thư viện cần thiết
```

## Cơ sở dữ liệu

Ứng dụng sử dụng SQLite làm cơ sở dữ liệu local, dễ dàng sao lưu và không cần kết nối mạng.

## Màu sắc

Ứng dụng sử dụng màu chủ đạo xanh lá (#4CAF50) và trắng (#FFFFFF) theo yêu cầu.
