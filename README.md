# Test-cursor

## Ứng dụng Quản lý công việc (To-Do List) CLI nâng cao

Ứng dụng chạy trên Terminal nhưng có giao diện menu dạng “app”, hỗ trợ lưu dữ liệu tự động ra file JSON.

### Chạy chương trình

```bash
python3 todo_cli.py
```

### Chức năng chính

- Thêm công việc mới (kèm mức ưu tiên).
- Sửa nội dung và ưu tiên của công việc.
- Đánh dấu công việc hoàn thành.
- Xóa một công việc theo STT.
- Xem toàn bộ danh sách với trạng thái rõ ràng.
- Tìm kiếm công việc theo từ khóa.
- Lọc danh sách theo trạng thái (đã/chưa hoàn thành).
- Thống kê nhanh tổng số việc.
- Xóa toàn bộ việc đã hoàn thành.

### Lưu dữ liệu

- Dữ liệu được lưu vào file: `todo_data.json`.
- Khi chạy lại app, danh sách công việc vẫn còn.
