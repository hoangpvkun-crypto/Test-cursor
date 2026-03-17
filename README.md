# Test-cursor

## To-Do List Web App (1 file `index.html`)

Phiên bản nâng cấp toàn diện với giao diện hiện đại, Dark/Light mode, ưu tiên công việc, tag danh mục, progress bar và hiệu ứng animation.

## Cách chạy

- Mở trực tiếp file `index.html` bằng trình duyệt.
- Hoặc chạy server tĩnh:

```bash
python3 -m http.server 8000
```

Sau đó truy cập: `http://localhost:8000`

## Tính năng chính

- Thêm công việc với **Mức ưu tiên**: Cao / Trung bình / Thấp.
- Gắn **Tag** cho từng công việc.
- Sắp xếp tự động theo ưu tiên (Cao luôn lên đầu).
- Đánh dấu hoàn thành, xóa công việc.
- Bộ lọc: Tất cả / Chưa xong / Đã xong.
- Thanh tiến trình hiển thị % hoàn thành.
- Bắn pháo giấy (confetti) khi hoàn thành 100% danh sách (có ít nhất 1 việc).
- Chuyển đổi **Light/Dark Mode**.
- Lưu toàn bộ trạng thái vào `localStorage` (tasks + theme).

## Ứng dụng Quản lý công việc (To-Do List) CLI

### Chạy chương trình

```bash
python3 todo_cli.py
```

### Tính năng

- Thêm công việc mới.
- Xóa công việc đã hoàn thành.
- In danh sách việc cần làm.
- Menu lặp cho đến khi người dùng chọn thoát.
 main
