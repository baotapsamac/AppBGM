# Công cụ soạn bài giảng (Hướng A)

Công cụ chạy trên máy cá nhân, không cần internet, không gửi tài liệu lên bất kỳ máy chủ nào bên ngoài.

Có hai cách dùng: **chạy trực tiếp bằng Python** (Cách 1), hoặc **đóng gói một lần thành file .exe** rồi từ đó chỉ cần bấm đúp (Cách 2). Cách 2 tiện hơn nếu thầy muốn dùng lâu dài hoặc chia sẻ cho đồng nghiệp không rành cài đặt.

---

## Cách 2 — Đóng gói thành file .exe, tự động cài mọi thứ (khuyên dùng)

### Đóng gói (chỉ làm một lần)

1. Chuột phải vào file **`DONG_GOI_EXE.bat`** → chọn **"Run as administrator"**
   (nên chạy quyền Administrator để việc tự cài Python/LibreOffice không bị chặn quyền; nếu bấm đúp thường không được thì thử lại bằng cách này).
2. Kịch bản sẽ tự động, theo thứ tự:
   - Kiểm tra và **tự cài Python** nếu máy chưa có
   - Kiểm tra và **tự cài LibreOffice** nếu máy chưa có (file khoảng 300MB, có thể mất vài phút)
   - Cài các thư viện Python cần thiết
   - Đóng gói thành file `.exe`
3. Đợi đến khi thấy dòng **"XONG!"**. Lần đầu có thể mất 5–15 phút tùy tốc độ mạng, chủ yếu do tải LibreOffice.
4. Kết quả nằm trong thư mục `dist` vừa được tạo ra: file **`CongCuSoanBaiGiang.exe`** và thư mục `sample`.

**Nếu máy vừa cài xong Python lần đầu mà kịch bản báo lỗi "chưa nhận được thay đổi":** đóng cửa sổ, bấm đúp lại file `.bat` một lần nữa là được — đây là giới hạn của Windows (cửa sổ dòng lệnh đang mở không tự thấy phần mềm vừa cài xong), không phải lỗi của công cụ.

### Từ lần sau trở đi

Chỉ cần bấm đúp vào **`dist\CongCuSoanBaiGiang.exe`** — không cần mở Command Prompt, không cần chạy lại file `.bat`. Trình duyệt tự mở.

Có thể copy cả thư mục `dist` sang máy khác (USB, máy đồng nghiệp) để dùng. Máy đó **không cần cài Python**, nhưng **vẫn cần có LibreOffice** (phần mềm ngoài, không nhét được vào trong file `.exe`) — nếu máy đó chưa có, chạy lại file `DONG_GOI_EXE.bat` một lần ở máy đó cũng được, nó sẽ tự cài LibreOffice rồi đóng gói lại.

Muốn tắt công cụ: đóng cửa sổ đen (cửa sổ dòng lệnh) hiện ra cùng lúc với trình duyệt.

---

## Cách 1 — Chạy trực tiếp bằng Python (không tạo file .exe)

Dùng cách này nếu chỉ muốn thử nhanh, không cần file .exe để mang đi.

### Chuẩn bị (một lần)

1. Cài **Python 3.9 trở lên**: https://www.python.org/downloads/
   (khi cài, nhớ tick ô "Add python.exe to PATH")
2. Cài **LibreOffice**: https://www.libreoffice.org/download/
3. Mở Command Prompt tại thư mục này, chạy:
   ```
   pip install -r requirements.txt
   ```

### Chạy

```
python app.py
```

Trình duyệt sẽ tự mở tới `http://127.0.0.1:5000`. Muốn tắt: đóng cửa sổ dòng lệnh hoặc bấm `Ctrl + C`.

---

## Cách dùng (áp dụng cho cả hai cách trên)

1. Mẫu Word đã có sẵn trong công cụ (không cần chọn) — chỉ cần tải lên file Excel dữ liệu (cột A là tên trường, cột B trở đi mỗi cột một bài giảng).
2. Bấm **Nạp dữ liệu**. Nếu Excel thiếu hoặc thừa trường so với mẫu Word, công cụ sẽ báo ngay, đồng thời tự dựng sẵn bản xem trước.
3. Chọn tab bài giảng (nếu Excel có nhiều cột — có ô tìm kiếm khi danh sách dài), sửa trực tiếp trên bảng bên trái. Tick chọn bài giảng nào cần đưa vào khi xuất hàng loạt.
4. Bấm **Cập nhật dữ liệu** để làm mới bản xem trước theo nội dung vừa sửa.
5. Ưng ý thì chọn định dạng Word/PDF rồi bấm **Xuất bài này** để tải về đúng một file, hoặc **Xuất tất cả** để tải về file `.zip` gồm các bài giảng đã tick chọn.

## Thư mục mẫu

`sample/MAU_BaiGiang.docx` là mẫu Word cố định mà công cụ dùng — sửa trực tiếp file này nếu cần đổi mẫu chung. `sample/DULIEU_BaiGiang.xlsx` là ví dụ cách tổ chức file Excel dữ liệu.

## Giới hạn hiện tại

- Chỉ chỉnh nội dung text, không chỉnh được bố cục/định dạng ngay trên bản xem trước (định dạng lấy nguyên từ mẫu Word).
- Công cụ dùng cho một người thao tác tại một thời điểm trên một máy; không phải dùng để nhiều người cùng sửa một lúc qua mạng.
- Nếu đổi thứ tự/tên cột trong Excel, tên bài giảng lấy theo dòng tiêu đề (dòng 1) của từng cột.
- Mẫu Word dùng chung cho mọi bài giảng — muốn đổi mẫu phải sửa trực tiếp `sample/MAU_BaiGiang.docx` rồi đóng gói lại (không có nút chọn mẫu khác trên giao diện).
- **LibreOffice vẫn là điều kiện bắt buộc** ở cả hai cách dùng, kể cả khi đã có file .exe — vì đây là chương trình ngoài phụ trách việc dựng ảnh xem trước, không gộp được vào trong file .exe.

