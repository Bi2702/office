# Add-in Word: Thể thức văn bản Đảng

Add-in Office.js chạy trong Word desktop, với một bảng công cụ trên thẻ Home. Quy tắc định dạng được xây dựng từ Quy định 399-QĐ/TW và Hướng dẫn 05-HD/VPTW mà người dùng cung cấp.

## Chức năng hiện có

- Đặt trang A4, chiều dọc; lề trên/dưới 20 mm, trái 30 mm, phải 15 mm trên các phiên bản Word hỗ trợ WordApiDesktop 1.3.
- Chuẩn hóa các đoạn nội dung ngoài bảng về Times New Roman 14 pt, màu đen, căn đều, thụt đầu dòng 10 mm, cách đoạn 6 pt và giãn dòng 20 pt.
- Nhận diện một số dòng thể thức phổ biến và đề mục “Phần”, “Chương”, “Mục”, “Điều” để căn chỉnh riêng.
- Định dạng dòng đang chọn thành tên loại văn bản hoặc trích yếu.
- Rà soát sự hiện diện của một số thành phần có thể đọc được từ nội dung tài liệu.

Rà soát chỉ tìm dấu hiệu văn bản, không xác nhận thẩm quyền ký, chữ ký/dấu, nội dung nghiệp vụ hay tính hợp lệ pháp lý. Công cụ không thay nội dung; thao tác chuẩn hóa thay đổi định dạng các đoạn ngoài bảng. Hãy lưu một bản sao trước khi áp dụng.

## Chạy trong Word trên Windows

1. Cài Node.js phiên bản LTS.
2. Mở thư mục này trong Terminal và chạy `npm install`.
3. Chạy `npm start`; chấp nhận cài chứng chỉ phát triển cục bộ nếu Windows yêu cầu.
4. Add-in được nạp vào Word desktop. Mở thẻ **Home** → nhóm **Văn bản Đảng** → **Chuẩn hóa thể thức**.
5. Dừng phiên bằng `Ctrl+C`; nếu add-in còn hiện trong Word, chạy `npm run stop`.

Word cần hỗ trợ WordApi 1.3. Tính năng đặt trang tự động cần Word desktop hỗ trợ WordApiDesktop 1.3; trên bản Word cũ hơn, định dạng đoạn vẫn chạy và add-in báo rõ rằng lề/trang chưa được đổi.

## Triển khai dùng chung

Đây là bản chạy phát triển trên `localhost`. Để phân phối cho nhiều người, cần lưu trữ thư mục web qua HTTPS ổn định, cập nhật URL trong `manifest.xml`, rồi triển khai manifest qua Microsoft 365 hoặc danh mục add-in nội bộ. Không dùng chứng chỉ phát triển cục bộ cho bản phát hành.

## Quy tắc tham chiếu

- Văn bản chính: Times New Roman, Unicode, màu đen; giấy A4; lề 20/20/30/15 mm.
- Nội dung: căn đều; dòng đầu lùi khoảng 10 mm; cách đoạn tối thiểu 6 pt; giãn dòng 18–22 pt.
- Cỡ chữ của tiêu đề, tên loại, trích yếu và các thành phần thể thức được áp dụng theo bảng tại Phụ lục 2 của Hướng dẫn 05.

Trong bản này, giãn dòng nội dung được đặt ở 20 pt (nằm trong khoảng 18–22 pt); dòng tên loại được đặt 15 pt đậm, còn trích yếu có nút riêng đặt 14 pt đậm. Có thể chỉnh trực tiếp trong `public/taskpane.js` nếu cơ quan yêu cầu lựa chọn khác trong khoảng cho phép.
