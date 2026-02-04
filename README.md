# OfficeUpdaterProject
<!-- items that need to be updated release to release -->
[ptUserAddin]: https://github.com/SanbiVN/OfficeUpdaterProject/releases/download/v1.0/OfficeUpdaterProject.xlam

|  Thông tin   | Tải xuống | Lượt tải |
|--------------|-----------|----------|
| OfficeUpdaterProject | [OfficeUpdaterProject.xlam][ptUserAddin] | [![Lượt tải](https://img.shields.io/github/downloads/SanbiVN/OfficeUpdaterProject/total.svg)](https://github.com/SanbiVN/OfficeUpdaterProject/releases/download/v1.0/OfficeUpdaterProject.xlam)  |

(Pass VBA là 1)\
Sao chép toàn bộ Module và Class Module vào dự án, đổi thông tin tại Module A_Center.

#### Tự động cập nhật chính dự án Microsoft Office chứa mã VBA
Tại sao cần có phương pháp này để VBA tự động cập nhật cho chính ứng dụng chứa nó?
1. VBA không thể ghi đè tệp Excel đang mở nếu không xử lý mã đúng phương pháp.​
2. VBA có thể tự mở chính nó ở phiên ReadOnly và cập nhật ngược lại cho tệp chính tệp đang mở. Tuy nhiên không tách VBA ra khỏi tiến trình chính thì sẽ tự kết thúc tiến trình theo tệp chính nếu đóng tệp chính để cập nhật.​
3. Có thể khởi tạo một Excel ở tiến trình mới tuy nhiên sẽ có cửa sổ trên màn hình, nếu ẩn bởi VBA vẫn tạo ra một nhấp nháy rất khó chịu cho người dùng.​
4. Nếu không bỏ unlock tệp tải về thì không thể cập nhật để chạy được.​

#### Giải pháp tận dụng Win32 API để thực hiện các bước là:
1. Tiến trình chính chạy http để tìm kiếm phiên bản mới.​
2. Nếu có bản mới thì khởi chạy một ứng dụng Excel mới với CreateObject("Excel.Application") và mở tệp chính với ReadOnly.​
3. Đăng ký đối tượng Workbook đó vào ROT để tiến trình tách rời khỏi ứng dụng Excel chính, để không bị đóng theo tiến trình chính.​
4. Ở tiến trình mới, chạy mã quét lấy phiên bản mới tải về (giải nén, hoặc truy xuất binary).​
5. Dùng API Win32 để bỏ unblock tệp đã giải nén để chạy được khi cài đặt vào Excel.​
6. Ở tiến trình mới truy cập tất cả các tiến trình Excel để thoát tệp chính, gỡ cài đặt, ghi đè hoặc tạo bản lưu trữ.​
7. Mở tệp phiên bản mới ở tất cả các tiến trình Excel. Và thông báo để hoàn thành.​
​
​
### Hướng dẫn​
Trong tệp add-in OfficeUpdaterProject.xlam bên dưới trong VBA có module A_Center chứa thông tin tải bản cập nhật, và khởi chạy lớp cập nhật.
Trong tệp bao gồm các lớp quan trọng để xử lý với tệp tải về như: hash file, ghi, đọc, giải nén, msgbox, inputBox hỗ trợ tiếng Việt, Dictionary, clsUnicode hỗ trợ nhập tiếng Việt, JsonConverter để phân tích Json.

Sau khi tải về mở lên, các nút thao tác sẽ nằm trên Ribbon có tên là Thử nghiệm
Nhấn nút Kiểm tra cập nhật, sẽ có thông báo Ballon tại hệ thống. Nhấn Cập nhật ngay, dự án sẽ được cập nhật, và có thông báo nếu có bản mới.

<img width="333" height="132" alt="1770199388913" src="https://github.com/user-attachments/assets/372aebbf-8822-482e-8096-19315aa854e9" />

#### Các khóa trong json để lấy thông tin cần thiết để thực hiện kiểm tra và cập nhật:​
 - name: tên dự án không có khoảng trắng
 - filename: tên tệp
 - version: phiên bản để đối chiếu
 - installer_url: đường dẫn tải thông tin hoặc tệp cần cập nhật
 - url_type: kiểu đường dẫn tải thông tin là gì, ví dụ: github.api.releases.lastest có nghĩa là tải từ api github tại phiên bản cuối cùng.
 - file_compress: tệp có được nén hay không
 - path_compress: đường dẫn thư mục trong tệp nén dẫn đến vị trí chứa tệp cần cập nhật
 - check_version_in: kiểm tra ở đâu trong dữ liệu được tải về từ github.api.releases.lastest, gồm: tag, title, asset (hoặc tự viết mã thêm)
 - file_type: kiểu tệp cần cập nhật là gì, gồm: Excel, Word, Access, PowerPoint, file, ... (hoặc tự viết mã thêm)
 - path_location: đường dẫn gốc thay cho đường dẫn tuyệt đối (nếu cần), gồm:
> - directory: thư mục cùng với thư mục chứa ứng dụng
> - temp: thư mục temp
> - local: thư mục local
> - programData: thư mục programData
> - UserProfiles: thư mục User
> - Windows: thư mục Windows
 - path_target: đường dẫn dẫn đến vị trí ghi tệp cần cập nhật (nếu có), nếu đường dẫn không có tiền tố thì lấy path_location nối vào.
 - path_registry: đường dẫn registry, dùng để ghi và xuất thông tin (nếu cần)
 - enabled: on/off, xét xem có cần cập nhật hay không nếu để on.
 - modules: các dự án con (nếu có), modules chứa các thông tin như đã liệt kê cho các dự án con cần cập nhật.

#### Các khóa phụ để dự án được chi tiết hơn:​
 - title: tiêu đề dự án
 - description: thông tin dự án
 - group_id: id nhóm
 - group_name: tên nhóm
 - root: tên gốc chứa dự án
 - hash: mã hash của tệp nếu có để đối chiếu
 - website: trang web chính
 - helps_file: tệp trợ giúp
 - helps_url: đường dẫn trợ giúp
 - author: người tạo
 - status: trạng thái
