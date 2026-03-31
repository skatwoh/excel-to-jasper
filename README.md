# Excel to Jasper (JRXML) Converter Web App

Ứng dụng web được xây dựng trên Spring Boot giúp chuyển đổi file Excel (.xlsx) sang định dạng thiết kế JasperReports (.jrxml).

## Tính năng
- Tải lên file Excel và phân tích danh sách các sheet.
- Cấu hình linh hoạt: chọn sheet, dòng bắt đầu tiêu đề (header), và số lượng dòng tiêu đề.
- Tạo cấu trúc báo cáo JRXML tự động dựa trên định dạng của file Excel (độ rộng cột, gộp ô, kiểu chữ cơ bản).
- Giao diện web thân thiện, dễ sử dụng với Bootstrap 5.

## Yêu cầu hệ thống
- **Java 8** trở lên (Khuyên dùng Java 11 hoặc 17).
- **Maven 3.6** trở lên.

## Cách chạy ứng dụng cục bộ
1. Clone repository này về máy.
2. Mở terminal tại thư mục gốc của dự án.
3. Chạy lệnh:
   ```bash
   mvn spring-boot:run
   ```
4. Truy cập trình duyệt tại địa chỉ: `http://localhost:8080`

## Cách đẩy mã nguồn lên GitHub (Deploy to GitHub)
Để lưu trữ và chia sẻ mã nguồn trên GitHub, bạn thực hiện các bước sau:

1. **Tạo Repository mới trên GitHub**: Truy cập [github.com/new](https://github.com/new) và tạo một repo mới (ví dụ: `excel-to-jasper-web`).
2. **Kết nối Local Repo với GitHub**:
   Nếu bạn đã có git khởi tạo:
   ```bash
   git remote add origin https://github.com/USERNAME/REPO_NAME.git
   git branch -M main
   git push -u origin main
   ```
   *(Thay `USERNAME` và `REPO_NAME` bằng thông tin của bạn)*

## Cách triển khai ứng dụng lên Internet (Cloud Deployment)
Vì đây là ứng dụng Spring Boot (Java), bạn có thể triển khai lên các nền tảng sau:

1. **Render.com** (Khuyên dùng):
   - Tạo Web Service mới, kết nối với GitHub repo.
   - Build Command: `mvn clean install -DskipTests`
   - Start Command: `java -jar target/excel-to-jasper-1.0-SNAPSHOT.jar`
   - Chọn môi trường là Docker hoặc Java.

2. **Railway.app**:
   - Chỉ cần kết nối GitHub repo, Railway sẽ tự động nhận diện Spring Boot và deploy.

3. **Heroku**:
   - Sử dụng Heroku CLI hoặc kết nối GitHub để deploy.

---
*Phát triển bởi Jules - Hỗ trợ chuyển đổi Excel sang JasperReports nhanh chóng.*
