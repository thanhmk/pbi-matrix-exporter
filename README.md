# 📊 PBI Matrix Exporter

Một công cụ giúp giải quyết "nỗi đau" khi xuất dữ liệu từ các bảng Matrix phức tạp trong Power BI Desktop. 

Thay vì phải xuất dữ liệu thô (flat data) và dùng Pivot Table trong Excel để dựng lại một cách thủ công, công cụ này cho phép bạn xuất trực tiếp dữ liệu từ Power BI ra file Excel mà vẫn **giữ nguyên hoàn toàn định dạng phân cấp (hierarchy)** của hàng và cột, nhìn giống như bảng trên visualize.

## ✨ Tính năng nổi bật
* **Hỗ trợ phân cấp linh hoạt:** Xử lý tốt phân cấp nhiều tầng ở cả Hàng (Rows) và Cột (Columns - ví dụ: phân cấp theo Thời gian).
* **Tuỳ chọn lặp lại tiêu đề:** Cho phép chọn giữa việc gom ô (trống nhãn ở các dòng con) giống giao diện Power BI, hoặc lặp lại tiêu đề cha trên từng dòng chi tiết để dễ xử lý dữ liệu về sau.
* **Đảo chiều chỉ tiêu (Transpose):** Hỗ trợ đưa các chỉ tiêu (Measures) làm hàng và các chiều (Dimensions) làm cột.
* **Cơ chế Fallback an toàn:** Nếu cấu trúc DAX quá phức tạp hoặc bị lỗi, phần mềm sẽ tự động chuyển sang chế độ xuất dữ liệu phẳng (flat dump) để đảm bảo không bao giờ bị crash và bạn luôn có dữ liệu để sử dụng.
* **Giao diện Web trực quan:** Thao tác hoàn toàn trên giao diện Streamlit thân thiện.

## 🚀 Hướng dẫn Cài đặt
Dự án đã được tối ưu hóa để quá trình cài đặt diễn ra tự động. Bạn không cần phải thiết lập môi trường phức tạp.

1. Tải toàn bộ mã nguồn dự án này về máy tính của bạn (Code -> Download ZIP) và giải nén.
2. Đảm bảo máy tính của bạn đã cài đặt sẵn **Python**.
3. Click đúp vào file `run.bat`.
4. Hệ thống sẽ tự động kiểm tra, cài đặt các thư viện cần thiết (có trong `requirements.txt`) và mở ứng dụng ngay trên trình duyệt web của bạn.

## 💡 Hướng dẫn Sử dụng
Để sử dụng công cụ, bạn cần lấy câu lệnh DAX của bảng Matrix từ Power BI Desktop đang mở.

**Bước 1: Lấy DAX Query từ Power BI**
1. Mở file `.pbix` của bạn bằng phần mềm Power BI Desktop.
2. Trên thanh công cụ, chọn thẻ **Optimize** -> Chọn **Performance Analyzer**.
3. Bấm **Start recording**, sau đó bấm **Refresh visuals** (hoặc tương tác vào bảng Matrix bạn muốn xuất).
4. Trong bảng Performance Analyzer, mở rộng dòng chứa visual Matrix của bạn.
5. Bấm **Copy query**.

**Bước 2: Xuất dữ liệu**
1. Mở ứng dụng PBI Matrix Exporter trên trình duyệt.
2. Dán đoạn DAX query vừa copy vào ô **"Dán DAX query"**.
3. Lựa chọn cấu hình bảng:
   - *Hướng bảng:* Chiều làm hàng (mặc định) hoặc Chỉ tiêu làm hàng (nếu bạn muốn đảo bảng).
   - *Tùy chọn:* Tích vào ô "Lặp lại tiêu đề dòng" nếu bạn muốn điền đầy đủ nhãn phân cấp ở mọi dòng.
4. Bấm **🚀 Tạo & Export**.
5. Nhấn nút **⬇️ Tải file Excel** để lưu file với cấu trúc hoàn chỉnh về máy.

## 🔗 Chia sẻ và Đóng góp
Dự án được phát triển nhằm mục đích tự động hóa các thao tác thủ công và hỗ trợ cộng đồng làm phân tích dữ liệu làm việc hiệu quả hơn. Bạn có thể tìm hiểu thêm các bài viết chuyên sâu về kỹ thuật xử lý dữ liệu, DAX, và định hướng phát triển trong ngành Data tại [phantichdulieu.vn](https://phantichdulieu.vn).

Nếu bạn thấy công cụ này hữu ích, hãy cho dự án một ⭐️ và thoải mái đóng góp (Pull Request) để công cụ ngày càng hoàn thiện hơn!
