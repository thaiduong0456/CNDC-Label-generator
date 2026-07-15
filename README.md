# Carton Label Generator

Web app đọc packing list Excel, gom các dòng SKU theo carton, tạo barcode Code 128 và xuất PDF vector nhiều trang (mỗi carton một trang 4×6 inch).

## Cài đặt

Yêu cầu Python 3.10 trở lên.

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

Sau đó mở địa chỉ Streamlit hiển thị trong Terminal (thường là `http://localhost:8501`).

## Định dạng Excel

- Dòng 1 và 2 là tiêu đề; dữ liệu bắt đầu từ dòng 3.
- Cột H có giá trị sẽ bắt đầu carton mới.
- Các dòng có cột H trống thuộc carton gần nhất phía trên.
- Các dòng tổng kết có cột A là `TOTAL`, `GRAND TOTAL`, `合计` hoặc `總計` được bỏ qua.
- Packaging ID lấy từ cột I; PO lấy từ B; SKU từ D; UPC từ E; Quantity từ G; tiêu đề từ Q.
- Nên định dạng các cột mã như UPC/Packaging code là **Text** trong Excel nếu cần giữ số 0 ở đầu. Excel không lưu lại số 0 đầu của một ô đã được nhập dưới dạng Number.

## Chức năng

- Chọn sheet trong workbook.
- Cảnh báo dữ liệu thiếu hoặc dòng sản phẩm xuất hiện trước carton đầu tiên.
- Giữ nguyên thứ tự carton và SKU.
- Cho sửa tiêu đề từng carton trước khi xuất.
- Tùy chọn gộp SKU trùng (theo SKU + UPC) và cộng Quantity.
- Xem trước PDF ngay trong trình duyệt.
- Barcode Code 128 và nội dung PDF được vẽ vector bằng ReportLab.
- Hiển thị `OR Code` từ cột Q và `Carton#` ở góc trên bên phải của mọi label.
- PDF khổ dọc 4×6 inch (288×432 points), sẵn sàng cho máy in nhiệt.

## Lưu ý in

Chọn khổ giấy 4×6 inch (101,6×152,4 mm), tỷ lệ 100% / Actual size, không chọn “Fit to page”. Nên quét thử một trang trước khi in số lượng lớn.

## Đưa app lên web bằng Streamlit Community Cloud

1. Tạo một repository mới trên GitHub.
2. Đưa các file `app.py`, `requirements.txt`, `README.md` và thư mục `.streamlit` lên repository.
3. Đăng nhập [Streamlit Community Cloud](https://share.streamlit.io/) bằng GitHub.
4. Chọn **Create app**, chọn repository vừa tạo, branch `main` và main file `app.py`.
5. Nhấn **Deploy**. Sau vài phút, Streamlit cung cấp một đường link dạng `https://ten-app.streamlit.app`.

Không đưa packing list hoặc PDF thật lên GitHub. File Excel chỉ được tải lên khi sử dụng app và được xử lý trong phiên chạy hiện tại.
