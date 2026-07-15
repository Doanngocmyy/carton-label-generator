# Carton Label Generator (Web)

Web app cho `carton_label_generator.py`: upload file Packing List Excel, web tự tạo file PDF nhãn thùng 4x6 inch + file CSV audit để tải về.

## File trong thư mục này

- `app.py` — giao diện web (Streamlit)
- `carton_label_generator.py` — script gốc, không đổi logic
- `requirements.txt` — danh sách thư viện cần cài
- `PUBLISH_GUIDE.md` — hướng dẫn đưa web này lên mạng (từng bước)

## Chạy thử trên máy mình trước khi publish (không bắt buộc)

```
pip install -r requirements.txt
streamlit run app.py
```

Mở trình duyệt tại `http://localhost:8501`.

## File Excel input có cần chỉnh sửa trước không?

Không cần. Script tự động:
- Dò tìm dòng tiêu đề thật (dòng chứa đủ các cột PO No., Packaging code, SKU#, BarCode/UPC, Quantity) — nên phần header thông tin khách hàng/shipper phía trên bị bỏ qua tự động.
- Bỏ dòng tiêu đề tiếng Trung lặp lại ngay dưới header tiếng Anh.
- Tự dừng tại dòng "TOTAL" ở cuối bảng, không đọc nhầm thành 1 dòng hàng.
- Tự điền lại Packaging code / PO No. cho các ô bị merge.

Đã test trực tiếp với file `PL_CN-4785_TOTAL.xlsx` thật: ra đúng 162 carton, 171 trang nhãn, khớp 100% với Package Total (162) và Quantity Total (4782) ghi trong chính file Excel. Chỉ cần upload nguyên file, không cần xóa/cắt gì cả.
