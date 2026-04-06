# NEU Curriculum (CTĐT) Extractor

Parser cho chương trình đào tạo NEU. Trích xuất toàn bộ thông tin theo Schema 1.

## Cài đặt
```bash
pip install -r requirements.txt
```

## Chạy

```bash
# Test local (không cần MinIO) — đặt các file .docx vào ./doccur/
python main.py --doc ./doccur --local ./output

# MinIO: đọc từ courses-raw/curriculum/ → ghi vào courses-raw/qbcur/
python main.py

# Test 5 file đầu từ MinIO, lưu local để kiểm tra
python main.py --test --local ./output_test

# Overwrite file đã xử lý
python main.py --no-skip

# Chỉ liệt kê file không xử lý
python main.py --dry-run
```

## Cấu trúc JSON đầu ra (Schema 1)
```
training_program          — tên chương trình, bằng cấp, trình độ
training_program_version  — phiên bản, mã ngành, TC, năm áp dụng, ...
objectives[]              — mục tiêu đào tạo PO1..PO8
plos[]                    — chuẩn đầu ra PLO1.1..PLO3.4
po_plo_maps[]             — ma trận PO↔PLO
sections[]                — các mục văn bản (triết lý, cơ hội việc làm, ...)
components[]              — cây cấu trúc chương trình (1, 1.1, 2.2.1, ...)
courses[]                 — danh sách học phần + phân bổ học kỳ
career_paths[]            — cơ hội nghề nghiệp sau tốt nghiệp
graduation_requirements[] — điều kiện tốt nghiệp
_qa                       — completeness_score, issues
```

## Format hỗ trợ
- Chính quy tiếng Việt (Bảo hiểm, KDQT, Thống kê kinh tế, ...)
- AEP/CLC tiếng Anh (FE, EPMP, EDA, IHME, ...)
- Form mới 2025 (EDA, IHME)
