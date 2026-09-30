# OCARCARE Grand Prix — Dashboard Sales Contest

Bảng đua trực tuyến cho chương trình **Chào sân OCARCARE — Launch Sprint** (21/09 – 31/10/2026).
Quy tắc tính và thiết kế: xem `SPEC.md`.

```
OCarCare_Dashboard/
├─ input/                  ← thả file "báo cáo 17h ngày ddmmyyyy VPBank.xlsx" (lấy file MỚI NHẤT)
├─ mapping/                ← file Danh_sach_Mapping_*.xlsx (thay khi có bản mới)
├─ docs/                   ← trang web (GitHub Pages) — chỉ chứa số liệu ĐÃ TỔNG HỢP
│   ├─ index.html
│   ├─ data.js             ← script tạo tự động, không sửa tay
│   └─ history/            ← lưu số liệu từng ngày
├─ tools/build_data.py     ← đọc Excel → data.js (thể lệ nằm ở CONFIG đầu file)
├─ ket_noi_github_lan_dau.bat ← chạy 1 lần để nối với repo
├─ update.bat              ← CẬP NHẬT + ĐẨY LÊN GITHUB (nhấp đúp)
├─ xem_thu.bat             ← tạo dữ liệu + mở trang trên máy, không đẩy lên
└─ setup_lich_tu_dong.bat  ← tạo lịch tự chạy 17:30 hằng ngày
```

`input/`, `mapping/` và mọi file Excel đã được chặn trong `.gitignore` → **dữ liệu gốc (tên KH, SĐT, biển số) không bao giờ lên GitHub.**

---

## A. Cài đặt lần đầu (khoảng 15 phút)

**1. Phần mềm trên máy**
- Python 3.9 trở lên: https://www.python.org/downloads/ (khi cài, tick **"Add Python to PATH"**)
- Git: https://git-scm.com/download/win
- Mở Command Prompt, chạy: `pip install openpyxl`

**2. Kết nối với repo** `PVI-AviationDivision/OCARCARE_SALECONTEST2026`
- Nhấp đúp **`ket_noi_github_lan_dau.bat`** (chạy 1 lần). Script sẽ cài openpyxl, tạo dữ liệu, nối thư mục với repo, xoá file `index.html` cũ ở gốc repo và đẩy cấu trúc mới (`docs/`) lên.
- Lần đầu Git mở cửa sổ đăng nhập GitHub. Tài khoản cần có quyền ghi vào repo của tổ chức PVI-AviationDivision.

**3. Chuyển GitHub Pages sang thư mục /docs**
- Repo → **Settings → Pages** → Source: *Deploy from a branch* → Branch `main`, thư mục **`/docs`** → Save.
- Link cố định: **https://pvi-aviationdivision.github.io/OCARCARE_SALECONTEST2026/**

**4. Kiểm tra**: mở link, xem dòng "Cập nhật 17:00 · dd/mm/yyyy" đúng ngày báo cáo.

**5. (Tuỳ chọn) Chạy tự động 17:30 hằng ngày**: nhấp đúp `setup_lich_tu_dong.bat` (máy cần bật lúc đó).

## B. Cập nhật hằng ngày

1. Lưu file báo cáo 17h vào `input/` (giữ nguyên tên dạng `báo cáo 17h ngày 29092026 VPBank.xlsx` — ngày báo cáo đọc từ tên file).
2. Nhấp đúp **`update.bat`** → đọc kết quả kiểm tra trên màn hình → xong.
3. Khoảng 1 phút sau, trang web cập nhật (người xem nhấn tải lại).

Màn hình kiểm tra sẽ cảnh báo nếu:
- Có **Mã chi nhánh Bank chưa có trong mapping** → bổ sung mapping rồi chạy lại.
- Tổng doanh thu các đơn vị lệch với tổng file gốc.

## C. Khi cần chỉnh
| Việc | Sửa ở đâu |
|---|---|
| Ngưỡng, giải thưởng, mốc Hub, thời gian | `CONFIG` đầu file `tools/build_data.py` |
| Khi có dữ liệu cán bộ bán (Top 3 cá nhân) | Gửi file cho Claude để bổ sung bước ghép thứ 3 |
| Giao diện, chữ | `docs/index.html` |

## D. Lưu ý
- Số trên trang là **số tạm tính** để theo dõi, kết quả chính thức theo Thể lệ.
- Doanh thu = cột TỔNG PHÍ BH (chưa VAT) = 100% phí OPES + BH PVI, theo ngày cấp trong kỳ, kênh "CT ĐỒNG BH VCX OPES".
- Repo Public ⇒ ai có link đều xem được xếp hạng theo chi nhánh (không có dữ liệu cá nhân).
