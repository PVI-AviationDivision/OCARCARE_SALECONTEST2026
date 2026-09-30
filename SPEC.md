# SPEC — OCARCARE Grand Prix Dashboard

Dashboard theo dõi Sales Contest **"Chào sân OCARCARE — Launch Sprint"** (21/09 – 31/10/2026)
dành cho cán bộ bán VPBank. Đơn vị vận hành: P.TMĐT, Ban HK&KHDN, Tổng Công ty Bảo hiểm PVI.

## 1. Mục tiêu
- Một link duy nhất (GitHub Pages), mở tốt trên điện thoại, cập nhật 1 lần/ngày sau báo cáo 17h.
- Cảm giác "đường đua": mỗi đơn vị là một chiếc xe trên làn, có vạch điều kiện, hạng, tăng/giảm hạng.
- Người xem tự tìm được đơn vị mình (kể cả gõ tên PGD) và biết còn thiếu bao nhiêu.

## 2. Nguồn dữ liệu
| File | Vai trò |
|---|---|
| `input/báo cáo 17h ngày ddmmyyyy VPBank.xlsx` | Báo cáo chi tiết nghiệp vụ XCG (sheet TK_NVU_OTO). Lấy file mới nhất trong thư mục |
| `mapping/Danh_sach_Mapping_*.xlsx` | Mapping Branch code → CN/PGD/TTTC/TTKD, Miền, Tỉnh, Đơn vị BH PVI phụ trách |

Cột dùng từ báo cáo (dò theo **tên tiêu đề**, không cố định chữ cái cột):
`Số đơn bảo hiểm`, `Số đơn sửa đổi bổ sung`, `Ngày cấp`, `Danh mục Kênh bán hàng`, `TỔNG PHÍ BH → Phí` (chưa VAT), `Mã chi nhánh Bank` (cột BX).

## 3. Quy tắc tính (chốt với anh Tùng 29/09/2026)
- **Doanh thu thi đua** = cột TỔNG PHÍ BH / Phí (chưa VAT) = **100% phí OPES + BH PVI**. Không quy đổi.
- **Kỳ tính**: Ngày cấp từ 21/09/2026 đến 31/10/2026. Chỉ lấy kênh chứa "OPES".
- **Số HĐBH** = số "Số đơn bảo hiểm" duy nhất, không tính dòng sửa đổi bổ sung. Doanh thu cộng cả dòng SĐBS (phí điều chỉnh).
- **Quy đổi đơn vị thi đua** theo cột "CN/PGD" của mapping:
  - `CN`, `PGD` → gộp về **Chi nhánh quản lý** (cột "Tên CN quản lý" / "Mã CN mẹ").
  - `TTTC` → là một đơn vị thi đua riêng (Trung tâm thế chấp).
  - `TTKD OTO MIEN BAC/NAM` → **Hub xe** (không tính vào bảng Chi nhánh/TTTC).
- **"Hôm nay"** = các đơn có Ngày cấp = ngày báo cáo. Hạng hôm qua = xếp hạng khi bỏ các đơn đó → tính mũi tên ▲▼.

### Hạng mục & thể lệ hiển thị
| Hạng mục | Cách hiển thị |
|---|---|
| 1a. 600 HĐBH đầu tiên (300/200/100 nghìn) | Thanh 600 ô toàn chương trình, xe ở vị trí đơn thứ N, còn bao nhiêu suất mỗi mức |
| 1b. Top 3 cá nhân (≥125tr) | Khung khoá: **"Sẽ cập nhật sau 31/10/2026"** (chưa có dữ liệu cán bộ bán) |
| 2. Chi nhánh / TTTC — Top 3: 7/5/3 triệu, ngưỡng 250tr | Đường đua chính, vạch 250tr, Top 10 + xem tất cả, ô tìm kiếm |
| 3a. GĐ Hub xe — Top 2: 15/10 triệu, ngưỡng 1,25 tỷ; +5tr ở 1,825 tỷ; +10tr ở 2,5 tỷ | Đua đối đầu Miền Bắc vs Miền Nam, 3 vạch mốc |
| 3b. Trưởng phòng Hub (Top 4, ngưỡng 400tr) | Ghi chú: chưa phân tách được cấp phòng/nhóm, công bố theo kết quả chính thức |

## 4. Bảo mật
- Chỉ `docs/` (trang + `data.js` **đã tổng hợp**) được đẩy lên GitHub.
- `input/`, `mapping/` nằm trong `.gitignore`. Không đưa tên khách hàng, SĐT, biển số, tên/email cán bộ lên web.

## 5. Giao diện
Một cột, mobile-first, nền navy/cam theo poster. Các khối: Hero (đếm ngược, dòng thời gian 41 ngày, 4 KPI) → Ticker tin nhanh → Khối 1 → Khối 2 → Khối 3 → Toàn cảnh (luỹ kế theo ngày, theo Miền, theo đơn vị BH PVI) → Chú thích thể lệ.
Hiệu ứng: xe chạy vào vị trí khi cuộn tới, số đếm tăng dần, pháo giấy khi có đơn vị vừa vượt ngưỡng trong ngày.

## 6. Quy trình cập nhật
Thả file báo cáo vào `input/` → chạy `update.bat` (hoặc Task Scheduler 17h30) → script build `docs/data.js` + lưu bản `docs/history/` → `git commit` + `git push` → GitHub Pages cập nhật sau ~1 phút.
