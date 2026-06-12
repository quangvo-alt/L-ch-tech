================================================
  TECH SHIFT DASHBOARD — Hướng dẫn sử dụng
================================================

FILE QUAN TRỌNG
---------------
TechShiftDashboard_v7.gs  ← DÙNG FILE NÀY (phiên bản mới nhất)
TechShiftDashboard_v6.gs  ← backup (không cần dùng)

CẤU TRÚC GOOGLE SHEET
----------------------
Tab: "Filter"
  Cột G = Tên nhân viên
  Cột H = Department (Tech / QC / BA / T lobby / T build ...)
  Cột I = Shift (C1 / C2 / C3 / ME / NP / HO / OFF)
  Row 1 : Header — cell I1 chứa ngày hiện tại (theo hàm)
  Row 2 : Tiêu đề cột
  Row 3+: Data nhân viên

Ý NGHĨA CÁC CA
--------------
C1  = Ca 1 — 7:00 AM – 3:00 PM    — Công chuẩn
C2  = Ca 2 — 3:00 PM – 11:00 PM   — Công chuẩn
C3  = Ca 3 — 9:00 AM – 4:00 PM    — Công chuẩn
ME  = Mac Energy                   — Nghỉ có công
NP  = Nghỉ phép năm                — Có công
HO  = Nghỉ lễ / Nghỉ bù lễ        — Có công (màu đỏ)
OFF = Nghỉ thường

All times: CST (Central Standard Time, UTC-6)

CÁCH DEPLOY LẦN ĐẦU
--------------------
1. Mở Google Sheet → Extensions → Apps Script
2. Xóa toàn bộ code cũ trong editor
3. Paste toàn bộ nội dung file TechShiftDashboard_v7.gs vào
4. Nhấn Save (Ctrl+S)
5. Nhấn Deploy → New deployment
6. Chọn type: Web app
   - Execute as: Me
   - Who has access: Anyone  ← QUAN TRỌNG (ai cũng vào được, không cần login)
7. Nhấn Deploy → Copy URL

CÁCH CẬP NHẬT KHI SỬA CODE
----------------------------
1. Apps Script → Deploy → Manage deployments
2. Chọn deployment hiện tại → Edit (bút chì)
3. Version: chọn "New version"
4. Deploy → URL cũ vẫn hoạt động, không cần chia sẻ lại

DÙNG HÀNG NGÀY (30 giây)
--------------------------
1. Mở URL web app trên Chrome (máy tính, không cần zoom)
2. Data tự cập nhật theo ngày từ sheet
3. Nhấn nút "⬇ Tải PNG" → file tech-shift-DDMMYYYY.png
4. Gửi file PNG lên Telegram → hiển thị ảnh inline luôn

YÊU CẦU KỸ THUẬT
-----------------
- Browser: Chrome (desktop), zoom 100%
- Sheet phải share "Anyone with link can view" để script đọc được
- Script cần quyền: SpreadsheetApp (tự cấp khi deploy lần đầu)

LINK GOOGLE SHEET GỐC
----------------------
https://docs.google.com/spreadsheets/d/1tqYgUWg8VOmgwexkdsEJoumhiU_qetagI1pSJtXc018/edit?gid=895580574

================================================
