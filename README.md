# bayich2_doichieutoa — Đối Chiếu Toa

Nhập hoặc chụp toa nhập hàng → so với tab **Retail** (sheet bayich2) → ghi **Giá Nhập Sỉ** mới và tô màu dòng:
đỏ = Giá Làm Tròn đổi (đổi giá bán), xanh = chỉ đổi giá nhập. Xong thì sang
[Đồng Bộ Giá](https://bayich2-dongbogia.vercel.app/) để đẩy giá sang Menu và Vân Bao Bì.

Thay cho `bayich2_doichieugia` (bản cũ dùng script gắn với Sheet).

## Cấu hình — tab `DoiChieu` trong sheet bayich2

| Khối | Tiêu đề | Nội dung |
|---|---|---|
| Cấu hình | Trường \| Giá trị \| Ghi chú | câu lệnh đọc toa, cảnh báo giá nhảy (%), bước làm tròn, 2 màu, xoá màu cũ, link Đồng Bộ Giá, link Tính Giá |
| Biệt danh | Cách viết trên toa \| Mặt Hàng (Retail) | "ly 720 vp" → Ly Trơn 720ml… — thêm/sửa thẳng trong Sheet |

Các khối và cột được tìm theo chữ tiêu đề, không theo vị trí.

## Cấu trúc

```
index.html        trang web (Vercel: bayich2-doichieutoa.vercel.app)
og.png            ảnh xem trước khi chia sẻ (1200×630)
.clasp.json       Script ID của backend, rootDir = appsscript
appsscript/       backend — Apps Script STANDALONE "bayich2_doichieutoa" (không deploy lên Vercel)
  Code.js
  appsscript.json
```

- Script: https://script.google.com/d/1G2zgC6wwXEevS8wBfch2mA8wMeEZT2RqdJ5ZWWLYo7q9yzBWoxtW8qVs/edit
- Deployment đang dùng (`API` trong index.html): `AKfycbyijaex0js9jjT42MDT1Z1JmQ8zdHFjwFiMY6UOjM_F0aSL3KQv46ZzF-n-KpIGUlxa`
- Cập nhật backend: `clasp push` → `clasp deploy -i <deploymentId>` (link /exec giữ nguyên).
- Lần đầu: mở script → chọn hàm `caiDat` → Run → cấp quyền Sheets (có tab DoiChieu rồi thì chỉ đọc thử).

Backend là web app quyền "Bất kỳ ai": chỉ ghi được cột Giá Nhập Sỉ, giá phải là số nguyên 1–100.000.000,
và chỉ khi giá trong Sheet vẫn đúng như lúc đối chiếu.
