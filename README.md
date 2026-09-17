# bayich2_doichieutoa — Đối Chiếu Toa

Nhập hoặc chụp toa nhập hàng → so với tab **Retail** (sheet bayich2) → ghi **Giá Nhập Sỉ** mới và tô màu dòng:
đỏ = Giá Làm Tròn đổi (đổi giá bán), xanh = chỉ đổi giá nhập. Ghi xong, màn kết quả tự tính phần
**đồng bộ sang Menu** (gọi thẳng backend của [Đồng Bộ Giá](https://bayich2-dongbogia.vercel.app/),
cùng luật tích sẵn và cách xác minh khi Google lỗi) — bấm "Đồng bộ N món" là xong cả luồng.

Bước làm tròn chỉ khai một chỗ: dòng "Bước làm tròn" trong tab `CauHinh`. Cột Giá Làm Tròn của Retail đọc dòng đó
(`caiDat` tự đổi công thức, so từng dòng trước/sau; lệch thì trả lại công thức cũ).

Thay cho `bayich2_doichieugia` (bản cũ dùng script gắn với Sheet).

## Cấu hình và biệt danh (sheet bayich2)

| Tab | Tiêu đề | Nội dung |
|---|---|---|
| `CauHinh` (dùng chung, riêng tư) | Trường \| Giá trị \| Ghi chú \| Dùng cho | câu lệnh đọc toa, cảnh báo giá nhảy (%), bước làm tròn, 2 màu, xoá màu cũ, link Đồng Bộ Giá, link Tính Giá |
| `DoiChieu` | Cách viết trên toa \| Mặt Hàng (Retail) | "ly 720 vp" → Ly Trơn 720ml… — thêm/sửa thẳng trong Sheet |

Bước làm tròn, cảnh báo, màu và xoá màu cũ là dòng chung với Đồng Bộ Giá — sửa một chỗ, hai trang cùng theo.

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
- Lần đầu / nâng cấp: mở script → chọn hàm `caiDat` → Run. Tạo tab còn thiếu, thêm trường còn thiếu vào `CauHinh`;
  bản cũ để cấu hình trong tab `DoiChieu` thì chép sang `CauHinh` rồi xoá khối cũ.

Backend là web app quyền "Bất kỳ ai": chỉ ghi được cột Giá Nhập Sỉ, giá phải là số nguyên 1–100.000.000,
và chỉ khi giá trong Sheet vẫn đúng như lúc đối chiếu. Mọi lệnh ghi (kể cả nút "Đồng bộ N món") cần **Mã PIN chung**
trong tab CauHinh — nhập một lần trên mỗi máy; đọc (tải bảng Retail, đối chiếu) cũng cần PIN.
