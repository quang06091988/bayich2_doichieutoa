/**
 * Apps Script STANDALONE "bayich2_doichieutoa" — backend của trang Đối Chiếu Toa
 * (bayich2-doichieutoa.vercel.app). Thay cho script gắn với sheet bayich2 (bản bayich2_doichieugia).
 *
 *  ĐỌC : doPost {hanhDong:'khoiTao', pin} → tab Retail + cấu hình + biệt danh (tab DoiChieu) — cần Mã PIN chung
 *  GHI : doPost {hanhDong:'capNhat', pin, retail:[{ten, giaNhapSi, giaNhapSiCu}]} — cần Mã PIN chung (tab CauHinh)
 *          → chỉ ghi cột "Giá Nhập Sỉ", rồi tô màu dòng:
 *            ĐỎ   = Giá Làm Tròn đổi sau khi Sheet tính lại
 *            XANH = Giá Làm Tròn giữ nguyên
 *          Các cột công thức (Giá Nhập Lẻ, Giá Bán Lẻ, Giá Làm Tròn) không bao giờ bị ghi đè.
 *
 * Cấu hình (câu lệnh đọc toa, ngưỡng cảnh báo, bước làm tròn, màu, link…) nằm ở tab CauHinh — tab riêng tư dùng
 * chung cho mọi công cụ (Trường | Giá trị | Ghi chú | Dùng cho). Biệt danh nằm ở tab DoiChieu, sửa thẳng trong Sheet.
 * Cột trong các tab được tìm theo CHỮ TIÊU ĐỀ, không theo vị trí.
 *
 * CÀI / CẤP QUYỀN / NÂNG CẤP: chọn hàm caiDat → Run. Tạo tab còn thiếu, thêm trường còn thiếu vào CauHinh; bản cũ
 * để cấu hình trong tab DoiChieu thì chép sang CauHinh rồi xoá khối cũ (biệt danh giữ nguyên). Chạy lại không sao.
 * CẬP NHẬT CODE: clasp push → clasp deploy -i <deploymentId> (link /exec giữ nguyên).
 */

var ID_BAYICH2 = '1Wd4Zvq2xiIiEzou_dvE2YtOk-bJJhAD7se0yREYe9c8';   // sheet có tab Retail + DoiChieu

var TAB_RETAIL = 'Retail';
var TAB_DOICHIEU = 'DoiChieu';

var COT_RETAIL = {
  ten: 'Mặt Hàng', donViSi: 'Đơn Vị Sỉ', giaNhapSi: 'Giá Nhập Sỉ', soLuong: 'Số Lượng',
  donViLe: 'Đơn Vị Lẻ', loiNhuan: '% Lợi Nhuận', giaLamTron: 'Giá Làm Tròn'
};
var TD_GIA_BAN_LE = 'Giá Bán Lẻ';              // cột công thức Retail mà Giá Làm Tròn làm tròn theo
var TRUONG_LAM_TRON = 'Bước làm tròn';        // tên trường trong CauHinh — công thức Retail tìm theo tên này

/* Tab CauHinh: Trường | Giá trị | Ghi chú | Dùng cho — riêng tư, KHÔNG xuất bản lên web.
   Tab DoiChieu: khối biệt danh (Cách viết trên toa | Mặt Hàng (Retail)). */
var TAB_CAU_HINH = 'CauHinh';
var TD_TRUONG = 'Trường', TD_GIA_TRI = 'Giá trị', TD_GHI_CHU = 'Ghi chú', TD_DUNG_CHO = 'Dùng cho';
var TD_VIET = 'Cách viết trên toa', TD_RETAIL = 'Mặt Hàng (Retail)';

/* Tên trường trong tab CauHinh → khoá gửi cho trang web (bietDanh: tên cũ vẫn nhận) */
var TRUONG = [
  { khoa: 'lenhDocToa',  ten: 'Câu lệnh đọc toa',      kieu: 'chu' },
  { khoa: 'nguongNhay',  ten: 'Cảnh báo giá nhảy (%)', kieu: 'so' },
  { khoa: 'buocLamTron', ten: 'Bước làm tròn',         kieu: 'so' },
  { khoa: 'mauDo',       ten: 'Màu đổi giá bán',       kieu: 'mau' },
  { khoa: 'mauXanh',     ten: 'Màu chỉ đổi giá nhập',  kieu: 'mau', bietDanh: ['Màu chỉ đổi giá sỉ'] },
  { khoa: 'xoaMauCu',    ten: 'Xoá màu cũ khi ghi',    kieu: 'coKhong' },
  { khoa: 'linkDongBo',  ten: 'Link Đồng Bộ Giá',      kieu: 'link' },
  { khoa: 'linkTinhGia', ten: 'Link Tính Giá',         kieu: 'link' }
];

/* Giới hạn kỹ thuật */
var TOI_DA_DONG_GHI = 200;
var GIA_TOI_DA = 100000000;

/* ══════════════════ ĐỌC ══════════════════ */
function doGet(e) {
  try {
    var viec = (e && e.parameter && e.parameter.viec) || 'ping';
    /* Đọc giá đã chuyển sang doPost {hanhDong:'khoiTao', pin} — GET không trả giá nữa (trang bản cũ thì báo tải lại) */
    if (viec === 'khoiTao') return traLoi({ ok: false, maLoi: 'CU', loi: 'Trang đang là bản cũ — tải lại trang (xem giá giờ cần Mã PIN chung)' });
    return traLoi({ ok: true, ten: 'bayich2_doichieutoa', thoiGian: new Date().toISOString() });
  } catch (err) {
    return traLoi({ ok: false, loi: String(err) });
  }
}

function khoiTao() {
  var ss = SpreadsheetApp.openById(ID_BAYICH2);
  var canhBao = [];
  var r = docRetail(ss);
  if (r.loi) return { ok: false, loi: r.loi };
  var dc = docDoiChieu(ss, canhBao);

  var coRetail = {};
  r.ds.forEach(function (x) { coRetail[chuanHoa(x.ten)] = x.ten; });
  var bietDanh = [];
  dc.bietDanh.forEach(function (b) {
    var ten = coRetail[chuanHoa(b.retail)];
    if (ten) bietDanh.push({ viet: b.viet, retail: ten });
    else canhBao.push('Biệt danh "' + b.viet + '" trỏ tới "' + b.retail + '" — không có trong Retail');
  });

  return {
    ok: true,
    retail: r.ds.map(function (x) {
      return { ten: x.ten, donViSi: x.donViSi, donViLe: x.donViLe, soLuong: x.soLuong,
               loiNhuan: x.loiNhuan, giaNhapSi: x.giaNhapSi, giaLamTron: x.giaLamTron };
    }),
    cauHinh: dc.cauHinh, bietDanh: bietDanh, canhBao: canhBao, thoiGian: new Date().toISOString()
  };
}

/* Đọc tab Retail, tìm cột theo tiêu đề. Trả {sh, hang, cot, ds} hoặc {loi} */
function docRetail(ss) {
  var sh = ss.getSheetByName(TAB_RETAIL);
  if (!sh) return { loi: 'Không tìm thấy tab ' + TAB_RETAIL };
  var hang = sh.getDataRange().getValues();
  if (!hang.length) return { loi: 'Tab ' + TAB_RETAIL + ' đang trống' };

  var cot = {}, thieu = [];
  Object.keys(COT_RETAIL).forEach(function (k) {
    cot[k] = timCot(hang[0], COT_RETAIL[k]);
    if (cot[k] < 0) thieu.push(COT_RETAIL[k]);
  });
  if (thieu.length) return { loi: 'Tab ' + TAB_RETAIL + ' thiếu cột: ' + thieu.join(', ') };

  var ds = [];
  for (var i = 1; i < hang.length; i++) {
    var d = hang[i], ten = String(d[cot.ten] || '').trim();
    if (!ten) continue;
    var ln = soThuc(d[cot.loiNhuan]);
    ds.push({
      ten: ten,
      donViSi: String(d[cot.donViSi] || '').trim(),
      donViLe: String(d[cot.donViLe] || '').trim(),
      soLuong: soThuc(d[cot.soLuong]) || 1,
      loiNhuan: ln == null ? 0 : (ln > 1 ? ln / 100 : ln),
      giaNhapSi: soHoa(d[cot.giaNhapSi]),
      giaLamTron: soHoa(d[cot.giaLamTron]),
      _dong: i + 1
    });
  }
  return { sh: sh, hang: hang, cot: cot, ds: ds };
}

/* Cấu hình từ tab CauHinh (bản cũ còn khối cấu hình trong tab DoiChieu thì đọc tạm từ đó) + biệt danh từ tab DoiChieu.
   Thiếu gì thì ghi vào canhBao, trường thiếu nhận null (tính năng đó tự tắt ở trang web). */
function docDoiChieu(ss, canhBao) {
  var cauHinh = {};
  TRUONG.forEach(function (t) { cauHinh[t.khoa] = null; });
  var kq = { cauHinh: cauHinh, bietDanh: [] };

  var sh = ss.getSheetByName(TAB_DOICHIEU);
  var chung = docCauHinhChung(ss), cu = khoiCu(sh), conCu = false;
  TRUONG.forEach(function (t) {
    var v = layTheoTen(chung, tenCuaTruong(t));
    if (v === undefined && cu) { v = layTheoTen(cu.gt, tenCuaTruong(t)); if (v !== undefined) conCu = true; }
    if (v === undefined) { canhBao.push('Tab ' + TAB_CAU_HINH + ' thiếu trường "' + t.ten + '" — mở Apps Script, chạy hàm caiDat một lần để thêm'); return; }
    cauHinh[t.khoa] = docGiaTri(v, t.kieu);
    if (cauHinh[t.khoa] == null && String(v).trim() !== '')
      canhBao.push('Trường "' + t.ten + '" có giá trị không hợp lệ: ' + v);
  });
  if (conCu) canhBao.push('Cấu hình còn nằm ở tab ' + TAB_DOICHIEU + ' — mở Apps Script, chạy hàm caiDat một lần để chuyển sang tab ' + TAB_CAU_HINH);

  if (!sh) { canhBao.push('Chưa có tab ' + TAB_DOICHIEU + ' — mở Apps Script, chạy hàm caiDat một lần'); return kq; }
  var k = khoiDoiChieu(sh);
  if (k.cotViet < 0 || k.cotRetail < 0) canhBao.push('Tab ' + TAB_DOICHIEU + ' thiếu tiêu đề "' + TD_VIET + '" / "' + TD_RETAIL + '"');
  else {
    for (var j = 1; j < k.hang.length; j++) {
      var viet = String(k.hang[j][k.cotViet] || '').trim(), retail = String(k.hang[j][k.cotRetail] || '').trim();
      if (viet && retail) kq.bietDanh.push({ viet: viet, retail: retail });
    }
  }
  return kq;
}

function khoiDoiChieu(sh) {
  var hang = sh.getDataRange().getValues();
  var tieuDe = hang[0] || [];
  return { hang: hang, cotViet: timCot(tieuDe, TD_VIET), cotRetail: timCot(tieuDe, TD_RETAIL) };
}

function docGiaTri(v, kieu) {
  var s = String(v == null ? '' : v).trim();
  if (!s) return null;
  if (kieu === 'so') { var n = soThuc(v); return n != null && n > 0 ? n : null; }
  if (kieu === 'mau') return /^#[0-9a-f]{6}$/i.test(s) ? s : null;
  if (kieu === 'coKhong') { var c = chuanHoa(s); return c === 'co' ? true : c === 'khong' ? false : null; }
  if (kieu === 'link') return /^https:\/\//i.test(s) ? s : null;
  return s;
}

/* ══════════════════ GHI ══════════════════ */
/* Mọi lệnh GHI cần Mã PIN chung (tab CauHinh). Kiểm PIN TRƯỚC khi giữ khoá ghi — sai PIN (chờ 2 giây) không chặn người khác.
   {hanhDong:'kiemPin', pin} để trang kiểm PIN ngay lúc nhập. {hanhDong:'khoiTao', pin} = ĐỌC — cũng cần PIN, không giữ khoá ghi. */
function doPost(e) {
  try {
    var d = JSON.parse(e.postData.contents);
    var p = kiemPin(SpreadsheetApp.openById(ID_BAYICH2), d.pin);
    if (!p.ok) return traLoi(p);
    if (d.hanhDong === 'kiemPin') return traLoi({ ok: true });
    if (d.hanhDong === 'khoiTao') return traLoi(khoiTao());
    if (d.hanhDong !== 'capNhat') return traLoi({ ok: false, loi: 'Hành động không hợp lệ' });
    var khoa = LockService.getScriptLock();
    if (!khoa.tryLock(15000)) return traLoi({ ok: false, loi: 'Đang có một lần ghi khác, thử lại sau ít giây' });
    try { return traLoi(capNhat(d)); } finally { khoa.releaseLock(); }
  } catch (err) {
    return traLoi({ ok: false, loi: String(err) });
  }
}

function capNhat(d) {
  if (!Array.isArray(d.retail) || !d.retail.length) return { ok: false, loi: 'Không có dòng nào để ghi' };
  if (d.retail.length > TOI_DA_DONG_GHI) return { ok: false, loi: 'Quá nhiều dòng trong một lần ghi' };

  var ss = SpreadsheetApp.openById(ID_BAYICH2);
  var r = docRetail(ss);
  if (r.loi) return { ok: false, loi: r.loi };
  var cauHinh = docDoiChieu(ss, []).cauHinh;

  var theoTen = {};
  r.ds.forEach(function (x) { theoTen[chuanHoa(x.ten)] = x; });

  /* Bước 1: kiểm từng dòng — chỉ dòng hợp lệ và giá trong Sheet chưa bị ai sửa mới được ghi */
  var can = [], boQua = [], daChon = {};
  d.retail.forEach(function (it) {
    it = it || {};
    var ten = String(it.ten || '').trim();
    var x = theoTen[chuanHoa(ten)];
    if (!x) { boQua.push({ ten: ten, lyDo: 'Không có trong Retail' }); return; }
    if (daChon[x._dong]) { boQua.push({ ten: x.ten, lyDo: 'Trùng với dòng khác trong cùng lần ghi' }); return; }
    var gia = it.giaNhapSi;
    if (typeof gia !== 'number' || gia !== Math.round(gia) || gia < 1 || gia > GIA_TOI_DA) {
      boQua.push({ ten: x.ten, lyDo: 'Giá nhập sỉ không hợp lệ' }); return;
    }
    if (it.giaNhapSiCu != null && x.giaNhapSi != null && Math.abs(x.giaNhapSi - it.giaNhapSiCu) >= 1) {
      boQua.push({ ten: x.ten, lyDo: 'Giá trong Sheet đã khác lúc đối chiếu', trongSheet: x.giaNhapSi }); return;
    }
    daChon[x._dong] = true;
    can.push({ x: x, gia: gia });
  });
  if (!can.length) return { ok: false, loi: 'Không ghi được dòng nào', boQua: boQua };

  /* Tô từ cột A tới cột Giá Làm Tròn */
  var soCotTo = r.cot.giaLamTron + 1;
  if (cauHinh.xoaMauCu && r.hang.length > 1) r.sh.getRange(2, 1, r.hang.length - 1, soCotTo).setBackground(null);

  /* Bước 2: ghi Giá Nhập Sỉ — và chỉ cột đó */
  can.forEach(function (c) { r.sh.getRange(c.x._dong, r.cot.giaNhapSi + 1).setValue(c.gia); });
  SpreadsheetApp.flush();   // để Sheet tính lại các cột công thức

  /* Bước 3: đọc lại Giá Làm Tròn, so trước/sau để quyết màu */
  var ketQua = [], soDo = 0, soXanh = 0;
  can.forEach(function (c) {
    var sau = soHoa(r.sh.getRange(c.x._dong, r.cot.giaLamTron + 1).getValue());
    var doi = sau !== c.x.giaLamTron;
    var mau = doi ? cauHinh.mauDo : cauHinh.mauXanh;
    if (mau) r.sh.getRange(c.x._dong, 1, 1, soCotTo).setBackground(mau);
    if (doi) soDo++; else soXanh++;
    ketQua.push({ ten: c.x.ten, giaNhapSiCu: c.x.giaNhapSi, giaNhapSi: c.gia, tronTruoc: c.x.giaLamTron, tronSau: sau, doi: doi });
  });
  SpreadsheetApp.flush();

  return { ok: true, soDo: soDo, soXanh: soXanh, ketQua: ketQua, boQua: boQua, khoiTao: khoiTao() };
}

/* ══════════════════ CÀI LẦN ĐẦU ══════════════════ */
/* Trường cấu hình của công cụ này trong tab CauHinh: [Trường, Giá trị mặc định, Ghi chú, Dùng cho].
   Chỉ dùng khi thêm trường còn thiếu — sau đó sửa trong Sheet, không sửa ở đây. Trường dùng chung với Đồng bộ giá
   (bước làm tròn, cảnh báo, màu, xoá màu cũ) có cùng tên và cùng dòng ở cả 2 script. */
var DUNG_CHUNG = 'Đối chiếu toa, Đồng bộ giá';
var CAU_HINH_BAN_DAU = [
  ['Câu lệnh đọc toa',
   'Đọc hóa đơn bán lẻ viết tay trong ảnh này. Trả về DUY NHẤT một mảng JSON, không kèm lời dẫn hay dấu ```.\n' +
   'Mỗi phần tử: {"ten": tên hàng đúng như trên toa, "dvt": đơn vị, "sl": số lượng, "donGia": đơn giá, "thanhTien": thành tiền}\n' +
   'Giữ nguyên con số như người viết ghi (216 thì để 216, đừng tự nhân nghìn).\n' +
   'Kiểm tra thành tiền = số lượng × đơn giá; nếu lệch thì đọc lại nét chữ cho khớp.\n' +
   'Dòng nào không đọc chắc chắn thì để "ten": "?" và vẫn giữ các số đọc được.',
   'Gửi kèm ảnh toa cho Claude', 'Đối chiếu toa'],
  ['Cảnh báo giá nhảy (%)', 30, 'Giá đổi từ mức này trở lên → bỏ tích sẵn để xem lại', DUNG_CHUNG],
  ['Bước làm tròn', 1000, 'Giá bán làm tròn tới bội số này — cột Giá Làm Tròn của Retail, Đối chiếu toa và Đồng bộ giá cùng theo dòng này. Ghi số liền, không dấu chấm (vd 1000)', 'Retail, ' + DUNG_CHUNG],
  ['Màu đổi giá bán', '#f4cccc', 'Tô dòng có đổi giá bán. Để trống = không tô', DUNG_CHUNG],
  ['Màu chỉ đổi giá nhập', '#d9ead3', 'Tô dòng chỉ đổi giá nhập / giá sỉ. Để trống = không tô', DUNG_CHUNG],
  ['Xoá màu cũ khi ghi', 'Có', 'Có / Không — để màu chỉ phản ánh lần ghi gần nhất', DUNG_CHUNG],
  ['Link Đồng Bộ Giá', 'https://bayich2-dongbogia.vercel.app/', 'Nút hiện sau khi ghi toa xong. Để trống = ẩn nút', 'Đối chiếu toa'],
  ['Link Tính Giá', 'https://bayich2-tinhgia.vercel.app/', 'Link "Tính giá bán" ở mục Hàng mới. Để trống = ẩn link', 'Đối chiếu toa']
];
var BIET_DANH_BAN_DAU = [
  ['Ly Trơn 360ml', ['ly 360', 'ly 360 noni', '360 noni', 'ly 360 doni', '360 doni']],
  ['Ly Trơn 500ml', ['ly 500', 'ly 500 noni', '500 noni', 'ly 500 doni', '500 doni']],
  ['Ly Trơn 650ml', ['ly 650', 'ly 650 noni', '650 noni', 'ly 650 doni', '650 doni']],
  ['Ly Trơn 720ml', ['ly 720', 'ly 720 vp', 'ly 800 vp', 'ly 800', '800 vp']],
  ['Nắp Hữu Phong 95mm', ['nap hp 95', 'hp 95', 'nap hp95']],
  ['Nắp Cầu Tròn 95mm', ['nap cau tron 95', 'cau tron 95', 'nap cau tron']],
  ['Nắp Bằng Cao 95mm', ['nap cau bang 95', 'nap bang 95', 'cau bang 95', 'nap cb 95', 'nap cau 17 95']],
  ['Nắp Hữu Phong 116mm', ['nap hp 116', 'hp 116', 'nap hp116']],
  ['Ống Hút Trong 6mm, 8mm', ['ong 6', 'ong 8', 'oc 6', 'oc 8', 'ong hut', 'ong hut trong']],
  ['Muỗng GT 150mm', ['gt muong ngan', 'muong ngan', 'gt ngan', 'gt 150', 'muong gt ngan']],
  ['Muỗng GT 200mm', ['gt muong dai', 'muong dai', 'gt dai', 'gt 200', 'muong gt dai']],
  ['Bị 1 Ly, 2 Ly Lớn', ['bi 1 ly', 'bi 2 ly', 'bi 1 2 ly', 'bi ly lon']],
  ['Bị Ngang Lớn', ['bi ngang', 'bi ngang lon']],
  ['Bị 40 Dương', ['bi 40d', '40d', 'bi 40 duong', '40 d']],
  ['Bị 50 Dương', ['bi 50d', '50d', 'bi 50 duong', '50 d']]
];

function caiDat() {
  var ss = SpreadsheetApp.openById(ID_BAYICH2);
  var sh = ss.getSheetByName(TAB_DOICHIEU);
  if (sh) {
    Logger.log('Tab ' + TAB_DOICHIEU + ' đã có — giữ nguyên biệt danh');
  } else {
    sh = ss.insertSheet(TAB_DOICHIEU, ss.getNumSheets());
    var bd = [[TD_VIET, TD_RETAIL]];
    BIET_DANH_BAN_DAU.forEach(function (m) { m[1].forEach(function (v) { bd.push([v, m[0]]); }); });
    sh.getRange(1, 1, bd.length, 2).setValues(bd);
    sh.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#2f5233').setFontColor('#f6f1e4');
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 170); sh.setColumnWidth(2, 190);
    Logger.log('Đã tạo tab ' + TAB_DOICHIEU + ': ' + (bd.length - 1) + ' biệt danh');
  }
  chuyenCauHinh(layTabCauHinh(ss), sh, CAU_HINH_BAN_DAU);
  noiCongThucLamTron(ss);
  var kq = khoiTao();
  Logger.log(kq.ok ? ('Đọc thử: ' + kq.retail.length + ' mặt hàng Retail, ' + kq.bietDanh.length + ' biệt danh') : ('Lỗi: ' + kq.loi));
  (kq.canhBao || []).forEach(function (c) { Logger.log('  ! ' + c); });
}

/* Một chỗ khai bước làm tròn: cột Giá Làm Tròn của Retail đọc "Bước làm tròn" từ tab CauHinh (tìm theo TÊN trường,
   nên sắp xếp lại CauHinh không sao). Chỉ thay dòng đang có công thức; tính lại rồi so từng dòng — lệch dù một dòng
   là trả lại toàn bộ công thức cũ. Dòng đã nối rồi thì bỏ qua, nên chạy lại không sao. */
function noiCongThucLamTron(ss) {
  var r = docRetail(ss);
  if (r.loi) { Logger.log('Chưa nối công thức làm tròn: ' + r.loi); return false; }
  var cBan = timCot(r.hang[0], TD_GIA_BAN_LE);
  if (cBan < 0) { Logger.log('Tab ' + TAB_RETAIL + ' thiếu cột "' + TD_GIA_BAN_LE + '" — chưa nối công thức làm tròn'); return false; }
  var shCH = ss.getSheetByName(TAB_CAU_HINH), tdCH = shCH ? (shCH.getDataRange().getValues()[0] || []) : [];
  var cTr = timCot(tdCH, TD_TRUONG), cGt = timCot(tdCH, TD_GIA_TRI);
  if (cTr < 0 || cGt < 0 || layTheoTen(docCauHinhChung(ss), [TRUONG_LAM_TRON]) === undefined) {
    Logger.log('Tab ' + TAB_CAU_HINH + ' chưa có dòng "' + TRUONG_LAM_TRON + '" — chưa nối công thức làm tròn'); return false;
  }
  var cot = r.cot.giaLamTron + 1, soDong = r.hang.length - 1;
  if (soDong < 1) return false;
  var o = r.sh.getRange(2, cot, soDong, 1), cu = o.getFormulas(), truoc = o.getValues();
  var cotGt = TAB_CAU_HINH + '!$' + chuCot(cGt + 1) + ':$' + chuCot(cGt + 1), cotTr = TAB_CAU_HINH + '!$' + chuCot(cTr + 1) + ':$' + chuCot(cTr + 1);
  var doi = [];
  cu.forEach(function (row, i) {
    var f = String(row[0] || '');
    if (!f || f.indexOf(TAB_CAU_HINH) >= 0) return;   // ô nhập tay, hoặc đã nối rồi
    doi.push({ dong: i + 2, cu: f, ref: chuCot(cBan + 1) + (i + 2) });
  });
  if (!doi.length) { Logger.log('Cột Giá Làm Tròn: không còn công thức nào cần nối (đã nối rồi hoặc nhập tay)'); return true; }

  /* Công thức theo ngôn ngữ của Sheet: bản tiếng Việt ngăn đối số bằng ";" (dấu "," là dấu thập phân) — lấy dấu ngăn
     từ công thức đang có. Thử trước trên 1 dòng; Sheet không nhận thì thử dấu còn lại, rồi mới làm cả cột. */
  var congThuc = function (ref, d) {
    return '=IF(' + ref + '=""' + d + '""' + d + 'MROUND(' + ref + d + 'VALUE(INDEX(' + cotGt + d + 'MATCH("' + TRUONG_LAM_TRON + '"' + d + cotTr + d + '0)))))';
  };
  var coChamPhay = doi.some(function (x) { return x.cu.indexOf(';') >= 0; });
  var thu = doi[0], oThu = r.sh.getRange(thu.dong, cot), dau = null, daThu = [];
  (coChamPhay ? [';', ','] : [',', ';']).some(function (d) {
    oThu.setFormula(congThuc(thu.ref, d));
    SpreadsheetApp.flush();
    var v = oThu.getValue();
    if (soHoa(v) === soHoa(truoc[thu.dong - 2][0])) { dau = d; return true; }
    daThu.push('dấu "' + d + '" → ' + JSON.stringify(v));
    oThu.setFormula(thu.cu);
    return false;
  });
  if (!dau) {
    SpreadsheetApp.flush();
    Logger.log('Chưa nối công thức làm tròn: thử ở dòng ' + thu.dong + ' không ra đúng số (' + daThu.join(' · ') + ') — giữ công thức cũ: ' + thu.cu);
    return false;
  }
  doi.forEach(function (x) { r.sh.getRange(x.dong, cot).setFormula(congThuc(x.ref, dau)); });
  SpreadsheetApp.flush();
  var sau = o.getValues(), lech = [];
  doi.forEach(function (x) {
    var a = soHoa(truoc[x.dong - 2][0]), b = sau[x.dong - 2][0];
    if (a !== soHoa(b)) lech.push({ x: x, a: a, b: b });
  });
  if (lech.length) {
    doi.forEach(function (x) { r.sh.getRange(x.dong, cot).setFormula(x.cu); });
    SpreadsheetApp.flush();
    var l = lech[0];
    Logger.log('Chưa nối công thức làm tròn: ' + lech.length + ' dòng ra số khác (vd dòng ' + l.x.dong + ': ' + l.a + ' → ' + JSON.stringify(l.b) +
      ') — đã trả lại công thức cũ. Công thức cũ dòng ' + l.x.dong + ': ' + l.x.cu);
    return false;
  }
  Logger.log('Cột Giá Làm Tròn của Retail đã đọc "' + TRUONG_LAM_TRON + '" từ tab ' + TAB_CAU_HINH + ' (' + doi.length + ' dòng, giá không đổi)');
  return true;
}
function chuCot(n) { var s = ''; while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }

/* ══════════════════ TAB CauHinh (dùng chung cho mọi công cụ — cùng đoạn code ở mọi script) ══════════════════ */
/* { chuanHoa(tên trường): giá trị } — chưa có tab thì rỗng */
function docCauHinhChung(ss) {
  var sh = ss.getSheetByName(TAB_CAU_HINH), gt = {};
  if (!sh) return gt;
  var hang = sh.getDataRange().getValues(), td = hang[0] || [];
  var cTr = timCot(td, TD_TRUONG), cGt = timCot(td, TD_GIA_TRI);
  if (cTr < 0 || cGt < 0) return gt;
  for (var i = 1; i < hang.length; i++) { var t = chuanHoa(hang[i][cTr]); if (t && !(t in gt)) gt[t] = hang[i][cGt]; }
  return gt;
}

function tenCuaTruong(t) { return [t.ten].concat(t.bietDanh || []); }

function timTruong(ten) {
  for (var i = 0; i < TRUONG.length; i++) if (chuanHoa(TRUONG[i].ten) === chuanHoa(ten)) return TRUONG[i];
  return null;
}

function layTheoTen(gt, ds) {
  for (var i = 0; i < ds.length; i++) { var k = chuanHoa(ds[i]); if (k in gt) return gt[k]; }
  return undefined;
}

/* Khối cấu hình kiểu cũ (Trường | Giá trị | Ghi chú) nằm trong tab riêng của công cụ — null nếu không có */
function khoiCu(sh) {
  if (!sh) return null;
  var hang = sh.getDataRange().getValues(), td = hang[0] || [];
  var k = { cTr: timCot(td, TD_TRUONG), cGt: timCot(td, TD_GIA_TRI), cGc: timCot(td, TD_GHI_CHU), gt: {}, cuoi: 1 };
  if (k.cTr < 0 || k.cGt < 0) return null;
  for (var i = 1; i < hang.length; i++) { var t = chuanHoa(hang[i][k.cTr]); if (t) { k.gt[t] = hang[i][k.cGt]; k.cuoi = i + 1; } }
  return k;
}

function layTabCauHinh(ss) {
  var sh = ss.getSheetByName(TAB_CAU_HINH);
  if (sh) return sh;
  sh = ss.insertSheet(TAB_CAU_HINH, ss.getNumSheets());
  sh.getRange(1, 1, 1, 4).setValues([[TD_TRUONG, TD_GIA_TRI, TD_GHI_CHU, TD_DUNG_CHO]])
    .setFontWeight('bold').setBackground('#2f5233').setFontColor('#f6f1e4');
  sh.setFrozenRows(1);
  sh.setColumnWidth(1, 170); sh.setColumnWidth(2, 320); sh.setColumnWidth(3, 420); sh.setColumnWidth(4, 170);
  Logger.log('Đã tạo tab ' + TAB_CAU_HINH + ' — KHÔNG đưa tab này vào "Xuất bản lên web"');
  return sh;
}

/* Đưa các trường của công cụ này vào tab CauHinh:
   - CauHinh chưa có trường → thêm dòng, giá trị lấy từ khối cũ (nếu có) hoặc mặc định
   - CauHinh đã có (do công cụ khác thêm) → giữ nguyên giá trị; khối cũ khác giá trị thì ghi log cho biết.
     Ghi chú / Dùng cho thì cập nhật theo bản mới (trường dùng chung có cùng chữ ở mọi script)
   Xong thì xoá khối cũ trong tab của công cụ để chỉ còn một chỗ sửa. Dòng khác trong CauHinh (PIN…) không đụng tới. */
function chuyenCauHinh(shCH, shCu, banDau) {
  var hang = shCH.getDataRange().getValues(), td = hang[0] || [];
  var cTr = timCot(td, TD_TRUONG), cGt = timCot(td, TD_GIA_TRI), cGc = timCot(td, TD_GHI_CHU), cDc = timCot(td, TD_DUNG_CHO);
  if (cTr < 0 || cGt < 0) { Logger.log('Tab ' + TAB_CAU_HINH + ' thiếu tiêu đề Trường / Giá trị — chưa chuyển được cấu hình'); return false; }
  var coCH = {}, dongCH = {}, cuoi = 1;
  for (var i = 1; i < hang.length; i++) { var t = String(hang[i][cTr] || '').trim(); if (t) { coCH[chuanHoa(t)] = hang[i][cGt]; dongCH[chuanHoa(t)] = i; cuoi = i + 1; } }
  var cu = khoiCu(shCu), soThem = 0;
  banDau.forEach(function (row) {
    var tr = timTruong(row[0]), ds = tr ? tenCuaTruong(tr) : [row[0]];
    var vCH = layTheoTen(coCH, ds), vCu = cu ? layTheoTen(cu.gt, ds) : undefined;
    if (vCH !== undefined) {
      if (vCu !== undefined && String(vCu).trim() !== String(vCH).trim())
        Logger.log('  ! "' + row[0] + '": tab ' + TAB_CAU_HINH + ' đang là "' + vCH + '", khối cũ là "' + vCu + '" — giữ giá trị ở ' + TAB_CAU_HINH);
      var d = layTheoTen(dongCH, ds);
      [[cGc, row[2]], [cDc, row[3]]].forEach(function (x) {
        if (x[0] >= 0 && String(hang[d][x[0]] == null ? '' : hang[d][x[0]]) !== String(x[1])) {
          shCH.getRange(d + 1, x[0] + 1).setValue(x[1]);
          Logger.log('  ~ ' + row[0] + ': cập nhật cột ' + td[x[0]]);
        }
      });
      return;
    }
    var r = cuoi + 1 + soThem++;
    shCH.getRange(r, cTr + 1).setValue(row[0]);
    shCH.getRange(r, cGt + 1).setNumberFormat('@').setValue(vCu !== undefined ? vCu : row[1]);
    if (cGc >= 0) shCH.getRange(r, cGc + 1).setValue(row[2]);
    if (cDc >= 0) shCH.getRange(r, cDc + 1).setValue(row[3]);
    Logger.log('  + ' + row[0] + (vCu !== undefined ? ' — chép từ khối cũ' : ' — giá trị mặc định'));
  });
  if (!soThem) Logger.log('Tab ' + TAB_CAU_HINH + ' đủ trường của công cụ này');
  if (cu) {
    xoaKhoiCu(shCu, cu);
    Logger.log('Đã xoá khối cấu hình cũ ở tab ' + shCu.getName() + ' — từ nay sửa ở tab ' + TAB_CAU_HINH);
  }
  return true;
}

function xoaKhoiCu(sh, k) {
  [k.cTr, k.cGt, k.cGc].forEach(function (c) {
    if (c >= 0) sh.getRange(1, c + 1, k.cuoi, 1).clearContent().setBackground(null).setFontWeight('normal');
  });
  sh.getRange(1, k.cTr + 1).setValue('Cấu hình → tab ' + TAB_CAU_HINH).setFontWeight('bold');
}

/* Mã PIN chung (tab CauHinh) — cần cho mọi lệnh đọc và ghi. Sai / thiếu thì chờ 2 giây như sổ bán hàng (chống dò PIN). */
var TRUONG_PIN = 'Mã PIN chung';
function kiemPin(ss, pin) {
  var dung = layTheoTen(docCauHinhChung(ss), [TRUONG_PIN]);
  if (dung === undefined || chuanPin(dung) === '')
    return { ok: false, maLoi: 'THIEU_PIN', loi: 'Chưa có "' + TRUONG_PIN + '" trong tab ' + TAB_CAU_HINH + ' — chưa ghi được' };
  if (chuanPin(pin) !== chuanPin(dung)) {
    Utilities.sleep(2000);
    return { ok: false, maLoi: 'PIN', loi: 'Sai mã PIN — xem ô "' + TRUONG_PIN + '" ở tab ' + TAB_CAU_HINH };
  }
  return { ok: true };
}

/* So PIN giống sổ bán hàng (bayich2_pos/appsscript): bỏ mọi khoảng trắng và số 0 đầu — ô PIN bị Sheet đổi thành số vẫn khớp */
function chuanPin(s) { return String(s == null ? '' : s).replace(/\s+/g, '').replace(/^0+(?=\d)/, ''); }

/* ══════════════════ phụ trợ ══════════════════ */
/* Số từ ô Sheet: số giữ nguyên, chữ kiểu "17.000" / "0,5" / "10%" thì bóc ra. Trống → null */
function soThuc(v) {
  if (typeof v === 'number') return v;
  var s = String(v == null ? '' : v).replace(/[^\d,.-]/g, '').replace(/\./g, '').replace(',', '.');
  var n = parseFloat(s);
  return isNaN(n) ? null : n;
}

function soHoa(v) {
  var n = soThuc(v);
  return n == null ? null : Math.round(n);
}

function chuanHoa(s) {
  return String(s == null ? '' : s)
    .toLowerCase()
    .normalize('NFD').replace(/[̀-ͯ]/g, '')
    .replace(/đ/g, 'd')
    .replace(/[^a-z0-9]+/g, ' ')
    .trim();
}

function timCot(tieuDe, ten) {
  var can = chuanHoa(ten);
  for (var i = 0; i < tieuDe.length; i++) if (chuanHoa(tieuDe[i]) === can) return i;
  return -1;
}

function traLoi(o) {
  return ContentService.createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}
