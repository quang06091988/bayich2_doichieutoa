/**
 * Apps Script STANDALONE "bayich2_doichieutoa" — backend của trang Đối Chiếu Toa
 * (bayich2-doichieutoa.vercel.app). Thay cho script gắn với sheet bayich2 (bản bayich2_doichieugia).
 *
 *  ĐỌC : doGet?viec=khoiTao → tab Retail + cấu hình + biệt danh (tab DoiChieu)
 *  GHI : doPost {hanhDong:'capNhat', retail:[{ten, giaNhapSi, giaNhapSiCu}]}
 *          → chỉ ghi cột "Giá Nhập Sỉ", rồi tô màu dòng:
 *            ĐỎ   = Giá Làm Tròn đổi sau khi Sheet tính lại
 *            XANH = Giá Làm Tròn giữ nguyên
 *          Các cột công thức (Giá Nhập Lẻ, Giá Bán Lẻ, Giá Làm Tròn) không bao giờ bị ghi đè.
 *
 * Mọi thông tin nghiệp vụ (câu lệnh đọc toa, ngưỡng cảnh báo, màu, biệt danh…) nằm trong tab DoiChieu —
 * biệt danh sửa thẳng trong Sheet. Cột trong cả 2 tab được tìm theo CHỮ TIÊU ĐỀ, không theo vị trí.
 *
 * CÀI / CẤP QUYỀN / NÂNG CẤP: chọn hàm caiDat → Run (chưa có tab DoiChieu thì tạo; có rồi thì chỉ thêm
 * trường cấu hình mới còn thiếu, rồi đọc thử).
 * CẬP NHẬT CODE: clasp push → clasp deploy -i <deploymentId> (link /exec giữ nguyên).
 */

var ID_BAYICH2 = '1Wd4Zvq2xiIiEzou_dvE2YtOk-bJJhAD7se0yREYe9c8';   // sheet có tab Retail + DoiChieu

var TAB_RETAIL = 'Retail';
var TAB_DOICHIEU = 'DoiChieu';

var COT_RETAIL = {
  ten: 'Mặt Hàng', donViSi: 'Đơn Vị Sỉ', giaNhapSi: 'Giá Nhập Sỉ', soLuong: 'Số Lượng',
  donViLe: 'Đơn Vị Lẻ', loiNhuan: '% Lợi Nhuận', giaLamTron: 'Giá Làm Tròn'
};

/* Tab DoiChieu: khối cấu hình (Trường | Giá trị) và khối biệt danh (Cách viết trên toa | Mặt Hàng) */
var TD_TRUONG = 'Trường', TD_GIA_TRI = 'Giá trị', TD_GHI_CHU = 'Ghi chú';
var TD_VIET = 'Cách viết trên toa', TD_RETAIL = 'Mặt Hàng (Retail)';

/* Tên trường trong tab DoiChieu → khoá gửi cho trang web */
var TRUONG = [
  { khoa: 'lenhDocToa',  ten: 'Câu lệnh đọc toa',      kieu: 'chu' },
  { khoa: 'nguongNhay',  ten: 'Cảnh báo giá nhảy (%)', kieu: 'so' },
  { khoa: 'buocLamTron', ten: 'Bước làm tròn',         kieu: 'so' },
  { khoa: 'mauDo',       ten: 'Màu đổi giá bán',       kieu: 'mau' },
  { khoa: 'mauXanh',     ten: 'Màu chỉ đổi giá nhập',  kieu: 'mau' },
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
    if (viec === 'khoiTao') return traLoi(khoiTao());
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

/* Đọc tab DoiChieu. Thiếu gì thì ghi vào canhBao, trường thiếu nhận null (tính năng đó tự tắt ở trang web). */
function docDoiChieu(ss, canhBao) {
  var cauHinh = {};
  TRUONG.forEach(function (t) { cauHinh[t.khoa] = null; });
  var kq = { cauHinh: cauHinh, bietDanh: [] };

  var sh = ss.getSheetByName(TAB_DOICHIEU);
  if (!sh) { canhBao.push('Chưa có tab ' + TAB_DOICHIEU + ' — mở Apps Script, chạy hàm caiDat một lần'); return kq; }
  var k = khoiDoiChieu(sh);

  if (k.cotTruong < 0 || k.cotGiaTri < 0) canhBao.push('Tab ' + TAB_DOICHIEU + ' thiếu tiêu đề "' + TD_TRUONG + '" / "' + TD_GIA_TRI + '"');
  else {
    var theoTen = {};
    for (var i = 1; i < k.hang.length; i++) {
      var tr = chuanHoa(k.hang[i][k.cotTruong]);
      if (tr) theoTen[tr] = k.hang[i][k.cotGiaTri];
    }
    TRUONG.forEach(function (t) {
      var v = theoTen[chuanHoa(t.ten)];
      if (v === undefined) { canhBao.push('Tab ' + TAB_DOICHIEU + ' thiếu trường "' + t.ten + '" — mở Apps Script, chạy hàm caiDat một lần để thêm'); return; }
      cauHinh[t.khoa] = docGiaTri(v, t.kieu);
      if (cauHinh[t.khoa] == null && String(v).trim() !== '')
        canhBao.push('Trường "' + t.ten + '" có giá trị không hợp lệ: ' + v);
    });
  }

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
  return {
    hang: hang,
    cotTruong: timCot(tieuDe, TD_TRUONG), cotGiaTri: timCot(tieuDe, TD_GIA_TRI),
    cotViet: timCot(tieuDe, TD_VIET), cotRetail: timCot(tieuDe, TD_RETAIL)
  };
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
function doPost(e) {
  var khoa = LockService.getScriptLock();
  if (!khoa.tryLock(15000)) return traLoi({ ok: false, loi: 'Đang có một lần ghi khác, thử lại sau ít giây' });
  try {
    var d = JSON.parse(e.postData.contents);
    if (d.hanhDong === 'capNhat') return traLoi(capNhat(d));
    return traLoi({ ok: false, loi: 'Hành động không hợp lệ' });
  } catch (err) {
    return traLoi({ ok: false, loi: String(err) });
  } finally {
    khoa.releaseLock();
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
/* Dữ liệu điền sẵn khi tạo tab DoiChieu. Sau khi tạo, trang web chỉ đọc từ Sheet — sửa trong Sheet, không sửa ở đây. */
var CAU_HINH_BAN_DAU = [
  ['Câu lệnh đọc toa',
   'Đọc hóa đơn bán lẻ viết tay trong ảnh này. Trả về DUY NHẤT một mảng JSON, không kèm lời dẫn hay dấu ```.\n' +
   'Mỗi phần tử: {"ten": tên hàng đúng như trên toa, "dvt": đơn vị, "sl": số lượng, "donGia": đơn giá, "thanhTien": thành tiền}\n' +
   'Giữ nguyên con số như người viết ghi (216 thì để 216, đừng tự nhân nghìn).\n' +
   'Kiểm tra thành tiền = số lượng × đơn giá; nếu lệch thì đọc lại nét chữ cho khớp.\n' +
   'Dòng nào không đọc chắc chắn thì để "ten": "?" và vẫn giữ các số đọc được.',
   'Gửi kèm ảnh toa cho Claude'],
  ['Cảnh báo giá nhảy (%)', 30, 'Giá nhập đổi từ mức này trở lên → bỏ tích sẵn để xem lại'],
  ['Bước làm tròn', 1000, 'Chỉ để xem trước — công thức cột Giá Làm Tròn trong Retail mới là chuẩn'],
  ['Màu đổi giá bán', '#f4cccc', 'Tô dòng Retail khi Giá Làm Tròn đổi. Để trống = không tô'],
  ['Màu chỉ đổi giá nhập', '#d9ead3', 'Tô dòng Retail khi Giá Làm Tròn giữ nguyên. Để trống = không tô'],
  ['Xoá màu cũ khi ghi', 'Có', 'Có / Không — để màu chỉ phản ánh lần nhập gần nhất'],
  ['Link Đồng Bộ Giá', 'https://bayich2-dongbogia.vercel.app/', 'Nút hiện sau khi ghi xong. Để trống = ẩn nút'],
  ['Link Tính Giá', 'https://bayich2-tinhgia.vercel.app/', 'Link "Tính giá bán" ở mục Hàng mới. Để trống = ẩn link']
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
    Logger.log('Tab ' + TAB_DOICHIEU + ' đã có — giữ nguyên dữ liệu, chỉ thêm trường cấu hình còn thiếu');
    themTruongThieu(sh);
  } else {
    sh = ss.insertSheet(TAB_DOICHIEU, ss.getNumSheets());
    var bd = [];
    BIET_DANH_BAN_DAU.forEach(function (m) { m[1].forEach(function (v) { bd.push([v, m[0]]); }); });
    var soDong = Math.max(CAU_HINH_BAN_DAU.length, bd.length) + 1;
    var o = [];
    for (var i = 0; i < soDong; i++) {
      if (i === 0) { o.push([TD_TRUONG, TD_GIA_TRI, TD_GHI_CHU, '', TD_VIET, TD_RETAIL]); continue; }
      var c = CAU_HINH_BAN_DAU[i - 1] || ['', '', ''], b = bd[i - 1] || ['', ''];
      o.push([c[0], c[1], c[2], '', b[0], b[1]]);
    }
    sh.getRange(1, 1, soDong, 6).setValues(o);
    sh.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#2f5233').setFontColor('#f6f1e4');
    sh.getRange(1, 5, 1, 2).setFontWeight('bold').setBackground('#2f5233').setFontColor('#f6f1e4');
    sh.getRange(2, 2, CAU_HINH_BAN_DAU.length, 1).setWrap(true).setVerticalAlignment('top');
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 170); sh.setColumnWidth(2, 320); sh.setColumnWidth(3, 280);
    sh.setColumnWidth(4, 24);  sh.setColumnWidth(5, 170); sh.setColumnWidth(6, 190);
    Logger.log('Đã tạo tab ' + TAB_DOICHIEU + ': ' + CAU_HINH_BAN_DAU.length + ' trường cấu hình, ' + bd.length + ' biệt danh');
  }
  var kq = khoiTao();
  Logger.log(kq.ok ? ('Đọc thử: ' + kq.retail.length + ' mặt hàng Retail, ' + kq.bietDanh.length + ' biệt danh') : ('Lỗi: ' + kq.loi));
  (kq.canhBao || []).forEach(function (c) { Logger.log('  ! ' + c); });
}

/* Tab đã có (tạo bởi bản cũ): thêm các trường cấu hình mới còn thiếu vào ngay dưới khối cấu hình.
   Chỉ ghi vào các cột Trường | Giá trị | Ghi chú — không đụng dòng cũ, không đụng khối biệt danh. */
function themTruongThieu(sh) {
  var k = khoiDoiChieu(sh);
  if (k.cotTruong < 0 || k.cotGiaTri < 0) { Logger.log('Không thấy khối cấu hình — bỏ qua'); return; }
  var co = {}, dongCuoi = 1;
  for (var i = 1; i < k.hang.length; i++) {
    var tr = String(k.hang[i][k.cotTruong] || '').trim();
    if (tr) { co[chuanHoa(tr)] = true; dongCuoi = i + 1; }
  }
  var them = CAU_HINH_BAN_DAU.filter(function (c) { return !co[chuanHoa(c[0])]; });
  if (!them.length) { Logger.log('Khối cấu hình đủ trường'); return; }
  var cotGhiChu = timCot(k.hang[0], TD_GHI_CHU);
  them.forEach(function (c, j) {
    var dong = dongCuoi + 1 + j;
    sh.getRange(dong, k.cotTruong + 1).setValue(c[0]);
    sh.getRange(dong, k.cotGiaTri + 1).setValue(c[1]);
    if (cotGhiChu >= 0) sh.getRange(dong, cotGhiChu + 1).setValue(c[2]);
  });
  Logger.log('Đã thêm ' + them.length + ' trường: ' + them.map(function (c) { return c[0]; }).join(', '));
}

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
