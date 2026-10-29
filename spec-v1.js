/* Specification completion layer: dashboard, store settings, validated Excel and A4 slips. */
state.settings = {};
state.recentTab = 'IN';
state.chartItem = '';
state.importPlan = null;
state.loadedSlipKinds = { IN: false, OUT: false };
state.pendingSlipRequestId = '';

async function logout(call = true) {
  if (call && state.token) { try { await request('logout'); } catch {} }
  state.token = ''; state.user = null; state.items = []; state.inbound = []; state.outbound = [];
  state.users = []; state.dashboard = {}; state.settings = {}; state.loadedSlipKinds = { IN: false, OUT: false };
  sessionStorage.removeItem('vfs-token'); sessionStorage.removeItem('vfs-user');
  showLogin();
}

async function request(action, data = {}) {
  if (!apiUrl()) throw Error('Chưa cấu hình URL Apps Script. Mở Cài đặt để nhập URL Web App.');
  let response;
  try {
    response = await fetch(apiUrl(), { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ action, token: state.token, ...data }), redirect: 'follow' });
  } catch (e) {
    throw Error('Trình duyệt không đọc được phản hồi từ Apps Script (có thể do CORS/redirect). Kiểm tra URL và quyền triển khai. ' + e.message);
  }
  if (!response.ok) throw Error('Máy chủ trả HTTP ' + response.status);
  let result;
  try { result = await response.json(); } catch { throw Error('Phản hồi Apps Script không phải JSON. Kiểm tra URL Web App /exec.'); }
  if (!result.ok) {
    if (result.error?.code === 'UNAUTHORIZED') logout(false);
    const details = result.error?.details;
    throw Error((result.error?.message || result.error?.code || 'Yêu cầu thất bại') + (Array.isArray(details) ? '\n' + details.map(x => `Dòng ${x.row}: ${x.message}`).join('\n') : ''));
  }
  return result;
}

async function refresh() {
  const content = $('#content');
  if (content && !state.items.length) content.innerHTML = '<div class="loading-state">Đang tải dữ liệu kho…</div>';
  try {
    const snapshot = await request('getAppData');
    state.items = snapshot.data?.items || [];
    state.inbound = groupSlips(snapshot.data?.inbound || [], 'IN');
    state.outbound = groupSlips(snapshot.data?.outbound || [], 'OUT');
    state.loadedSlipKinds = { IN: false, OUT: false };
    state.dashboard = snapshot.data?.dashboard || {};
    state.settings = snapshot.data?.settings || {};
    state.users = [];
    if (state.user?.role === 'ADMIN') {
      try { state.users = (await request('listUsers')).data || []; }
      catch (userError) { notify('Không tải được danh sách tài khoản: ' + userError.message, true); }
    }
    if (state.page === 'inbound' || state.page === 'outbound') await ensureSlipData(state.page === 'inbound' ? 'IN' : 'OUT');
    render();
  } catch (e) {
    if (content && !state.items.length) content.innerHTML = `<div class="error-panel">${esc(e.message)}</div>`;
    notify(e.message, true);
  }
}

async function ensureSlipData(kind) {
  if (state.loadedSlipKinds[kind]) return;
  const response = await request('listSlips', { kind });
  const grouped = groupSlips(response.data || [], kind);
  if (kind === 'IN') state.inbound = grouped; else state.outbound = grouped;
  state.loadedSlipKinds[kind] = true;
}

function setPage(page) {
  state.page = page;
  render();
  if (page === 'inbound' || page === 'outbound') {
    ensureSlipData(page === 'inbound' ? 'IN' : 'OUT').then(render).catch(e => notify(e.message, true));
  }
}

function formatDateVi(value) {
  if (!value) return '—';
  const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(String(value));
  return m ? `${m[3]}/${m[2]}/${m[1]}` : String(value);
}

function chartForSelection() {
  const d = state.dashboard || {};
  if (!state.chartItem) return d.chart || [];
  const code = state.chartItem;
  const item = state.items.find(x => x.MA_HANG === code);
  if (!item) return d.chart || [];
  const inRows = state.inbound.flatMap(s => s.lines.map(l => ({ date: String(s.date).slice(0, 10), ...l })));
  const outRows = state.outbound.flatMap(s => s.lines.map(l => ({ date: String(s.date).slice(0, 10), ...l })));
  return (d.chart || []).map(point => {
    const date = point.date;
    const inboundQty = inRows.filter(x => x.code === code && x.date === date).reduce((a, x) => a + x.qty, 0);
    const outboundQty = outRows.filter(x => x.code === code && x.date === date).reduce((a, x) => a + x.qty, 0);
    const stockQty = Number(item.TON_DAU || 0)
      + inRows.filter(x => x.code === code && x.date <= date).reduce((a, x) => a + x.qty, 0)
      - outRows.filter(x => x.code === code && x.date <= date).reduce((a, x) => a + x.qty, 0);
    return { date, inboundQty, outboundQty, stockQty };
  });
}

function dashboardChart() {
  const data = chartForSelection();
  const w = 760, h = 230, left = 42, right = 18, top = 12, bottom = 34;
  const plotW = w - left - right, plotH = h - top - bottom;
  const maxFlow = Math.max(1, ...data.flatMap(x => [x.inboundQty, x.outboundQty]));
  const maxStock = Math.max(1, ...data.map(x => x.stockQty));
  const step = plotW / Math.max(data.length, 1), barW = Math.min(22, step * .24);
  const bars = data.map((p, i) => {
    const x = left + i * step + step / 2;
    const hi = p.inboundQty / maxFlow * plotH, ho = p.outboundQty / maxFlow * plotH;
    return `<rect x="${x - barW - 2}" y="${top + plotH - hi}" width="${barW}" height="${hi}" rx="3" fill="#087653"><title>Nhập ${num(p.inboundQty)}</title></rect><rect x="${x + 2}" y="${top + plotH - ho}" width="${barW}" height="${ho}" rx="3" fill="#247cf0"><title>Xuất ${num(p.outboundQty)}</title></rect><text x="${x}" y="${h - 9}" text-anchor="middle">${esc(String(p.date || '').slice(5).replace('-', '/'))}</text>`;
  }).join('');
  const points = data.map((p, i) => `${left + i * step + step / 2},${top + plotH - p.stockQty / maxStock * plotH}`).join(' ');
  return `<svg class="dash-chart" viewBox="0 0 ${w} ${h}" role="img" aria-label="Biểu đồ nhập xuất tồn 7 ngày"><line x1="${left}" y1="${top + plotH}" x2="${w - right}" y2="${top + plotH}" stroke="var(--line)"/><polyline points="${points}" fill="none" stroke="#f58220" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"/>${bars}<text x="7" y="18" class="axis-label">Nhập / xuất</text><text x="${w - 50}" y="18" class="axis-label">Tồn</text></svg>`;
}

function dashboardView() {
  const d = state.dashboard || {};
  const low = state.items.filter(x => (x.TRANG_THAI === 'Đang dùng' || x.TRANG_THAI === 'ACTIVE') && Number(x.TON_HIEN_TAI) <= Number(x.TON_TOI_THIEU));
  const stats = [
    ['Mặt hàng đang dùng', num(d.totalItems), 'catalog', 'green'],
    ['Phiếu nhập tháng này', num(d.inboundCount), 'inbound', 'blue'],
    ['Phiếu xuất tháng này', num(d.outboundCount), 'outbound', 'orange'],
    ['Mặt hàng tồn thấp', num(d.lowStockCount), 'stock', 'red']
  ];
  const top = (d.topItems || []).map((x, i) => `<tr><td>${i + 1}</td><td>${esc(x.MA_HANG)}</td><td>${esc(x.TEN_HANG)}</td><td>${num(x.TON_HIEN_TAI)}</td><td>${esc(x.DON_VI)}</td></tr>`);
  const recent = (state.recentTab === 'IN' ? d.recentInbound : d.recentOutbound) || [];
  const kindKey = state.recentTab === 'IN' ? 'inbound' : 'outbound';
  const recentRows = recent.map(x => {
    return `<tr><td>${esc(x.no)}</td><td>${formatDateVi(x.date)}</td><td>${esc(x.party || '—')}</td><td>${num(x.qty)}</td>${state.recentTab === 'IN' ? `<td>${money(x.total)}</td>` : ''}<td><button data-action="open-recent-slip" data-kind="${state.recentTab}" data-id="${esc(x.id)}" data-operation="view">Chi tiết</button><button data-action="open-recent-slip" data-kind="${state.recentTab}" data-id="${esc(x.id)}" data-operation="print">In</button></td></tr>`;
  });
  const lowRows = low.slice(0, 10).map(x => `<tr><td>${esc(x.MA_HANG)}</td><td>${esc(x.TEN_HANG)}</td><td>${num(x.TON_HIEN_TAI)}</td><td>${esc(x.DON_VI)}</td><td>${num(x.TON_TOI_THIEU)}</td></tr>`);
  return `<div class="grid stats">${stats.map(([label, value, page, color]) => `<button class="stat stat-${color}" data-go="${page}"><span class="muted">${label}</span><strong class="value">${value}</strong><small>Dữ liệu đã xác nhận từ máy chủ</small></button>`).join('')}</div>
    <div class="grid two dash-grid"><div class="panel"><div class="panel-title-row"><div><h2>Nhập – xuất – tồn 7 ngày gần nhất</h2><small class="muted">${state.chartItem ? `Mặt hàng ${esc(state.chartItem)}` : esc(d.chartNote || '')}</small></div><select id="chartItem" aria-label="Lọc biểu đồ theo mặt hàng"><option value="">Tất cả mặt hàng · chỉ báo số học</option>${state.items.filter(x => x.TRANG_THAI === 'Đang dùng').map(x => `<option value="${esc(x.MA_HANG)}" ${state.chartItem === x.MA_HANG ? 'selected' : ''}>${esc(x.MA_HANG)} · ${esc(x.TEN_HANG)}</option>`).join('')}</select></div><div class="chart-legend"><span><i class="legend-in"></i>Nhập</span><span><i class="legend-out"></i>Xuất</span><span><i class="legend-stock"></i>Tồn cuối ngày</span></div>${dashboardChart()}</div>
    <div class="panel"><div class="panel-title-row"><h2>Top 5 mặt hàng tồn nhiều nhất</h2><button data-go="stock">Xem tồn kho</button></div><p class="footnote">Xếp hạng theo số lượng; đơn vị các mặt hàng có thể khác nhau.</p>${table(['#','Mã hàng','Tên hàng','Tồn kho','ĐVT'], top)}</div></div>
    <div class="panel dashboard-recent"><div class="panel-title-row"><div><h2>Phiếu gần đây</h2><div class="tabs"><button data-action="recent-tab" data-kind="IN" class="${state.recentTab === 'IN' ? 'active' : ''}">Phiếu nhập</button><button data-action="recent-tab" data-kind="OUT" class="${state.recentTab === 'OUT' ? 'active' : ''}">Phiếu xuất</button></div></div><button data-go="${kindKey}">Xem tất cả</button></div>${table(['Số phiếu','Ngày','Nhà cung cấp / Nơi nhận','Tổng số lượng',state.recentTab === 'IN' ? 'Tổng tiền' : '', 'Thao tác'].filter(Boolean), recentRows)}</div>
    <div class="panel"><div class="panel-title-row"><h2>Hàng tồn thấp</h2><button data-go="stock" data-low="1">Xem tất cả (${low.length})</button></div>${table(['Mã hàng','Tên hàng','Tồn kho','ĐVT','Tồn tối thiểu'], lowRows)}</div>`;
}

function render() {
  document.body.classList.toggle('dark', state.theme === 'dark');
  $('#title').textContent = pages[state.page] || 'Tổng quan';
  $('#who').textContent = `${state.user?.fullName || state.user?.username || ''} · ${state.user?.role || ''}`;
  updateProfile();
  const todayPill = $('#today-pill');
  if (todayPill) todayPill.textContent = new Date().toLocaleDateString('vi-VN', { weekday: 'long', day: '2-digit', month: '2-digit', year: 'numeric' });
  document.querySelectorAll('nav [data-page]').forEach(b => b.classList.toggle('active', b.dataset.page === state.page));
  let content = '';
  const low = state.items.filter(x => x.TRANG_THAI === 'Đang dùng' && Number(x.TON_HIEN_TAI) <= Number(x.TON_TOI_THIEU));
  if (state.page === 'dashboard') content = dashboardView();
  if (state.page === 'catalog' || state.page === 'stock') {
    const isCatalog = state.page === 'catalog';
    let items = state.items.filter(x => `${x.MA_HANG} ${x.TEN_HANG}`.toLowerCase().includes(state.query.toLowerCase()));
    if (state.lowOnly) items = items.filter(x => Number(x.TON_HIEN_TAI) <= Number(x.TON_TOI_THIEU));
    const rows = items.map(x => isCatalog
      ? `<tr><td>${esc(x.MA_HANG)}</td><td>${esc(x.TEN_HANG)}</td><td>${esc(x.DON_VI)}</td><td>${num(x.TON_DAU)}</td><td>${num(x.TON_HIEN_TAI)}</td><td>${num(x.TON_TOI_THIEU)}</td><td>${money(x.GIA_MUA)}</td><td><span class="tag ${x.TRANG_THAI === 'Ngừng dùng' ? 'warn' : ''}">${esc(x.TRANG_THAI)}</span></td><td class="actions">${state.user?.role === 'ADMIN' ? `<button data-action="edit-item" data-code="${esc(x.MA_HANG)}">Sửa</button><button data-action="toggle-item" data-code="${esc(x.MA_HANG)}">${x.TRANG_THAI === 'Ngừng dùng' ? 'Dùng lại' : 'Ngừng dùng'}</button>` : ''}</td></tr>`
      : `<tr><td>${esc(x.MA_HANG)}</td><td>${esc(x.TEN_HANG)}</td><td>${esc(x.DON_VI)}</td><td>${num(x.TON_DAU)}</td><td>${num(x.TONG_NHAP)}</td><td>${num(x.TONG_XUAT)}</td><td>${num(x.TON_HIEN_TAI)}</td><td>${num(x.TON_TOI_THIEU)}</td><td>${money(x.GIA_MUA)}</td><td>${money(x.GIA_TRI_TON)}</td><td>${esc(x.TRANG_THAI)}</td></tr>`);
    content = `${toolbar('Tìm mã hoặc tên hàng', `${isCatalog && state.user?.role === 'ADMIN' ? '<button class="primary" data-action="new-item">+ Thêm mặt hàng</button>' : ''}<button data-action="toggle-low">${state.lowOnly ? 'Hiện tất cả' : 'Chỉ tồn thấp'}</button>`)}${panel(isCatalog ? 'Danh mục hàng hóa' : 'Tồn kho hiện tại', table(isCatalog ? ['Mã hàng','Tên hàng','Đơn vị','Tồn đầu','Tồn hiện tại','Tối thiểu','Giá mua','Trạng thái','Thao tác'] : ['Mã hàng','Tên hàng','ĐVT','Tồn đầu','Tổng nhập','Tổng xuất','Tồn hiện tại','Tối thiểu','Giá mua','Giá trị tồn','Trạng thái'], rows))}`;
  }
  if (state.page === 'inbound' || state.page === 'outbound') {
    const k = state.page, slips = state[k].filter(x => `${x.no} ${x.party}`.toLowerCase().includes(state.query.toLowerCase()));
    const rows = slips.map(x => `<tr><td>${esc(x.no)}</td><td>${formatDateVi(x.date)}</td><td>${esc(x.party || '—')}</td><td>${x.lines.length} dòng · ${num(x.qty)}</td><td>${esc(x.creator)}</td><td>${k === 'inbound' ? money(x.total) : '—'}</td><td class="actions"><button data-action="view-slip" data-kind="${k}" data-id="${esc(x.id)}">Xem</button><button data-action="edit-slip" data-kind="${k}" data-id="${esc(x.id)}">Sửa</button><button class="delete" data-action="delete-slip" data-kind="${k}" data-id="${esc(x.id)}">Xóa</button><button data-action="print-slip" data-kind="${k}" data-id="${esc(x.id)}">In</button></td></tr>`);
    content = `${toolbar('Tìm số phiếu hoặc đối tượng', `<button class="primary" data-action="new-slip" data-kind="${k}">+ Lập phiếu ${k === 'inbound' ? 'nhập' : 'xuất'}</button>`)}${panel('Danh sách phiếu', state.loadedSlipKinds[k === 'inbound' ? 'IN' : 'OUT'] ? table(['Số phiếu','Ngày','Nhà cung cấp / Nơi nhận','Dòng hàng / SL','Người lập','Tổng tiền nhập','Thao tác'], rows) : '<div class="loading-state">Đang tải phiếu…</div>')}`;
  }
  if (state.page === 'accounts') {
    content = state.user?.role !== 'ADMIN' ? panel('Quản lý tài khoản', 'Chức năng này dành cho ADMIN.') : `${toolbar('Tìm tài khoản', '<button class="primary" data-action="new-user">+ Tạo tài khoản User</button>')}${panel('Tài khoản đơn vị', table(['Tên đăng nhập','Họ tên','Vai trò','Trạng thái','Thao tác'], state.users.filter(u => `${u.USERNAME} ${u.HO_TEN}`.toLowerCase().includes(state.query.toLowerCase())).map(u => `<tr><td>${esc(u.USERNAME)}</td><td>${esc(u.HO_TEN)}</td><td>${esc(u.VAI_TRO)}</td><td>${esc(u.TRANG_THAI)}</td><td>${u.VAI_TRO === 'USER' ? `<button data-action="user-status" data-user="${esc(u.USERNAME)}" data-status="${String(u.TRANG_THAI).toUpperCase() === 'ACTIVE' ? 'LOCKED' : 'ACTIVE'}">${String(u.TRANG_THAI).toUpperCase() === 'ACTIVE' ? 'Khóa' : 'Mở khóa'}</button> <button data-action="delete-user" data-user="${esc(u.USERNAME)}">Xóa</button>` : ''}</td></tr>`)))}`;
  }
  if (state.page === 'excel') content = excelPanel();
  if (state.page === 'settings') content = settingsPanel();
  $('#content').innerHTML = content;
}

function excelPanel() {
  const plan = state.importPlan;
  const errors = plan?.errors?.length ? `<div class="import-errors"><b>Chưa thể nhập dữ liệu (${plan.errors.length} lỗi):</b><ul>${plan.errors.map(x => `<li>Dòng ${x.row}: ${esc(x.message)}</li>`).join('')}</ul></div>` : '';
  const preview = plan ? `<p>Đã kiểm tra ${plan.rows.length} dòng. ${plan.validCount} dòng hợp lệ; ${plan.errors.length} dòng lỗi.</p>${errors}${plan.kind === 'items' && plan.existingCount ? `<label class="checkline"><input id="updateExisting" type="checkbox"> Cập nhật mặt hàng trùng mã sau khi xác nhận (giữ nguyên tồn đầu)</label>` : ''}${table(plan.kind === 'items' ? ['Mã hàng','Tên hàng','Đơn vị','Tồn đầu','Tồn tối thiểu','Giá mua','Trạng thái','Kết quả'] : ['Số phiếu','Ngày','Mã hàng','Tên hàng','Số lượng','Đơn giá / Đối tượng','Ghi chú','Kết quả'], plan.rows.slice(0, 40).map(x => `<tr class="${x._error ? 'row-error' : ''}">${plan.kind === 'items' ? `<td>${esc(x.MA_HANG)}</td><td>${esc(x.TEN_HANG)}</td><td>${esc(x.DON_VI)}</td><td>${esc(x.TON_DAU)}</td><td>${esc(x.TON_TOI_THIEU)}</td><td>${esc(x.GIA_MUA)}</td><td>${esc(x.TRANG_THAI)}</td>` : `<td>${esc(x.SO_PHIEU)}</td><td>${esc(x.NGAY)}</td><td>${esc(x.MA_HANG)}</td><td>${esc(x.TEN_HANG)}</td><td>${esc(x.SO_LUONG)}</td><td>${esc(x.DON_GIA ?? x.DOI_TUONG)}</td><td>${esc(x.GHI_CHU)}</td>`}<td>${esc(x._error || (x._exists ? 'Mã đã có' : 'Hợp lệ'))}</td></tr>`))}${plan.rows.length > 40 ? `<p class="footnote">Đang hiển thị 40/${plan.rows.length} dòng đầu.</p>` : ''}` : '<p class="muted">Chọn tệp rồi xem trước. Ứng dụng chưa ghi dữ liệu cho tới khi anh xác nhận.</p>';
  return `<div class="panel"><h2>Nhập / Xuất Excel</h2><p>Nhập danh mục hoặc phiếu nhập/xuất theo lô. Bản xem trước chỉ ra lỗi theo dòng; dữ liệu chỉ được gửi sau khi xác nhận.</p><div class="formgrid"><label>Loại dữ liệu<select id="importKind"><option value="items">Danh mục hàng hóa</option><option value="IN">Phiếu nhập</option><option value="OUT">Phiếu xuất</option></select></label><label>Tệp Excel<input type="file" id="importFile" accept=".xlsx,.xls,.csv"></label></div><p class="row-actions"><button data-action="preview-import">Kiểm tra và xem trước</button><button class="primary" data-action="commit-import" ${plan?.errors?.length || !plan ? 'disabled' : ''}>Xác nhận nhập</button><button data-action="download-template">Tải mẫu danh mục 20 dòng</button></p><div id="importPreview">${preview}</div><hr><label>Loại dữ liệu xuất<select id="exportKind"><option value="items">Danh mục và tồn kho</option><option value="IN">Phiếu nhập</option><option value="OUT">Phiếu xuất</option></select></label> <button data-action="export-v1">Tải Excel</button><p class="footnote">Tệp xuất chỉ có dữ liệu nghiệp vụ; không bao gồm mật khẩu, mã băm, token hay bí mật.</p></div>`;
}

function settingsPanel() {
  const admin = state.user?.role === 'ADMIN';
  const tabs = [['connection','Kết nối'], ...(admin ? [['store','Thông tin cửa hàng']] : []), ['security','Bảo mật'], ['appearance','Giao diện'], ['password','Đổi mật khẩu']];
  let body = '';
  if (state.settingsTab === 'connection') body = `<div class="settings-table"><div class="settings-row"><label for="settingsUrl">Apps Script Web App URL</label><div class="settings-control"><input class="field" id="settingsUrl" value="${esc(apiUrl())}" placeholder="https://script.google.com/macros/s/.../exec"><small>URL chỉ lưu trong trình duyệt đang dùng.</small></div></div></div><div class="settings-actions"><button class="primary" data-action="save-url">Lưu URL</button></div>`;
  if (state.settingsTab === 'store' && admin) body = `<form id="storeSettingsForm" class="settings-table"><div class="settings-row"><label for="storeName">Tên cửa hàng / đơn vị <span class="required">*</span></label><div class="settings-control"><input id="storeName" name="storeName" required maxlength="120" value="${esc(state.settings.storeName)}"></div></div><div class="settings-row"><label for="storeAddress">Địa chỉ</label><div class="settings-control"><input id="storeAddress" name="address" maxlength="200" value="${esc(state.settings.address)}"></div></div><div class="settings-row"><label for="storePhone">Điện thoại</label><div class="settings-control"><input id="storePhone" name="phone" maxlength="40" value="${esc(state.settings.phone)}"></div></div><div class="settings-row"><label for="storeTaxCode">Mã số thuế</label><div class="settings-control"><input id="storeTaxCode" name="taxCode" maxlength="40" value="${esc(state.settings.taxCode)}"></div></div><div class="settings-row"><label for="storeFooterNote">Ghi chú chân trang</label><div class="settings-control"><input id="storeFooterNote" name="footerNote" maxlength="200" value="${esc(state.settings.footerNote)}"></div></div></form><div class="settings-actions"><button class="primary" data-action="save-store-settings">Lưu thông tin cửa hàng</button></div>`;
  if (state.settingsTab === 'security') body = `<div class="settings-table"><div class="settings-row"><b>Phiên đăng nhập</b><div class="settings-description">Phiên và quyền thao tác được kiểm tra tại Apps Script.</div></div><div class="settings-row"><b>Bảo vệ tài khoản</b><div class="settings-description">Không chia sẻ mật khẩu hoặc mã phiên với người khác.</div></div></div>`;
  if (state.settingsTab === 'appearance') body = `<div class="settings-table"><div class="settings-row"><label>Giao diện cá nhân</label><div class="settings-control"><button data-action="theme">${state.theme === 'dark' ? 'Chuyển sang sáng' : 'Chuyển sang tối'}</button><small>Áp dụng cho trình duyệt đang sử dụng.</small></div></div></div>`;
  if (state.settingsTab === 'password') body = `<div class="settings-table"><div class="settings-row"><div><b>Mật khẩu tài khoản</b><small class="settings-hint">Đổi mật khẩu đăng nhập hiện tại.</small></div><div class="settings-control"><button class="primary" data-action="change-password">Đổi mật khẩu</button></div></div></div>`;
  return panel('Cài đặt', `<div class="tabs">${tabs.map(([id, label]) => `<button data-setting="${id}" class="${state.settingsTab === id ? 'active' : ''}">${label}</button>`).join('')}</div>${body}`);
}

function lineHTML(line, kind) {
  const items = state.items.filter(x => x.TRANG_THAI === 'Đang dùng' || x.TRANG_THAI === 'ACTIVE');
  const found = items.find(x => x.MA_HANG === line.code);
  const price = line.price !== undefined ? line.price : Number(found?.GIA_MUA || 0);
  return `<tr><td><select name="code" required><option value="">Chọn hàng</option>${items.map(x => `<option value="${esc(x.MA_HANG)}" ${line.code === x.MA_HANG ? 'selected' : ''}>${esc(x.MA_HANG)} — ${esc(x.TEN_HANG)}</option>`).join('')}</select></td><td><input name="qty" type="number" min="0.000001" step="any" required value="${Number(line.qty) || 1}"></td>${kind === 'inbound' ? `<td><input name="price" type="number" min="0" step="any" value="${Number(price) || 0}"></td>` : ''}<td><button type="button" data-action="remove-line">Xóa</button></td></tr>`;
}

function slipForm(kind, slip) {
  state.pendingSlipRequestId = slip ? '' : crypto.randomUUID().replaceAll('-', '');
  const lines = slip?.lines?.length ? slip.lines : [{ code: '', qty: 1 }];
  const fields = `<form id="slipForm"><div class="formgrid"><label>Ngày<input name="date" type="date" required value="${esc(String(slip?.date || today()).slice(0, 10))}"></label><label class="span2">${kind === 'inbound' ? 'Nhà cung cấp' : 'Nơi nhận / Người nhận'}<input name="counterparty" required value="${esc(slip?.party || '')}"></label><label class="span2">Ghi chú<input name="note" value="${esc(slip?.note || '')}"></label></div><div class="tablewrap"><table><thead><tr><th>Mặt hàng</th><th>Số lượng</th>${kind === 'inbound' ? '<th>Đơn giá</th>' : ''}<th></th></tr></thead><tbody id="lineRows">${lines.map(x => lineHTML(x, kind)).join('')}</tbody></table></div><p id="slipTotal" class="slip-total"></p><button type="button" data-action="add-line" data-kind="${kind}">+ Thêm dòng</button></form>`;
  modal(slip ? `Sửa phiếu ${esc(slip.no)}` : `Lập phiếu ${kind === 'inbound' ? 'nhập kho' : 'xuất kho'}`, fields, `<button data-action="close">Hủy</button><button class="primary" data-action="save-slip-v1" data-kind="${kind}" data-id="${esc(slip?.id || '')}">Lưu phiếu</button>`);
  updateSlipTotal();
}

function updateSlipTotal() {
  const rows = [...($('#lineRows')?.rows || [])];
  const qty = rows.reduce((sum, row) => sum + (Number($('[name=qty]', row)?.value) || 0), 0);
  const total = rows.reduce((sum, row) => sum + (Number($('[name=qty]', row)?.value) || 0) * (Number($('[name=price]', row)?.value) || 0), 0);
  const output = $('#slipTotal');
  if (output) output.textContent = `Tổng số lượng: ${num(qty)}${$('[name=price]') ? ` · Tổng tiền: ${money(total)}` : ''}`;
}

function printSlip(kind, id) {
  const slip = currentSlip(kind, id);
  if (!slip) return notify('Không tìm thấy phiếu đã lưu.', true);
  const store = state.settings || {};
  if (!store.storeName) { notify('Admin cần nhập tên cửa hàng trong Cài đặt trước khi in.', true); setPage('settings'); state.settingsTab = 'store'; render(); return; }
  const inbound = kind === 'inbound';
  const bodyRows = slip.lines.map((line, i) => `<tr><td>${i + 1}</td><td>${esc(line.code)}</td><td>${esc(line.name)}</td><td>${esc(line.unit)}</td><td class="num">${num(line.qty)}</td>${inbound ? `<td class="num">${money(line.price)}</td><td class="num">${money(line.qty * line.price)}</td>` : ''}</tr>`).join('');
  const qty = slip.lines.reduce((sum, x) => sum + Number(x.qty || 0), 0);
  const amount = slip.lines.reduce((sum, x) => sum + Number(x.qty || 0) * Number(x.price || 0), 0);
  const w = window.open('', '_blank');
  if (!w) return notify('Trình duyệt đã chặn cửa sổ in.', true);
  w.document.write(`<!doctype html><html lang="vi"><head><meta charset="utf-8"><title>${esc(slip.no)}</title><style>
    @page{size:A4 portrait;margin:16mm}*{box-sizing:border-box}body{font:13px/1.45 Arial,"Segoe UI",sans-serif;color:#172a24;margin:0}.store{text-align:center;border-bottom:2px solid #087653;padding-bottom:10px;margin-bottom:16px}.store h2{font-size:18px;margin:0 0 4px;color:#075744}.store p{margin:1px}.title{text-align:center;font-size:21px;margin:18px 0 14px}.meta{display:grid;grid-template-columns:1fr 1fr;gap:4px 20px;margin:10px 0}table{width:100%;border-collapse:collapse;margin:12px 0}thead{display:table-header-group}th,td{border:1px solid #444;padding:6px 7px;vertical-align:top}th{text-align:center;background:#edf5f1}td.num{text-align:right}.total{text-align:right;font-weight:bold}.note{min-height:28px}.signatures{display:grid;grid-template-columns:repeat(4,1fr);gap:8px;text-align:center;margin-top:42px;page-break-inside:avoid}.signatures b{display:block}.signatures small{display:block;font-style:italic;margin-top:3px}.signature-space{height:70px}tr{page-break-inside:avoid}.footer{text-align:center;margin-top:24px;color:#52675d}@media print{button{display:none}}
  </style></head><body><header class="store"><h2>${esc(store.storeName)}</h2>${store.address ? `<p>${esc(store.address)}</p>` : ''}${store.phone ? `<p>Điện thoại: ${esc(store.phone)}</p>` : ''}${store.taxCode ? `<p>Mã số thuế: ${esc(store.taxCode)}</p>` : ''}</header><h1 class="title">${inbound ? 'PHIẾU NHẬP KHO' : 'PHIẾU XUẤT KHO'}</h1><div class="meta"><div><b>Số phiếu:</b> ${esc(slip.no)}</div><div><b>Ngày:</b> ${formatDateVi(slip.date)}</div>${inbound ? `<div><b>Nhà cung cấp:</b> ${esc(slip.party || '')}</div>` : `<div><b>Nơi nhận:</b> ${esc(slip.party || '')}</div>`}<div><b>Người lập:</b> ${esc(slip.creator || '')}</div></div><table><thead><tr><th>STT</th><th>Mã hàng</th><th>Tên hàng</th><th>Đơn vị</th><th>Số lượng</th>${inbound ? '<th>Đơn giá</th><th>Thành tiền</th>' : ''}</tr></thead><tbody>${bodyRows}</tbody></table><p class="total">Tổng số lượng: ${num(qty)}</p>${inbound ? `<p class="total">Tổng tiền: ${money(amount)}</p>` : ''}<p class="note"><b>Ghi chú:</b> ${esc(slip.note || '')}</p><div class="signatures"><div><b>Người lập phiếu</b><small>(Ký, họ tên)</small><div class="signature-space"></div>${esc(slip.creator || '')}</div><div><b>${inbound ? 'Người giao hàng' : 'Người nhận hàng'}</b><small>(Ký, họ tên)</small><div class="signature-space"></div></div><div><b>Thủ kho</b><small>(Ký, họ tên)</small><div class="signature-space"></div></div><div><b>Người duyệt</b><small>(Ký, họ tên)</small><div class="signature-space"></div></div></div>${store.footerNote ? `<p class="footer">${esc(store.footerNote)}</p>` : ''}<script>window.addEventListener('load',()=>window.print())<\/script></body></html>`);
  w.document.close();
}

function excelHeaderKey(value) {
  let key = String(value || '').trim().toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/Đ/g, 'D').replace(/[^A-Z0-9]+/g, '_').replace(/^_|_$/g, '');
  const aliases = { MA_HANG_HOA: 'MA_HANG', SO_PHIEU_NHAP: 'SO_PHIEU', SO_PHIEU_XUAT: 'SO_PHIEU', NGAY_NHAP_KHO: 'NGAY', NGAY_XUAT_KHO: 'NGAY', NGAY_NHAP: 'NGAY', NGAY_XUAT: 'NGAY', DON_VI_TINH: 'DON_VI', MUC_TON_TOI_THIEU: 'TON_TOI_THIEU', DOI_TUONG: 'DOI_TUONG', NOI_NHAN: 'DOI_TUONG', NGUOI_NHAN: 'DOI_TUONG', NHA_CUNG_CAP: 'DOI_TUONG' };
  return aliases[key] || key;
}

function validatedExcelDate(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return `${value.getFullYear()}-${String(value.getMonth() + 1).padStart(2, '0')}-${String(value.getDate()).padStart(2, '0')}`;
  }
  if (typeof value === 'number' && Number.isFinite(value) && value > 0 && window.XLSX) {
    const result = XLSX.SSF.format('yyyy-mm-dd', value);
    return /^\d{4}-\d{2}-\d{2}$/.test(result) ? result : null;
  }
  const text = String(value || '').trim();
  let m = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(text);
  if (!m) { const dmy = /^(\d{1,2})[/.\-](\d{1,2})[/.\-](\d{4})$/.exec(text); if (dmy) m = [dmy[0], dmy[3], dmy[2], dmy[1]]; }
  if (!m) return null;
  const year = Number(m[1]), month = Number(m[2]), day = Number(m[3]);
  const check = new Date(Date.UTC(year, month - 1, day));
  if (check.getUTCFullYear() !== year || check.getUTCMonth() !== month - 1 || check.getUTCDate() !== day) return null;
  return `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
}

async function readImportRows() {
  if (!window.XLSX) throw Error('Thư viện Excel chưa tải. Hãy tải lại trang khi có mạng.');
  const file = $('#importFile')?.files?.[0];
  if (!file) throw Error('Anh hãy chọn tệp Excel trước.');
  const wb = XLSX.read(await file.arrayBuffer(), { type: 'array', cellDates: true });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(sheet, { defval: '' }).map((raw, i) => {
    const row = { _row: i + 2 };
    Object.entries(raw).forEach(([key, value]) => { row[excelHeaderKey(key)] = value; });
    return row;
  }).filter(r => Object.keys(r).some(k => !k.startsWith('_') && r[k] !== ''));
}

function validateImportRows(rows, kind) {
  const errors = [], seen = {}, itemMap = new Map(state.items.map(x => [String(x.MA_HANG).toLowerCase(), x]));
  const normalized = rows.map((r, i) => {
    const rowErrors = [];
    if (kind === 'items') {
      ['MA_HANG','TEN_HANG','DON_VI'].forEach(k => { if (!String(r[k] || '').trim()) rowErrors.push(`thiếu ${k}`); });
      ['TON_DAU','GIA_MUA','TON_TOI_THIEU'].forEach(k => { if (r[k] === '' || !Number.isFinite(Number(r[k])) || Number(r[k]) < 0) rowErrors.push(`${k} phải là số không âm`); });
      if (!['Đang dùng','Ngừng dùng'].includes(String(r.TRANG_THAI || '').trim())) rowErrors.push('TRANG_THAI phải là Đang dùng hoặc Ngừng dùng');
      const key = String(r.MA_HANG || '').trim().toLowerCase();
      if (key && seen[key]) rowErrors.push('mã hàng bị trùng trong tệp');
      if (key) seen[key] = true;
      r._exists = itemMap.has(key);
    } else {
      if (!validatedExcelDate(r.NGAY)) rowErrors.push('thiếu ngày hợp lệ');
      if (!String(r.DOI_TUONG || '').trim()) rowErrors.push('thiếu nhà cung cấp/nơi nhận');
      if (!String(r.MA_HANG || '').trim()) rowErrors.push('thiếu mã hàng');
      if (!Number.isFinite(Number(r.SO_LUONG)) || Number(r.SO_LUONG) <= 0) rowErrors.push('số lượng phải lớn hơn 0');
      if (kind === 'IN' && r.DON_GIA !== '' && (!Number.isFinite(Number(r.DON_GIA)) || Number(r.DON_GIA) < 0)) rowErrors.push('đơn giá phải là số không âm');
      const item = itemMap.get(String(r.MA_HANG || '').trim().toLowerCase());
      if (!item) rowErrors.push('mã hàng chưa có trong danh mục');
      else if (item.TRANG_THAI !== 'Đang dùng' && item.TRANG_THAI !== 'ACTIVE') rowErrors.push('mặt hàng đã ngừng dùng');
    }
    r._error = rowErrors.join('; ');
    if (rowErrors.length) errors.push({ row: r._row || i + 2, message: r._error });
    return r;
  });
  const existingCount = kind === 'items' ? normalized.filter(x => x._exists).length : 0;
  return { kind: kind === 'items' ? 'items' : kind, rows: normalized, errors, validCount: normalized.length - errors.length, existingCount };
}

function groupImportedSlips(rows, kind) {
  const groups = new Map();
  rows.forEach((r, i) => {
    const key = String(r.SO_PHIEU || r.EXTERNAL_SLIP_NO || `DONG-${i + 1}`).trim();
    if (!groups.has(key)) groups.set(key, { date: validatedExcelDate(r.NGAY), counterparty: String(r.DOI_TUONG || '').trim(), note: String(r.GHI_CHU || ''), externalSlipNo: key, lines: [] });
    const line = { MA_HANG: String(r.MA_HANG || '').trim(), SO_LUONG: Number(r.SO_LUONG) };
    if (kind === 'IN' && r.DON_GIA !== '') line.DON_GIA = Number(r.DON_GIA);
    groups.get(key).lines.push(line);
  });
  return [...groups.values()];
}

function downloadItemTemplate() {
  if (!window.XLSX) return notify('Thư viện Excel chưa tải.', true);
  const names = ['Giấy A4','Bút bi xanh','Bút bi đỏ','Bút chì HB','Tẩy trắng','Bìa hồ sơ','Kẹp giấy nhỏ','Kim bấm số 10','Băng keo trong','Kéo văn phòng','Dao rọc giấy','Mực in đen','Mực in màu','Sổ tay A5','Bìa lá','Phong bì A5','Phong bì A4','Khăn giấy','Nước rửa tay','Pin AA'];
  const units = ['Ram','Cây','Cây','Cây','Cục','Cái','Hộp','Hộp','Cuộn','Cái','Cái','Hộp','Hộp','Cuốn','Cái','Cái','Cái','Gói','Chai','Vỉ'];
  const rows = names.map((name, i) => ({ 'Mã hàng': `HH${String(i + 1).padStart(3, '0')}`, 'Tên hàng': name, 'Đơn vị': units[i], 'Tồn đầu': 0, 'Giá mua': 0, 'Tồn tối thiểu': 0, 'Trạng thái': 'Đang dùng' }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), 'Danh mục');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['DỮ LIỆU MẪU — thay bằng dữ liệu thật trước khi nhập. Tệp này không tự ghi vào hệ thống.'], ['Cột bắt buộc: Mã hàng, Tên hàng, Đơn vị, Tồn đầu, Giá mua, Tồn tối thiểu, Trạng thái.']]), 'Hướng dẫn');
  XLSX.writeFile(wb, `VIFUSOFT_Mau_Danh_Muc_20_dong.xlsx`);
}

function excelDateValue(value) {
  const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(String(value || ''));
  return m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])) : String(value || '');
}

async function handleSpecAction(action, button) {
  if (action === 'recent-tab') { state.recentTab = button.dataset.kind; render(); return true; }
  if (action === 'open-recent-slip') {
    const kind = button.dataset.kind, page = kind === 'IN' ? 'inbound' : 'outbound';
    await ensureSlipData(kind);
    const slip = currentSlip(page, button.dataset.id);
    if (!slip) throw Error('Không tìm thấy phiếu đã lưu.');
    if (button.dataset.operation === 'print') printSlip(page, slip.id);
    else modal(`Phiếu ${esc(slip.no)}`, `<p>${esc(slip.party || '')} · ${formatDateVi(slip.date)} · Người lập: ${esc(slip.creator || '')}</p>${table(['Mã hàng','Tên hàng','ĐVT','Số lượng',kind === 'IN' ? 'Đơn giá' : ''].filter(Boolean), slip.lines.map(x => `<tr><td>${esc(x.code)}</td><td>${esc(x.name)}</td><td>${esc(x.unit)}</td><td>${num(x.qty)}</td>${kind === 'IN' ? `<td>${money(x.price)}</td>` : ''}</tr>`))}`, '<button data-action="close">Đóng</button>');
    return true;
  }
  if (action === 'save-slip-v1') {
    const form = $('#slipForm'); if (!form.reportValidity()) return true;
    const lines = [...$('#lineRows').rows].map(row => ({ MA_HANG: $('[name="code"]', row).value,
      SO_LUONG: $('[name="qty"]', row).value,
      ...(button.dataset.kind === 'inbound' ? { DON_GIA: $('[name="price"]', row).value } : {}) }));
    if (!lines.length || lines.some(x => !x.MA_HANG || Number(x.SO_LUONG) <= 0)) throw Error('Phiếu cần có dòng hàng hợp lệ.');
    const data = Object.fromEntries(new FormData(form));
    const kind = button.dataset.kind === 'inbound' ? 'IN' : 'OUT';
    const slip = { date: data.date, counterparty: data.counterparty, note: data.note, lines };
    const originalText = button.textContent; button.disabled = true; button.textContent = 'Đang lưu…';
    try {
      if (button.dataset.id) await request('updateSlip', { kind, slipId: button.dataset.id, slip });
      else await request('createSlip', { kind, requestId: state.pendingSlipRequestId, slip });
      state.pendingSlipRequestId = ''; $('#modal').innerHTML = ''; await refresh(); notify('Phiếu đã được máy chủ xác nhận và lưu.');
    } finally { button.disabled = false; button.textContent = originalText; }
    return true;
  }
  if (action === 'save-store-settings') {
    const form = $('#storeSettingsForm'); if (!form.reportValidity()) return true;
    const result = await request('updateStoreSettings', { settings: Object.fromEntries(new FormData(form)) });
    state.settings = result.data || {}; render(); notify('Đã lưu thông tin cửa hàng.'); return true;
  }
  if (action === 'preview-import') {
    const kind = $('#importKind').value;
    const rows = await readImportRows();
    if (!rows.length) throw Error('Tệp không có dòng dữ liệu.');
    state.importPlan = validateImportRows(rows, kind); render();
    notify(state.importPlan.errors.length ? `Có ${state.importPlan.errors.length} dòng cần sửa.` : `Đã kiểm tra ${rows.length} dòng.` , !!state.importPlan.errors.length); return true;
  }
  if (action === 'commit-import') {
    const plan = state.importPlan;
    if (!plan || plan.errors.length) throw Error('Cần kiểm tra và sửa hết lỗi trước khi nhập.');
    if (plan.kind === 'items') {
      const updateExisting = !!$('#updateExisting')?.checked;
      if (plan.existingCount && !updateExisting) throw Error('Bỏ chọn mã hàng trùng hoặc đánh dấu xác nhận cập nhật mã đã có.');
      if (!confirm(`Xác nhận nhập ${plan.rows.length} mặt hàng?${plan.existingCount ? ` Sẽ cập nhật ${plan.existingCount} mã trùng; tồn đầu hiện tại được giữ nguyên.` : ''}`)) return true;
      const items = plan.rows.map(r => ({ MA_HANG: String(r.MA_HANG).trim(), TEN_HANG: String(r.TEN_HANG).trim(), DON_VI: String(r.DON_VI).trim(), TON_DAU: Number(r.TON_DAU), TON_TOI_THIEU: Number(r.TON_TOI_THIEU), GIA_MUA: Number(r.GIA_MUA), TRANG_THAI: String(r.TRANG_THAI).trim() }));
      const result = await request('importItems', { items, updateExisting });
      state.importPlan = null; await refresh(); notify(`Đã thêm ${result.data.added}, cập nhật ${result.data.updated} mặt hàng.`); return true;
    }
    const slips = groupImportedSlips(plan.rows, plan.kind);
    if (!confirm(`Xác nhận nhập ${slips.length} phiếu gồm ${plan.rows.length} dòng? Máy chủ sẽ kiểm tra lại tồn kho trước khi ghi.`)) return true;
    plan.requestId = plan.requestId || crypto.randomUUID().replaceAll('-', '');
    const result = await request('importSlips', { kind: plan.kind, requestId: plan.requestId, slips });
    state.importPlan = null; await refresh(); notify(`Đã nhập ${result.data.slipCount} phiếu, ${result.data.lineCount} dòng.`); return true;
  }
  if (action === 'export-v1') {
    if (!window.XLSX) throw Error('Thư viện Excel chưa tải.');
    const kind = $('#exportKind').value, wb = XLSX.utils.book_new();
    if (kind === 'items') {
      const rows = state.items.map(x => ({ 'Mã hàng': x.MA_HANG, 'Tên hàng': x.TEN_HANG, 'Đơn vị': x.DON_VI, 'Tồn đầu': x.TON_DAU, 'Tổng nhập': x.TONG_NHAP, 'Tổng xuất': x.TONG_XUAT, 'Tồn hiện tại': x.TON_HIEN_TAI, 'Tồn tối thiểu': x.TON_TOI_THIEU, 'Giá mua': x.GIA_MUA, 'Giá trị tồn': x.GIA_TRI_TON, 'Trạng thái': x.TRANG_THAI }));
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), 'Ton kho');
    } else {
      await ensureSlipData(kind);
      const slips = kind === 'IN' ? state.inbound : state.outbound;
      const rows = slips.flatMap(s => slips.length ? s.lines.map(line => kind === 'IN'
        ? { 'Số phiếu': s.no, 'Ngày nhập': excelDateValue(s.date), 'Mã hàng': line.code, 'Tên hàng': line.name, 'Đơn vị': line.unit, 'Số lượng': line.qty, 'Đơn giá': line.price, 'Thành tiền': line.qty * line.price, 'Nhà cung cấp': s.party, 'Người lập': s.creator, 'Ghi chú': s.note }
        : { 'Số phiếu': s.no, 'Ngày xuất': excelDateValue(s.date), 'Mã hàng': line.code, 'Tên hàng': line.name, 'Đơn vị': line.unit, 'Số lượng': line.qty, 'Nơi nhận': s.party, 'Người lập': s.creator, 'Ghi chú': s.note }) : []);
      const sheet = XLSX.utils.json_to_sheet(rows);
      for (let r = 2; r <= rows.length + 1; r++) if (sheet[`B${r}`]) sheet[`B${r}`].z = 'dd/mm/yyyy';
      XLSX.utils.book_append_sheet(wb, sheet, kind === 'IN' ? 'Nhap kho' : 'Xuat kho');
    }
    XLSX.writeFile(wb, `VIFUSOFT_${kind}_${today()}.xlsx`); return true;
  }
  if (action === 'download-template') { downloadItemTemplate(); return true; }
  return false;
}

document.addEventListener('click', async e => {
  const button = e.target.closest('[data-action]');
  if (!button) return;
  try { await handleSpecAction(button.dataset.action, button); }
  catch (err) { notify(err.message || 'Thao tác không thành công.', true); }
});

document.addEventListener('change', e => {
  if (e.target.id === 'chartItem') {
    state.chartItem = e.target.value;
    if (state.chartItem) Promise.all([ensureSlipData('IN'), ensureSlipData('OUT')]).then(render).catch(err => notify(err.message, true));
    else render();
  }
  if (e.target.matches('#lineRows [name="code"]')) {
    const item = state.items.find(x => x.MA_HANG === e.target.value);
    const price = $('[name="price"]', e.target.closest('tr'));
    if (price && item) price.value = Number(item.GIA_MUA || 0);
    updateSlipTotal();
  }
  if (e.target.id === 'importFile' || e.target.id === 'importKind') state.importPlan = null;
});
document.addEventListener('input', e => { if (e.target.matches('#lineRows input')) updateSlipTotal(); });

if (state.token && state.user) { showApp(); refresh(); } else showLogin();
document.addEventListener('click', e => {
  const b = e.target.closest('[data-go][data-low]');
  if (b) { state.lowOnly = true; state.page = 'stock'; render(); }
});
