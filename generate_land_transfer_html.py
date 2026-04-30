from __future__ import annotations

import base64
import html
import json
import re
import shutil
import unicodedata
from pathlib import Path

import openpyxl
from openpyxl.utils import get_column_letter


BASE_DIR = Path(r"D:\Codex\Tools")
SOURCE = Path(r"C:\Users\QUANGHUY\Downloads\Bieu_chu_chuyen_dat_dai_mau_cong_thuc.xlsx")
OUT = BASE_DIR / "public" / "index.html"
JSZIP = Path(r"C:\Users\QUANGHUY\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\node_modules\jszip\dist\jszip.min.js")
LOGO = Path(r"C:\Users\QUANGHUY\Downloads\482087578_122221961630205345_1940337838885474762_n.jpg")
HOME_BACKGROUND = Path(r"C:\Users\QUANGHUY\Downloads\ChatGPT Image 12_00_04 30 thg 4, 2026.png")

LAND_NAME_FIXES = {
    "Đất côn trình thủy lợi": "Đất công trình thủy lợi",
}

STT_FIXES_BY_CODE = {
    "TIN": "2.10",
}

HEADER_ROW = 3
CURRENT_COL = 4
MATRIX_START_COL = 5
MATRIX_END_COL = 66
DECREASE_COL = 67
CHANGE_COL = 68
PLAN_COL = 69
PREVIOUS_PLAN_COL = 70
TOTAL_INCREASE_ROW = 67
PLAN_ROW = 68
PREVIOUS_PLAN_DIR = BASE_DIR / "Dulieu"
SAMPLE_DIR = BASE_DIR / "public" / "samples"
LEGACY_SAMPLE_DIR = BASE_DIR / "samples"
SAMPLE_FILES = [
    ("Dữ liệu hiện trạng mẫu.xlsx", "hien-trang-mau.xlsx", "Dữ liệu hiện trạng mẫu"),
    ("Kết quả thực hiện quy hoạch năm mẫu.xlsx", "ket-qua-quy-hoach-nam-mau.xlsx", "Kết quả thực hiện quy hoạch năm mẫu"),
    ("Bảng quy hoạch mẫu.xlsx", "bang-quy-hoach-mau.xlsx", "Bảng quy hoạch mẫu"),
]
WEBGIS_SAMPLE_DATA = {
    "type": "FeatureCollection",
    "features": [
        {
            "type": "Feature",
            "properties": {"layer": "administrative", "ten": "Ranh giới xã mẫu", "ma_dv": "XA-001", "ghi_chu": "Ranh giới hành chính phục vụ demo WebGIS"},
            "geometry": {"type": "Polygon", "coordinates": [[[105.8422, 21.0474], [105.8616, 21.0474], [105.8616, 21.0330], [105.8422, 21.0330], [105.8422, 21.0474]]]},
        },
        {
            "type": "Feature",
            "properties": {"layer": "parcels", "ma_thua": "TD-101", "chu_su_dung": "Nguyễn Văn A", "loai_dat": "ONT", "dien_tich": 1240.5, "muc_dich": "Đất ở tại nông thôn", "quy_hoach": "Đất ở", "dia_danh": "Thôn Đông", "ghi_chu": "Thửa đất mẫu"},
            "geometry": {"type": "Polygon", "coordinates": [[[105.8481, 21.0424], [105.8515, 21.0423], [105.8513, 21.0398], [105.8478, 21.0399], [105.8481, 21.0424]]]},
        },
        {
            "type": "Feature",
            "properties": {"layer": "parcels", "ma_thua": "TD-102", "chu_su_dung": "Trần Thị B", "loai_dat": "LUC", "dien_tich": 3560.0, "muc_dich": "Đất trồng lúa nước", "quy_hoach": "Đất nông nghiệp", "dia_danh": "Cánh đồng Bắc", "ghi_chu": "Giữ nguyên hiện trạng"},
            "geometry": {"type": "Polygon", "coordinates": [[[105.8517, 21.0422], [105.8562, 21.0421], [105.8560, 21.0395], [105.8515, 21.0397], [105.8517, 21.0422]]]},
        },
        {
            "type": "Feature",
            "properties": {"layer": "landuse", "ma_khoanh": "HT-01", "loai_dat": "LUC", "dien_tich": 6.42, "muc_dich": "Đất trồng lúa", "quy_hoach": "Một phần chuyển sang giao thông", "ghi_chu": "Hiện trạng sử dụng đất"},
            "geometry": {"type": "Polygon", "coordinates": [[[105.8460, 21.0384], [105.8542, 21.0381], [105.8540, 21.0346], [105.8458, 21.0349], [105.8460, 21.0384]]]},
        },
        {
            "type": "Feature",
            "properties": {"layer": "planning", "ma_khoanh": "QH-02", "loai_dat": "DGT", "dien_tich": 1.18, "muc_dich": "Đất giao thông", "quy_hoach": "Tuyến đường quy hoạch", "loai_quy_hoach": "Hạ tầng giao thông", "ghi_chu": "Vùng quy hoạch mẫu"},
            "geometry": {"type": "Polygon", "coordinates": [[[105.8450, 21.0400], [105.8600, 21.0395], [105.8600, 21.0387], [105.8450, 21.0392], [105.8450, 21.0400]]]},
        },
        {
            "type": "Feature",
            "properties": {"layer": "roads", "ten": "Đường trục xã", "loai_dat": "DGT", "dien_tich": 0.84, "muc_dich": "Giao thông", "quy_hoach": "Nâng cấp mở rộng", "ghi_chu": "Tuyến đường mẫu"},
            "geometry": {"type": "LineString", "coordinates": [[105.8442, 21.0436], [105.8490, 21.0411], [105.8548, 21.0390], [105.8608, 21.0362]]},
        },
        {
            "type": "Feature",
            "properties": {"layer": "water", "ten": "Kênh tiêu nội đồng", "loai_dat": "DTL", "dien_tich": 0.52, "muc_dich": "Thủy lợi", "quy_hoach": "Giữ nguyên", "ghi_chu": "Tuyến thủy hệ mẫu"},
            "geometry": {"type": "LineString", "coordinates": [[105.8435, 21.0360], [105.8497, 21.0375], [105.8552, 21.0370], [105.8610, 21.0350]]},
        },
        {
            "type": "Feature",
            "properties": {"layer": "public", "ten": "Trụ sở UBND xã", "loai_dat": "TSC", "dien_tich": 0.32, "muc_dich": "Đất trụ sở cơ quan", "quy_hoach": "Giữ nguyên công trình công cộng", "ghi_chu": "Điểm công trình công cộng"},
            "geometry": {"type": "Point", "coordinates": [105.8520, 21.0437]},
        },
        {
            "type": "Feature",
            "properties": {"layer": "public", "ten": "Trường tiểu học", "loai_dat": "DGD", "dien_tich": 0.78, "muc_dich": "Đất giáo dục", "quy_hoach": "Mở rộng khuôn viên", "ghi_chu": "Điểm công trình công cộng"},
            "geometry": {"type": "Point", "coordinates": [105.8570, 21.0413]},
        },
    ],
}

WEBGIS_CSS = r"""
.webgis-page {
  overflow: hidden;
  background: linear-gradient(135deg, #f8fbff 0%, #eef6ff 52%, #f8fafc 100%);
}
.webgis-shell {
  display: flex;
  flex-direction: column;
  min-height: calc(100vh - 102px);
  border-radius: 8px;
  overflow: hidden;
}
.webgis-topbar {
  display: grid;
  grid-template-columns: minmax(260px, 1fr) minmax(220px, 460px) auto;
  gap: 12px;
  align-items: center;
  padding: 12px 14px;
  border-bottom: 1px solid #dbe7f3;
  background: linear-gradient(135deg, rgba(255,255,255,0.98), rgba(237,248,255,0.96));
}
.webgis-title {
  min-width: 0;
}
.webgis-title strong {
  display: block;
  color: #0f2f57;
  font-size: 18px;
  letter-spacing: 0;
}
.webgis-title span {
  display: block;
  margin-top: 3px;
  color: #55708d;
  font-size: 12px;
}
.webgis-search {
  position: relative;
  display: flex;
  gap: 8px;
  min-width: 0;
}
.webgis-search input,
.webgis-panel input,
.webgis-panel select,
.webgis-panel textarea,
.webgis-attr-tools input,
.webgis-attr-tools select {
  height: 34px;
  min-width: 0;
  border: 1px solid #bdd0e3;
  border-radius: 7px;
  padding: 6px 9px;
  background: #fff;
  color: #102033;
  font-size: 13px;
}
.webgis-search input {
  flex: 1;
}
.webgis-actions,
.webgis-map-tools,
.webgis-attr-tools {
  display: flex;
  align-items: center;
  flex-wrap: wrap;
  gap: 8px;
}
.webgis-save-status {
  min-height: 26px;
  display: inline-flex;
  align-items: center;
  border: 1px solid #bfdbfe;
  border-radius: 999px;
  padding: 3px 9px;
  background: #eff6ff;
  color: #1d4ed8;
  font-size: 12px;
  font-weight: 700;
}
.webgis-save-status.error {
  border-color: #fecaca;
  background: #fff1f2;
  color: #991b1b;
}
.webgis-actions button,
.webgis-map-tools button,
.webgis-search button,
.webgis-panel button,
.webgis-attr-tools button {
  height: 34px;
  border: 1px solid #b7c8da;
  border-radius: 7px;
  padding: 0 11px;
  background: #fff;
  color: #102033;
  font-size: 12px;
  font-weight: 700;
  cursor: pointer;
  box-shadow: 0 2px 8px rgba(15, 47, 87, 0.08);
}
.webgis-actions .primary,
.webgis-search .primary,
.webgis-panel .primary {
  border-color: #0f766e;
  background: #0f766e;
  color: #fff;
}
.webgis-search-results {
  position: absolute;
  z-index: 540;
  top: calc(100% + 7px);
  left: 0;
  right: 0;
  max-height: 260px;
  overflow: auto;
  border: 1px solid #bdd0e3;
  border-radius: 8px;
  background: #fff;
  box-shadow: 0 18px 34px rgba(15, 47, 87, 0.18);
}
.webgis-search-results[hidden] {
  display: none;
}
.webgis-result-item {
  width: 100%;
  min-height: 42px;
  border: 0;
  border-bottom: 1px solid #edf2f7;
  padding: 8px 10px;
  background: #fff;
  color: #102033;
  text-align: left;
  cursor: pointer;
}
.webgis-result-item:hover {
  background: #eff6ff;
}
.webgis-result-item strong {
  display: block;
  font-size: 13px;
}
.webgis-result-item span {
  display: block;
  margin-top: 2px;
  color: #64748b;
  font-size: 12px;
}
.webgis-workspace {
  display: grid;
  grid-template-columns: 304px minmax(360px, 1fr) 320px;
  gap: 10px;
  min-height: 0;
  flex: 1;
  padding: 10px;
}
.webgis-sidebar,
.webgis-info {
  display: flex;
  flex-direction: column;
  gap: 10px;
  min-width: 0;
  overflow: auto;
}
.webgis-panel {
  border: 1px solid #d4e2ef;
  border-radius: 8px;
  background: rgba(255,255,255,0.96);
  box-shadow: 0 10px 24px rgba(15, 47, 87, 0.08);
}
.webgis-panel-head {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 8px;
  padding: 10px 12px;
  border-bottom: 1px solid #e4edf6;
}
.webgis-panel-head h2,
.webgis-panel-head h3 {
  margin: 0;
  color: #0f2f57;
  font-size: 14px;
}
.webgis-panel-body {
  padding: 10px 12px;
}
.webgis-layer-list {
  display: flex;
  flex-direction: column;
  gap: 8px;
}
.webgis-layer-item {
  display: grid;
  grid-template-columns: auto 1fr auto;
  gap: 8px;
  align-items: center;
  padding: 9px;
  border: 1px solid #e3edf6;
  border-radius: 8px;
  background: #f8fbff;
}
.webgis-layer-main {
  min-width: 0;
}
.webgis-layer-main label {
  display: flex;
  align-items: center;
  gap: 7px;
  color: #102033;
  font-size: 13px;
  font-weight: 700;
}
.webgis-layer-tools {
  grid-column: 2 / 4;
  display: flex;
  align-items: center;
  gap: 8px;
}
.webgis-layer-tools input[type="range"] {
  flex: 1;
  accent-color: #2563eb;
}
.webgis-symbol {
  width: 18px;
  height: 18px;
  border: 1px solid rgba(15,47,87,0.20);
  border-radius: 5px;
  flex: 0 0 auto;
}
.webgis-map-panel {
  position: relative;
  min-height: 620px;
  overflow: hidden;
  border: 1px solid #cbdcec;
  border-radius: 8px;
  background: #eaf2f8;
  box-shadow: inset 0 0 0 1px rgba(255,255,255,0.65), 0 12px 30px rgba(15, 47, 87, 0.12);
}
.webgis-map {
  width: 100%;
  height: 100%;
  min-height: 620px;
}
.webgis-map-tools {
  position: absolute;
  z-index: 500;
  top: 12px;
  left: 12px;
  max-width: calc(100% - 24px);
  padding: 8px;
  border: 1px solid rgba(183, 200, 218, 0.72);
  border-radius: 10px;
  background: rgba(255,255,255,0.94);
  box-shadow: 0 14px 30px rgba(15,47,87,0.14);
  backdrop-filter: blur(8px);
}
.webgis-coordinate-bar,
.webgis-measure-badge {
  position: absolute;
  z-index: 500;
  right: 12px;
  padding: 7px 10px;
  border-radius: 8px;
  background: rgba(15, 47, 87, 0.86);
  color: #fff;
  font-size: 12px;
  box-shadow: 0 12px 28px rgba(15,47,87,0.16);
}
.webgis-coordinate-bar {
  bottom: 12px;
}
.webgis-measure-badge {
  left: 12px;
  right: auto;
  bottom: 12px;
  max-width: 420px;
}
.webgis-detail-empty {
  min-height: 140px;
  display: grid;
  place-items: center;
  color: #64748b;
  text-align: center;
  font-size: 13px;
}
.webgis-detail-table {
  width: 100%;
  border-collapse: collapse;
  font-size: 13px;
}
.webgis-detail-table th,
.webgis-detail-table td {
  padding: 7px 5px;
  border-bottom: 1px solid #edf2f7;
  text-align: left;
  vertical-align: top;
}
.webgis-detail-table th {
  width: 112px;
  color: #64748b;
  font-weight: 700;
}
.webgis-legend {
  display: grid;
  grid-template-columns: 1fr;
  gap: 7px;
  font-size: 13px;
}
.webgis-legend-item {
  display: flex;
  align-items: center;
  gap: 8px;
}
.webgis-admin-panel[hidden],
.webgis-attr-panel[hidden] {
  display: none;
}
.webgis-admin-grid {
  display: grid;
  gap: 8px;
}
.webgis-admin-grid label {
  display: grid;
  gap: 4px;
  color: #334155;
  font-size: 12px;
  font-weight: 700;
}
.webgis-admin-note {
  margin: 0;
  color: #64748b;
  font-size: 12px;
  line-height: 1.4;
}
.webgis-attr-panel {
  margin: 0 10px 10px;
  border: 1px solid #cbdcec;
  border-radius: 8px;
  background: #fff;
  box-shadow: 0 14px 32px rgba(15,47,87,0.12);
}
.webgis-attr-tools {
  padding: 10px 12px;
  border-bottom: 1px solid #e4edf6;
}
.webgis-attr-wrap {
  max-height: 260px;
  overflow: auto;
}
.webgis-attr-table {
  width: 100%;
  border-collapse: collapse;
  min-width: 900px;
  font-size: 12px;
}
.webgis-attr-table th,
.webgis-attr-table td {
  padding: 8px;
  border: 1px solid #e4edf6;
  background: #fff;
  text-align: left;
  white-space: nowrap;
}
.webgis-attr-table th {
  position: sticky;
  top: 0;
  z-index: 1;
  background: #eaf4ff;
  color: #0f2f57;
  cursor: pointer;
}
.webgis-attr-table tr:hover td,
.webgis-attr-table tr.selected td {
  background: #fff7ed;
}
.webgis-popup {
  min-width: 220px;
  font-size: 13px;
}
.webgis-popup strong {
  display: block;
  margin-bottom: 6px;
  color: #0f2f57;
}
body.webgis-mode .leaflet-control-layers {
  border-radius: 8px;
  border-color: #b7c8da;
  box-shadow: 0 12px 28px rgba(15,47,87,0.15);
}
@media print {
  body.webgis-mode .appbar,
  body.webgis-mode .webgis-sidebar,
  body.webgis-mode .webgis-info,
  body.webgis-mode .webgis-topbar,
  body.webgis-mode .webgis-attr-panel,
  body.webgis-mode .webgis-map-tools {
    display: none !important;
  }
  body.webgis-mode .webgis-page,
  body.webgis-mode .webgis-shell,
  body.webgis-mode .webgis-workspace,
  body.webgis-mode .webgis-map-panel,
  body.webgis-mode .webgis-map {
    display: block !important;
    margin: 0 !important;
    width: 100% !important;
    height: 100vh !important;
    min-height: 100vh !important;
    box-shadow: none !important;
  }
}
@media (max-width: 1180px) {
  .webgis-workspace {
    grid-template-columns: 280px minmax(320px, 1fr);
  }
  .webgis-info {
    grid-column: 1 / -1;
    display: grid;
    grid-template-columns: repeat(2, minmax(0, 1fr));
  }
}
@media (max-width: 820px) {
  .webgis-topbar {
    grid-template-columns: 1fr;
  }
  .webgis-workspace {
    grid-template-columns: 1fr;
  }
  .webgis-sidebar,
  .webgis-info {
    max-height: none;
  }
  .webgis-info {
    display: flex;
  }
  .webgis-map-panel,
  .webgis-map {
    min-height: 520px;
  }
}
"""

WEBGIS_HTML = r"""
<main id="webgisPage" class="webgis-page" aria-label="WebGis">
  <section class="webgis-shell">
    <header class="webgis-topbar">
      <div class="webgis-title">
        <strong>WEBGIS QUẢN LÝ DỮ LIỆU ĐẤT ĐAI</strong>
        <span>Hiển thị, tra cứu và quản lý dữ liệu bản đồ đất đai/quy hoạch</span>
      </div>
      <div class="webgis-search">
        <input id="webgisSearchInput" type="search" placeholder="Tìm mã thửa, chủ sử dụng, mã đất, địa danh, quy hoạch">
        <button id="webgisSearchBtn" class="primary" type="button">Tìm</button>
        <div id="webgisSearchResults" class="webgis-search-results" hidden></div>
      </div>
      <div class="webgis-actions">
        <span id="webgisSaveStatus" class="webgis-save-status">Chưa nạp dữ liệu</span>
        <button id="webgisOpenTableBtn" type="button">Bảng thuộc tính</button>
        <button id="webgisAdminBtn" class="primary" type="button">Quản trị dữ liệu</button>
        <button id="webgisHomeBtn" type="button">Màn chính</button>
      </div>
    </header>
    <div class="webgis-workspace">
      <aside class="webgis-sidebar">
        <section class="webgis-panel">
          <div class="webgis-panel-head">
            <h2>Lớp bản đồ</h2>
            <button id="webgisFitAllBtn" type="button">Toàn bộ</button>
          </div>
          <div class="webgis-panel-body">
            <div id="webgisLayerList" class="webgis-layer-list"></div>
          </div>
        </section>
        <section class="webgis-panel">
          <div class="webgis-panel-head">
            <h3>Chú giải</h3>
          </div>
          <div class="webgis-panel-body">
            <div id="webgisLegend" class="webgis-legend"></div>
          </div>
        </section>
        <section id="webgisAdminPanel" class="webgis-panel webgis-admin-panel" hidden>
          <div class="webgis-panel-head">
            <h3>Quản trị dữ liệu</h3>
            <button id="webgisCloseAdminBtn" type="button">Đóng</button>
          </div>
          <div class="webgis-panel-body webgis-admin-grid">
            <p class="webgis-admin-note">Bản demo xử lý GeoJSON trên trình duyệt. Khi nâng cấp backend có thể lưu vào PostGIS và cấp quyền admin.</p>
            <label>Tên layer mới
              <input id="webgisNewLayerName" type="text" placeholder="Ví dụ: Quy hoạch khu dân cư">
            </label>
            <label>Màu ký hiệu
              <input id="webgisNewLayerColor" type="color" value="#2563eb">
            </label>
            <label>File GeoJSON
              <input id="webgisImportInput" type="file" accept=".geojson,.json,application/geo+json,application/json">
            </label>
            <button id="webgisImportBtn" class="primary" type="button">Thêm layer GeoJSON</button>
            <label>Thuộc tính đối tượng đang chọn
              <textarea id="webgisFeatureEditor" placeholder="Chọn một đối tượng trên bản đồ để sửa thuộc tính JSON"></textarea>
            </label>
            <button id="webgisSaveFeatureBtn" type="button">Lưu thuộc tính</button>
          </div>
        </section>
      </aside>
      <section class="webgis-map-panel">
        <div id="webgisMap" class="webgis-map" role="application" aria-label="Bản đồ WebGIS"></div>
        <div class="webgis-map-tools" aria-label="Công cụ bản đồ">
          <button id="webgisLocateBtn" type="button">Vị trí</button>
          <button id="webgisMeasureDistanceBtn" type="button">Đo dài</button>
          <button id="webgisMeasureAreaBtn" type="button">Đo diện tích</button>
          <button id="webgisClearMeasureBtn" type="button">Xóa đo</button>
          <button id="webgisPrintBtn" type="button">In</button>
          <button id="webgisShotBtn" type="button">Chụp ảnh</button>
          <button id="webgisFullscreenBtn" type="button">Toàn màn hình</button>
        </div>
        <div id="webgisMeasureBadge" class="webgis-measure-badge">Sẵn sàng tra cứu bản đồ</div>
        <div id="webgisCoordinateBar" class="webgis-coordinate-bar">Tọa độ: --</div>
      </section>
      <aside class="webgis-info">
        <section class="webgis-panel">
          <div class="webgis-panel-head">
            <h2>Thông tin đối tượng</h2>
          </div>
          <div id="webgisFeatureDetail" class="webgis-panel-body">
            <div class="webgis-detail-empty">Bấm vào thửa đất, vùng quy hoạch, tuyến hoặc điểm công trình để xem thông tin.</div>
          </div>
        </section>
        <section class="webgis-panel">
          <div class="webgis-panel-head">
            <h3>Hướng dẫn nhanh</h3>
          </div>
          <div class="webgis-panel-body">
            <p class="webgis-admin-note">Bật/tắt lớp ở sidebar, dùng thanh trong suốt để so sánh nền bản đồ và dữ liệu. Ô tìm kiếm hỗ trợ mã thửa, mã đất, chủ sử dụng, địa danh và loại quy hoạch.</p>
          </div>
        </section>
      </aside>
    </div>
    <section id="webgisAttributePanel" class="webgis-attr-panel" hidden>
      <div class="webgis-attr-tools">
        <strong>Bảng thuộc tính</strong>
        <select id="webgisAttrLayer"></select>
        <input id="webgisAttrSearch" type="search" placeholder="Lọc thuộc tính">
        <button id="webgisCloseTableBtn" type="button">Đóng bảng</button>
      </div>
      <div class="webgis-attr-wrap">
        <table id="webgisAttrTable" class="webgis-attr-table"></table>
      </div>
    </section>
  </section>
</main>
"""

WEBGIS_JS = r"""
const webgisLayerDefs = [
  { id: 'administrative', label: 'Ranh giới hành chính', color: '#2563eb', visible: true },
  { id: 'landuse', label: 'Hiện trạng sử dụng đất', color: '#22c55e', visible: true },
  { id: 'planning', label: 'Quy hoạch sử dụng đất', color: '#f59e0b', visible: true },
  { id: 'roads', label: 'Giao thông', color: '#6b7280', visible: true },
  { id: 'water', label: 'Thủy hệ', color: '#0ea5e9', visible: true },
  { id: 'parcels', label: 'Thửa đất', color: '#ef4444', visible: true },
  { id: 'public', label: 'Công trình công cộng', color: '#8b5cf6', visible: true }
];

const webgisLandColors = {
  ONT: '#fca5a5',
  ODT: '#fb7185',
  LUC: '#86efac',
  LUA: '#4ade80',
  HNK: '#bbf7d0',
  CLN: '#65a30d',
  DGT: '#9ca3af',
  DTL: '#38bdf8',
  TSC: '#c084fc',
  DGD: '#fde68a',
  DKV: '#facc15'
};

let webgisState = {
  initialized: false,
  initializing: null,
  map: null,
  layerDefs: [],
  overlayLayers: new Map(),
  featureLayers: new Map(),
  features: [],
  selectedFeatureId: null,
  selectedVector: null,
  attrSortKey: '',
  attrSortDir: 1,
  measureMode: null,
  measurePoints: [],
  measureLayer: null
};

const webgisStorageKey = 'webgis-state-v1';
const webgisProjectId = 'webgis-default';
const webgisApiBase = '/api/webgis';
let webgisSaveTimer = 0;

function webgisEl(id) {
  return document.getElementById(id);
}

function webgisEscape(value) {
  return String(value ?? '').replace(/[&<>"']/g, ch => ({
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#039;'
  }[ch]));
}

function webgisFeatureTitle(feature) {
  const props = feature.properties || {};
  return props.ma_thua || props.ma_khoanh || props.ten || props.ma_dv || `Đối tượng ${props.__id || ''}`;
}

function webgisLayerLabel(id) {
  return webgisState.layerDefs.find(layer => layer.id === id)?.label || id || 'Khác';
}

function webgisLayerColor(id, feature) {
  const code = String(feature?.properties?.loai_dat || '').toUpperCase();
  return webgisLandColors[code] || webgisState.layerDefs.find(layer => layer.id === id)?.color || '#2563eb';
}

function webgisNormalizeFeatures(collection, defaultLayer = '') {
  const features = Array.isArray(collection?.features) ? collection.features : [];
  return features
    .filter(feature => feature && feature.geometry)
    .map((feature, index) => {
      const props = { ...(feature.properties || {}) };
      props.layer = props.layer || defaultLayer || 'imported';
      props.__id = props.__id || `${props.layer}-${Date.now()}-${index}-${Math.random().toString(36).slice(2, 7)}`;
      return { ...feature, properties: props };
    });
}

function webgisNormalizeLayerDefs(savedDefs = []) {
  const savedById = new Map(
    (Array.isArray(savedDefs) ? savedDefs : [])
      .filter(def => def && def.id)
      .map(def => [String(def.id), def])
  );
  const defaultIds = new Set(webgisLayerDefs.map(def => def.id));
  const defaults = webgisLayerDefs.map(def => {
    const saved = savedById.get(def.id) || {};
    return {
      ...def,
      ...saved,
      id: def.id,
      label: saved.label || def.label,
      color: saved.color || def.color,
      opacity: Number.isFinite(Number(saved.opacity)) ? Number(saved.opacity) : 78,
      visible: saved.visible !== false
    };
  });
  const custom = Array.from(savedById.values())
    .filter(def => !defaultIds.has(String(def.id)))
    .map(def => ({
      id: String(def.id),
      label: String(def.label || def.id),
      color: String(def.color || '#2563eb'),
      visible: def.visible !== false,
      opacity: Number.isFinite(Number(def.opacity)) ? Number(def.opacity) : 78,
      custom: true
    }));
  return [...defaults, ...custom];
}

function webgisStatePayload() {
  return {
    version: 1,
    savedAt: new Date().toISOString(),
    layerDefs: webgisState.layerDefs.map(def => ({
      id: def.id,
      label: def.label,
      color: def.color,
      visible: def.visible !== false,
      opacity: Number(def.opacity || 78),
      custom: Boolean(def.custom)
    })),
    features: webgisState.features
  };
}

function webgisValidPayload(data) {
  return data && typeof data === 'object' && Array.isArray(data.layerDefs) && Array.isArray(data.features);
}

function webgisSetSaveStatus(text, isError = false) {
  const status = webgisEl('webgisSaveStatus');
  if (!status) return;
  status.textContent = text;
  status.classList.toggle('error', Boolean(isError));
}

function webgisSaveLocal(data) {
  try {
    localStorage.setItem(webgisStorageKey, JSON.stringify(data));
    return true;
  } catch (error) {
    return false;
  }
}

function webgisLoadLocal() {
  try {
    const raw = localStorage.getItem(webgisStorageKey);
    if (!raw) return null;
    const data = JSON.parse(raw);
    return webgisValidPayload(data) ? data : null;
  } catch (error) {
    return null;
  }
}

async function webgisLoadSavedData() {
  const localData = webgisLoadLocal();
  try {
    const response = await fetch(`${webgisApiBase}/${encodeURIComponent(webgisProjectId)}`, { cache: 'no-store' });
    if (response.ok) {
      const payload = await response.json();
      if (webgisValidPayload(payload.data)) {
        webgisSaveLocal(payload.data);
        webgisSetSaveStatus(payload.storage === 'supabase' || payload.storage === 'supabase-migrated' ? 'Đã nạp từ Supabase' : 'Đã nạp dữ liệu đã lưu');
        return payload.data;
      }
    }
    if (response.status !== 404) throw new Error('Không nạp được dữ liệu WebGIS từ server.');
  } catch (error) {
    if (localData) webgisSetSaveStatus('Đang dùng bản lưu tạm', true);
  }
  if (localData) return localData;
  webgisSetSaveStatus('Đang dùng dữ liệu mẫu');
  return null;
}

async function webgisSaveNow() {
  if (!webgisState.initialized) return;
  const data = webgisStatePayload();
  const savedLocal = webgisSaveLocal(data);
  webgisSetSaveStatus('Đang tự lưu...');
  try {
    const response = await fetch(`${webgisApiBase}/${encodeURIComponent(webgisProjectId)}`, {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ data })
    });
    if (!response.ok) throw new Error(await response.text());
    const result = await response.json().catch(() => ({}));
    const target = result.storage === 'supabase' ? 'Supabase' : 'server';
    webgisSetSaveStatus(`Đã tự lưu ${target} ${new Date().toLocaleTimeString('vi-VN', { hour: '2-digit', minute: '2-digit' })}`);
  } catch (error) {
    webgisSetSaveStatus(savedLocal ? 'Đã lưu tạm trên trình duyệt' : 'Không lưu được dữ liệu', true);
  }
}

function webgisScheduleSave() {
  if (!webgisState.initialized) return;
  clearTimeout(webgisSaveTimer);
  webgisSaveTimer = setTimeout(() => webgisSaveNow(), 600);
}

function webgisStyle(feature) {
  const layerId = feature?.properties?.layer;
  const color = webgisLayerColor(layerId, feature);
  const opacity = Number(webgisEl(`webgisOpacity_${layerId}`)?.value || 78) / 100;
  const isLine = ['LineString', 'MultiLineString'].includes(feature.geometry?.type);
  return {
    color,
    weight: layerId === 'administrative' ? 3 : isLine ? 4 : 1.6,
    dashArray: layerId === 'administrative' ? '8 5' : '',
    opacity: Math.max(0.15, opacity),
    fillColor: color,
    fillOpacity: isLine ? 0 : Math.min(0.55, opacity * 0.55)
  };
}

function webgisPopupHtml(feature) {
  const props = feature.properties || {};
  const rows = [
    ['Mã thửa/khoanh', props.ma_thua || props.ma_khoanh || props.ma_dv || ''],
    ['Loại đất', props.loai_dat || ''],
    ['Diện tích', props.dien_tich ? `${props.dien_tich} ha/m2` : ''],
    ['Chủ sử dụng', props.chu_su_dung || ''],
    ['Hiện trạng', props.muc_dich || ''],
    ['Quy hoạch', props.quy_hoach || props.loai_quy_hoach || ''],
    ['Ghi chú', props.ghi_chu || '']
  ].filter(([, value]) => value !== '');
  return `<div class="webgis-popup"><strong>${webgisEscape(webgisFeatureTitle(feature))}</strong>${
    rows.map(([key, value]) => `<div><b>${webgisEscape(key)}:</b> ${webgisEscape(value)}</div>`).join('')
  }</div>`;
}

function webgisRenderFeatureDetail(feature) {
  const detail = webgisEl('webgisFeatureDetail');
  if (!feature) {
    detail.innerHTML = '<div class="webgis-detail-empty">Bấm vào thửa đất, vùng quy hoạch, tuyến hoặc điểm công trình để xem thông tin.</div>';
    webgisEl('webgisFeatureEditor').value = '';
    return;
  }
  const props = feature.properties || {};
  const rows = Object.entries(props).filter(([key]) => key !== '__id');
  detail.innerHTML = `<table class="webgis-detail-table"><tbody>${
    rows.map(([key, value]) => `<tr><th>${webgisEscape(key)}</th><td>${webgisEscape(value)}</td></tr>`).join('')
  }</tbody></table>`;
  webgisEl('webgisFeatureEditor').value = JSON.stringify(Object.fromEntries(rows), null, 2);
}

function webgisSelectFeature(feature, vectorLayer, openPopup = true) {
  if (!feature) return;
  if (webgisState.selectedVector?.setStyle && webgisState.selectedFeatureId) {
    const previousFeature = webgisState.features.find(item => item.properties.__id === webgisState.selectedFeatureId);
    webgisState.selectedVector.setStyle(webgisStyle(previousFeature));
  }
  webgisState.selectedFeatureId = feature.properties.__id;
  webgisState.selectedVector = vectorLayer || webgisState.featureLayers.get(feature.properties.__id);
  if (webgisState.selectedVector?.setStyle) {
    webgisState.selectedVector.setStyle({ color: '#f97316', weight: 4, fillOpacity: 0.48 });
    webgisState.selectedVector.bringToFront?.();
  }
  webgisRenderFeatureDetail(feature);
  if (openPopup && webgisState.selectedVector?.bindPopup) {
    webgisState.selectedVector.bindPopup(webgisPopupHtml(feature)).openPopup();
  }
  webgisHighlightAttrRow(feature.properties.__id);
}

function webgisBuildOverlayLayer(def) {
  const collection = {
    type: 'FeatureCollection',
    features: webgisState.features.filter(feature => feature.properties?.layer === def.id)
  };
  return L.geoJSON(collection, {
    style: webgisStyle,
    pointToLayer(feature, latlng) {
      const color = webgisLayerColor(def.id, feature);
      return L.circleMarker(latlng, { radius: 8, color, weight: 2, fillColor: color, fillOpacity: 0.82 });
    },
    onEachFeature(feature, vectorLayer) {
      webgisState.featureLayers.set(feature.properties.__id, vectorLayer);
      vectorLayer.on('click', () => webgisSelectFeature(feature, vectorLayer));
      vectorLayer.bindTooltip(webgisFeatureTitle(feature), { sticky: true });
    }
  });
}

function webgisRebuildOverlays() {
  if (!webgisState.map) return;
  webgisState.overlayLayers.forEach(layer => webgisState.map.removeLayer(layer));
  webgisState.overlayLayers.clear();
  webgisState.featureLayers.clear();
  webgisState.layerDefs.forEach(def => {
    const layer = webgisBuildOverlayLayer(def);
    webgisState.overlayLayers.set(def.id, layer);
    if (def.visible !== false) layer.addTo(webgisState.map);
  });
  if (webgisState.selectedFeatureId && !webgisState.features.some(feature => feature.properties.__id === webgisState.selectedFeatureId)) {
    webgisState.selectedFeatureId = null;
    webgisState.selectedVector = null;
    webgisRenderFeatureDetail(null);
  }
  webgisRenderLayerList();
  webgisRenderLegend();
  webgisPopulateAttrLayerSelect();
}

function webgisRenderLayerList() {
  const root = webgisEl('webgisLayerList');
  root.innerHTML = webgisState.layerDefs.map(def => `
    <div class="webgis-layer-item" data-layer="${webgisEscape(def.id)}">
      <span class="webgis-symbol" style="background:${webgisEscape(def.color)}"></span>
      <div class="webgis-layer-main">
        <label><input type="checkbox" data-webgis-layer-toggle="${webgisEscape(def.id)}" ${def.visible === false ? '' : 'checked'}> ${webgisEscape(def.label)}</label>
      </div>
      <button type="button" data-webgis-layer-zoom="${webgisEscape(def.id)}">Zoom</button>
      <div class="webgis-layer-tools">
        <span>Trong suốt</span>
        <input id="webgisOpacity_${webgisEscape(def.id)}" type="range" min="10" max="100" value="${def.opacity || 78}" data-webgis-layer-opacity="${webgisEscape(def.id)}">
      </div>
    </div>
  `).join('');
}

function webgisRenderLegend() {
  const legend = webgisEl('webgisLegend');
  legend.innerHTML = [
    ['#fca5a5', 'Đất ở'],
    ['#86efac', 'Đất nông nghiệp'],
    ['#9ca3af', 'Đất giao thông'],
    ['#38bdf8', 'Đất thủy lợi, sông suối'],
    ['#c084fc', 'Đất công cộng'],
    ['#facc15', 'Đất quy hoạch/hạ tầng']
  ].map(([color, label]) => `<div class="webgis-legend-item"><span class="webgis-symbol" style="background:${color}"></span>${label}</div>`).join('');
}

function webgisFitLayer(layerId) {
  const layer = webgisState.overlayLayers.get(layerId);
  if (!layer) return;
  const bounds = layer.getBounds?.();
  if (bounds?.isValid?.()) webgisState.map.fitBounds(bounds.pad(0.14));
}

function webgisFitAll() {
  const group = L.featureGroup(Array.from(webgisState.overlayLayers.values()));
  const bounds = group.getBounds();
  if (bounds.isValid()) webgisState.map.fitBounds(bounds.pad(0.12));
}

function webgisUpdateLayerStyle(layerId) {
  const layer = webgisState.overlayLayers.get(layerId);
  if (!layer) return;
  layer.eachLayer(child => {
    if (child.feature && child.setStyle) child.setStyle(webgisStyle(child.feature));
  });
}

function webgisTextForFeature(feature) {
  const props = feature.properties || {};
  return [props.ma_thua, props.ma_khoanh, props.ma_dv, props.ten, props.chu_su_dung, props.loai_dat, props.muc_dich, props.quy_hoach, props.loai_quy_hoach, props.dia_danh, props.ghi_chu, webgisLayerLabel(props.layer)].join(' ').toLowerCase();
}

function webgisSearch() {
  const query = webgisEl('webgisSearchInput').value.trim().toLowerCase();
  const box = webgisEl('webgisSearchResults');
  if (!query) {
    box.hidden = true;
    box.innerHTML = '';
    return;
  }
  const results = webgisState.features.filter(feature => webgisTextForFeature(feature).includes(query)).slice(0, 30);
  if (!results.length) {
    box.innerHTML = '<div class="webgis-result-item"><strong>Không tìm thấy dữ liệu phù hợp</strong><span>Thử mã đất, mã thửa hoặc tên địa danh khác.</span></div>';
    box.hidden = false;
    return;
  }
  box.innerHTML = results.map(feature => `
    <button type="button" class="webgis-result-item" data-feature-id="${webgisEscape(feature.properties.__id)}">
      <strong>${webgisEscape(webgisFeatureTitle(feature))}</strong>
      <span>${webgisEscape(webgisLayerLabel(feature.properties.layer))} - ${webgisEscape(feature.properties.loai_dat || feature.properties.quy_hoach || '')}</span>
    </button>
  `).join('');
  box.hidden = false;
  webgisZoomToFeature(results[0].properties.__id);
}

function webgisZoomToFeature(featureId) {
  const feature = webgisState.features.find(item => item.properties.__id === featureId);
  const vector = webgisState.featureLayers.get(featureId);
  if (!feature || !vector) return;
  if (vector.getBounds) {
    const bounds = vector.getBounds();
    if (bounds.isValid()) webgisState.map.fitBounds(bounds.pad(0.35));
  } else if (vector.getLatLng) {
    webgisState.map.setView(vector.getLatLng(), 17);
  }
  webgisSelectFeature(feature, vector);
}

function webgisAllPropertyKeys(features) {
  const keys = new Set(['ma_thua', 'ma_khoanh', 'ten', 'loai_dat', 'dien_tich', 'chu_su_dung', 'muc_dich', 'quy_hoach', 'ghi_chu']);
  features.forEach(feature => Object.keys(feature.properties || {}).forEach(key => key !== '__id' && keys.add(key)));
  return Array.from(keys);
}

function webgisPopulateAttrLayerSelect() {
  const select = webgisEl('webgisAttrLayer');
  select.innerHTML = webgisState.layerDefs.map(def => `<option value="${webgisEscape(def.id)}">${webgisEscape(def.label)}</option>`).join('');
}

function webgisRenderAttributeTable() {
  const layerId = webgisEl('webgisAttrLayer').value || webgisState.layerDefs[0]?.id || '';
  const query = webgisEl('webgisAttrSearch').value.trim().toLowerCase();
  let features = webgisState.features.filter(feature => feature.properties.layer === layerId);
  if (query) features = features.filter(feature => JSON.stringify(feature.properties).toLowerCase().includes(query));
  if (webgisState.attrSortKey) {
    const key = webgisState.attrSortKey;
    features = features.slice().sort((a, b) => String(a.properties[key] ?? '').localeCompare(String(b.properties[key] ?? ''), 'vi') * webgisState.attrSortDir);
  }
  const keys = webgisAllPropertyKeys(features);
  const table = webgisEl('webgisAttrTable');
  table.innerHTML = `
    <thead><tr>${keys.map(key => `<th data-webgis-sort="${webgisEscape(key)}">${webgisEscape(key)}</th>`).join('')}</tr></thead>
    <tbody>${features.map(feature => `
      <tr data-feature-id="${webgisEscape(feature.properties.__id)}">${keys.map(key => `<td>${webgisEscape(feature.properties[key] ?? '')}</td>`).join('')}</tr>
    `).join('')}</tbody>
  `;
  webgisHighlightAttrRow(webgisState.selectedFeatureId);
}

function webgisHighlightAttrRow(featureId) {
  const table = webgisEl('webgisAttrTable');
  if (!table) return;
  table.querySelectorAll('tr.selected').forEach(row => row.classList.remove('selected'));
  if (!featureId || !window.CSS?.escape) return;
  table.querySelector(`tr[data-feature-id="${CSS.escape(featureId)}"]`)?.classList.add('selected');
}

function webgisFormatDistance(meters) {
  return meters >= 1000 ? `${(meters / 1000).toFixed(2)} km` : `${meters.toFixed(1)} m`;
}

function webgisFormatArea(squareMeters) {
  return squareMeters >= 10000 ? `${(squareMeters / 10000).toFixed(2)} ha` : `${squareMeters.toFixed(1)} m²`;
}

function webgisSphericalArea(latlngs) {
  if (latlngs.length < 3) return 0;
  const radius = 6378137;
  const rad = Math.PI / 180;
  let area = 0;
  for (let i = 0; i < latlngs.length; i += 1) {
    const p1 = latlngs[i];
    const p2 = latlngs[(i + 1) % latlngs.length];
    area += (p2.lng - p1.lng) * rad * (2 + Math.sin(p1.lat * rad) + Math.sin(p2.lat * rad));
  }
  return Math.abs(area * radius * radius / 2);
}

function webgisSetMeasureMode(mode) {
  webgisState.measureMode = mode;
  webgisState.measurePoints = [];
  webgisClearMeasure(false);
  webgisEl('webgisMeasureBadge').textContent = mode === 'distance' ? 'Đo khoảng cách: bấm các điểm trên bản đồ.' : 'Đo diện tích: bấm các đỉnh vùng trên bản đồ.';
}

function webgisClearMeasure(resetMode = true) {
  if (webgisState.measureLayer) {
    webgisState.map.removeLayer(webgisState.measureLayer);
    webgisState.measureLayer = null;
  }
  webgisState.measurePoints = [];
  if (resetMode) {
    webgisState.measureMode = null;
    webgisEl('webgisMeasureBadge').textContent = 'Sẵn sàng tra cứu bản đồ';
  }
}

function webgisHandleMeasureClick(latlng) {
  if (!webgisState.measureMode) return;
  webgisState.measurePoints.push(latlng);
  if (webgisState.measureLayer) webgisState.map.removeLayer(webgisState.measureLayer);
  if (webgisState.measureMode === 'distance') {
    webgisState.measureLayer = L.polyline(webgisState.measurePoints, { color: '#f97316', weight: 4 }).addTo(webgisState.map);
    const total = webgisState.measurePoints.slice(1).reduce((sum, point, index) => sum + webgisState.map.distance(webgisState.measurePoints[index], point), 0);
    webgisEl('webgisMeasureBadge').textContent = `Chiều dài: ${webgisFormatDistance(total)}. Bấm Xóa đo để kết thúc.`;
  } else {
    webgisState.measureLayer = L.polygon(webgisState.measurePoints, { color: '#f97316', weight: 3, fillOpacity: 0.18 }).addTo(webgisState.map);
    webgisEl('webgisMeasureBadge').textContent = `Diện tích: ${webgisFormatArea(webgisSphericalArea(webgisState.measurePoints))}. Bấm Xóa đo để kết thúc.`;
  }
}

async function webgisTakeScreenshot() {
  if (!window.html2canvas) {
    alert('Chưa tải được thư viện chụp ảnh bản đồ. Vui lòng thử lại sau.');
    return;
  }
  try {
    const canvas = await window.html2canvas(webgisEl('webgisMap'), { useCORS: true, backgroundColor: '#eef6ff' });
    const link = document.createElement('a');
    link.download = 'webgis-ban-do.png';
    link.href = canvas.toDataURL('image/png');
    link.click();
  } catch (error) {
    alert('Không chụp được ảnh do trình duyệt chặn ảnh nền bản đồ từ nguồn ngoài. Có thể dùng công cụ In bản đồ để lưu PDF.');
  }
}

async function webgisImportGeoJson() {
  const file = webgisEl('webgisImportInput').files?.[0];
  const name = webgisEl('webgisNewLayerName').value.trim() || file?.name?.replace(/\.(geojson|json)$/i, '') || 'Layer GeoJSON';
  if (!file) {
    alert('Hãy chọn file GeoJSON trước.');
    return;
  }
  const data = JSON.parse(await file.text());
  const layerId = `custom_${Date.now()}`;
  const color = webgisEl('webgisNewLayerColor').value || '#2563eb';
  const features = webgisNormalizeFeatures(data, layerId).map(feature => ({ ...feature, properties: { ...feature.properties, layer: layerId } }));
  if (!features.length) {
    alert('File GeoJSON không có đối tượng hợp lệ.');
    return;
  }
  webgisState.layerDefs.push({ id: layerId, label: name, color, visible: true, opacity: 78, custom: true });
  webgisState.features.push(...features);
  webgisRebuildOverlays();
  webgisFitLayer(layerId);
  webgisScheduleSave();
  alert(`Đã thêm ${features.length} đối tượng vào layer "${name}".`);
}

function webgisSaveSelectedFeatureProps() {
  const id = webgisState.selectedFeatureId;
  if (!id) {
    alert('Hãy chọn một đối tượng trên bản đồ trước.');
    return;
  }
  let props;
  try {
    props = JSON.parse(webgisEl('webgisFeatureEditor').value || '{}');
  } catch (error) {
    alert('Nội dung thuộc tính phải là JSON hợp lệ.');
    return;
  }
  const feature = webgisState.features.find(item => item.properties.__id === id);
  if (!feature) return;
  feature.properties = { ...props, layer: props.layer || feature.properties.layer, __id: id };
  webgisRebuildOverlays();
  webgisZoomToFeature(id);
  webgisRenderAttributeTable();
  webgisScheduleSave();
}

function webgisBindEvents() {
  webgisEl('webgisLayerList').addEventListener('change', event => {
    const toggleId = event.target?.dataset?.webgisLayerToggle;
    const opacityId = event.target?.dataset?.webgisLayerOpacity;
    if (toggleId) {
      const def = webgisState.layerDefs.find(layer => layer.id === toggleId);
      if (def) def.visible = event.target.checked;
      const layer = webgisState.overlayLayers.get(toggleId);
      if (event.target.checked) layer?.addTo(webgisState.map);
      else if (layer) webgisState.map.removeLayer(layer);
      webgisScheduleSave();
    }
    if (opacityId) {
      const def = webgisState.layerDefs.find(layer => layer.id === opacityId);
      if (def) def.opacity = Number(event.target.value);
      webgisUpdateLayerStyle(opacityId);
      webgisScheduleSave();
    }
  });
  webgisEl('webgisLayerList').addEventListener('input', event => {
    const opacityId = event.target?.dataset?.webgisLayerOpacity;
    if (!opacityId) return;
    const def = webgisState.layerDefs.find(layer => layer.id === opacityId);
    if (def) def.opacity = Number(event.target.value);
    webgisUpdateLayerStyle(opacityId);
    webgisScheduleSave();
  });
  webgisEl('webgisLayerList').addEventListener('click', event => {
    const layerId = event.target?.dataset?.webgisLayerZoom;
    if (layerId) webgisFitLayer(layerId);
  });
  webgisEl('webgisSearchBtn').addEventListener('click', webgisSearch);
  webgisEl('webgisSearchInput').addEventListener('keydown', event => {
    if (event.key === 'Enter') webgisSearch();
  });
  webgisEl('webgisSearchResults').addEventListener('click', event => {
    const button = event.target.closest('[data-feature-id]');
    if (!button) return;
    webgisZoomToFeature(button.dataset.featureId);
    webgisEl('webgisSearchResults').hidden = true;
  });
  webgisEl('webgisFitAllBtn').addEventListener('click', webgisFitAll);
  webgisEl('webgisHomeBtn').addEventListener('click', showHomePage);
  webgisEl('webgisOpenTableBtn').addEventListener('click', () => {
    webgisEl('webgisAttributePanel').hidden = false;
    webgisRenderAttributeTable();
  });
  webgisEl('webgisCloseTableBtn').addEventListener('click', () => webgisEl('webgisAttributePanel').hidden = true);
  webgisEl('webgisAttrLayer').addEventListener('change', webgisRenderAttributeTable);
  webgisEl('webgisAttrSearch').addEventListener('input', webgisRenderAttributeTable);
  webgisEl('webgisAttrTable').addEventListener('click', event => {
    const sortKey = event.target?.dataset?.webgisSort;
    if (sortKey) {
      webgisState.attrSortDir = webgisState.attrSortKey === sortKey ? -webgisState.attrSortDir : 1;
      webgisState.attrSortKey = sortKey;
      webgisRenderAttributeTable();
      return;
    }
    const row = event.target.closest('tr[data-feature-id]');
    if (row) webgisZoomToFeature(row.dataset.featureId);
  });
  webgisEl('webgisAdminBtn').addEventListener('click', () => webgisEl('webgisAdminPanel').hidden = !webgisEl('webgisAdminPanel').hidden);
  webgisEl('webgisCloseAdminBtn').addEventListener('click', () => webgisEl('webgisAdminPanel').hidden = true);
  webgisEl('webgisImportBtn').addEventListener('click', () => webgisImportGeoJson().catch(error => alert(error.message || String(error))));
  webgisEl('webgisSaveFeatureBtn').addEventListener('click', webgisSaveSelectedFeatureProps);
  webgisEl('webgisLocateBtn').addEventListener('click', () => webgisState.map.locate({ setView: true, maxZoom: 17 }));
  webgisEl('webgisMeasureDistanceBtn').addEventListener('click', () => webgisSetMeasureMode('distance'));
  webgisEl('webgisMeasureAreaBtn').addEventListener('click', () => webgisSetMeasureMode('area'));
  webgisEl('webgisClearMeasureBtn').addEventListener('click', () => webgisClearMeasure(true));
  webgisEl('webgisPrintBtn').addEventListener('click', () => window.print());
  webgisEl('webgisShotBtn').addEventListener('click', () => webgisTakeScreenshot());
  webgisEl('webgisFullscreenBtn').addEventListener('click', () => webgisEl('webgisPage').requestFullscreen?.());
}

async function initializeWebGIS() {
  if (webgisState.initialized) {
    setTimeout(() => webgisState.map?.invalidateSize(), 80);
    return;
  }
  if (webgisState.initializing) {
    await webgisState.initializing;
    setTimeout(() => webgisState.map?.invalidateSize(), 80);
    return;
  }
  if (!window.L) {
    webgisEl('webgisMap').innerHTML = '<div class="webgis-detail-empty">Không tải được Leaflet.js. Vui lòng kiểm tra kết nối mạng hoặc CDN.</div>';
    return;
  }
  webgisState.initializing = (async () => {
    try {
      const sample = JSON.parse(document.getElementById('webgisSampleData').textContent);
      const savedData = await webgisLoadSavedData();
      webgisState.layerDefs = webgisNormalizeLayerDefs(savedData?.layerDefs);
      webgisState.features = webgisNormalizeFeatures(
        { type: 'FeatureCollection', features: savedData?.features || sample.features },
        ''
      );
      const map = L.map('webgisMap', { zoomControl: false, preferCanvas: true }).setView([21.0405, 105.8520], 15);
      webgisState.map = map;
      L.control.zoom({ position: 'bottomright' }).addTo(map);
      const osm = L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', { maxZoom: 20, attribution: '&copy; OpenStreetMap contributors' }).addTo(map);
      const satellite = L.tileLayer('https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}', { maxZoom: 19, attribution: 'Tiles &copy; Esri' });
      const terrain = L.tileLayer('https://{s}.tile.opentopomap.org/{z}/{x}/{y}.png', { maxZoom: 17, attribution: '&copy; OpenTopoMap contributors' });
      L.control.layers({ OpenStreetMap: osm, 'Ảnh vệ tinh': satellite, 'Địa hình': terrain }, null, { position: 'topright', collapsed: true }).addTo(map);
      webgisRebuildOverlays();
      webgisBindEvents();
      map.on('mousemove', event => {
        webgisEl('webgisCoordinateBar').textContent = `Tọa độ: ${event.latlng.lat.toFixed(6)}, ${event.latlng.lng.toFixed(6)}`;
      });
      map.on('click', event => webgisHandleMeasureClick(event.latlng));
      map.on('locationfound', event => {
        L.circleMarker(event.latlng, { radius: 8, color: '#0f766e', fillColor: '#14b8a6', fillOpacity: 0.85 }).addTo(map)
          .bindPopup('Vị trí hiện tại của bạn').openPopup();
      });
      map.on('locationerror', () => alert('Không xác định được vị trí. Hãy cho phép trình duyệt truy cập vị trí nếu cần.'));
      webgisFitAll();
      webgisState.initialized = true;
    } finally {
      webgisState.initializing = null;
    }
  })();
  await webgisState.initializing;
}
"""


def normalize_key(value) -> str:
    text = unicodedata.normalize("NFD", str(value or "").strip().lower())
    text = "".join(ch for ch in text if unicodedata.category(ch) != "Mn")
    text = text.replace("đ", "d")
    return re.sub(r"\s+", " ", text)


def parse_number(value) -> float | None:
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(" ", "")
    if "," in text and "." not in text:
        text = text.replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return None


def format_ha(value: float | None) -> str:
    if value is None:
        return ""
    return f"{value:.2f}".replace(".", ",")


def read_previous_plan_values() -> dict[str, float]:
    files = [p for p in PREVIOUS_PLAN_DIR.glob("*.xlsx") if not p.name.startswith("~$")]
    if not files:
        return {}
    wb = openpyxl.load_workbook(files[0], data_only=True)
    ws = wb[wb.sheetnames[0]]
    code_col = None
    area_col = None
    for row in range(1, min(ws.max_row, 25) + 1):
        for col in range(1, ws.max_column + 1):
            text = normalize_key(ws.cell(row, col).value)
            if text in {"mã", "ma", "mã đất", "ma dat"}:
                code_col = col
            if "diện tích" in text or "dien tich" in text:
                area_col = col
    if not area_col:
        for row in range(1, min(ws.max_row, 25) + 1):
            for col in range(1, ws.max_column + 1):
                if "quy hoạch" in normalize_key(ws.cell(row, col).value) and col <= ws.max_column:
                    area_col = col
                    break
            if area_col:
                break
    if not code_col or not area_col:
        return {}

    values: dict[str, float] = {}
    for row in range(1, ws.max_row + 1):
        code = str(ws.cell(row, code_col).value or "").strip().upper()
        name = normalize_key(ws.cell(row, max(1, code_col - 1)).value)
        if not code and "tổng diện tích tự nhiên" in name:
            code = "DTTN"
        area = parse_number(ws.cell(row, area_col).value)
        if code and area is not None:
            values[code] = area
    return values


def read_previous_plan_values_clean() -> dict[str, float]:
    files = [p for p in PREVIOUS_PLAN_DIR.glob("*.xlsx") if not p.name.startswith("~$")]
    if not files:
        return {}
    wb = openpyxl.load_workbook(files[0], data_only=True)
    ws = wb[wb.sheetnames[0]]
    code_col = None
    area_col = None
    for row in range(1, min(ws.max_row, 25) + 1):
        for col in range(1, ws.max_column + 1):
            text = normalize_key(ws.cell(row, col).value)
            if text in {"ma", "ma dat"}:
                code_col = col
            if "dien tich" in text and (code_col is None or col > code_col):
                area_col = col
    if not area_col:
        for row in range(1, min(ws.max_row, 25) + 1):
            for col in range(1, ws.max_column + 1):
                if "quy hoach" in normalize_key(ws.cell(row, col).value):
                    area_col = col
                    break
            if area_col:
                break
    if not code_col or not area_col:
        return {}

    values: dict[str, float] = {}
    for row in range(1, ws.max_row + 1):
        code = str(ws.cell(row, code_col).value or "").strip().upper()
        name = normalize_key(ws.cell(row, max(1, code_col - 1)).value)
        if not code and "tong dien tich tu nhien" in name:
            code = "DTTN"
        area = parse_number(ws.cell(row, area_col).value)
        if code and area is not None:
            values[code] = area
    return values
def color(value) -> str | None:
    if not value:
        return None
    if getattr(value, "type", None) == "rgb" and value.rgb:
        rgb = value.rgb[-6:]
        if rgb == "000000" and str(value.rgb).startswith("00"):
            return None
        return f"#{rgb}"
    return None


def border_css(side) -> str:
    if side is None or side.style is None:
        return "1px solid #c8d0d9"
    width = "2px" if side.style in {"medium", "thick", "double"} else "1px"
    clr = color(side.color) or "#2f3640"
    return f"{width} solid {clr}"


def style_key(cell) -> str:
    fill = None
    if cell.fill and cell.fill.fill_type == "solid":
        fill = color(cell.fill.fgColor)
    font_color = color(cell.font.color)
    horizontal = cell.alignment.horizontal or "center"
    if horizontal == "centerContinuous":
        horizontal = "center"
    if horizontal in {"general", "distributed", "justify"}:
        horizontal = "left"
    parts = [
        f"background:{fill}" if fill else "",
        f"font-weight:{'700' if cell.font.bold else '400'}",
        f"font-style:{'italic' if cell.font.italic else 'normal'}",
        f"font-size:{int(cell.font.sz or 11)}pt",
        f"color:{font_color or '#17202a'}",
        f"text-align:{horizontal}",
        f"vertical-align:{cell.alignment.vertical or 'middle'}",
        f"white-space:{'normal' if cell.alignment.wrap_text else 'nowrap'}",
        f"border-top:{border_css(cell.border.top)}",
        f"border-right:{border_css(cell.border.right)}",
        f"border-bottom:{border_css(cell.border.bottom)}",
        f"border-left:{border_css(cell.border.left)}",
    ]
    return ";".join(p for p in parts if p)


def display_value(value, code: str = "", col: int | None = None) -> str:
    if value is None:
        return ""
    if isinstance(value, str) and value.startswith("="):
        return ""
    if col == 1 and code in STT_FIXES_BY_CODE:
        return STT_FIXES_BY_CODE[code]
    text = str(value)
    return LAND_NAME_FIXES.get(text.strip(), text)


def main() -> None:
    wb = openpyxl.load_workbook(SOURCE, data_only=False)
    ws = wb["Sheet1"]
    previous_plan_values = {}
    total_columns = ws.max_column + 1

    merged_parent = {}
    merged_skip = set()
    for rng in ws.merged_cells.ranges:
        min_col, min_row, max_col, max_row = rng.bounds
        merged_parent[(min_row, min_col)] = (max_row - min_row + 1, max_col - min_col + 1)
        for row in range(min_row, max_row + 1):
            for col in range(min_col, max_col + 1):
                if (row, col) != (min_row, min_col):
                    merged_skip.add((row, col))

    code_rows = {}
    for row in range(1, ws.max_row + 1):
        code = ws.cell(row, 3).value
        if code is not None:
            code_rows[str(code).strip()] = row

    code_cols = {}
    for col in range(MATRIX_START_COL, MATRIX_END_COL + 1):
        code = ws.cell(HEADER_ROW, col).value
        if code is not None:
            code_cols[str(code).strip()] = col

    direct_children: dict[str, list[str]] = {}
    for code, row in code_rows.items():
        value = ws.cell(row, CURRENT_COL).value
        if not (isinstance(value, str) and value.startswith("=")):
            continue
        child_codes = []
        for child_row in [int(x) for x in re.findall(r"D(\d+)", value)]:
            child_code = ws.cell(child_row, 3).value
            if child_code is not None:
                child_codes.append(str(child_code).strip())
        if child_codes:
            direct_children[code] = child_codes

    all_data_codes = [
        str(ws.cell(row, 3).value).strip()
        for row in range(5, TOTAL_INCREASE_ROW)
        if ws.cell(row, 3).value is not None
    ]
    parent_codes = set(direct_children)
    input_codes = [code for code in all_data_codes if code not in parent_codes and code in code_cols]
    missing_codes: list[str] = []

    styles = {}
    style_names = {}
    css_rules = []
    for row in range(1, ws.max_row + 1):
        for col in range(1, ws.max_column + 1):
            key = style_key(ws.cell(row, col))
            if key not in style_names:
                name = f"xl{len(style_names) + 1}"
                style_names[key] = name
                css_rules.append(f".{name}{{{key}}}")

    colgroup = []
    for col in range(1, ws.max_column + 1):
        letter = get_column_letter(col)
        width = ws.column_dimensions[letter].width or 8
        px = int(width * 8)
        colgroup.append(f'<col style="width:{px}px;min-width:{px}px">')
    colgroup.append('<col style="width:112px;min-width:112px">')

    rows_html = []
    for row in range(1, ws.max_row + 1):
        height = ws.row_dimensions[row].height or 30
        if row == 1:
            rows_html.append(
                f'<tr style="height:{max(height, 42)}px">'
                f'<td class="sheet-title" data-addr="A1" data-row="1" data-col="1" colspan="{total_columns}">'
                'BẢNG CHU CHUYỂN ĐẤT ĐAI'
                '</td></tr>'
            )
            continue
        cells = []
        for col in range(1, ws.max_column + 1):
            if (row, col) in merged_skip:
                continue
            cell = ws.cell(row, col)
            rowspan, colspan = merged_parent.get((row, col), (1, 1))
            cls = style_names[style_key(cell)]
            addr = f"{get_column_letter(col)}{row}"
            code = str(ws.cell(row, 3).value or "").strip()
            col_code = str(ws.cell(HEADER_ROW, col).value or "").strip()
            is_current_input = col == CURRENT_COL and code in input_codes
            is_matrix_input = row in [code_rows[c] for c in input_codes] and col in [code_cols[c] for c in input_codes]
            is_input = is_current_input or is_matrix_input
            attrs = [
                f'class="{cls}"',
                f'data-addr="{addr}"',
                f'data-row="{row}"',
                f'data-col="{col}"',
            ]
            if rowspan > 1:
                attrs.append(f'rowspan="{rowspan}"')
            if colspan > 1:
                attrs.append(f'colspan="{colspan}"')
            if code:
                attrs.append(f'data-code="{html.escape(code)}"')
            if col_code:
                attrs.append(f'data-col-code="{html.escape(col_code)}"')

            text = html.escape(display_value(cell.value, code, col))
            if is_input:
                attrs.append('data-input="1"')
                value = "" if cell.value is None or (isinstance(cell.value, str) and cell.value.startswith("=")) else html.escape(str(cell.value))
                content = f'<input inputmode="decimal" value="{value}" aria-label="{html.escape(addr)}">'
            elif col >= CURRENT_COL and row >= 4:
                attrs.append('data-auto="1"')
                content = f'<span class="value">{text}</span>'
            else:
                content = text
            cells.append(f"<td {' '.join(attrs)}>{content}</td>")
        if row == 2:
            cells.append(
                f'<td class="xl3" data-addr="{get_column_letter(PREVIOUS_PLAN_COL)}{row}" '
                f'data-row="{row}" data-col="{PREVIOUS_PLAN_COL}" rowspan="2">'
                'Quy hoạch kỳ trước</td>'
            )
        elif row >= 4:
            previous_code = str(ws.cell(row, 3).value or "").strip().upper()
            if row == 4:
                previous_code = "DTTN"
            previous_text = html.escape(format_ha(previous_plan_values.get(previous_code)))
            cells.append(
                f'<td class="xl7" data-addr="{get_column_letter(PREVIOUS_PLAN_COL)}{row}" '
                f'data-row="{row}" data-col="{PREVIOUS_PLAN_COL}" data-previous-plan="1" data-auto="1">'
                f'<span class="value">{previous_text}</span></td>'
            )
        rows_html.append(f'<tr style="height:{height}px">{"".join(cells)}</tr>')

    meta = {
        "inputCodes": input_codes,
        "missingCodes": missing_codes,
        "directChildren": direct_children,
        "codeRows": code_rows,
        "codeCols": code_cols,
        "dttnRow": 4,
        "currentCol": CURRENT_COL,
        "matrixStartCol": MATRIX_START_COL,
        "matrixEndCol": MATRIX_END_COL,
        "decreaseCol": DECREASE_COL,
        "changeCol": CHANGE_COL,
        "planCol": PLAN_COL,
        "previousPlanCol": PREVIOUS_PLAN_COL,
        "totalIncreaseRow": TOTAL_INCREASE_ROW,
        "planRow": PLAN_ROW,
        "tolerance": 0.0001,
    }

    meta_json = json.dumps(meta, ensure_ascii=False).replace("</", "<\\/")

    jszip_js = JSZIP.read_text(encoding="utf-8")
    logo_data_url = ""
    if LOGO.exists():
        logo_data_url = "data:image/jpeg;base64," + base64.b64encode(LOGO.read_bytes()).decode("ascii")
    home_bg_data_url = ""
    if HOME_BACKGROUND.exists():
        home_bg_data_url = "data:image/png;base64," + base64.b64encode(HOME_BACKGROUND.read_bytes()).decode("ascii")
    SAMPLE_DIR.mkdir(parents=True, exist_ok=True)
    LEGACY_SAMPLE_DIR.mkdir(parents=True, exist_ok=True)
    sample_links = []
    for source_name, public_name, label in SAMPLE_FILES:
        source_path = PREVIOUS_PLAN_DIR / source_name
        if not source_path.exists():
            continue
        shutil.copy2(source_path, SAMPLE_DIR / public_name)
        shutil.copy2(source_path, LEGACY_SAMPLE_DIR / public_name)
        payload = base64.b64encode(source_path.read_bytes()).decode("ascii")
        href = f"data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{payload}"
        sample_links.append(
            f'<a href="{href}" download="{html.escape(public_name, quote=True)}">{html.escape(label)}</a>'
        )
    sample_links_html = "\n      ".join(sample_links)
    webgis_data_dir = OUT.parent / "webgis"
    webgis_data_dir.mkdir(parents=True, exist_ok=True)
    webgis_sample_json_pretty = json.dumps(WEBGIS_SAMPLE_DATA, ensure_ascii=False, indent=2)
    (webgis_data_dir / "sample-land-data.geojson").write_text(webgis_sample_json_pretty, encoding="utf-8")
    webgis_sample_json = json.dumps(WEBGIS_SAMPLE_DATA, ensure_ascii=False).replace("</", "<\\/")
    doc = f"""<!doctype html>
<html lang="vi">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Phần mềm chu chuyển đất đai</title>
<link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css">
<style>
:root {{
  --bg: #eef4f1;
  --panel: #ffffff;
  --ink: #17202a;
  --muted: #64748b;
  --line: #d5dde8;
  --accent: #0f766e;
  --accent-2: #2563eb;
  --surface: rgba(255, 255, 255, 0.88);
  --warn: #b42318;
  --input: #fff8d9;
  --diagonal: #dcfce7;
  --auto: #f8fafc;
  --header: #e8f1f7;
  --locked: #f7f9fc;
}}
* {{ box-sizing: border-box; }}
html {{ min-height: 100%; }}
body {{
  margin: 0;
  min-height: 100vh;
  background:
    linear-gradient(180deg, rgba(255, 255, 255, 0.66), rgba(255, 255, 255, 0.28)),
    linear-gradient(135deg, #eef4f1 0%, #f8fafc 52%, #edf3f7 100%);
  background-attachment: fixed;
  color: var(--ink);
  font-family: Arial, Helvetica, sans-serif;
}}
.appbar {{
  position: sticky;
  top: 0;
  z-index: 50;
  display: flex;
  flex-wrap: wrap;
  align-items: flex-start;
  gap: 10px 14px;
  padding: 10px 14px;
  background: linear-gradient(135deg, rgba(255, 255, 255, 0.96), rgba(244, 250, 247, 0.92));
  border-bottom: 1px solid rgba(148, 163, 184, 0.42);
  box-shadow: 0 10px 28px rgba(15, 23, 42, 0.10);
  backdrop-filter: blur(12px);
}}
.title {{
  font-size: 15px;
  font-weight: 700;
  text-transform: uppercase;
  color: #0f3d31;
}}
.subtitle {{
  color: #1e4d5f;
  font-size: 12px;
  font-weight: 700;
}}
.brand {{
  display: flex;
  align-items: center;
  gap: 9px;
  flex: 0 1 310px;
  min-width: 250px;
}}
.brand-logo {{
  width: 40px;
  height: 40px;
  border-radius: 50%;
  object-fit: cover;
  border: 2px solid #ffffff;
  box-shadow: 0 4px 12px rgba(15, 23, 42, 0.20);
  flex: 0 0 auto;
}}
.brand-text {{
  display: flex;
  flex-direction: column;
  gap: 1px;
  line-height: 1.15;
}}
.designer {{
  color: #64748b;
  font-size: 11px;
  font-weight: 600;
}}
.status {{
  display: flex;
  flex: 1 1 260px;
  min-width: 220px;
  gap: 8px;
  flex-wrap: wrap;
  align-items: center;
  color: var(--muted);
  font-size: 12px;
}}
.quick-save {{
  flex: 0 0 auto;
  margin-left: auto;
}}
.main-menu {{
  position: relative;
  flex: 0 0 auto;
}}
.menu-trigger {{
  min-width: 92px;
  font-weight: 700;
}}
.menu-list {{
  position: absolute;
  top: calc(100% + 6px);
  left: 0;
  z-index: 80;
  min-width: 220px;
  padding: 6px;
  border: 1px solid rgba(100, 116, 139, 0.34);
  background: rgba(255, 255, 255, 0.98);
  box-shadow: 0 18px 36px rgba(15, 23, 42, 0.18);
}}
.menu-list[hidden] {{
  display: none;
}}
.menu-list button {{
  width: 100%;
  justify-content: flex-start;
  text-align: left;
  border-color: transparent;
  background: transparent;
  box-shadow: none;
}}
.menu-list button:hover {{
  background: #f0fdfa;
  filter: none;
}}
.home-page {{
  min-height: calc(100vh - 74px);
  margin: 14px;
  border: 1px solid rgba(148, 163, 184, 0.24);
  border-radius: 8px;
  background:
    linear-gradient(180deg, rgba(6, 25, 65, 0.05), rgba(6, 25, 65, 0.10)),
    url("{home_bg_data_url}") center / cover no-repeat,
    linear-gradient(135deg, #0752b7, #52c7e8);
  box-shadow: inset 0 0 0 1px rgba(255, 255, 255, 0.12), 0 18px 42px rgba(15, 23, 42, 0.18);
}}
body.home-mode .module-only,
body.home-mode .table-wrap,
body.home-mode #importLog {{
  display: none;
}}
body.docs-mode .module-only,
body.webgis-mode .module-only,
body.docs-mode .table-wrap,
body.webgis-mode .table-wrap,
body.docs-mode #importLog,
body.webgis-mode #importLog {{
  display: none;
}}
body.module-mode .home-page,
body.module-mode .docs-page,
body.module-mode .webgis-page,
body.home-mode .docs-page,
body.home-mode .webgis-page,
body.docs-mode .home-page,
body.docs-mode .webgis-page,
body.webgis-mode .home-page,
body.webgis-mode .docs-page {{
  display: none;
}}
.docs-page {{
  min-height: calc(100vh - 74px);
  margin: 14px;
  border: 1px solid rgba(148, 163, 184, 0.28);
  border-radius: 8px;
  background: rgba(255, 255, 255, 0.92);
  box-shadow: 0 18px 42px rgba(15, 23, 42, 0.12);
}}
.webgis-page {{
  min-height: calc(100vh - 74px);
  margin: 14px;
  border: 1px solid rgba(148, 163, 184, 0.28);
  border-radius: 8px;
  background: #ffffff;
  box-shadow: 0 18px 42px rgba(15, 23, 42, 0.10);
}}
{WEBGIS_CSS}
.library-shell {{
  min-height: calc(100vh - 104px);
  padding: 18px;
}}
.library-head {{
  display: flex;
  align-items: flex-start;
  justify-content: space-between;
  gap: 14px;
  padding: 16px;
  border: 1px solid rgba(148, 163, 184, 0.28);
  border-radius: 10px;
  background: linear-gradient(135deg, rgba(240, 253, 250, 0.98), rgba(239, 246, 255, 0.98));
}}
.library-head h1 {{
  margin: 0;
  color: #0f172a;
  font-size: 23px;
  line-height: 1.2;
}}
.library-head p {{
  margin: 6px 0 0;
  max-width: 760px;
  color: #475569;
  font-size: 13px;
}}
.library-head-actions,
.reader-tools,
.library-admin-actions {{
  display: flex;
  align-items: center;
  flex-wrap: wrap;
  gap: 8px;
}}
.library-session-badge {{
  display: inline-flex;
  align-items: center;
  min-height: 28px;
  padding: 4px 9px;
  border: 1px solid #99f6e4;
  border-radius: 999px;
  color: #115e59;
  background: #f0fdfa;
  font-size: 12px;
  font-weight: 700;
}}
.library-controls {{
  display: grid;
  grid-template-columns: minmax(220px, 1fr) minmax(150px, 220px) minmax(130px, 180px) auto;
  gap: 9px;
  margin: 14px 0;
  padding: 10px;
  border: 1px solid rgba(148, 163, 184, 0.28);
  border-radius: 10px;
  background: #fff;
}}
.library-controls input,
.library-controls select,
.library-access input,
.library-admin input,
.library-admin select,
.library-admin textarea,
.reader-page-input {{
  height: 34px;
  border: 1px solid #cbd5e1;
  border-radius: 7px;
  padding: 5px 9px;
  color: #0f172a;
  background: #fff;
  font-size: 13px;
}}
.library-admin textarea {{
  min-height: 72px;
  resize: vertical;
}}
.library-grid {{
  display: grid;
  grid-template-columns: repeat(auto-fill, minmax(230px, 260px));
  gap: 14px;
  align-items: stretch;
}}
.library-card {{
  display: flex;
  flex-direction: column;
  width: 100%;
  min-height: 358px;
  min-width: 0;
  overflow: hidden;
  border: 1px solid rgba(148, 163, 184, 0.30);
  border-radius: 10px;
  background: #fff;
  box-shadow: 0 12px 28px rgba(15, 23, 42, 0.10);
}}
.library-cover {{
  position: relative;
  height: 164px;
  overflow: hidden;
  background: linear-gradient(135deg, #e0f2fe, #dcfce7);
  display: grid;
  place-items: center;
}}
.library-cover img {{
  display: block;
  width: 100%;
  height: 100%;
  object-fit: cover;
}}
.library-cover::after {{
  content: "";
  position: absolute;
  inset: 0;
  pointer-events: none;
  background: linear-gradient(180deg, rgba(255,255,255,0.04), rgba(15,23,42,0.08));
}}
.library-cover-placeholder {{
  max-width: 100%;
  padding: 12px;
  color: #0f172a;
  font-size: 15px;
  font-weight: 700;
  line-height: 1.3;
  text-align: center;
  overflow: hidden;
  overflow-wrap: anywhere;
  display: -webkit-box;
  -webkit-line-clamp: 3;
  -webkit-box-orient: vertical;
}}
.library-card-body {{
  display: flex;
  flex-direction: column;
  gap: 7px;
  flex: 1;
  min-width: 0;
  padding: 12px;
}}
.library-card h3 {{
  margin: 0;
  color: #0f172a;
  font-size: 15px;
  line-height: 1.3;
  min-height: 39px;
  overflow: hidden;
  overflow-wrap: anywhere;
  display: -webkit-box;
  -webkit-line-clamp: 2;
  -webkit-box-orient: vertical;
}}
.library-meta {{
  display: flex;
  flex-wrap: wrap;
  gap: 5px;
  color: #475569;
  font-size: 12px;
  min-width: 0;
}}
.library-author {{
  display: block;
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
}}
.library-pill {{
  display: inline-flex;
  align-items: center;
  max-width: 100%;
  min-height: 22px;
  padding: 2px 7px;
  border: 1px solid #cbd5e1;
  border-radius: 999px;
  background: #f8fafc;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}}
.library-description {{
  flex: 1;
  color: #475569;
  font-size: 13px;
  line-height: 1.45;
  overflow: hidden;
  overflow-wrap: anywhere;
  display: -webkit-box;
  -webkit-line-clamp: 3;
  -webkit-box-orient: vertical;
}}
.library-read-btn {{
  margin-top: auto;
}}
.library-empty {{
  padding: 28px;
  border: 1px dashed #cbd5e1;
  border-radius: 10px;
  color: #64748b;
  text-align: center;
  background: #f8fafc;
}}
.library-admin,
.library-access,
.pdf-reader {{
  position: fixed;
  inset: 88px 18px 18px;
  z-index: 120;
  overflow: auto;
  border: 1px solid rgba(100, 116, 139, 0.34);
  border-radius: 12px;
  background: #ffffff;
  box-shadow: 0 24px 70px rgba(15, 23, 42, 0.30);
}}
.library-admin[hidden],
.library-access[hidden],
.pdf-reader[hidden] {{
  display: none;
}}
.library-access {{
  z-index: 130;
  display: grid;
  place-items: center;
  background: rgba(15, 23, 42, 0.22);
}}
.library-access-card {{
  width: min(420px, calc(100vw - 32px));
  padding: 18px;
  border: 1px solid rgba(148, 163, 184, 0.36);
  border-radius: 12px;
  background: #ffffff;
  box-shadow: 0 24px 70px rgba(15, 23, 42, 0.26);
}}
.library-access-card h2 {{
  margin: 0;
  color: #0f172a;
  font-size: 18px;
}}
.library-access-card p {{
  margin: 8px 0 12px;
  color: #475569;
  font-size: 13px;
  line-height: 1.45;
}}
.library-access-form {{
  display: grid;
  gap: 10px;
}}
.library-access-form label {{
  display: grid;
  gap: 4px;
  color: #475569;
  font-size: 12px;
}}
.library-access-hint {{
  padding: 10px 12px;
  border: 1px solid #dbeafe;
  border-radius: 8px;
  background: #eff6ff;
}}
.library-admin-inner {{
  display: grid;
  grid-template-columns: minmax(240px, 300px) minmax(300px, 420px) minmax(0, 1fr);
  gap: 14px;
  padding: 14px;
}}
.library-admin-toolbar {{
  position: sticky;
  top: 0;
  z-index: 2;
  display: flex;
  align-items: center;
  gap: 12px;
  padding: 14px;
  border-bottom: 1px solid #e2e8f0;
  background: rgba(255, 255, 255, 0.96);
}}
.library-admin-toolbar h2 {{
  flex: 1;
  margin: 0;
  color: #0f172a;
  font-size: 18px;
}}
.library-admin-card {{
  border: 1px solid rgba(148, 163, 184, 0.30);
  border-radius: 10px;
  padding: 12px;
  background: #f8fafc;
}}
.library-admin-card[hidden] {{
  display: none;
}}
.library-admin-status {{
  padding: 10px 12px;
  border: 1px solid #bbf7d0;
  border-radius: 8px;
  color: #166534;
  background: #f0fdf4;
  font-size: 13px;
  line-height: 1.45;
}}
.library-admin-status.error {{
  border-color: #fecaca;
  color: #991b1b;
  background: #fff1f2;
}}
.library-admin-card h2,
.library-admin-card h3 {{
  margin: 0 0 10px;
  color: #0f172a;
  font-size: 16px;
}}
.library-admin-form {{
  display: grid;
  gap: 9px;
}}
.library-admin-form label {{
  display: grid;
  gap: 4px;
  color: #475569;
  font-size: 12px;
}}
.library-admin-table {{
  width: 100%;
  border-collapse: collapse;
  font-size: 12px;
}}
.library-admin-table th,
.library-admin-table td {{
  border-bottom: 1px solid #e2e8f0;
  padding: 8px 6px;
  text-align: left;
  vertical-align: top;
}}
.library-admin-table th {{
  color: #334155;
  background: #f1f5f9;
}}
.pdf-reader {{
  display: flex;
  flex-direction: column;
  background: #f8fafc;
}}
.reader-topbar {{
  position: sticky;
  top: 0;
  z-index: 2;
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 10px;
  padding: 10px;
  border-bottom: 1px solid #cbd5e1;
  background: rgba(255, 255, 255, 0.96);
}}
.reader-title {{
  min-width: 180px;
  color: #0f172a;
  font-size: 14px;
  font-weight: 700;
}}
.reader-page-input {{
  width: 72px;
  text-align: center;
}}
.reader-notice {{
  padding: 9px 14px;
  color: #7a271a;
  border-bottom: 1px solid #fed7aa;
  background: #fff7ed;
  font-size: 13px;
}}
.pdf-stage {{
  flex: 1;
  overflow: auto;
  display: grid;
  place-items: start center;
  padding: 18px;
  user-select: none;
  -webkit-user-select: none;
}}
.pdf-canvas-wrap {{
  position: relative;
  max-width: 100%;
  padding: 10px;
  border: 1px solid #cbd5e1;
  border-radius: 8px;
  background: #fff;
  box-shadow: 0 16px 40px rgba(15, 23, 42, 0.18);
}}
#pdfCanvas {{
  display: block;
  max-width: 100%;
  height: auto;
  user-select: none;
  -webkit-user-select: none;
}}
@media (max-width: 820px) {{
  .library-head {{
    flex-direction: column;
  }}
  .library-controls {{
    grid-template-columns: 1fr;
  }}
  .library-admin,
  .library-access,
  .pdf-reader {{
    inset: 74px 8px 8px;
  }}
  .library-admin-inner {{
    grid-template-columns: 1fr;
  }}
  .reader-topbar {{
    align-items: flex-start;
    flex-direction: column;
  }}
}}
.badge {{
  display: inline-flex;
  align-items: center;
  min-height: 24px;
  padding: 4px 8px;
  border: 1px solid var(--line);
  background: rgba(248, 250, 252, 0.88);
  border-radius: 6px;
}}
.badge.warn {{
  color: #7a271a;
  border-color: #f4b0a1;
  background: #fff1ed;
}}
.actions {{
  display: flex;
  flex: 1 1 100%;
  flex-wrap: wrap;
  justify-content: flex-start;
  gap: 8px;
  align-items: flex-start;
}}
.tool-group {{
  position: relative;
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  gap: 0;
  min-height: 36px;
  padding: 0;
  border: 1px solid rgba(148, 163, 184, 0.36);
  border-radius: 8px;
  background: rgba(255, 255, 255, 0.72);
}}
.tool-group-title {{
  height: 34px;
  border: 0;
  border-radius: 8px;
  background: transparent;
  color: #334155;
  font-size: 12px;
  font-weight: 700;
  line-height: 34px;
  text-transform: uppercase;
  padding: 0 11px;
  cursor: pointer;
  box-shadow: none;
}}
.tool-group-title::after {{
  content: "▾";
  margin-left: 7px;
  font-size: 10px;
  color: #64748b;
}}
.tool-group.open .tool-group-title::after {{
  content: "▴";
}}
.tool-items {{
  position: absolute;
  top: calc(100% + 6px);
  left: 0;
  z-index: 85;
  display: none;
  flex-wrap: wrap;
  gap: 6px;
  min-width: 260px;
  max-width: min(520px, calc(100vw - 28px));
  padding: 8px;
  border: 1px solid rgba(100, 116, 139, 0.34);
  border-radius: 8px;
  background: rgba(255, 255, 255, 0.98);
  box-shadow: 0 18px 36px rgba(15, 23, 42, 0.18);
}}
.tool-group.open .tool-items {{
  display: flex;
}}
.tool-group:nth-last-child(-n+2) .tool-items {{
  left: auto;
  right: 0;
}}
.project-items {{
  min-width: min(420px, calc(100vw - 28px));
  gap: 10px;
}}
.project-section {{
  display: grid;
  grid-template-columns: 1fr;
  gap: 7px;
  width: 100%;
  padding: 8px;
  border: 1px solid rgba(226, 232, 240, 0.9);
  border-radius: 8px;
  background: #f8fafc;
}}
.project-section strong {{
  color: #0f172a;
  font-size: 12px;
}}
.project-field {{
  display: grid;
  grid-template-columns: 135px minmax(0, 1fr);
  gap: 8px;
  align-items: center;
  color: #475569;
  font-size: 12px;
}}
.project-field input {{
  min-width: 0;
  width: 100%;
  height: 30px;
  border: 1px solid #cbd5e1;
  border-radius: 6px;
  padding: 4px 8px;
  background: #fff;
  color: #0f172a;
  font-size: 12px;
}}
.project-actions {{
  display: flex;
  justify-content: flex-end;
  align-items: center;
  gap: 8px;
  width: 100%;
}}
.project-actions button {{
  min-width: 112px;
}}
.project-db-status {{
  flex: 1 1 auto;
  min-width: 160px;
  color: #64748b;
  font-size: 12px;
}}
.sample-downloads {{
  position: relative;
  display: flex;
  flex: 0 0 auto;
  justify-content: flex-start;
  flex-wrap: nowrap;
  gap: 0;
  align-items: center;
  font-size: 12px;
  color: #475569;
}}
.sample-downloads > span {{
  height: 34px;
  line-height: 34px;
  padding: 0 11px;
  border: 1px solid rgba(148, 163, 184, 0.36);
  border-radius: 8px;
  background: rgba(255, 255, 255, 0.72);
  color: #334155;
  font-size: 12px;
  font-weight: 700;
  text-transform: uppercase;
  cursor: pointer;
}}
.sample-downloads > span::after {{
  content: "▾";
  margin-left: 7px;
  font-size: 10px;
  color: #64748b;
}}
.sample-downloads.open > span::after {{
  content: "▴";
}}
.sample-items {{
  position: absolute;
  top: calc(100% + 6px);
  right: 0;
  z-index: 85;
  display: none;
  flex-wrap: wrap;
  gap: 6px;
  min-width: 280px;
  max-width: min(520px, calc(100vw - 28px));
  padding: 8px;
  border: 1px solid rgba(100, 116, 139, 0.34);
  border-radius: 8px;
  background: rgba(255, 255, 255, 0.98);
  box-shadow: 0 18px 36px rgba(15, 23, 42, 0.18);
}}
.sample-downloads.open .sample-items {{
  display: flex;
}}
.sample-downloads a {{
  color: #0f766e;
  text-decoration: none;
  border: 1px solid rgba(15, 118, 110, 0.24);
  background: rgba(240, 253, 250, 0.82);
  padding: 4px 8px;
}}
.sample-downloads a:hover {{
  border-color: rgba(15, 118, 110, 0.52);
  background: #ccfbf1;
}}
.search-box {{
  display: flex;
  align-items: center;
  gap: 4px;
  padding: 0;
  border: 1px solid rgba(100, 116, 139, 0.42);
  background: rgba(255, 255, 255, 0.68);
  border-radius: 6px;
}}
.search-box input {{
  width: 96px;
  height: 28px;
  min-height: 28px;
  border: 0;
  background: #ffffff;
  padding: 0 8px;
  text-align: left;
  text-transform: uppercase;
}}
.search-box button {{
  height: 28px;
  padding: 0 8px;
  border-top: 0;
  border-right: 0;
  border-bottom: 0;
}}
.import-options {{
  display: flex;
  align-items: center;
  gap: 6px;
  font-size: 12px;
  color: var(--muted);
}}
.import-options input,
.view-options input,
.report-option input {{
  width: auto;
  height: auto;
  min-height: 0;
}}
select {{
  height: 32px;
  border: 1px solid rgba(100, 116, 139, 0.62);
  border-radius: 6px;
  background: rgba(255, 255, 255, 0.92);
  color: #0f172a;
  padding: 0 8px;
  font-size: 13px;
}}
button {{
  height: 32px;
  border: 1px solid rgba(100, 116, 139, 0.62);
  border-radius: 6px;
  background: linear-gradient(180deg, #ffffff, #f8fafc);
  color: #0f172a;
  padding: 0 11px;
  font-size: 13px;
  cursor: pointer;
  box-shadow: 0 1px 2px rgba(15, 23, 42, 0.08);
}}
button.primary {{
  border-color: var(--accent);
  background: linear-gradient(180deg, #158176, #0f766e);
  color: #ffffff;
  font-weight: 700;
  box-shadow: 0 6px 14px rgba(15, 118, 110, 0.20);
}}
button:hover {{ filter: brightness(0.97); }}
.table-toolbar {{
  display: flex;
  flex-wrap: wrap;
  justify-content: space-between;
  align-items: center;
  gap: 8px 14px;
  margin: 12px 14px -4px;
  padding: 8px 10px;
  border: 1px solid rgba(148, 163, 184, 0.34);
  border-radius: 8px;
  background: rgba(255, 255, 255, 0.86);
  box-shadow: 0 8px 22px rgba(15, 23, 42, 0.08);
}}
.legend {{
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  gap: 8px 12px;
  color: #334155;
  font-size: 12px;
}}
.legend-item {{
  display: inline-flex;
  align-items: center;
  gap: 5px;
  white-space: nowrap;
}}
.swatch {{
  width: 16px;
  height: 16px;
  border: 1px solid #cbd5e1;
  border-radius: 4px;
}}
.swatch.input {{ background: var(--input); }}
.swatch.diagonal {{ background: var(--diagonal); }}
.swatch.auto {{ background: var(--auto); }}
.swatch.locked {{ background: #ffffff; }}
.view-options {{
  display: flex;
  align-items: center;
  gap: 8px;
  color: #334155;
  font-size: 12px;
}}
.table-wrap {{
  height: calc(100vh - 184px);
  min-height: 420px;
  overflow: auto;
  margin: 14px;
  border: 1px solid rgba(148, 163, 184, 0.44);
  border-radius: 8px;
  background:
    linear-gradient(rgba(255, 255, 255, 0.94), rgba(255, 255, 255, 0.94)),
    repeating-linear-gradient(135deg, rgba(15, 118, 110, 0.05) 0 12px, rgba(37, 99, 235, 0.04) 12px 24px);
  box-shadow: 0 18px 40px rgba(15, 23, 42, 0.14);
  scrollbar-color: #94a3b8 #e2e8f0;
  scrollbar-width: thin;
}}
.table-wrap::-webkit-scrollbar {{
  width: 14px;
  height: 14px;
}}
.table-wrap::-webkit-scrollbar-track {{
  background: #e2e8f0;
}}
.table-wrap::-webkit-scrollbar-thumb {{
  background: #94a3b8;
  border: 3px solid #e2e8f0;
  border-radius: 999px;
}}
table {{
  border-collapse: collapse;
  table-layout: fixed;
  width: max-content;
  background: #ffffff;
}}
.sheet-title {{
  height: 42px;
  background: #f8fafc;
  color: #0f3d31;
  font-size: 18pt;
  font-weight: 700;
  text-align: center;
  vertical-align: middle;
  letter-spacing: 0;
  border: 1px solid #c8d0d9;
}}
td {{
  position: relative;
  background: #ffffff;
  padding: 4px 6px;
  line-height: 1.25;
  overflow: hidden;
  font-size: 12px;
}}
td[data-row="2"], td[data-row="3"] {{
  position: sticky;
  top: 0;
  z-index: 12;
  font-weight: 700;
  background: var(--header) !important;
  background-clip: padding-box;
}}
td[data-row="2"] {{ top: 0; }}
td[data-row="3"] {{ top: 30px; }}
td[data-col="1"], td[data-col="2"], td[data-col="3"], td[data-col="4"] {{
  position: sticky;
  z-index: 14;
  background: #ffffff;
  background-clip: padding-box;
}}
td[data-col="1"] {{ left: 0; }}
td[data-col="2"] {{ left: 48px; }}
td[data-col="3"] {{ left: 336px; }}
td[data-col="4"] {{ left: 400px; }}
td[data-row="2"][data-col="1"], td[data-row="2"][data-col="2"], td[data-row="2"][data-col="3"], td[data-row="2"][data-col="4"],
td[data-row="3"][data-col="1"], td[data-row="3"][data-col="2"], td[data-row="3"][data-col="3"], td[data-row="3"][data-col="4"] {{
  z-index: 30;
}}
td input {{
  width: 100%;
  height: 100%;
  min-height: 24px;
  border: 0;
  outline: 1px solid transparent;
  background: var(--input);
  text-align: right;
  font: inherit;
  color: #111827;
}}
td input:focus {{
  outline: 2px solid var(--accent);
  background: #ffffff;
}}
td[data-input="1"] {{ background: var(--input) !important; }}
td[data-auto="1"] {{ background-color: var(--locked); }}
td[data-auto="1"][style*="background"], td[data-input="1"] {{
  background-clip: padding-box;
}}
td.diagonal {{
  background: var(--diagonal) !important;
}}
td.diagonal input {{
  background: var(--diagonal) !important;
  font-weight: 700;
}}
td.diagonal input:focus {{
  background: #f1fff4 !important;
}}
.value {{
  display: block;
  text-align: right;
}}
body.hide-zero td.zero-cell .value {{
  visibility: hidden;
}}
body.hide-zero td.zero-cell input:not(:focus) {{
  color: transparent;
  caret-color: #111827;
}}
td.hover-row::after,
td.hover-col::after {{
  content: "";
  position: absolute;
  inset: 0;
  pointer-events: none;
  background: rgba(37, 99, 235, 0.055);
}}
td.hover-cell {{
  outline: 2px solid rgba(37, 99, 235, 0.75);
  outline-offset: -2px;
}}
td.warn {{
  background: #ffe4e6 !important;
  color: #7f1d1d !important;
}}
td.search-hit {{
  outline: 3px solid #2563eb !important;
  outline-offset: -3px;
  box-shadow: inset 0 0 0 2px rgba(255, 255, 255, 0.86), 0 0 0 2px rgba(37, 99, 235, 0.24);
  z-index: 35;
}}
.hidden-input {{ display: none; }}
.import-log {{
  border-bottom: 1px solid var(--line);
  background: rgba(248, 250, 252, 0.92);
  color: #17202a;
  padding: 8px 14px;
  font-size: 13px;
  line-height: 1.4;
}}
.import-log strong {{ font-weight: 700; }}
.import-log ul {{
  margin: 4px 0 0;
  padding-left: 18px;
}}
.report-panel {{
  position: fixed;
  inset: 64px 18px auto auto;
  z-index: 60;
  width: min(620px, calc(100vw - 36px));
}}
.report-card {{
  border: 1px solid rgba(100, 116, 139, 0.45);
  border-radius: 8px;
  background: #ffffff;
  box-shadow: 0 24px 48px rgba(15, 23, 42, 0.20);
  overflow: hidden;
}}
.report-head {{
  display: flex;
  justify-content: space-between;
  align-items: center;
  gap: 10px;
  padding: 10px 12px;
  border-bottom: 1px solid var(--line);
  background: #f8fafc;
}}
.report-controls {{
  display: flex;
  flex-wrap: wrap;
  gap: 6px;
  padding: 10px 12px;
  border-bottom: 1px solid var(--line);
}}
.report-controls input {{
  width: 180px;
  height: 30px;
  min-height: 30px;
  border: 1px solid rgba(100, 116, 139, 0.62);
  background: #ffffff;
  padding: 0 8px;
  text-align: left;
}}
.report-controls input[type="number"] {{
  width: 112px;
}}
.report-options {{
  display: grid;
  grid-template-columns: repeat(auto-fill, minmax(178px, 1fr));
  gap: 6px;
  max-height: 360px;
  overflow: auto;
  padding: 10px 12px;
}}
.report-option {{
  display: flex;
  gap: 6px;
  align-items: flex-start;
  border: 1px solid #e2e8f0;
  background: #f8fafc;
  padding: 6px;
  font-size: 12px;
  line-height: 1.25;
}}
.report-option input {{
  width: auto;
  height: auto;
  min-height: 0;
  margin-top: 2px;
}}
.report-option span {{
  display: block;
}}
.ai-panel {{
  position: fixed;
  inset: 64px 18px auto auto;
  z-index: 70;
  width: min(520px, calc(100vw - 36px));
}}
.ai-card {{
  border: 1px solid rgba(100, 116, 139, 0.45);
  border-radius: 8px;
  background: #ffffff;
  box-shadow: 0 24px 48px rgba(15, 23, 42, 0.22);
  overflow: hidden;
}}
.ai-head {{
  display: flex;
  justify-content: space-between;
  align-items: center;
  gap: 10px;
  padding: 10px 12px;
  border-bottom: 1px solid var(--line);
  background: #f8fafc;
}}
.ai-messages {{
  display: flex;
  flex-direction: column;
  gap: 8px;
  max-height: 340px;
  overflow: auto;
  padding: 12px;
  background: #f8fafc;
}}
.ai-message {{
  border: 1px solid #e2e8f0;
  background: #ffffff;
  padding: 8px 10px;
  line-height: 1.45;
  white-space: pre-wrap;
}}
.ai-message.user {{
  border-color: rgba(15, 118, 110, 0.28);
  background: #f0fdfa;
}}
.ai-controls {{
  display: flex;
  gap: 8px;
  padding: 10px 12px;
  border-top: 1px solid var(--line);
}}
.ai-controls textarea {{
  flex: 1 1 auto;
  min-height: 66px;
  resize: vertical;
  border: 1px solid rgba(100, 116, 139, 0.62);
  padding: 8px;
  font-family: Arial, Helvetica, sans-serif;
  font-size: 13px;
}}
.ai-controls button {{
  height: 34px;
  align-self: flex-end;
}}
@media print {{
  .appbar, .import-log {{ display: none; }}
  .table-wrap {{ height: auto; overflow: visible; }}
  tr:nth-child(2) td, tr:nth-child(3) td, td:nth-child(1), td:nth-child(2), td:nth-child(3), td:nth-child(4) {{
    position: static;
  }}
}}
{chr(10).join(css_rules)}
.table-wrap td {{
  font-size: 12px !important;
  min-height: 28px;
}}
.table-wrap tr {{
  height: 30px;
}}
.table-wrap td[data-row="2"],
.table-wrap td[data-row="3"] {{
  background: var(--header) !important;
  color: #0f172a !important;
  font-weight: 700 !important;
}}
.table-wrap td[data-col="1"],
.table-wrap td[data-col="2"],
.table-wrap td[data-col="3"],
.table-wrap td[data-col="4"] {{
  background-clip: padding-box !important;
}}
.table-wrap td[data-input="1"] {{
  background: var(--input) !important;
}}
.table-wrap td.diagonal,
.table-wrap td.diagonal input {{
  background: var(--diagonal) !important;
}}
.table-wrap td[data-auto="1"]:not(.diagonal) {{
  background-color: var(--locked) !important;
}}
</style>
</head>
<body class="home-mode">
<header class="appbar">
  <div class="brand">
    <img class="brand-logo" src="{logo_data_url}" alt="Logo Nguyễn Quang Huy">
    <div class="brand-text">
      <div class="title">PHẦN MỀM ĐẤT ĐAI</div>
      <div class="subtitle">Biểu chu chuyển sử dụng đất</div>
      <div class="designer">Designed by Nguyễn Quang Huy</div>
    </div>
  </div>
  <nav class="main-menu" aria-label="Menu chức năng">
    <button id="menuBtn" class="menu-trigger" type="button" aria-expanded="false">Menu</button>
    <div id="menuList" class="menu-list" hidden>
      <button id="openLandTransferBtn" type="button">Chu chuyển đất đai</button>
      <button id="openDocumentLibraryBtn" type="button">Thư viện tài liệu</button>
      <button id="openWebGisBtn" type="button">WebGis</button>
    </div>
  </nav>
  <div class="status module-only">
    <span id="statusTotal" class="badge">Đang tính</span>
    <span id="statusRows" class="badge">0 lệch hàng</span>
    <span id="statusMissing" class="badge"></span>
  </div>
  <button class="primary quick-save module-only" id="saveBtn" type="button">Lưu</button>
  <div class="actions module-only">
    <div class="tool-group">
      <button class="tool-group-title" type="button">Dự án</button>
      <div class="tool-items project-items">
        <div class="project-section">
          <strong>Thiết lập đơn vị hành chính</strong>
          <label class="project-field">
            <span>Tên xã</span>
            <input id="projectCommune" type="text" placeholder="Ví dụ: xã An Bình">
          </label>
          <label class="project-field">
            <span>Tỉnh/thành</span>
            <input id="projectProvince" type="text" placeholder="Ví dụ: tỉnh Bắc Ninh">
          </label>
        </div>
        <div class="project-section">
          <strong>Thông tin quy hoạch</strong>
          <label class="project-field">
            <span>Quy hoạch kỳ trước</span>
            <input id="projectPreviousPlanYear" type="text" placeholder="Ví dụ: 2015-2025">
          </label>
          <label class="project-field">
            <span>Năm hiện trạng</span>
            <input id="projectCurrentYear" type="number" min="1900" max="2200" value="2020">
          </label>
          <label class="project-field">
            <span>Kỳ quy hoạch</span>
            <input id="projectPlanYear" type="text" placeholder="Ví dụ: 2025-2035" value="2020-2030">
          </label>
        </div>
        <div class="project-section">
          <strong>Cơ sở dữ liệu dự án (*.gtp)</strong>
          <button id="gtpOpenBtn" type="button">Add file GTP</button>
          <button id="gtpSetupBtn" type="button">Thiết lập nơi lưu GTP</button>
          <button id="gtpSaveBtn" type="button">Lưu vào file GTP</button>
          <span id="gtpStatus" class="project-db-status">Chưa thiết lập file GTP</span>
        </div>
        <div class="project-actions">
          <button id="projectConfirmBtn" class="primary" type="button">Xác nhận</button>
        </div>
      </div>
    </div>
    <div class="tool-group">
      <button class="tool-group-title" type="button">Nhập dữ liệu</button>
      <div class="tool-items">
        <button id="importCurrentBtn" type="button">Nhập hiện trạng XLSX</button>
        <button id="importGisBtn" type="button">Import bảng chồng xếp GIS</button>
        <button id="importPreviousPlanBtn" type="button">Import quy hoạch kỳ trước</button>
        <button id="loadBtn" type="button">Nhập JSON</button>
      </div>
    </div>
    <div class="tool-group">
      <button class="tool-group-title" type="button">Xử lý</button>
      <div class="tool-items">
        <button id="reportBtn" type="button">Xuất tăng/giảm</button>
        <select id="gisImportMode" title="Chế độ xử lý mã đất lạ">
          <option value="add" selected>Tự thêm mã mới</option>
          <option value="known">Chỉ mã đã có</option>
        </select>
        <label class="import-options" title="Áp dụng khi cột diện tích là m2">
          <input id="gisM2ToHa" type="checkbox" checked>
          m2 -> ha
        </label>
        <button id="clearBtn" type="button">Xóa nhập</button>
      </div>
    </div>
    <div class="tool-group">
      <button class="tool-group-title" type="button">Xuất file</button>
      <div class="tool-items">
        <button id="jsonBtn" type="button">Xuất JSON</button>
        <button id="xlsxBtn" type="button">Xuất XLSX</button>
        <button id="csvBtn" type="button">Xuất CSV</button>
        <button id="printBtn" type="button">In</button>
      </div>
    </div>
    <div class="tool-group">
      <button class="tool-group-title" type="button">Công cụ</button>
      <div class="tool-items">
        <button id="homeBtn" type="button">Màn chính</button>
        <div class="search-box">
          <input id="codeSearch" type="search" placeholder="Tìm mã" aria-label="Tìm mã đất">
          <button id="codeSearchBtn" type="button">Tìm</button>
        </div>
      </div>
    </div>
    <div class="tool-group">
      <button class="tool-group-title" type="button">AI</button>
      <div class="tool-items">
        <button id="aiBtn" type="button">Trợ lý AI</button>
      </div>
    </div>
    <div class="sample-downloads">
      <span>Tải file mẫu</span>
      <div class="sample-items">
        {sample_links_html}
      </div>
    </div>
  </div>
  <input id="fileInput" class="hidden-input" type="file" accept="application/json">
  <input id="currentXlsxInput" class="hidden-input" type="file" accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet">
  <input id="previousPlanXlsxInput" class="hidden-input" type="file" accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet">
  <input id="gisXlsxInput" class="hidden-input" type="file" accept=".xlsx,.xls,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel">
  <input id="gtpInput" class="hidden-input" type="file" accept=".gtp,application/json">
</header>
<main id="homePage" class="home-page" aria-label="Trang chính"></main>
<main id="documentLibraryPage" class="docs-page" aria-label="Thư viện tài liệu">
  <section class="library-shell">
    <div class="library-head">
      <div>
        <h1>Thư viện số tài liệu PDF</h1>
        <p>Tài liệu được đọc trực tuyến trong trình xem riêng của phần mềm. File PDF được lưu ở vùng bảo vệ trên server, không đặt trong thư mục public.</p>
      </div>
      <div class="library-head-actions">
        <span id="librarySessionBadge" class="library-session-badge" hidden></span>
        <button id="libraryHomeBtn" type="button">Màn chính</button>
        <button id="libraryLogoutBtn" type="button" hidden>Đăng xuất</button>
        <button id="libraryAdminOpenBtn" class="primary" type="button">Quản trị</button>
      </div>
    </div>
    <div class="library-controls">
      <input id="librarySearch" type="search" placeholder="Tìm theo tên, tác giả, năm, danh mục">
      <select id="libraryCategoryFilter"><option value="">Tất cả danh mục</option></select>
      <select id="libraryYearFilter"><option value="">Tất cả năm</option></select>
      <button id="libraryRefreshBtn" type="button">Làm mới</button>
    </div>
    <div id="libraryGrid" class="library-grid"></div>
    <div id="libraryEmpty" class="library-empty" hidden>Chưa có tài liệu phù hợp.</div>
  </section>
</main>
{WEBGIS_HTML}
<section id="libraryAccessPanel" class="library-access" hidden>
  <div class="library-access-card">
    <div class="library-admin-actions">
      <h2 style="flex:1">&#272;&#259;ng nh&#7853;p th&#432; vi&#7879;n</h2>
      <button id="libraryAccessCloseBtn" type="button">&#272;&#243;ng</button>
    </div>
    <p>Vui l&#242;ng &#273;&#259;ng nh&#7853;p tr&#432;&#7899;c khi v&#224;o th&#432; vi&#7879;n t&#224;i li&#7879;u.</p>
    <div class="library-access-form">
      <label>T&#224;i kho&#7843;n
        <input id="libraryAccessUser" type="text" autocomplete="username">
      </label>
      <label>M&#7853;t kh&#7849;u
        <input id="libraryAccessPassword" type="password" autocomplete="current-password">
      </label>
      <button id="libraryAccessLoginBtn" class="primary" type="button">&#272;&#259;ng nh&#7853;p</button>
      <div id="libraryAccessMsg" class="library-empty" hidden></div>
    </div>
    <p class="library-access-hint">N&#7871;u ch&#432;a c&#243; t&#224;i kho&#7843;n li&#234;n h&#7879; tr&#7921;c ti&#7871;p admin &#273;&#7875; &#273;&#432;&#7907;c cung c&#7845;p!</p>
  </div>
</section>
<section id="libraryAdminPanel" class="library-admin" hidden>
  <div class="library-admin-toolbar">
    <h2>Qu&#7843;n tr&#7883; th&#432; vi&#7879;n PDF</h2>
    <button id="libraryAdminCloseBtn" type="button">&#272;&#243;ng</button>
  </div>
  <div class="library-admin-inner">
    <div class="library-admin-card">
      <h3>&#272;&#259;ng nh&#7853;p qu&#7843;n tr&#7883;</h3>
      <div id="libraryLoginBox" class="library-admin-form">
        <label>Tài khoản quản trị
          <input id="libraryAdminUser" type="text" autocomplete="username">
        </label>
        <label>Mật khẩu
          <input id="libraryAdminPassword" type="password" autocomplete="current-password">
        </label>
        <button id="libraryLoginBtn" class="primary" type="button">Đăng nhập</button>
        <div id="libraryLoginMsg" class="library-empty" hidden></div>
      </div>
      <div id="libraryLoginStatus" class="library-admin-status" hidden></div>
    </div>
    <div id="libraryUploadCard" class="library-admin-card" hidden>
      <h3>Upload t&#224;i li&#7879;u</h3>
      <form id="libraryDocForm" class="library-admin-form" hidden>
        <input id="libraryDocId" type="hidden">
        <label>Tên tài liệu
          <input id="libraryDocTitle" type="text" required>
        </label>
        <label>Tác giả / đơn vị biên soạn
          <input id="libraryDocAuthor" type="text">
        </label>
        <label>Năm xuất bản
          <input id="libraryDocYear" type="number" min="1800" max="2300">
        </label>
        <label>Danh mục tài liệu
          <input id="libraryDocCategory" type="text" list="libraryCategorySuggestions">
        </label>
        <datalist id="libraryCategorySuggestions"></datalist>
        <label>Mô tả ngắn
          <textarea id="libraryDocDescription"></textarea>
        </label>
        <label>File PDF
          <input id="libraryDocPdf" type="file" accept="application/pdf">
        </label>
        <label>Ảnh bìa nếu có
          <input id="libraryDocCover" type="file" accept="image/png,image/jpeg,image/webp,image/svg+xml">
        </label>
        <label class="import-options">
          <input id="libraryDocVisible" type="checkbox" checked>
          Hiển thị tài liệu
        </label>
        <div class="library-admin-actions">
          <button id="libraryDocSaveBtn" class="primary" type="submit">Lưu tài liệu</button>
          <button id="libraryDocNewBtn" type="button">Tạo mới</button>
        </div>
        <div id="libraryAdminMsg" class="library-empty" hidden></div>
      </form>
    </div>
    <div class="library-admin-card">
      <div class="library-admin-actions">
        <h3 style="flex:1">Danh sách tài liệu</h3>
        <button id="libraryAdminReloadBtn" type="button">Tải lại</button>
      </div>
      <div style="overflow:auto">
        <table class="library-admin-table">
          <thead>
            <tr>
              <th>Tên tài liệu</th>
              <th>Danh mục</th>
              <th>Năm</th>
              <th>Trạng thái</th>
              <th>Thao tác</th>
            </tr>
          </thead>
          <tbody id="libraryAdminRows"></tbody>
        </table>
      </div>
    </div>
  </div>
</section>
<section id="pdfReader" class="pdf-reader" hidden>
  <div class="reader-topbar">
    <div id="readerTitle" class="reader-title">Tài liệu PDF</div>
    <div class="reader-tools">
      <button id="readerPrevBtn" type="button">Trang trước</button>
      <input id="readerPageInput" class="reader-page-input" type="number" min="1" value="1">
      <span id="readerPageTotal">/ 1</span>
      <button id="readerNextBtn" type="button">Trang sau</button>
      <button id="readerZoomOutBtn" type="button">Thu nhỏ</button>
      <button id="readerZoomInBtn" type="button">Phóng to</button>
      <button id="readerFullscreenBtn" type="button">Toàn màn hình</button>
      <button id="readerCloseBtn" type="button">Đóng</button>
    </div>
  </div>
  <div class="reader-notice">Tài liệu chỉ được phép đọc trực tuyến, không được sao chép hoặc tải xuống.</div>
  <div id="pdfStage" class="pdf-stage">
    <div id="pdfCanvasWrap" class="pdf-canvas-wrap">
      <canvas id="pdfCanvas"></canvas>
    </div>
  </div>
</section>
<section id="importLog" class="import-log" hidden></section>
<section id="reportPanel" class="report-panel" hidden>
  <div class="report-card">
    <div class="report-head">
      <strong>Xuất thuyết minh cộng tăng/cộng giảm</strong>
      <button id="reportCloseBtn" type="button">Đóng</button>
    </div>
    <div class="report-controls">
      <input id="reportFilter" type="search" placeholder="Lọc mã hoặc tên đất">
      <input id="reportCurrentYear" type="number" min="1900" max="2200" value="2020" title="Năm hiện trạng">
      <input id="reportPlanYear" type="number" min="1900" max="2200" value="2030" title="Năm quy hoạch">
      <button id="reportSelectActiveBtn" type="button">Chọn mã có dữ liệu</button>
      <button id="reportClearBtn" type="button">Bỏ chọn</button>
      <button class="primary" id="reportExportBtn" type="button">Xuất Word</button>
    </div>
    <div id="reportOptions" class="report-options"></div>
  </div>
</section>
<section id="aiPanel" class="ai-panel" hidden>
  <div class="ai-card">
    <div class="ai-head">
      <strong>Trợ lý AI</strong>
      <button id="aiCloseBtn" type="button">Đóng</button>
    </div>
    <div id="aiMessages" class="ai-messages">
      <div class="ai-message">Anh có thể hỏi: “Kiểm tra giúp tôi bảng này có lệch tổng không?”, “LUC tăng giảm thế nào?”, hoặc “Viết nhận xét ngắn về biến động đất”.</div>
    </div>
    <div class="ai-controls">
      <textarea id="aiQuestion" placeholder="Nhập câu hỏi cho AI"></textarea>
      <button id="aiSendBtn" class="primary" type="button">Gửi</button>
    </div>
  </div>
</section>
<section class="table-toolbar module-only">
  <div class="legend" aria-label="Chú giải màu">
    <span class="legend-item"><span class="swatch input"></span>Ô nhập liệu</span>
    <span class="legend-item"><span class="swatch diagonal"></span>Ô giữ nguyên loại đất / đường chéo</span>
    <span class="legend-item"><span class="swatch auto"></span>Ô công thức / tổng hợp</span>
    <span class="legend-item"><span class="swatch locked"></span>Ô khóa không nhập</span>
  </div>
  <label class="view-options">
    <input id="hideZeroToggle" type="checkbox">
    Ẩn ô 0,00
  </label>
</section>
<main class="table-wrap">
<table id="landTable">
<colgroup>{''.join(colgroup)}</colgroup>
<tbody>
{''.join(rows_html)}
</tbody>
</table>
</main>
<script id="meta" type="application/json">{meta_json}</script>
<script id="webgisSampleData" type="application/json">{webgis_sample_json}</script>
<script>{jszip_js}</script>
<script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js" referrerpolicy="no-referrer"></script>
<script>
const meta = JSON.parse(document.getElementById('meta').textContent);
const $ = (sel, root = document) => root.querySelector(sel);
const $$ = (sel, root = document) => Array.from(root.querySelectorAll(sel));
const storageKey = 'land-transfer-html-v1';
const hideZeroKey = 'land-transfer-hide-zero';
const projectId = 'default';
const apiBase = '/api/projects';
const libraryApiBase = '/api/library';
const inputCodes = meta.inputCodes;
const inputSet = new Set(inputCodes);
const rowsByCode = meta.codeRows;
const colsByCode = meta.codeCols;
const rowCodes = Object.fromEntries(Object.entries(rowsByCode).map(([code, row]) => [String(row), code]));
const colCodes = Object.fromEntries(Object.entries(colsByCode).map(([code, col]) => [String(col), code]));
const directChildren = meta.directChildren || {{}};
let matrixCodes = Object.keys(colsByCode);
let calcRowEntries = Object.entries(rowsByCode).filter(([, row]) => row >= meta.dttnRow && row < meta.totalIncreaseRow);
const inputCells = Array.from(document.querySelectorAll('td[data-input="1"]'));
const inputTds = new Map();
const inputEls = new Map();
const cellsByKey = new Map();
const autoSpans = new Map();
const inputKeys = new Set();
const previousWarnCells = new Set();
const previousPlanValues = {{}};
let projectTitlesConfirmed = false;
let gtpFileHandle = null;
let gtpFileName = '';
{WEBGIS_JS}
let libraryDocuments = [];
const librarySessionTokenKey = 'library-session-token';
const librarySessionRoleKey = 'library-session-role';
let librarySessionToken = localStorage.getItem(librarySessionTokenKey) || localStorage.getItem('library-admin-token') || '';
let librarySessionRole = localStorage.getItem(librarySessionRoleKey) || (librarySessionToken ? 'admin' : '');
let libraryAdminToken = librarySessionRole === 'admin' ? librarySessionToken : '';
let activePdf = null;
let activePdfPage = 1;
let activePdfScale = 1.2;
let activePdfRenderTask = null;
let activePdfRenderSerial = 0;
let nextDynamicRow = meta.planRow + 1;
let nextDynamicCol = (meta.previousPlanCol || meta.planCol) + 1;

function isDiagonalMatrixCell(td) {{
  const row = Number(td.dataset.row || 0);
  const col = Number(td.dataset.col || 0);
  if (row < meta.dttnRow || col < meta.matrixStartCol || col > meta.matrixEndCol) return false;
  const rowCode = rowCodes[String(row)];
  const colCode = colCodes[String(col)];
  return Boolean(rowCode && colCode && rowCode === colCode);
}}

function registerCell(td) {{
  const key = `${{td.dataset.row}}:${{td.dataset.col}}`;
  cellsByKey.set(key, td);
  if (td.dataset.input === '1') inputKeys.add(key);
  td.classList.toggle('diagonal', isDiagonalMatrixCell(td));
  const span = td.querySelector('.value');
  if (span) autoSpans.set(key, span);
  const input = td.querySelector('input');
  if (input) {{
    inputTds.set(td.dataset.addr, td);
    inputEls.set(td.dataset.addr, input);
    input.addEventListener('input', scheduleRecalc);
    input.addEventListener('blur', () => {{
      normalizeInputElement(input);
      recalc();
    }});
  }}
}}

document.querySelectorAll('td[data-row][data-col]').forEach(registerCell);

inputCells.forEach(td => {{
  const input = td.querySelector('input');
  inputTds.set(td.dataset.addr, td);
  if (input) inputEls.set(td.dataset.addr, input);
}});

function addr(col, row) {{
  let n = col, s = '';
  while (n > 0) {{
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - m) / 26);
  }}
  return s + row;
}}

function createCell(row, col, content, options = {{}}) {{
  const td = document.createElement('td');
  td.className = options.className || 'xl8';
  td.dataset.addr = addr(col, row);
  td.dataset.row = String(row);
  td.dataset.col = String(col);
  if (options.code) td.dataset.code = options.code;
  if (options.colCode) td.dataset.colCode = options.colCode;
  if (options.input) {{
    td.dataset.input = '1';
    td.innerHTML = `<input inputmode="decimal" value="${{content || ''}}" aria-label="${{td.dataset.addr}}">`;
  }} else if (options.auto) {{
    td.dataset.auto = '1';
    td.innerHTML = '<span class="value"></span>';
  }} else {{
    td.textContent = content || '';
  }}
  registerCell(td);
  return td;
}}

function refreshCalcEntries() {{
  matrixCodes = Object.keys(colsByCode);
  calcRowEntries = Object.entries(rowsByCode).filter(([, row]) => row >= meta.dttnRow);
}}

function addMatrixColumn(code) {{
  const col = nextDynamicCol++;
  colsByCode[code] = col;
  matrixCodes.push(code);
  const colgroup = document.querySelector('#landTable colgroup');
  const colEl = document.createElement('col');
  colEl.style.width = '64px';
  colEl.style.minWidth = '64px';
  colgroup.appendChild(colEl);

  document.querySelectorAll('#landTable tbody tr').forEach(tr => {{
    const row = Number(tr.querySelector('td[data-row]')?.dataset.row || 0);
    let cell;
    if (row === 3) {{
      cell = createCell(row, col, code, {{ className: 'xl3', colCode: code }});
    }} else if (row >= 4) {{
      const rowCode = rowCodes[String(row)];
      const isInputRow = rowCode && inputSet.has(rowCode);
      cell = createCell(row, col, '', {{ input: isInputRow, auto: !isInputRow, colCode: code }});
    }} else {{
      cell = createCell(row, col, '', {{ colCode: code }});
    }}
    tr.appendChild(cell);
  }});
  return col;
}}

function addMatrixRow(code) {{
  const row = nextDynamicRow++;
  rowsByCode[code] = row;
  rowCodes[String(row)] = code;
  inputSet.add(code);
  if (!inputCodes.includes(code)) inputCodes.push(code);
  leavesCache.set(code, [code]);

  const tr = document.createElement('tr');
  tr.style.height = '30px';
  tr.appendChild(createCell(row, 1, '', {{ className: 'xl8' }}));
  tr.appendChild(createCell(row, 2, `Mã mới ${{code}}`, {{ className: 'xl8' }}));
  tr.appendChild(createCell(row, 3, code, {{ className: 'xl8', code }}));
  tr.appendChild(createCell(row, meta.currentCol, '', {{ className: 'xl8', input: true, code }}));
  for (const colCode of matrixCodes) {{
    tr.appendChild(createCell(row, colsByCode[colCode], '', {{
      className: 'xl8',
      input: true,
      code,
      colCode
    }}));
  }}
  for (let col = meta.decreaseCol; col <= (meta.previousPlanCol || meta.planCol); col++) {{
    tr.appendChild(createCell(row, col, '', {{ className: 'xl8', auto: true, code }}));
  }}
  document.querySelector('#landTable tbody').appendChild(tr);
  calcRowEntries.push([code, row]);
  return row;
}}

function addMissingLandCode(code) {{
  const normalized = normalizeLandCode(code);
  if (!normalized) return false;
  if (!colsByCode[normalized]) addMatrixColumn(normalized);
  if (!rowsByCode[normalized]) addMatrixRow(normalized);
  refreshCalcEntries();
  return true;
}}

function readProjectSettings() {{
  return {{
    commune: ($('#projectCommune')?.value || '').trim(),
    province: ($('#projectProvince')?.value || '').trim(),
    previousPlanYear: ($('#projectPreviousPlanYear')?.value || '').trim(),
    currentYear: ($('#projectCurrentYear')?.value || '').trim(),
    planYear: ($('#projectPlanYear')?.value || '').trim(),
    confirmed: projectTitlesConfirmed
  }};
}}

function extractYears(text) {{
  return String(text || '').match(/(?:19|20|21|22)\\d{{2}}/g) || [];
}}

function yearFromPlanPeriod(period, fallback = 2030) {{
  const years = extractYears(period);
  return Number(years[years.length - 1]) || fallback;
}}

function syncProjectYearsToReport() {{
  const currentYear = ($('#projectCurrentYear')?.value || '').trim();
  const planYear = yearFromPlanPeriod($('#projectPlanYear')?.value, 2030);
  if ($('#reportCurrentYear') && currentYear) $('#reportCurrentYear').value = currentYear;
  if ($('#reportPlanYear') && planYear) $('#reportPlanYear').value = planYear;
}}

function syncReportYearsToProject() {{
  const currentYear = ($('#reportCurrentYear')?.value || '').trim();
  const planYear = ($('#reportPlanYear')?.value || '').trim();
  if ($('#projectCurrentYear') && currentYear) $('#projectCurrentYear').value = currentYear;
  if ($('#projectPlanYear') && planYear && !extractYears($('#projectPlanYear').value).length) $('#projectPlanYear').value = `${{currentYear || '2020'}}-${{planYear}}`;
}}

function planningPeriodLabel(period) {{
  const years = extractYears(period);
  if (years.length >= 2) return `năm ${{years[0]}} đến năm ${{years[years.length - 1]}}`;
  if (years.length === 1) return `đến năm ${{years[0]}}`;
  return 'theo kỳ quy hoạch đã thiết lập';
}}

function updateProjectTitles() {{
  const settings = readProjectSettings();
  const commune = (settings.commune || '...').replace(/^xã\\s+/i, '').trim() || '...';
  const period = planningPeriodLabel(settings.planYear);
  const titleCell = document.querySelector('[data-row="1"][data-col="1"]');
  const matrixTitleCell = document.querySelector('[data-addr="E2"]');
  if (titleCell) titleCell.textContent = `Chu chuyển đất đai trong kỳ quy hoạch sử dụng đất của xã ${{commune}}`;
  if (matrixTitleCell) matrixTitleCell.textContent = `Chu chuyển đất đai ${{period}}`;
}}

function resetProjectTitles() {{
  const titleCell = document.querySelector('[data-row="1"][data-col="1"]');
  const matrixTitleCell = document.querySelector('[data-addr="E2"]');
  if (titleCell) titleCell.textContent = 'BẢNG CHU CHUYỂN ĐẤT ĐAI';
  if (matrixTitleCell) matrixTitleCell.textContent = 'Chu chuyển các loại đất';
}}

function applyProjectSettings(settings = {{}}) {{
  const safe = settings && typeof settings === 'object' ? settings : {{}};
  projectTitlesConfirmed = Boolean(safe.confirmed);
  if ($('#projectCommune')) $('#projectCommune').value = safe.commune || '';
  if ($('#projectProvince')) $('#projectProvince').value = safe.province || '';
  if ($('#projectPreviousPlanYear')) $('#projectPreviousPlanYear').value = safe.previousPlanYear || '';
  if ($('#projectCurrentYear')) $('#projectCurrentYear').value = safe.currentYear || $('#projectCurrentYear').value || '2020';
  if ($('#projectPlanYear')) $('#projectPlanYear').value = safe.planYear || $('#projectPlanYear').value || '2020-2030';
  syncProjectYearsToReport();
  if (projectTitlesConfirmed) updateProjectTitles();
  else resetProjectTitles();
}}

function readInputs() {{
  normalizeAllInputs();
  const data = {{}};
  inputEls.forEach((input, address) => {{
    data[address] = input.value.trim();
  }});
  data.__previousPlan = {{ ...previousPlanValues }};
  data.__projectSettings = readProjectSettings();
  return data;
}}

function applyInputs(data) {{
  if (!data || typeof data !== 'object') return;
  if (data && data.__previousPlan && typeof data.__previousPlan === 'object') {{
    applyPreviousPlanValues(data.__previousPlan);
  }}
  applyProjectSettings(data.__projectSettings);
  inputEls.forEach((input, address) => {{
    if (Object.prototype.hasOwnProperty.call(data, address)) {{
      input.value = data[address];
      normalizeInputElement(input);
    }}
  }});
}}

function gtpPayload() {{
  return {{
    format: 'gtp-land-transfer',
    version: 1,
    savedAt: new Date().toISOString(),
    data: readInputs()
  }};
}}

function gtpDataFromPayload(payload) {{
  if (!payload || typeof payload !== 'object') throw new Error('File GTP không hợp lệ.');
  if (payload.format === 'gtp-land-transfer' && payload.data && typeof payload.data === 'object') return payload.data;
  return payload;
}}

function updateGtpStatus(message = '') {{
  const label = gtpFileName || 'chưa chọn file';
  $('#gtpStatus').textContent = message || `File GTP: ${{label}}`;
}}

function applyProjectData(data) {{
  applyInputs(data);
  localStorage.setItem(storageKey, JSON.stringify(readInputs()));
  normalizeAllInputs();
  recalc();
}}

async function saveGtpFile({{ choose = false, silent = false }} = {{}}) {{
  const text = JSON.stringify(gtpPayload(), null, 2);
  if (window.showSaveFilePicker) {{
    if (choose || !gtpFileHandle) {{
      gtpFileHandle = await window.showSaveFilePicker({{
        suggestedName: gtpFileName || 'du_an_chu_chuyen_dat_dai.gtp',
        types: [{{
          description: 'Dữ liệu dự án GTP',
          accept: {{ 'application/json': ['.gtp'] }}
        }}]
      }});
      gtpFileName = gtpFileHandle.name || gtpFileName || 'du_an_chu_chuyen_dat_dai.gtp';
    }}
    const writable = await gtpFileHandle.createWritable();
    await writable.write(text);
    await writable.close();
    updateGtpStatus(silent ? '' : `Đã lưu: ${{gtpFileName}}`);
    return true;
  }}
  download(gtpFileName || 'du_an_chu_chuyen_dat_dai.gtp', 'application/json;charset=utf-8', text);
  updateGtpStatus('Trình duyệt tải xuống file GTP mới');
  return true;
}}

async function openGtpProjectFile(file) {{
  const payload = JSON.parse(await file.text());
  const data = gtpDataFromPayload(payload);
  gtpFileHandle = null;
  gtpFileName = file.name || 'du_an_chu_chuyen_dat_dai.gtp';
  applyProjectData(data);
  updateGtpStatus(`Đã nạp: ${{gtpFileName}}`);
}}

function xmlText(node) {{
  return node ? node.textContent || '' : '';
}}

function columnIndexFromCellRef(ref) {{
  const letters = String(ref || '').replace(/[0-9]/g, '');
  let n = 0;
  for (const ch of letters) n = n * 26 + ch.charCodeAt(0) - 64;
  return n;
}}

function rowIndexFromCellRef(ref) {{
  const match = String(ref || '').match(/\\d+/);
  return match ? Number(match[0]) : 0;
}}

async function parseXlsxRows(file) {{
  const zip = await JSZip.loadAsync(await file.arrayBuffer());
  const parser = new DOMParser();
  const workbookXml = parser.parseFromString(await zip.file('xl/workbook.xml').async('text'), 'application/xml');
  const firstSheet = workbookXml.querySelector('sheet');
  const relId = firstSheet?.getAttribute('r:id');
  let sheetPath = 'xl/worksheets/sheet1.xml';
  const relsFile = zip.file('xl/_rels/workbook.xml.rels');
  if (relId && relsFile) {{
    const relsXml = parser.parseFromString(await relsFile.async('text'), 'application/xml');
    const rel = Array.from(relsXml.querySelectorAll('Relationship')).find(item => item.getAttribute('Id') === relId);
    const target = rel?.getAttribute('Target');
    if (target) sheetPath = target.startsWith('/') ? target.slice(1) : 'xl/' + target.replace(/^\\.\\.\\//, '');
  }}

  const sharedStrings = [];
  const sharedFile = zip.file('xl/sharedStrings.xml');
  if (sharedFile) {{
    const sharedXml = parser.parseFromString(await sharedFile.async('text'), 'application/xml');
    Array.from(sharedXml.querySelectorAll('si')).forEach(si => {{
      sharedStrings.push(Array.from(si.querySelectorAll('t')).map(t => t.textContent || '').join(''));
    }});
  }}

  const sheetXml = parser.parseFromString(await zip.file(sheetPath).async('text'), 'application/xml');
  const rows = [];
  Array.from(sheetXml.querySelectorAll('row')).forEach(rowNode => {{
    const row = {{ number: Number(rowNode.getAttribute('r') || 0), cells: {{}} }};
    Array.from(rowNode.querySelectorAll('c')).forEach(cell => {{
      const ref = cell.getAttribute('r') || '';
      const col = columnIndexFromCellRef(ref);
      const type = cell.getAttribute('t');
      let value = '';
      if (type === 's') value = sharedStrings[Number(xmlText(cell.querySelector('v')))] || '';
      else if (type === 'inlineStr') value = xmlText(cell.querySelector('is t'));
      else value = xmlText(cell.querySelector('v'));
      row.cells[col] = value;
    }});
    rows.push(row);
  }});
  return rows;
}}

function normalizeHeader(text) {{
  return String(text || '').trim().toLowerCase();
}}

function normalizeNumber(text) {{
  const value = parseNumericText(text);
  return Number.isFinite(value) ? String(roundNumber(value)) : '';
}}

function setPreviousPlanCell(code, value) {{
  const row = code === 'DTTN' ? meta.dttnRow : rowsByCode[code];
  if (!row || !meta.previousPlanCol) return false;
  const td = cellsByKey.get(`${{row}}:${{meta.previousPlanCol}}`);
  const span = td?.querySelector('.value');
  if (!span) return false;
  const numeric = parseNumericText(value);
  if (!Number.isFinite(numeric)) return false;
  previousPlanValues[code] = formatNumber(numeric);
  span.textContent = formatNumber(numeric);
  return true;
}}

function applyPreviousPlanValues(values) {{
  Object.keys(previousPlanValues).forEach(code => delete previousPlanValues[code]);
  document.querySelectorAll('td[data-previous-plan="1"] .value').forEach(span => {{
    span.textContent = '';
  }});
  Object.entries(values || {{}}).forEach(([code, value]) => {{
    setPreviousPlanCell(normalizeLandCode(code), value);
  }});
}}

function detectPreviousPlanColumns(rows) {{
  let codeCol = null;
  let areaCol = null;
  let headerRow = 0;
  const codeNames = new Set(['ma', 'ma_dat', 'code', 'land_code', 'ma_loai_dat']);
  const areaNames = new Set(['dien_tich', 'dien_tich_ha', 'area', 'area_ha', 'quy_hoach', 'quy_hoach_ky_truoc']);
  for (const row of rows.slice(0, 30)) {{
    for (const [colText, value] of Object.entries(row.cells)) {{
      const key = normalizeHeaderKey(value);
      const col = Number(colText);
      if (codeNames.has(key)) {{
        codeCol = col;
        headerRow = Math.max(headerRow, row.number);
      }}
      if ((areaNames.has(key) || key.includes('dien_tich')) && (!codeCol || col > codeCol)) {{
        areaCol = col;
        headerRow = Math.max(headerRow, row.number);
      }}
    }}
    if (codeCol && areaCol) break;
  }}
  if (!areaCol && codeCol) {{
    for (const row of rows.slice(0, 30)) {{
      for (const [colText, value] of Object.entries(row.cells)) {{
        const key = normalizeHeaderKey(value);
        const col = Number(colText);
        if (col > codeCol && key.includes('quy_hoach')) {{
          areaCol = col;
          headerRow = Math.max(headerRow, row.number);
          break;
        }}
      }}
      if (areaCol) break;
    }}
  }}
  if (!codeCol || !areaCol) {{
    throw new Error('Không nhận diện được cột Mã đất và cột Diện tích quy hoạch kỳ trước.');
  }}
  return {{ codeCol, areaCol, headerRow }};
}}

async function importPreviousPlanExcel(file) {{
  if (/\\.xls$/i.test(file.name) && !/\\.xlsx$/i.test(file.name)) {{
    throw new Error('File .xls đời cũ chưa được trình đọc tích hợp hỗ trợ. Vui lòng lưu lại thành .xlsx rồi import.');
  }}
  const rows = await parseXlsxRows(file);
  const columns = detectPreviousPlanColumns(rows);
  const imported = {{}};
  const unknownCodes = new Set();
  let readRows = 0;
  let validRows = 0;
  let skippedRows = 0;
  for (const row of rows) {{
    if (row.number <= columns.headerRow) continue;
    readRows++;
    let code = normalizeLandCode(row.cells[columns.codeCol]);
    const nameKey = normalizeHeaderKey(row.cells[columns.codeCol - 1]);
    if (!code && nameKey.includes('tong_dien_tich_tu_nhien')) code = 'DTTN';
    const value = parseNumericText(row.cells[columns.areaCol]);
    if (!code || !Number.isFinite(value)) {{
      skippedRows++;
      continue;
    }}
    if (code !== 'DTTN' && !rowsByCode[code]) {{
      unknownCodes.add(code);
      skippedRows++;
      continue;
    }}
    imported[code] = formatNumber(value);
    validRows++;
  }}
  applyPreviousPlanValues(imported);
  localStorage.setItem(storageKey, JSON.stringify(readInputs()));
  const el = $('#importLog');
  el.hidden = false;
  el.innerHTML = `
    <strong>Log import quy hoạch kỳ trước</strong>
    <ul>
      <li>Tổng số dòng đã đọc: ${{readRows}}</li>
      <li>Số dòng hợp lệ: ${{validRows}}</li>
      <li>Số dòng bị bỏ qua: ${{skippedRows}}</li>
      <li>Mã đất lạ: ${{unknownCodes.size ? Array.from(unknownCodes).sort().join(', ') : 'Không có'}}</li>
    </ul>`;
  return {{ readRows, validRows, skippedRows, unknownCodes: Array.from(unknownCodes) }};
}}

async function importCurrentAreasFromXlsx(file) {{
  const rows = await parseXlsxRows(file);
  let codeCol = null;
  let areaCol = null;
  for (const row of rows) {{
    for (const [colText, value] of Object.entries(row.cells)) {{
      const col = Number(colText);
      const header = normalizeHeaderKey(value);
      if (['ma', 'ma_dat', 'code', 'land_code'].includes(header)) codeCol = col;
      if (header.includes('dien_tich') || header.includes('area')) areaCol = col;
    }}
    if (codeCol && areaCol) break;
  }}
  if (!codeCol || !areaCol) throw new Error('Không tìm thấy cột Mã và cột Diện tích trong file Excel.');

  let imported = 0;
  let matchedNoValue = 0;
  const unmatched = [];
  const currentAreasByCode = new Map();
  for (const row of rows) {{
    const code = normalizeLandCode(row.cells[codeCol]);
    if (!code || !rowsByCode[code]) continue;
    const value = normalizeNumber(row.cells[areaCol]);
    if (value !== '') currentAreasByCode.set(code, Number(value));
    const input = inputEls.get(`D${{rowsByCode[code]}}`);
    if (!input) {{
      if (!directChildren[code]) unmatched.push(code);
      continue;
    }}
    if (value === '') {{
      matchedNoValue++;
      continue;
    }}
    setInputNumber(`D${{rowsByCode[code]}}`, Number(value));
    imported++;
  }}
  const adjustments = reconcileCurrentAreaRounding(currentAreasByCode);
  recalc();
  return {{ imported, matchedNoValue, unmatched: Array.from(new Set(unmatched)), adjustments }};
}}

function reconcileCurrentAreaRounding(currentAreasByCode) {{
  const adjustments = [];
  const parentCodes = Object.keys(directChildren)
    .filter(code => currentAreasByCode.has(code) && !inputEls.has(`D${{rowsByCode[code]}}`))
    .sort((a, b) => leaves(a).length - leaves(b).length);

  parentCodes.forEach(parentCode => {{
    const leafCodes = leaves(parentCode).filter(code => inputEls.has(`D${{rowsByCode[code]}}`));
    if (!leafCodes.length) return;
    const parentValue = roundNumber(currentAreasByCode.get(parentCode));
    const childSum = roundNumber(leafCodes.reduce((sum, code) => sum + numberFromInputByAddr(`D${{rowsByCode[code]}}`), 0));
    const diff = roundNumber(parentValue - childSum);
    if (!diff || Math.abs(diff) > 0.05) return;

    const targetCode = leafCodes.reduce((best, code) => {{
      const bestValue = numberFromInputByAddr(`D${{rowsByCode[best]}}`);
      const codeValue = numberFromInputByAddr(`D${{rowsByCode[code]}}`);
      return codeValue > bestValue ? code : best;
    }}, leafCodes[0]);
    const targetAddress = `D${{rowsByCode[targetCode]}}`;
    setInputNumber(targetAddress, numberFromInputByAddr(targetAddress) + diff);
    adjustments.push({{ parentCode, targetCode, diff }});
  }});
  return adjustments;
}}

function normalizeHeaderKey(text) {{
  return String(text || '')
    .normalize('NFD')
    .replace(/[̀-ͯ]/g, '')
    .toLowerCase()
    .trim()
    .replace(/[^a-z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '');
}}

function normalizeLandCode(value) {{
  return String(value ?? '').trim().toUpperCase();
}}

function detectGISColumns(rows) {{
  const fromNames = new Set(['ma_hien_trang', 'ma_ht', 'hien_trang', 'from', 'from_code']);
  const toNames = new Set(['ma_quy_hoach', 'ma_qh', 'quy_hoach', 'to', 'to_code']);
  const areaNames = new Set(['dien_tich', 'area', 'area_ha', 'shape_area', 'shape_area_ha', 'area_m2', 'shape_area_m2', 'dt']);

  for (const row of rows.slice(0, 20)) {{
    const found = {{ headerRow: row.number, fromCode: null, toCode: null, area: null, areaHeader: '' }};
    for (const [colText, value] of Object.entries(row.cells)) {{
      const key = normalizeHeaderKey(value);
      const col = Number(colText);
      if (fromNames.has(key)) found.fromCode = col;
      if (toNames.has(key)) found.toCode = col;
      if (areaNames.has(key)) {{
        found.area = col;
        found.areaHeader = key;
      }}
    }}
    if (found.fromCode && found.toCode && found.area) return found;
  }}
  throw new Error('Không nhận diện được các cột Mã hiện trạng, Mã quy hoạch và Diện tích.');
}}

function areaUnitInfo(areaHeader) {{
  const key = normalizeHeaderKey(areaHeader);
  if (['shape_area', 'area_m2', 'shape_area_m2'].includes(key)) return {{ unit: 'm2', uncertain: false }};
  if (['dien_tich', 'area', 'area_ha', 'shape_area_ha', 'dt'].includes(key)) return {{ unit: 'ha', uncertain: false }};
  return {{ unit: 'unknown', uncertain: true }};
}}

function normalizeAreaValue(value) {{
  const raw = String(value ?? '').trim();
  if (!raw) return {{ value: 0, empty: true }};
  return {{ value: parseNumericText(raw), empty: false }};
}}

function aggregateOverlayRows(rows, columns, options = {{}}) {{
  const log = {{
    totalRows: 0,
    validRows: 0,
    skippedRows: 0,
    unknownCodes: new Set(),
    negativeRows: 0,
    totalArea: 0,
    warnings: [],
    addedCodes: new Set(),
    skippedUnknownRows: 0
  }};
  const matrix = {{}};
  const unit = areaUnitInfo(columns.areaHeader);
  const convertM2ToHa = unit.unit === 'm2' && options.convertM2ToHa;

  if (unit.unit === 'm2' && !convertM2ToHa) {{
    log.warnings.push('Cột diện tích có vẻ là m2; đang import nguyên giá trị. Hãy bật m2 -> ha nếu cần.');
  }}
  if (unit.uncertain) {{
    log.warnings.push('Không chắc đơn vị diện tích; vui lòng kiểm tra lại đơn vị sau khi import.');
  }}

  for (const row of rows) {{
    if (row.number <= columns.headerRow) continue;
    log.totalRows++;
    const fromCode = normalizeLandCode(row.cells[columns.fromCode]);
    const toCode = normalizeLandCode(row.cells[columns.toCode]);
    const parsedArea = normalizeAreaValue(row.cells[columns.area]);
    if (!fromCode || !toCode || !Number.isFinite(parsedArea.value)) {{
      log.skippedRows++;
      continue;
    }}
    if (parsedArea.value < 0) {{
      log.skippedRows++;
      log.negativeRows++;
      continue;
    }}
    const missing = [fromCode, toCode].filter(code => !rowsByCode[code] || !colsByCode[code]);
    if (missing.length) {{
      missing.forEach(code => log.unknownCodes.add(code));
      if (options.mode === 'add') {{
        missing.forEach(code => {{
          if (addMissingLandCode(code)) log.addedCodes.add(code);
        }});
      }} else {{
        log.skippedRows++;
        log.skippedUnknownRows++;
        continue;
      }}
    }}
    const area = convertM2ToHa ? parsedArea.value / 10000 : parsedArea.value;
    matrix[fromCode] ||= {{}};
    matrix[fromCode][toCode] = (matrix[fromCode][toCode] || 0) + area;
    log.validRows++;
    log.totalArea += area;
  }}

  log.unknownCodes = Array.from(log.unknownCodes).sort();
  log.addedCodes = Array.from(log.addedCodes).sort();
  return {{ matrix, log }};
}}

function clearGISMatrixInputs() {{
  inputCodes.forEach(code => {{
    matrixCodes.forEach(colCode => {{
      const input = inputEls.get(addr(colsByCode[colCode], rowsByCode[code]));
      if (input) input.value = '';
    }});
  }});
}}

function applyGISMatrix(matrix) {{
  clearGISMatrixInputs();
  let filledCells = 0;
  let skippedCells = 0;
  for (const [fromCode, rowValues] of Object.entries(matrix)) {{
    for (const [toCode, area] of Object.entries(rowValues)) {{
      const input = inputEls.get(addr(colsByCode[toCode], rowsByCode[fromCode]));
      if (input) {{
        setInputNumber(addr(colsByCode[toCode], rowsByCode[fromCode]), area);
        filledCells++;
      }} else {{
        skippedCells++;
      }}
    }}
  }}
  return {{ filledCells, skippedCells }};
}}

function calculateCurrentArea() {{}}
function calculateMatrixTotals() {{}}
function calculateDecrease() {{}}
function calculateIncrease() {{}}
function calculatePlanningArea() {{}}
function calculateChange() {{}}
function validateTable() {{}}
function renderTable() {{ recalc(); }}

function recalculateAfterImport() {{
  calculateCurrentArea();
  calculateMatrixTotals();
  calculateDecrease();
  calculateIncrease();
  calculatePlanningArea();
  calculateChange();
  validateTable();
  renderTable();
}}

function showImportLog(log) {{
  const el = $('#importLog');
  const warnings = [];
  if (log.negativeRows) warnings.push(`${{log.negativeRows}} dòng diện tích âm đã bị bỏ qua`);
  if (log.skippedUnknownRows) warnings.push(`${{log.skippedUnknownRows}} dòng có mã lạ đã bị bỏ qua`);
  warnings.push(...log.warnings);
  el.hidden = false;
  el.innerHTML = `
    <strong>Log import GIS</strong>
    <ul>
      <li>Tổng số dòng đã đọc: ${{log.totalRows}}</li>
      <li>Số dòng hợp lệ: ${{log.validRows}}</li>
      <li>Số dòng bị bỏ qua: ${{log.skippedRows}}</li>
      <li>Tổng diện tích đã import: ${{formatNumber(log.totalArea)}}</li>
      <li>Số ô vàng ma trận đã điền: ${{log.filledCells || 0}}</li>
      <li>Mã đất lạ: ${{log.unknownCodes.length ? log.unknownCodes.join(', ') : 'Không có'}}</li>
      <li>Mã đất đã tự thêm: ${{log.addedCodes.length ? log.addedCodes.join(', ') : 'Không có'}}</li>
      ${{warnings.length ? `<li>Cảnh báo: ${{warnings.join('; ')}}</li>` : ''}}
    </ul>`;
}}

async function importGISOverlayExcel(file) {{
  if (/\\.xls$/i.test(file.name) && !/\\.xlsx$/i.test(file.name)) {{
    throw new Error('File .xls đời cũ chưa được trình đọc tích hợp hỗ trợ. Vui lòng lưu lại thành .xlsx rồi import.');
  }}
  const rows = await parseXlsxRows(file);
  const columns = detectGISColumns(rows);
  const {{ matrix, log }} = aggregateOverlayRows(rows, columns, {{
    mode: $('#gisImportMode').value,
    convertM2ToHa: $('#gisM2ToHa').checked
  }});
  Object.assign(log, applyGISMatrix(matrix));
  recalculateAfterImport();
  showImportLog(log);
  localStorage.setItem(storageKey, JSON.stringify(readInputs()));
  return log;
}}



const displayDecimals = 2;
const displayFactor = 10 ** displayDecimals;

function roundNumber(value) {{
  if (!Number.isFinite(value)) return 0;
  const rounded = Math.round((value + Number.EPSILON) * displayFactor) / displayFactor;
  return Math.abs(rounded) < 0.0000001 ? 0 : rounded;
}}

function parseNumericText(text) {{
  const raw = String(text ?? '').trim();
  if (!raw) return NaN;
  let cleaned = raw.replace(/\\s/g, '');
  if (cleaned.includes(',') && cleaned.includes('.')) {{
    const lastComma = cleaned.lastIndexOf(',');
    const lastDot = cleaned.lastIndexOf('.');
    if (lastComma > lastDot) cleaned = cleaned.replace(/\\./g, '').replace(',', '.');
    else cleaned = cleaned.replace(/,/g, '');
  }} else if (cleaned.includes(',')) {{
    cleaned = cleaned.replace(',', '.');
  }}
  const value = Number(cleaned);
  return Number.isFinite(value) ? value : NaN;
}}

function formatInputValue(value) {{
  if (!Number.isFinite(value)) return '';
  const rounded = roundNumber(value);
  if (Math.abs(rounded) < 0.0000001) return '';
  return rounded.toFixed(displayDecimals).replace('.', ',');
}}

function normalizeInputElement(input) {{
  if (!input) return;
  const value = parseNumericText(input.value);
  input.value = Number.isFinite(value) ? formatInputValue(value) : '';
  updateInputZeroState(input);
}}

function normalizeAllInputs() {{
  inputEls.forEach(input => normalizeInputElement(input));
}}

function parseInputNumber(input) {{
  if (!input) return 0;
  const value = parseNumericText(input.value);
  return Number.isFinite(value) ? roundNumber(value) : 0;
}}

function setInputNumber(address, value) {{
  const input = inputEls.get(address);
  if (!input) return;
  input.value = formatInputValue(value);
  updateInputZeroState(input);
}}

function updateInputZeroState(input) {{
  if (!input) return;
  const td = input.closest('td');
  if (!td) return;
  const value = parseInputNumber(input);
  td.classList.toggle('zero-cell', Math.abs(value) <= meta.tolerance);
}}

function setDiagonalValue(code, value) {{
  const row = rowsByCode[code];
  const col = colsByCode[code];
  if (!row || !col) return;
  const address = addr(col, row);
  if (inputEls.has(address)) setInputNumber(address, value);
  else setAuto(row, col, value);
}}

function numberFromInputByAddr(address) {{
  return parseInputNumber(inputEls.get(address));
}}

function formatNumber(value) {{
  if (!Number.isFinite(value)) return '';
  const rounded = roundNumber(value);
  return rounded.toLocaleString('vi-VN', {{ maximumFractionDigits: displayDecimals, minimumFractionDigits: displayDecimals }});
}}

function setAuto(row, col, value) {{
  const key = `${{row}}:${{col}}`;
  const td = cellsByKey.get(key);
  if (!td || td.dataset.input === '1') return;
  const text = formatNumber(value);
  const span = autoSpans.get(key);
  if (span && span.textContent !== text) span.textContent = text;
  const raw = String(roundNumber(value));
  if (td.dataset.value !== raw) td.dataset.value = raw;
  td.classList.toggle('zero-cell', Math.abs(roundNumber(value)) <= meta.tolerance);
}}

function getAuto(row, col) {{
  const td = cellsByKey.get(`${{row}}:${{col}}`);
  if (!td) return 0;
  if (td.dataset.input === '1') return numberFromInputByAddr(td.dataset.addr);
  const value = Number(td.dataset.value);
  return Number.isFinite(value) ? roundNumber(value) : 0;
}}

const leavesCache = new Map();

function leaves(code) {{
  if (leavesCache.has(code)) return leavesCache.get(code);
  let result;
  if (code === 'DTTN') result = inputCodes.slice();
  else if (directChildren[code]) result = directChildren[code].flatMap(child => leaves(child));
  else result = inputSet.has(code) ? [code] : [];
  leavesCache.set(code, result);
  return result;
}}

leavesCache.set('DTTN', inputCodes.slice());

function diagonalCodes() {{
  return matrixCodes.filter(code => rowsByCode[code] && colsByCode[code]);
}}

function diagonalOutflowTotal(code, matrixValue) {{
  const ownLeafCodes = new Set(leaves(code));
  return inputCodes.reduce((sum, colCode) => {{
    if (ownLeafCodes.has(colCode)) return sum;
    return sum + matrixValue(code, colCode);
  }}, 0);
}}

function createCalcContext() {{
  const inputValues = new Map();
  inputEls.forEach((input, address) => inputValues.set(address, parseInputNumber(input)));
  const currentCache = new Map();
  const matrixCache = new Map();

  function currentArea(code) {{
    if (currentCache.has(code)) return currentCache.get(code);
    const value = inputSet.has(code)
      ? (inputValues.get('D' + rowsByCode[code]) || 0)
      : leaves(code).reduce((sum, leaf) => sum + currentArea(leaf), 0);
    currentCache.set(code, value);
    return value;
  }}

  function matrixLeaf(rowCode, colCode) {{
    if (!inputSet.has(rowCode) || !inputSet.has(colCode)) return 0;
    return inputValues.get(addr(colsByCode[colCode], rowsByCode[rowCode])) || 0;
  }}

  function matrixValue(rowCode, colCode) {{
    const key = rowCode + ':' + colCode;
    if (matrixCache.has(key)) return matrixCache.get(key);
    const rLeaves = leaves(rowCode);
    const cLeaves = leaves(colCode);
    let sum = 0;
    rLeaves.forEach(r => cLeaves.forEach(c => sum += matrixLeaf(r, c)));
    matrixCache.set(key, sum);
    return sum;
  }}

  return {{ currentArea, matrixLeaf, matrixValue }};
}}

function recalc() {{
  const {{ currentArea, matrixLeaf, matrixValue }} = createCalcContext();
  for (const code of Object.keys(rowsByCode)) {{
    setAuto(rowsByCode[code], meta.currentCol, currentArea(code));
  }}
  setAuto(meta.dttnRow, meta.currentCol, currentArea('DTTN'));

  for (const [code, row] of calcRowEntries) {{
    for (const colCode of matrixCodes) {{
      const col = colsByCode[colCode];
      if (inputKeys.has(`${{row}}:${{col}}`)) continue;
      setAuto(row, col, matrixValue(code, colCode));
    }}
  }}
  for (const colCode of matrixCodes) {{
    setAuto(meta.dttnRow, colsByCode[colCode], matrixValue('DTTN', colCode));
  }}

  for (const code of diagonalCodes()) {{
    const current = currentArea(code);
    const outflowTotal = diagonalOutflowTotal(code, matrixValue);
    if (current > meta.tolerance || outflowTotal > meta.tolerance) {{
      setDiagonalValue(code, Math.max(0, current - outflowTotal));
    }}
  }}

  const refreshedCalc = createCalcContext();
  for (const [code, row] of calcRowEntries) {{
    const current = row === meta.dttnRow ? refreshedCalc.currentArea('DTTN') : refreshedCalc.currentArea(code);
    const diagonal = refreshedCalc.matrixValue(code, code);
    const plan = refreshedCalc.matrixValue('DTTN', code);
    setAuto(row, meta.decreaseCol, current - diagonal);
    setAuto(row, meta.planCol, plan);
    setAuto(row, meta.changeCol, plan - current);
  }}
  setAuto(meta.dttnRow, meta.decreaseCol,
    ['NNP', 'PNN', 'CSD'].reduce((sum, code) => sum + getAuto(rowsByCode[code] || 0, meta.decreaseCol), 0)
  );
  setAuto(meta.dttnRow, meta.planCol, ['NNP', 'PNN', 'CSD'].reduce((sum, code) => sum + refreshedCalc.matrixValue('DTTN', code), 0));
  setAuto(meta.dttnRow, meta.changeCol, getAuto(meta.dttnRow, meta.planCol) - getAuto(meta.dttnRow, meta.currentCol));

  for (const colCode of matrixCodes) {{
    const col = colsByCode[colCode];
    const plan = refreshedCalc.matrixValue('DTTN', colCode);
    const diagonal = refreshedCalc.matrixValue(colCode, colCode);
    setAuto(meta.totalIncreaseRow, col, plan - diagonal);
    setAuto(meta.planRow, col, plan);
  }}
  setAuto(meta.totalIncreaseRow, meta.decreaseCol,
    ['NNP', 'PNN', 'CSD'].reduce((sum, code) => sum + getAuto(rowsByCode[code] || 0, meta.decreaseCol), 0)
  );

  updateWarnings({{ currentArea: refreshedCalc.currentArea, matrixLeaf: refreshedCalc.matrixLeaf }});
}}

function updateWarnings(calc) {{
  const tol = meta.tolerance;
  let rowErrors = 0;
  const nextWarnCells = new Set();
  for (const code of inputCodes) {{
    const row = rowsByCode[code];
    const rowSum = inputCodes.reduce((sum, colCode) => sum + calc.matrixLeaf(code, colCode), 0);
    const current = calc.currentArea(code);
    if (Math.abs(rowSum - current) > tol) {{
      rowErrors++;
      for (let col = 1; col <= (meta.previousPlanCol || meta.planCol); col++) {{
        const td = cellsByKey.get(`${{row}}:${{col}}`);
        if (td) nextWarnCells.add(td);
      }}
    }}
  }}
  previousWarnCells.forEach(td => {{
    if (!nextWarnCells.has(td)) td.classList.remove('warn');
  }});
  nextWarnCells.forEach(td => {{
    if (!previousWarnCells.has(td)) td.classList.add('warn');
  }});
  previousWarnCells.clear();
  nextWarnCells.forEach(td => previousWarnCells.add(td));
  const totalDiff = Math.abs(getAuto(meta.dttnRow, meta.currentCol) - getAuto(meta.dttnRow, meta.planCol));
  const totalBadge = $('#statusTotal');
  totalBadge.textContent = totalDiff > tol ? `DTTN lệch ${{formatNumber(totalDiff)}}` : 'DTTN cân bằng';
  totalBadge.classList.toggle('warn', totalDiff > tol);
  const rowBadge = $('#statusRows');
  rowBadge.textContent = `${{rowErrors}} lệch hàng`;
  rowBadge.classList.toggle('warn', rowErrors > 0);
}}

function download(name, type, text) {{
  const blob = new Blob([text], {{ type }});
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = name;
  a.click();
  URL.revokeObjectURL(a.href);
}}

function csvText(row, col) {{
  const td = cellsByKey.get(`${{row}}:${{col}}`);
  if (!td) return '';
  return td.dataset.input === '1' ? (inputEls.get(td.dataset.addr)?.value || '') : (td.textContent || '').trim();
}}

function exportCellText(row, col, renumberMap = new Map()) {{
  if (col === 1 && renumberMap.has(row)) return renumberMap.get(row);
  return csvText(row, col);
}}

function originalSttForCode(code) {{
  const row = rowsByCode[code];
  return row ? csvText(row, 1) : '';
}}

function csvEscape(text) {{
  return '"' + String(text).replaceAll('"', '""') + '"';
}}

function xmlEscape(text) {{
  return String(text ?? '')
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;');
}}

function exportActiveMatrixCodes(calc) {{
  const tol = meta.tolerance;
  return matrixCodes.filter(code => {{
    if (!rowsByCode[code] && !colsByCode[code]) return false;
    if (Math.abs(calc.currentArea(code)) > tol) return true;
    for (const otherCode of matrixCodes) {{
      if (Math.abs(calc.matrixValue(code, otherCode)) > tol) return true;
      if (Math.abs(calc.matrixValue(otherCode, code)) > tol) return true;
    }}
    return false;
  }});
}}

function exportCsv() {{
  normalizeAllInputs();
  recalc();
  const {{ exportCols, exportRows, renumberMap }} = exportMatrixShape();
  const rows = [];
  for (const row of exportRows) {{
    rows.push(exportCols.map(col => csvEscape(exportCellText(row, col, renumberMap))).join(','));
  }}
  download('chu_chuyen_dat_dai.csv', 'text/csv;charset=utf-8', '\\ufeff' + rows.join('\\n'));
}}

function exportMatrixShape() {{
  const calc = createCalcContext();
  const activeCodes = exportActiveMatrixCodes(calc);
  const activeSet = new Set(activeCodes);
  const exportCols = [
    1,
    2,
    3,
    meta.currentCol,
    ...activeCodes.map(code => colsByCode[code]).filter(Boolean),
    meta.decreaseCol,
    meta.changeCol,
    meta.planCol
  ];
  const exportDataRows = [];
  for (const [code, row] of calcRowEntries) {{
    if (row === meta.dttnRow || activeSet.has(code)) exportDataRows.push(row);
  }}
  exportDataRows.sort((a, b) => a - b);
  const exportRows = Array.from(new Set([1, 2, 3, meta.dttnRow, ...exportDataRows, meta.totalIncreaseRow, meta.planRow]))
    .sort((a, b) => a - b);
  const renumberMap = new Map();
  exportRows.forEach(row => {{
    if (row === meta.dttnRow || row === meta.totalIncreaseRow || row === meta.planRow) return;
    const code = rowCodes[String(row)];
    const stt = code ? originalSttForCode(code) : '';
    if (stt) renumberMap.set(row, stt);
  }});
  return {{ exportCols, exportRows, renumberMap }};
}}

function xlsxCellXml(cellRef, text, styleId = 0, forceText = false) {{
  const raw = String(text ?? '').trim();
  const numeric = parseNumericText(raw);
  if (!forceText && raw && Number.isFinite(numeric) && !/[A-Za-zÀ-ỹ]/.test(raw)) {{
    return `<c r="${{cellRef}}" s="${{styleId}}"><v>${{String(numeric)}}</v></c>`;
  }}
  return `<c r="${{cellRef}}" s="${{styleId}}" t="inlineStr"><is><t>${{xmlEscape(raw)}}</t></is></c>`;
}}

function xlsxStyleFor(row) {{
  const code = rowCodes[String(row)];
  if (row === 1) return 2;
  if (row <= 3 || row === meta.dttnRow || row === meta.totalIncreaseRow || row === meta.planRow) return 1;
  if (['NNP', 'PNN', 'CSD'].includes(code)) return 1;
  return 0;
}}

function xlsxCellStyleFor(row, col) {{
  if (col === 2 && row > 3) {{
    const baseStyle = xlsxStyleFor(row);
    return baseStyle === 1 ? 4 : 3;
  }}
  return xlsxStyleFor(row);
}}

function exportRowHeight(row) {{
  const name = csvText(row, 2);
  if (name.length > 48) return 36;
  if (name.length > 32) return 28;
  return row === 2 ? 31.2 : 18;
}}

function exportXlsx() {{
  normalizeAllInputs();
  recalc();
  const {{ exportCols, exportRows, renumberMap }} = exportMatrixShape();
  const sheetRows = exportRows.map((row, rowIndex) => {{
    const cells = exportCols.map((col, colIndex) => xlsxCellXml(addr(colIndex + 1, rowIndex + 1), exportCellText(row, col, renumberMap), xlsxCellStyleFor(row, col), col === 1)).join('');
    const height = exportRowHeight(row);
    return `<row r="${{rowIndex + 1}}" ht="${{height}}" customHeight="1">${{cells}}</row>`;
  }}).join('');
  const widths = exportCols.map((col, index) => {{
    const width = col === 2 ? 34 : (col === 1 || col === 3 ? 10 : 15);
    return `<col min="${{index + 1}}" max="${{index + 1}}" width="${{width}}" customWidth="1"/>`;
  }}).join('');
  const lastRef = addr(exportCols.length, 1);
  const matrixEnd = 4 + exportCols.filter(col => col >= meta.matrixStartCol && col <= meta.matrixEndCol).length;
  const mergeXml = matrixEnd > 5
    ? `<mergeCells count="2"><mergeCell ref="A1:${{lastRef}}"/><mergeCell ref="E2:${{addr(matrixEnd, 2)}}"/></mergeCells>`
    : `<mergeCells count="1"><mergeCell ref="A1:${{lastRef}}"/></mergeCells>`;
  const zip = new JSZip();
  zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>`);
  zip.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`);
  zip.file('xl/_rels/workbook.xml.rels', `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`);
  zip.file('xl/workbook.xml', `<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Chu chuyển đất đai" sheetId="1" r:id="rId1"/></sheets>
</workbook>`);
  zip.file('xl/styles.xml', `<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="3">
    <font><sz val="12"/><name val="Times New Roman"/></font>
    <font><b/><sz val="12"/><name val="Times New Roman"/></font>
    <font><b/><sz val="12"/><name val="Times New Roman"/></font>
  </fonts>
  <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
  <borders count="1"><border><left style="thin"><color auto="1"/></left><right style="thin"><color auto="1"/></right><top style="thin"><color auto="1"/></top><bottom style="thin"><color auto="1"/></bottom><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="5">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center" wrapText="1"/></xf>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`);
  zip.file('xl/worksheets/sheet1.xml', `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetViews><sheetView workbookViewId="0"><pane xSplit="4" ySplit="3" topLeftCell="E4" activePane="bottomRight" state="frozen"/></sheetView></sheetViews>
  <cols>${{widths}}</cols>
  <sheetData>${{sheetRows}}</sheetData>
  ${{mergeXml}}
</worksheet>`);
  zip.generateAsync({{ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }})
    .then(blob => {{
      const a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = 'chu_chuyen_dat_dai.xlsx';
      a.click();
      URL.revokeObjectURL(a.href);
    }});
}}

let searchHitTimer = 0;
function jumpToLandCode(rawCode) {{
  const code = normalizeLandCode(rawCode);
  if (!code) return;
  const row = rowsByCode[code];
  const col = colsByCode[code];
  if (!row || !col) {{
    alert(`Không tìm thấy mã đất: ${{code}}`);
    return;
  }}
  const td = cellsByKey.get(`${{row}}:${{col}}`);
  if (!td) return;
  td.scrollIntoView({{ block: 'center', inline: 'center', behavior: 'smooth' }});
  clearTimeout(searchHitTimer);
  document.querySelectorAll('.search-hit').forEach(el => el.classList.remove('search-hit'));
  td.classList.add('search-hit');
  searchHitTimer = setTimeout(() => td.classList.remove('search-hit'), 2600);
}}

function landName(code) {{
  const row = rowsByCode[code];
  const name = cellsByKey.get(`${{row}}:2`)?.textContent?.trim();
  return name || code;
}}

function ownLeafSet(code) {{
  return new Set(leaves(code));
}}

function reportIncomingEntries(code, calc) {{
  const ownLeaves = ownLeafSet(code);
  return inputCodes
    .filter(sourceCode => !ownLeaves.has(sourceCode))
    .map(sourceCode => [sourceCode, calc.matrixValue(sourceCode, code)])
    .filter(([, value]) => Math.abs(value) > meta.tolerance)
    .sort((a, b) => Math.abs(b[1]) - Math.abs(a[1]));
}}

function reportOutgoingEntries(code, calc) {{
  const ownLeaves = ownLeafSet(code);
  return inputCodes
    .filter(targetCode => !ownLeaves.has(targetCode))
    .map(targetCode => [targetCode, calc.matrixValue(code, targetCode)])
    .filter(([, value]) => Math.abs(value) > meta.tolerance)
    .sort((a, b) => Math.abs(b[1]) - Math.abs(a[1]));
}}

function hasReportData(code, calc) {{
  return Math.abs(calc.currentArea(code)) > meta.tolerance ||
    Math.abs(calc.matrixValue('DTTN', code)) > meta.tolerance ||
    reportIncomingEntries(code, calc).length > 0 ||
    reportOutgoingEntries(code, calc).length > 0;
}}

function reportCodeOptions() {{
  return matrixCodes
    .filter(code => rowsByCode[code] && colsByCode[code])
    .map(code => ({{ code, name: landName(code) }}));
}}

function renderReportOptions(filter = '') {{
  const q = normalizeHeaderKey(filter);
  const selected = new Set(Array.from(document.querySelectorAll('#reportOptions input:checked')).map(input => input.value));
  const options = reportCodeOptions().filter(item => {{
    if (!q) return true;
    return normalizeHeaderKey(item.code).includes(q) || normalizeHeaderKey(item.name).includes(q);
  }});
  $('#reportOptions').innerHTML = options.map(item => `
    <label class="report-option">
      <input type="checkbox" value="${{item.code}}" ${{selected.has(item.code) ? 'checked' : ''}}>
      <span><strong>${{item.code}}</strong><br>${{item.name}}</span>
    </label>
  `).join('');
}}

function selectedReportCodes() {{
  return Array.from(document.querySelectorAll('#reportOptions input:checked')).map(input => input.value);
}}

function reportLine(prefix, text, value, end = ';') {{
  return `<div class="line">${{prefix}} ${{text}}<span class="amount">: ${{formatNumber(value)}} ha${{end}}</span></div>`;
}}

function reportBlock(code, calc, years) {{
  const name = landName(code);
  const current = calc.currentArea(code);
  const plan = calc.matrixValue('DTTN', code);
  const natural = calc.currentArea('DTTN');
  const share = natural > meta.tolerance ? (plan / natural) * 100 : 0;
  const change = plan - current;
  const direction = change >= -meta.tolerance ? 'tăng' : 'giảm';
  const incoming = reportIncomingEntries(code, calc);
  const outgoing = reportOutgoingEntries(code, calc);
  const incomingTotal = incoming.reduce((sum, [, value]) => sum + value, 0);
  const outgoingTotal = outgoing.reduce((sum, [, value]) => sum + value, 0);
  if (Math.abs(change) <= meta.tolerance) {{
    return `
    <div class="block">
      <div class="title-line">* <strong><em>${{name}}:</em></strong></div>
      <div>Quy hoạch sử dụng đất đến năm ${{years.planYear}} là ${{formatNumber(plan)}} ha, chiếm ${{formatNumber(share)}}% tổng diện tích tự nhiên, không biến động so với năm ${{years.currentYear}}.</div>
    </div>`;
  }}
  const incomingLines = incoming.length
    ? incoming.map(([sourceCode, value], index) => reportLine('-', landName(sourceCode), value, index === incoming.length - 1 ? '.' : ';')).join('')
    : '';
  const outgoingLines = outgoing.length
    ? outgoing.map(([targetCode, value], index) => reportLine('-', landName(targetCode), value, index === outgoing.length - 1 ? '.' : ';')).join('')
    : '';
  const incomingSection = Math.abs(incomingTotal) > meta.tolerance
    ? `<div class="section">+ Cộng tăng ${{formatNumber(incomingTotal)}} ha do chuyển sang từ các loại đất sau:</div>${{incomingLines}}`
    : '';
  const outgoingSection = Math.abs(outgoingTotal) > meta.tolerance
    ? `<div class="section">+ Cộng giảm ${{formatNumber(outgoingTotal)}} ha, do chuyển sang các loại đất sau:</div>${{outgoingLines}}`
    : '';
  return `
    <div class="block">
      <div class="title-line">* <strong><em>${{name}}:</em></strong></div>
      <div>Diện tích năm ${{years.currentYear}} là ${{formatNumber(current)}} ha, quy hoạch sử dụng đất đến năm ${{years.planYear}} là ${{formatNumber(plan)}} ha, chiếm ${{formatNumber(share)}}% tổng diện tích tự nhiên, ${{direction}} ${{formatNumber(Math.abs(change))}} ha so với năm ${{years.currentYear}}, chi tiết như sau:</div>
      ${{incomingSection}}
      ${{outgoingSection}}
    </div>`;
}}

function reportYears() {{
  const currentYear = Number(($('#projectCurrentYear')?.value || $('#reportCurrentYear').value || '').trim()) || 2020;
  const planYear = yearFromPlanPeriod($('#projectPlanYear')?.value || $('#reportPlanYear').value, 2030);
  $('#reportCurrentYear').value = currentYear;
  $('#reportPlanYear').value = planYear;
  return {{ currentYear, planYear }};
}}

function exportReportWord() {{
  normalizeAllInputs();
  recalc();
  const codes = selectedReportCodes();
  if (!codes.length) {{
    alert('Hãy chọn ít nhất một loại đất để xuất.');
    return;
  }}
  const calc = createCalcContext();
  const years = reportYears();
  const body = codes.map(code => reportBlock(code, calc, years)).join('');
  const html = `<!doctype html>
<html>
<head>
<meta charset="utf-8">
<style>
@page {{ margin: 2cm; }}
body {{ font-family: "Times New Roman", serif; font-size: 13pt; line-height: 1.35; }}
.block {{ margin: 0 0 14pt 0; }}
.title-line {{ font-weight: 700; }}
.section {{ margin-top: 6pt; }}
.line {{ white-space: nowrap; }}
.amount {{ display: inline-block; min-width: 150px; text-align: left; }}
</style>
</head>
<body>${{body}}</body>
</html>`;
  download('thuyet_minh_cong_tang_cong_giam.doc', 'application/msword;charset=utf-8', '\\ufeff' + html);
}}

let pendingRecalc = 0;
function scheduleRecalc() {{
  if (pendingRecalc) return;
  pendingRecalc = requestAnimationFrame(() => {{
    pendingRecalc = 0;
    recalc();
  }});
}}

async function saveProjectToServer() {{
  const response = await fetch(`${{apiBase}}/${{encodeURIComponent(projectId)}}`, {{
    method: 'PUT',
    headers: {{ 'Content-Type': 'application/json' }},
    body: JSON.stringify({{ data: readInputs() }})
  }});
  if (!response.ok) {{
    const payload = await response.json().catch(() => ({{ error: response.statusText }}));
    throw new Error(payload.error || 'Không lưu được dữ liệu lên server');
  }}
  return response.json();
}}

async function loadProjectFromServer() {{
  try {{
    const response = await fetch(`${{apiBase}}/${{encodeURIComponent(projectId)}}`);
    if (response.status === 404) return false;
    if (!response.ok) throw new Error('Không đọc được dữ liệu từ server');
    const payload = await response.json();
    if (payload.data && typeof payload.data === 'object') {{
      applyInputs(payload.data);
      localStorage.setItem(storageKey, JSON.stringify(payload.data));
      normalizeAllInputs();
      recalc();
      return true;
    }}
  }} catch (error) {{
    console.warn(error.message || error);
  }}
  return false;
}}

function showLibraryMessage(target, message, isError = false) {{
  const box = $(target);
  if (!box) return;
  box.hidden = !message;
  box.textContent = message || '';
  box.style.borderColor = isError ? '#f4b0a1' : '#bbf7d0';
  box.style.background = isError ? '#fff1ed' : '#f0fdf4';
  box.style.color = isError ? '#7a271a' : '#166534';
}}

function libraryAuthHeaders() {{
  return librarySessionToken ? {{ Authorization: `Bearer ${{librarySessionToken}}` }} : {{}};
}}

function libraryAdminHeaders() {{
  return librarySessionRole === 'admin' && librarySessionToken ? {{ Authorization: `Bearer ${{librarySessionToken}}` }} : {{}};
}}

function updateLibrarySessionUi() {{
  const logged = Boolean(librarySessionToken);
  const isAdmin = logged && librarySessionRole === 'admin';
  const badge = $('#librarySessionBadge');
  if (badge) {{
    badge.hidden = !logged;
    badge.textContent = isAdmin ? 'Admin' : 'Khách';
  }}
  const logoutBtn = $('#libraryLogoutBtn');
  if (logoutBtn) logoutBtn.hidden = !logged;
  const adminBtn = $('#libraryAdminOpenBtn');
  if (adminBtn) adminBtn.hidden = !isAdmin;
}}

function setLibrarySession(payload) {{
  librarySessionToken = payload.token || '';
  librarySessionRole = payload.role || 'guest';
  libraryAdminToken = librarySessionRole === 'admin' ? librarySessionToken : '';
  localStorage.setItem(librarySessionTokenKey, librarySessionToken);
  localStorage.setItem(librarySessionRoleKey, librarySessionRole);
  if (libraryAdminToken) localStorage.setItem('library-admin-token', libraryAdminToken);
  else localStorage.removeItem('library-admin-token');
  updateLibrarySessionUi();
}}

function clearLibrarySession() {{
  librarySessionToken = '';
  librarySessionRole = '';
  libraryAdminToken = '';
  localStorage.removeItem(librarySessionTokenKey);
  localStorage.removeItem(librarySessionRoleKey);
  localStorage.removeItem('library-admin-token');
  $('#libraryAdminPanel').hidden = true;
  updateLibrarySessionUi();
}}

function showLibraryAccessPanel(message = '') {{
  closeMainMenu();
  $('#libraryAccessPanel').hidden = false;
  showLibraryMessage('#libraryAccessMsg', message);
  setTimeout(() => $('#libraryAccessUser')?.focus(), 0);
}}

function hideLibraryAccessPanel() {{
  $('#libraryAccessPanel').hidden = true;
  showLibraryMessage('#libraryAccessMsg', '');
}}

function escapeHtml(value) {{
  return String(value ?? '').replace(/[&<>"']/g, ch => ({{
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#39;'
  }}[ch]));
}}

function fileToDataUrl(file) {{
  return new Promise((resolve, reject) => {{
    if (!file) {{
      resolve('');
      return;
    }}
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result || ''));
    reader.onerror = () => reject(reader.error || new Error('Không đọc được file.'));
    reader.readAsDataURL(file);
  }});
}}

async function loadPdfJs() {{
  if (window.pdfjsLib) return window.pdfjsLib;
  await new Promise((resolve, reject) => {{
    const script = document.createElement('script');
    script.src = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js';
    script.onload = resolve;
    script.onerror = () => reject(new Error('Không tải được PDF.js. Hãy kiểm tra kết nối mạng.'));
    document.head.appendChild(script);
  }});
  window.pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
  return window.pdfjsLib;
}}

function libraryQueryString(includeHidden = false) {{
  const params = new URLSearchParams();
  const q = $('#librarySearch')?.value.trim();
  const category = $('#libraryCategoryFilter')?.value;
  const year = $('#libraryYearFilter')?.value;
  if (q) params.set('q', q);
  if (category) params.set('category', category);
  if (year) params.set('year', year);
  if (includeHidden) params.set('includeHidden', '1');
  const text = params.toString();
  return text ? `?${{text}}` : '';
}}

async function fetchLibraryDocuments(includeHidden = false) {{
  if (!librarySessionToken) {{
    showLibraryAccessPanel();
    throw new Error('Bạn cần đăng nhập để vào thư viện tài liệu.');
  }}
  const response = await fetch(`${{libraryApiBase}}/documents${{libraryQueryString(includeHidden)}}`, {{
    headers: includeHidden ? libraryAdminHeaders() : libraryAuthHeaders()
  }});
  const payload = await response.json().catch(() => ({{}}));
  if (response.status === 401) {{
    clearLibrarySession();
    showLibraryAccessPanel(payload.error || 'Phiên đăng nhập thư viện đã hết hạn. Vui lòng đăng nhập lại.');
  }}
  if (!response.ok) throw new Error(payload.error || 'Không tải được thư viện tài liệu.');
  libraryDocuments = payload.documents || [];
  renderLibraryFilters(payload);
  renderLibraryGrid(libraryDocuments);
  if (includeHidden) renderLibraryAdminRows(libraryDocuments);
  return payload;
}}

function renderLibraryFilters(payload = {{}}) {{
  const categorySelect = $('#libraryCategoryFilter');
  const yearSelect = $('#libraryYearFilter');
  const currentCategory = categorySelect.value;
  const currentYear = yearSelect.value;
  categorySelect.innerHTML = '<option value="">Tất cả danh mục</option>' +
    (payload.categories || []).map(item => `<option value="${{escapeHtml(item.category)}}">${{escapeHtml(item.category)}} (${{item.count}})</option>`).join('');
  yearSelect.innerHTML = '<option value="">Tất cả năm</option>' +
    (payload.years || []).map(year => `<option value="${{year}}">${{year}}</option>`).join('');
  categorySelect.value = currentCategory;
  yearSelect.value = currentYear;
  const datalist = $('#libraryCategorySuggestions');
  if (datalist) {{
    datalist.innerHTML = (payload.categories || []).map(item => `<option value="${{escapeHtml(item.category)}}"></option>`).join('');
  }}
}}

function renderLibraryGrid(documents) {{
  const grid = $('#libraryGrid');
  const empty = $('#libraryEmpty');
  grid.innerHTML = documents.map(doc => `
    <article class="library-card">
      <div class="library-cover">
        ${{doc.coverUrl ? `<img src="${{doc.coverUrl}}" alt="Bìa tài liệu ${{escapeHtml(doc.title)}}">` : `<div class="library-cover-placeholder">${{escapeHtml(doc.title)}}</div>`}}
      </div>
      <div class="library-card-body">
        <h3>${{escapeHtml(doc.title)}}</h3>
        <div class="library-meta">
          <span class="library-pill">${{escapeHtml(doc.category || 'Chưa phân loại')}}</span>
          <span class="library-pill">${{doc.year || 'Không rõ năm'}}</span>
        </div>
        <div class="library-meta library-author">${{escapeHtml(doc.author || 'Chưa rõ tác giả')}}</div>
        <div class="library-description">${{escapeHtml(doc.description || '')}}</div>
        <button class="primary library-read-btn" type="button" data-id="${{doc.id}}">Đọc trực tuyến</button>
      </div>
    </article>
  `).join('');
  empty.hidden = documents.length > 0;
}}

function renderLibraryAdminRows(documents) {{
  const tbody = $('#libraryAdminRows');
  tbody.innerHTML = documents.map(doc => `
    <tr>
      <td><strong>${{escapeHtml(doc.title)}}</strong><br><span class="library-meta">${{escapeHtml(doc.author || '')}}</span></td>
      <td>${{escapeHtml(doc.category || '')}}</td>
      <td>${{doc.year || ''}}</td>
      <td>${{doc.visible ? 'Hiển thị' : 'Đang ẩn'}}</td>
      <td>
        <button type="button" data-action="edit" data-id="${{doc.id}}">Sửa</button>
        <button type="button" data-action="toggle" data-id="${{doc.id}}">${{doc.visible ? 'Ẩn' : 'Hiện'}}</button>
        <button type="button" data-action="delete" data-id="${{doc.id}}">Xóa</button>
      </td>
    </tr>
  `).join('');
}}

function resetLibraryDocForm() {{
  $('#libraryDocForm').reset();
  $('#libraryDocId').value = '';
  $('#libraryDocVisible').checked = true;
  showLibraryMessage('#libraryAdminMsg', '');
}}

function fillLibraryDocForm(doc) {{
  $('#libraryDocId').value = doc.id;
  $('#libraryDocTitle').value = doc.title || '';
  $('#libraryDocAuthor').value = doc.author || '';
  $('#libraryDocYear').value = doc.year || '';
  $('#libraryDocCategory').value = doc.category || '';
  $('#libraryDocDescription').value = doc.description || '';
  $('#libraryDocVisible').checked = Boolean(doc.visible);
  $('#libraryDocPdf').value = '';
  $('#libraryDocCover').value = '';
  showLibraryMessage('#libraryAdminMsg', `Đang sửa: ${{doc.title}}`);
}}

async function openLibraryAdminPanel() {{
  if (librarySessionRole !== 'admin') {{
    if (!librarySessionToken) showLibraryAccessPanel('Vui lòng đăng nhập bằng tài khoản admin để quản trị thư viện.');
    else alert('Tài khoản khách chỉ được đọc tài liệu. Vui lòng đăng nhập admin để upload hoặc chỉnh sửa.');
    return;
  }}
  $('#libraryAdminPanel').hidden = false;
  showLibraryMessage('#libraryLoginMsg', '');
  showLibraryMessage('#libraryLoginStatus', '');
  showLibraryMessage('#libraryAdminMsg', '');
  const isLogged = Boolean(libraryAdminToken);
  $('#libraryLoginBox').hidden = isLogged;
  $('#libraryUploadCard').hidden = !isLogged;
  $('#libraryDocForm').hidden = !isLogged;
  if (isLogged) {{
    showLibraryMessage('#libraryLoginStatus', '\u0110\u00e3 \u0111\u0103ng nh\u1eadp th\u00e0nh c\u00f4ng. B\u1ea1n c\u00f3 th\u1ec3 upload ho\u1eb7c ch\u1ec9nh s\u1eeda t\u00e0i li\u1ec7u \u1edf khung b\u00ean c\u1ea1nh.');
    try {{
      await fetchLibraryDocuments(true);
    }} catch (error) {{
      clearLibrarySession();
      $('#libraryLoginBox').hidden = false;
      $('#libraryUploadCard').hidden = true;
      $('#libraryDocForm').hidden = true;
      showLibraryMessage('#libraryLoginStatus', '');
      showLibraryMessage('#libraryLoginMsg', error.message || String(error), true);
    }}
  }}
}}

async function libraryAccessLogin() {{
  const username = $('#libraryAccessUser').value.trim();
  const password = $('#libraryAccessPassword').value;
  const response = await fetch(`${{libraryApiBase}}/login`, {{
    method: 'POST',
    headers: {{ 'Content-Type': 'application/json' }},
    body: JSON.stringify({{ username, password }})
  }});
  const payload = await response.json().catch(() => ({{}}));
  if (!response.ok) throw new Error(payload.error || 'Không đăng nhập được thư viện.');
  setLibrarySession(payload);
  hideLibraryAccessPanel();
  showDocumentLibraryPage();
}}

async function libraryAdminLogin() {{
  const username = $('#libraryAdminUser').value.trim();
  const password = $('#libraryAdminPassword').value;
  const response = await fetch(`${{libraryApiBase}}/login`, {{
    method: 'POST',
    headers: {{ 'Content-Type': 'application/json' }},
    body: JSON.stringify({{ username, password }})
  }});
  const payload = await response.json().catch(() => ({{}}));
  if (!response.ok) throw new Error(payload.error || 'Không đăng nhập được.');
  if (payload.role !== 'admin') throw new Error('Tài khoản khách chỉ được đọc tài liệu, không thể quản trị.');
  setLibrarySession(payload);
  $('#libraryLoginBox').hidden = true;
  $('#libraryUploadCard').hidden = false;
  $('#libraryDocForm').hidden = false;
  showLibraryMessage('#libraryLoginMsg', '');
  showLibraryMessage('#libraryLoginStatus', '\u0110\u00e3 \u0111\u0103ng nh\u1eadp th\u00e0nh c\u00f4ng. B\u1ea1n c\u00f3 th\u1ec3 upload ho\u1eb7c ch\u1ec9nh s\u1eeda t\u00e0i li\u1ec7u \u1edf khung b\u00ean c\u1ea1nh.');
  resetLibraryDocForm();
  await fetchLibraryDocuments(true);
}}

async function saveLibraryDocument(event) {{
  event.preventDefault();
  const id = $('#libraryDocId').value;
  const pdfFile = $('#libraryDocPdf').files[0];
  if (!id && !pdfFile) {{
    showLibraryMessage('#libraryAdminMsg', 'Tài liệu mới cần có file PDF.', true);
    return;
  }}
  const coverFile = $('#libraryDocCover').files[0];
  const payload = {{
    title: $('#libraryDocTitle').value,
    author: $('#libraryDocAuthor').value,
    year: $('#libraryDocYear').value,
    category: $('#libraryDocCategory').value,
    description: $('#libraryDocDescription').value,
    visible: $('#libraryDocVisible').checked,
    pdfName: pdfFile?.name || '',
    coverName: coverFile?.name || '',
    pdfDataUrl: await fileToDataUrl(pdfFile),
    coverDataUrl: await fileToDataUrl(coverFile)
  }};
  const response = await fetch(`${{libraryApiBase}}/documents${{id ? `/${{id}}` : ''}}`, {{
    method: id ? 'PUT' : 'POST',
    headers: {{ 'Content-Type': 'application/json', ...libraryAdminHeaders() }},
    body: JSON.stringify(payload)
  }});
  const result = await response.json().catch(() => ({{}}));
  if (!response.ok) throw new Error(result.error || 'Không lưu được tài liệu.');
  showLibraryMessage('#libraryAdminMsg', 'Đã lưu tài liệu.');
  resetLibraryDocForm();
  await fetchLibraryDocuments(true);
}}

async function handleLibraryAdminAction(event) {{
  const button = event.target.closest('button[data-action]');
  if (!button) return;
  const id = Number(button.dataset.id);
  const doc = libraryDocuments.find(item => Number(item.id) === id);
  if (!doc) return;
  const action = button.dataset.action;
  if (action === 'edit') {{
    fillLibraryDocForm(doc);
    return;
  }}
  if (action === 'toggle') {{
    const response = await fetch(`${{libraryApiBase}}/documents/${{id}}/visibility`, {{
      method: 'PATCH',
      headers: {{ 'Content-Type': 'application/json', ...libraryAdminHeaders() }},
      body: JSON.stringify({{ visible: !doc.visible }})
    }});
    const payload = await response.json().catch(() => ({{}}));
    if (!response.ok) throw new Error(payload.error || 'Không đổi được trạng thái tài liệu.');
    await fetchLibraryDocuments(true);
    return;
  }}
  if (action === 'delete') {{
    if (!confirm(`Xóa tài liệu "${{doc.title}}"?`)) return;
    const response = await fetch(`${{libraryApiBase}}/documents/${{id}}`, {{
      method: 'DELETE',
      headers: libraryAdminHeaders()
    }});
    const payload = await response.json().catch(() => ({{}}));
    if (!response.ok) throw new Error(payload.error || 'Không xóa được tài liệu.');
    await fetchLibraryDocuments(true);
  }}
}}

function drawPdfWatermark(ctx, width, height) {{
  const text = 'Thư viện số - Chỉ đọc trực tuyến';
  ctx.save();
  ctx.globalAlpha = 0.09;
  ctx.fillStyle = '#0f766e';
  ctx.font = `${{Math.max(22, Math.round(width / 28))}}px Arial`;
  ctx.textAlign = 'center';
  ctx.translate(width / 2, height / 2);
  ctx.rotate(-Math.PI / 6);
  for (let y = -height; y <= height; y += 170) {{
    for (let x = -width; x <= width; x += 420) {{
      ctx.fillText(text, x, y);
    }}
  }}
  ctx.restore();
}}

async function renderPdfPage() {{
  if (!activePdf) return;
  const serial = ++activePdfRenderSerial;
  try {{
    if (activePdfRenderTask) activePdfRenderTask.cancel();
  }} catch (error) {{}}
  const page = await activePdf.getPage(activePdfPage);
  if (serial !== activePdfRenderSerial) return;
  const canvas = $('#pdfCanvas');
  const ctx = canvas.getContext('2d', {{ alpha: false }});
  const cssViewport = page.getViewport({{ scale: activePdfScale }});
  const pixelRatio = Math.min(window.devicePixelRatio || 1, 2.5);
  const renderViewport = page.getViewport({{ scale: activePdfScale * pixelRatio }});
  canvas.width = Math.floor(renderViewport.width);
  canvas.height = Math.floor(renderViewport.height);
  canvas.style.width = `${{Math.floor(cssViewport.width)}}px`;
  canvas.style.height = `${{Math.floor(cssViewport.height)}}px`;
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, canvas.width, canvas.height);
  activePdfRenderTask = page.render({{ canvasContext: ctx, viewport: renderViewport }});
  await activePdfRenderTask.promise.catch(error => {{
    if (error?.name !== 'RenderingCancelledException') throw error;
  }});
  if (serial !== activePdfRenderSerial) return;
  drawPdfWatermark(ctx, canvas.width, canvas.height);
  $('#readerPageInput').value = activePdfPage;
  $('#readerPageTotal').textContent = `/ ${{activePdf.numPages}}`;
  $('#readerPrevBtn').disabled = activePdfPage <= 1;
  $('#readerNextBtn').disabled = activePdfPage >= activePdf.numPages;
}}

async function openPdfReader(doc) {{
  $('#pdfReader').hidden = false;
  $('#readerTitle').textContent = doc.title;
  const tokenResponse = await fetch(`${{libraryApiBase}}/documents/${{doc.id}}/view-token`, {{
    method: 'POST',
    headers: libraryAuthHeaders()
  }});
  const tokenPayload = await tokenResponse.json().catch(() => ({{}}));
  if (!tokenResponse.ok) throw new Error(tokenPayload.error || 'Không tạo được phiên đọc tài liệu.');
  const pdfjs = await loadPdfJs();
  const url = `${{libraryApiBase}}/documents/${{doc.id}}/pdf?token=${{encodeURIComponent(tokenPayload.token)}}`;
  activePdf = await pdfjs.getDocument({{ url, disableAutoFetch: true, disableStream: true }}).promise;
  activePdfPage = 1;
  activePdfScale = 1.2;
  await renderPdfPage();
}}

function closePdfReader() {{
  $('#pdfReader').hidden = true;
  activePdf = null;
  activePdfRenderSerial++;
}}

async function changePdfPage(delta) {{
  if (!activePdf) return;
  activePdfPage = Math.min(activePdf.numPages, Math.max(1, activePdfPage + delta));
  await renderPdfPage();
}}

async function setPdfPage(value) {{
  if (!activePdf) return;
  const next = Number(value);
  if (!Number.isFinite(next)) return;
  activePdfPage = Math.min(activePdf.numPages, Math.max(1, Math.trunc(next)));
  await renderPdfPage();
}}

async function zoomPdf(delta) {{
  activePdfScale = Math.min(3, Math.max(0.6, activePdfScale + delta));
  await renderPdfPage();
}}

function closeMainMenu() {{
  $('#menuList').hidden = true;
  $('#menuBtn').setAttribute('aria-expanded', 'false');
}}

function closeToolDropdowns(except = null) {{
  $$('.tool-group.open, .sample-downloads.open').forEach(group => {{
    if (group !== except) group.classList.remove('open');
  }});
}}

function showHomePage() {{
  document.body.classList.add('home-mode');
  document.body.classList.remove('module-mode');
  document.body.classList.remove('docs-mode');
  document.body.classList.remove('webgis-mode');
  $('#reportPanel').hidden = true;
  $('#aiPanel').hidden = true;
  $('#libraryAccessPanel').hidden = true;
  $('#importLog').hidden = true;
  closeMainMenu();
}}

function showLandTransferPage() {{
  document.body.classList.add('module-mode');
  document.body.classList.remove('home-mode');
  document.body.classList.remove('docs-mode');
  document.body.classList.remove('webgis-mode');
  $('#libraryAccessPanel').hidden = true;
  closeMainMenu();
  recalc();
}}

function showDocumentLibraryPage() {{
  if (!librarySessionToken) {{
    showLibraryAccessPanel();
    return;
  }}
  document.body.classList.add('docs-mode');
  document.body.classList.remove('home-mode');
  document.body.classList.remove('module-mode');
  document.body.classList.remove('webgis-mode');
  $('#reportPanel').hidden = true;
  $('#aiPanel').hidden = true;
  $('#importLog').hidden = true;
  closeMainMenu();
  updateLibrarySessionUi();
  fetchLibraryDocuments().catch(error => alert(error.message || String(error)));
}}

function showWebGisPage() {{
  document.body.classList.add('webgis-mode');
  document.body.classList.remove('home-mode');
  document.body.classList.remove('module-mode');
  document.body.classList.remove('docs-mode');
  $('#reportPanel').hidden = true;
  $('#aiPanel').hidden = true;
  $('#libraryAccessPanel').hidden = true;
  $('#libraryAdminPanel').hidden = true;
  $('#pdfReader').hidden = true;
  $('#importLog').hidden = true;
  closeMainMenu();
  initializeWebGIS().catch(error => {{
    webgisSetSaveStatus('Không khởi động được WebGIS', true);
    alert(error.message || String(error));
  }});
}}

$('#menuBtn').addEventListener('click', event => {{
  event.stopPropagation();
  const menu = $('#menuList');
  menu.hidden = !menu.hidden;
  $('#menuBtn').setAttribute('aria-expanded', String(!menu.hidden));
}});
$('#openLandTransferBtn').addEventListener('click', showLandTransferPage);
$('#openDocumentLibraryBtn').addEventListener('click', showDocumentLibraryPage);
$('#openWebGisBtn').addEventListener('click', showWebGisPage);
$('#homeBtn').addEventListener('click', showHomePage);
$('#libraryHomeBtn').addEventListener('click', showHomePage);
$('#libraryLogoutBtn').addEventListener('click', () => {{
  clearLibrarySession();
  libraryDocuments = [];
  renderLibraryGrid([]);
  showLibraryAccessPanel('Đã đăng xuất. Vui lòng đăng nhập lại để vào thư viện.');
}});
$('#libraryAccessCloseBtn').addEventListener('click', () => $('#libraryAccessPanel').hidden = true);
$('#libraryAccessLoginBtn').addEventListener('click', () => {{
  libraryAccessLogin().catch(error => showLibraryMessage('#libraryAccessMsg', error.message || String(error), true));
}});
['libraryAccessUser', 'libraryAccessPassword'].forEach(id => {{
  $(`#${{id}}`).addEventListener('keydown', event => {{
    if (event.key === 'Enter') libraryAccessLogin().catch(error => showLibraryMessage('#libraryAccessMsg', error.message || String(error), true));
  }});
}});
$('#libraryAdminOpenBtn').addEventListener('click', openLibraryAdminPanel);
$('#libraryAdminCloseBtn').addEventListener('click', () => $('#libraryAdminPanel').hidden = true);
$('#libraryLoginBtn').addEventListener('click', () => {{
  libraryAdminLogin().catch(error => showLibraryMessage('#libraryLoginMsg', error.message || String(error), true));
}});
$('#libraryDocForm').addEventListener('submit', event => {{
  saveLibraryDocument(event).catch(error => showLibraryMessage('#libraryAdminMsg', error.message || String(error), true));
}});
$('#libraryDocNewBtn').addEventListener('click', resetLibraryDocForm);
$('#libraryAdminReloadBtn').addEventListener('click', () => {{
  fetchLibraryDocuments(true).catch(error => showLibraryMessage('#libraryAdminMsg', error.message || String(error), true));
}});
$('#libraryAdminRows').addEventListener('click', event => {{
  handleLibraryAdminAction(event).catch(error => showLibraryMessage('#libraryAdminMsg', error.message || String(error), true));
}});
$('#libraryGrid').addEventListener('click', event => {{
  const button = event.target.closest('.library-read-btn');
  if (!button) return;
  const doc = libraryDocuments.find(item => Number(item.id) === Number(button.dataset.id));
  if (doc) openPdfReader(doc).catch(error => alert(error.message || String(error)));
}});
['librarySearch', 'libraryCategoryFilter', 'libraryYearFilter'].forEach(id => {{
  const input = $(`#${{id}}`);
  input.addEventListener(id === 'librarySearch' ? 'input' : 'change', () => fetchLibraryDocuments().catch(error => alert(error.message || String(error))));
}});
$('#libraryRefreshBtn').addEventListener('click', () => fetchLibraryDocuments().catch(error => alert(error.message || String(error))));
$('#readerCloseBtn').addEventListener('click', closePdfReader);
$('#readerPrevBtn').addEventListener('click', () => changePdfPage(-1));
$('#readerNextBtn').addEventListener('click', () => changePdfPage(1));
$('#readerPageInput').addEventListener('change', event => setPdfPage(event.currentTarget.value));
$('#readerZoomOutBtn').addEventListener('click', () => zoomPdf(-0.15));
$('#readerZoomInBtn').addEventListener('click', () => zoomPdf(0.15));
$('#readerFullscreenBtn').addEventListener('click', () => {{
  const reader = $('#pdfReader');
  if (!document.fullscreenElement) reader.requestFullscreen?.();
  else document.exitFullscreen?.();
}});
$('#pdfReader').addEventListener('contextmenu', event => event.preventDefault());
$('#pdfReader').addEventListener('selectstart', event => event.preventDefault());
$('#pdfReader').addEventListener('dragstart', event => event.preventDefault());
document.addEventListener('keydown', event => {{
  if ($('#pdfReader').hidden) return;
  const key = event.key.toLowerCase();
  const blocked = event.key === 'F12' ||
    event.key === 'PrintScreen' ||
    ((event.ctrlKey || event.metaKey) && ['c', 'v', 's', 'p', 'a', 'u'].includes(key));
  if (blocked) {{
    event.preventDefault();
    event.stopPropagation();
  }}
}});
$$('.tool-group-title').forEach(button => {{
  button.addEventListener('click', event => {{
    event.stopPropagation();
    const group = event.currentTarget.closest('.tool-group');
    const willOpen = !group.classList.contains('open');
    closeToolDropdowns(group);
    group.classList.toggle('open', willOpen);
  }});
}});
$('.sample-downloads > span').addEventListener('click', event => {{
  event.stopPropagation();
  const group = event.currentTarget.closest('.sample-downloads');
  const willOpen = !group.classList.contains('open');
  closeToolDropdowns(group);
  group.classList.toggle('open', willOpen);
}});
$$('.tool-items, .sample-items').forEach(panel => {{
  panel.addEventListener('click', event => event.stopPropagation());
}});
document.addEventListener('click', event => {{
  if (!event.target.closest('.main-menu')) closeMainMenu();
  closeToolDropdowns();
}});
document.addEventListener('keydown', event => {{
  if (event.key === 'Escape') {{
    closeMainMenu();
    closeToolDropdowns();
  }}
}});

function landName(code) {{
  const row = rowsByCode[code];
  const cell = row ? cellsByKey.get(`${{row}}:2`) : null;
  return cell ? cell.textContent.trim() : code;
}}

function buildAiContext() {{
  normalizeAllInputs();
  recalc();
  const calc = createCalcContext();
  const codes = matrixCodes.filter(code => rowsByCode[code]);
  const landTypes = codes.map(code => {{
    const current = calc.currentArea(code);
    const plan = calc.matrixValue('DTTN', code);
    const diagonal = calc.matrixValue(code, code);
    return {{
      code,
      name: landName(code),
      current: roundNumber(current),
      planning: roundNumber(plan),
      decrease: roundNumber(current - diagonal),
      increase: roundNumber(plan - diagonal),
      change: roundNumber(plan - current)
    }};
  }}).filter(item =>
    Math.abs(item.current) > meta.tolerance ||
    Math.abs(item.planning) > meta.tolerance ||
    Math.abs(item.decrease) > meta.tolerance ||
    Math.abs(item.increase) > meta.tolerance ||
    Math.abs(item.change) > meta.tolerance
  );

  const transfers = [];
  inputCodes.forEach(fromCode => {{
    inputCodes.forEach(toCode => {{
      const value = calc.matrixLeaf(fromCode, toCode);
      if (Math.abs(value) > meta.tolerance) {{
        transfers.push({{
          fromCode,
          fromName: landName(fromCode),
          toCode,
          toName: landName(toCode),
          area: roundNumber(value)
        }});
      }}
    }});
  }});
  transfers.sort((a, b) => Math.abs(b.area) - Math.abs(a.area));

  const totalCurrent = calc.currentArea('DTTN');
  const totalPlanning = calc.matrixValue('DTTN', 'DTTN');
  return {{
    unit: 'ha',
    decimals: displayDecimals,
    tolerance: meta.tolerance,
    totals: {{
      current: roundNumber(totalCurrent),
      planning: roundNumber(totalPlanning),
      difference: roundNumber(totalPlanning - totalCurrent)
    }},
    landTypes,
    topTransfers: transfers.slice(0, 40)
  }};
}}

function appendAiMessage(type, text) {{
  const el = document.createElement('div');
  el.className = `ai-message ${{type || ''}}`.trim();
  el.textContent = text;
  $('#aiMessages').appendChild(el);
  $('#aiMessages').scrollTop = $('#aiMessages').scrollHeight;
  return el;
}}

async function sendAiQuestion() {{
  const input = $('#aiQuestion');
  const question = input.value.trim();
  if (!question) return;
  input.value = '';
  appendAiMessage('user', question);
  const waiting = appendAiMessage('', 'AI đang phân tích dữ liệu...');
  $('#aiSendBtn').disabled = true;
  try {{
    const response = await fetch('/api/ai', {{
      method: 'POST',
      headers: {{ 'Content-Type': 'application/json' }},
      body: JSON.stringify({{ question, context: buildAiContext() }})
    }});
    const payload = await response.json().catch(() => ({{ error: response.statusText }}));
    if (!response.ok) throw new Error(payload.error || 'Không gọi được AI.');
    waiting.textContent = payload.answer || 'AI không trả về nội dung.';
  }} catch (error) {{
    const message = error.message || String(error);
    waiting.textContent = message.includes('fetch')
      ? 'Không kết nối được server AI. Hãy chạy npm start và mở phần mềm tại http://127.0.0.1:3000.'
      : message;
  }} finally {{
    $('#aiSendBtn').disabled = false;
  }}
}}

inputEls.forEach(input => {{
  input.addEventListener('input', () => {{
    updateInputZeroState(input);
    scheduleRecalc();
  }});
}});

function applyHideZeroState(enabled) {{
  document.body.classList.toggle('hide-zero', enabled);
  $('#hideZeroToggle').checked = enabled;
  localStorage.setItem(hideZeroKey, enabled ? '1' : '0');
}}

$('#hideZeroToggle').addEventListener('change', event => {{
  applyHideZeroState(event.currentTarget.checked);
}});

let hoverRow = null;
let hoverCol = null;
let hoverCell = null;
function clearTableHover() {{
  if (hoverRow !== null) $$(`td[data-row="${{hoverRow}}"]`).forEach(td => td.classList.remove('hover-row'));
  if (hoverCol !== null) $$(`td[data-col="${{hoverCol}}"]`).forEach(td => td.classList.remove('hover-col'));
  if (hoverCell) hoverCell.classList.remove('hover-cell');
  hoverRow = null;
  hoverCol = null;
  hoverCell = null;
}}
$('#landTable').addEventListener('mouseover', event => {{
  const td = event.target.closest('td');
  if (!td || hoverCell === td) return;
  clearTableHover();
  hoverRow = td.dataset.row;
  hoverCol = td.dataset.col;
  hoverCell = td;
  $$(`td[data-row="${{hoverRow}}"]`).forEach(cell => cell.classList.add('hover-row'));
  $$(`td[data-col="${{hoverCol}}"]`).forEach(cell => cell.classList.add('hover-col'));
  td.classList.add('hover-cell');
}});
$('#landTable').addEventListener('mouseleave', clearTableHover);
$('#codeSearchBtn').addEventListener('click', () => jumpToLandCode($('#codeSearch').value));
$('#codeSearch').addEventListener('keydown', event => {{
  if (event.key === 'Enter') {{
    event.preventDefault();
    jumpToLandCode(event.currentTarget.value);
  }}
}});
$('#reportBtn').addEventListener('click', () => {{
  syncProjectYearsToReport();
  renderReportOptions($('#reportFilter').value);
  $('#reportPanel').hidden = false;
}});
$('#reportCloseBtn').addEventListener('click', () => $('#reportPanel').hidden = true);
$('#aiBtn').addEventListener('click', () => {{
  $('#aiPanel').hidden = false;
  $('#aiQuestion').focus();
}});
$('#aiCloseBtn').addEventListener('click', () => $('#aiPanel').hidden = true);
$('#aiSendBtn').addEventListener('click', sendAiQuestion);
$('#aiQuestion').addEventListener('keydown', event => {{
  if (event.key === 'Enter' && !event.shiftKey) {{
    event.preventDefault();
    sendAiQuestion();
  }}
}});
$('#reportFilter').addEventListener('input', event => renderReportOptions(event.currentTarget.value));
['projectCommune', 'projectProvince', 'projectPreviousPlanYear', 'projectCurrentYear', 'projectPlanYear'].forEach(id => {{
  const input = $(`#${{id}}`);
  if (!input) return;
  input.addEventListener('input', () => {{
    projectTitlesConfirmed = false;
    syncProjectYearsToReport();
  }});
  input.addEventListener('change', () => localStorage.setItem(storageKey, JSON.stringify(readInputs())));
}});
['reportCurrentYear', 'reportPlanYear'].forEach(id => {{
  const input = $(`#${{id}}`);
  if (!input) return;
  input.addEventListener('input', syncReportYearsToProject);
  input.addEventListener('change', () => localStorage.setItem(storageKey, JSON.stringify(readInputs())));
}});
$('#projectConfirmBtn').addEventListener('click', () => {{
  projectTitlesConfirmed = true;
  syncProjectYearsToReport();
  updateProjectTitles();
  localStorage.setItem(storageKey, JSON.stringify(readInputs()));
  $('#projectConfirmBtn').textContent = 'Đã xác nhận';
  setTimeout(() => $('#projectConfirmBtn').textContent = 'Xác nhận', 900);
}});
$('#gtpOpenBtn').addEventListener('click', () => $('#gtpInput').click());
$('#gtpInput').addEventListener('change', async event => {{
  const file = event.target.files[0];
  if (!file) return;
  try {{
    await openGtpProjectFile(file);
  }} catch (error) {{
    alert(error.message || String(error));
  }} finally {{
    event.target.value = '';
  }}
}});
$('#gtpSetupBtn').addEventListener('click', async () => {{
  try {{
    await saveGtpFile({{ choose: true }});
  }} catch (error) {{
    if (error && error.name === 'AbortError') return;
    alert(error.message || String(error));
  }}
}});
$('#gtpSaveBtn').addEventListener('click', async () => {{
  try {{
    await saveGtpFile({{ choose: !gtpFileHandle }});
  }} catch (error) {{
    if (error && error.name === 'AbortError') return;
    alert(error.message || String(error));
  }}
}});
$('#reportSelectActiveBtn').addEventListener('click', () => {{
  normalizeAllInputs();
  recalc();
  const calc = createCalcContext();
  renderReportOptions($('#reportFilter').value);
  document.querySelectorAll('#reportOptions input').forEach(input => {{
    input.checked = hasReportData(input.value, calc);
  }});
}});
$('#reportClearBtn').addEventListener('click', () => {{
  document.querySelectorAll('#reportOptions input').forEach(input => input.checked = false);
}});
$('#reportExportBtn').addEventListener('click', exportReportWord);
$('#saveBtn').addEventListener('click', async () => {{
  const data = readInputs();
  localStorage.setItem(storageKey, JSON.stringify(data));
  $('#saveBtn').disabled = true;
  $('#saveBtn').textContent = 'Đang lưu';
  const failures = [];
  if (gtpFileHandle) {{
    try {{
      await saveGtpFile({{ silent: true }});
    }} catch (error) {{
      failures.push(error.message || String(error));
    }}
  }}
  try {{
    await saveProjectToServer();
  }} catch (error) {{
    failures.push(error.message || String(error));
  }}
  if (failures.length) {{
    $('#saveBtn').textContent = 'Lưu lỗi';
    alert(failures.join('\\n'));
  }} else {{
    $('#saveBtn').textContent = 'Đã lưu';
  }}
  try {{
    setTimeout(() => {{
      $('#saveBtn').disabled = false;
      $('#saveBtn').textContent = 'Lưu';
    }}, 900);
  }} catch (error) {{}}
}});
$('#importGisBtn').addEventListener('click', () => $('#gisXlsxInput').click());
$('#gisXlsxInput').addEventListener('change', async event => {{
  const file = event.target.files[0];
  if (!file) return;
  try {{
    await importGISOverlayExcel(file);
  }} catch (error) {{
    alert(error.message || String(error));
  }} finally {{
    event.target.value = '';
  }}
}});
$('#importCurrentBtn').addEventListener('click', () => $('#currentXlsxInput').click());
$('#currentXlsxInput').addEventListener('change', async event => {{
  const file = event.target.files[0];
  if (!file) return;
  try {{
    const result = await importCurrentAreasFromXlsx(file);
    localStorage.setItem(storageKey, JSON.stringify(readInputs()));
    const msg = `Đã nhập ${{result.imported}} ô hiện trạng từ XLSX` +
      (result.matchedNoValue ? `; ${{result.matchedNoValue}} mã trùng nhưng trống diện tích` : '') +
      (result.adjustments.length ? `; cân sai số làm tròn: ${{result.adjustments.map(item => `${{item.parentCode}} -> ${{item.targetCode}} ${{formatNumber(item.diff)}}`).join(', ')}}` : '') +
      (result.unmatched.length ? `; bỏ qua mã không phải dòng nhập: ${{result.unmatched.slice(0, 8).join(', ')}}` : '');
    alert(msg);
  }} catch (error) {{
    alert(error.message || String(error));
  }} finally {{
    event.target.value = '';
  }}
}});
$('#importPreviousPlanBtn').addEventListener('click', () => $('#previousPlanXlsxInput').click());
$('#previousPlanXlsxInput').addEventListener('change', async event => {{
  const file = event.target.files[0];
  if (!file) return;
  try {{
    const result = await importPreviousPlanExcel(file);
    alert(`Đã nhập ${{result.validRows}} dòng quy hoạch kỳ trước từ XLSX`);
  }} catch (error) {{
    alert(error.message || String(error));
  }} finally {{
    event.target.value = '';
  }}
}});
$('#jsonBtn').addEventListener('click', () => download('du_lieu_chu_chuyen_dat_dai.json', 'application/json;charset=utf-8', JSON.stringify(readInputs(), null, 2)));
$('#xlsxBtn').addEventListener('click', exportXlsx);
$('#csvBtn').addEventListener('click', exportCsv);
$('#printBtn').addEventListener('click', () => window.print());
$('#clearBtn').addEventListener('click', () => {{
  if (!confirm('Xóa toàn bộ dữ liệu nhập trong trang?')) return;
  const projectSettings = readProjectSettings();
  inputEls.forEach(input => input.value = '');
  applyPreviousPlanValues({{}});
  applyProjectSettings(projectSettings);
  localStorage.setItem(storageKey, JSON.stringify(readInputs()));
  recalc();
}});
$('#loadBtn').addEventListener('click', () => $('#fileInput').click());
$('#fileInput').addEventListener('change', async event => {{
  const file = event.target.files[0];
  if (!file) return;
  applyProjectData(gtpDataFromPayload(JSON.parse(await file.text())));
  event.target.value = '';
}});

const saved = localStorage.getItem(storageKey);
if (saved) applyInputs(JSON.parse(saved));
normalizeAllInputs();
applyHideZeroState(localStorage.getItem(hideZeroKey) === '1');
updateLibrarySessionUi();
$('#statusMissing').textContent = meta.missingCodes.length ? `Thiếu mã: ${{meta.missingCodes.join(', ')}}` : 'Đủ mã nhập';
$('#statusMissing').classList.toggle('warn', meta.missingCodes.length > 0);
recalc();
loadProjectFromServer();
</script>
</body>
</html>
"""
    OUT.write_text(doc, encoding="utf-8")
    print(OUT)
    print("input_codes=", ",".join(input_codes))
    print("missing_codes=", ",".join(missing_codes))


if __name__ == "__main__":
    main()
