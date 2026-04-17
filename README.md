# Deflection Check Tool — SAP2000

Kiểm tra độ võng dầm kết cấu thép, kết nối trực tiếp SAP2000 qua COM API.

## Yêu cầu
- Windows, Python 3.8+, SAP2000 đang mở và đã Run Analysis
- `pip install comtypes openpyxl`

## Chạy Tool
```bash
python main.py       # hoặc double-click run.bat
```

---

## Hướng dẫn sử dụng

### ① Beam Groups
- **Load Groups**: Đọc toàn bộ Groups từ SAP2000
- **Select All / Select Beams / Clear**: Lọc nhanh
- **📌 From SAP**: Bôi đen Groups (G-...) chứa các frame đang được chọn trong SAP
- **⚙ Auto Create**: Tự động tạo Groups từ PMember detection
  - *Auto Mesh*: 1 group = 1 frame element (dùng khi đã meshed)
  - *No Auto Mesh*: Gom các frame thành Physical Member (dùng section + topology)
- **🗑 Remove**: Xóa Groups đã chọn khỏi SAP (chỉ xóa Group, không xóa member)
- **✏ Manual Create**: Tạo Group từ các frame đang select trong SAP (phải contiguous)
- Click vào Group trong danh sách → SAP2000 tự động highlight Group đó

### ② Load Cases / Combinations
- Chọn Load Cases hoặc Load Combinations cần kiểm tra
- **Quick prefix**: Gõ prefix và nhấn Select để chọn nhanh nhiều LCs cùng lúc

### ③ Settings — Chọn 1 trong 2 loại kiểm tra:
- **Relative check  L / [360]**: Kiểm tra L/δ (chord rotation method, 3D)
  - Cantilever: L/δ ≥ L/180 (tự động detect)
  - Regular beam: L/δ ≥ L/360 (hoặc giá trị tự đặt)
- **Absolute check  Limit: [25] mm**: Kiểm tra U3 (thẳng đứng) ≤ giới hạn mm

### ④ Results
- Bảng kết quả sau khi RUN CHECK
- **Show NG only**: Lọc chỉ hiện các beam bị NG
- Màu xanh = OK, đỏ = NG

### Export
- **📊 Excel**: Xuất file `.xlsx` với bảng summary
- **📄 TXT**: Xuất file text dạng cột
- Output: `[SAP model dir]/Deflection Check/`

---

## Thuật toán tính toán

### Relative check (chord rotation method)
1. Lấy toàn bộ nodes dọc beam (bao gồm auto-mesh nodes `~xxx`)
2. Tính **original chord**: vector từ node đầu đến node cuối theo geometry gốc
3. Với mỗi node trung gian:
   - Tính `t` = projection của node gốc lên chord gốc
   - Tính rigid-body motion = nội suy tuyến tính giữa displacement 2 đầu
   - Relative deformation = displacement thực − rigid-body motion
   - Perpendicular component = khoảng cách vuông góc với chord gốc
4. Max perpendicular distance = δ thực tế
5. So sánh L/δ với giới hạn

### Cantilever detection
- Build `node_degree_map`: đếm số FrameObj kết nối tại mỗi node
- Endpoint có degree = 1 → free end → is_cantilever = True
- Tính δ tại free end so với fixed end, chiếu vuông góc với trục beam gốc

### Absolute check
- Lấy U3 (displacement thẳng đứng) từ SAP2000 cho tất cả nodes
- Max |U3| = δ tuyệt đối
- So sánh với giới hạn mm đặt trước

### PMember detection (Auto Create Groups — No Auto Mesh)
- **Pass 1**: Phân loại COLUMN / BEAM / POST (vertical + có release → POST)
- **Pass 2**: Standalone VBR/HBR
- **Pass 3a**: Chain POST elements
- **Pass 3b**: Chain BEAM elements (dừng tại col_nodes, dừng khi section thay đổi)
- **Pass 4**: Fallback beam với col_nodes boundary

---

## Cấu trúc file

```
14 Deflection Check/
├── main.py                        # Entry point
├── run.bat                        # Shortcut chạy tool
├── requirements.txt
├── def.ico                        # Icon
├── ui/
│   └── main_window.py             # Toàn bộ GUI (tkinter)
├── utils/
│   ├── sap_connector.py           # SAP2000 COM API wrapper
│   ├── deflection_calc.py         # Engine tính deflection
│   └── pmember_detector.py        # PMember chain detection
├── data/
│   └── models.py                  # DataClasses: BeamGroup, BeamCheckSummary, ...
└── output/
    └── exporters.py               # ExcelExporter, TxtExporter
```

---

## Lưu ý SAP2000 COM API

| Tác vụ | API |
|---|---|
| Lấy mesh nodes của frame | `FrameObj.GetElm()` → `LineElm.GetPoints()` |
| Tọa độ virtual node (`~xxx`) | `PointElm.GetCoordCartesian()` |
| Displacement PointObj | `Results.JointDispl(name, 0)` |
| Displacement PointElm (`~xxx`) | `Results.JointDispl(name, 1)` |
| Displacement mesh node (frame element) | `Results.FrameJointDispl(elm, 1)` — index: ACase=ret[5], U1=ret[8], U2=ret[9], U3=ret[10] |
| Select group trong SAP | `SelectObj.Group(name, False)` |
| Refresh viewport | `View.RefreshView(0, False)` |
| Tạo/xóa group | `GroupDef.SetGroup()` / `GroupDef.Delete()` |
| Assign frame vào group | `FrameObj.SetGroupAssign(frame, group)` |

---

## License
Expiry: 01/04/2027 — Liên hệ Roberto để gia hạn.
