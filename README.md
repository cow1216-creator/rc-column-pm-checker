RC Column P-M Interaction Checker / RC 矩形柱 P-M 互交曲線檢核工具
> **Excel-based P-M interaction analysis for rectangular RC columns with arbitrary seismic load angles.** No VBA, no macros — pure Excel formulas using Sutherland-Hodgman polygon clipping for the compression zone geometry.
---
English
Overview
This Excel workbook performs P-M interaction diagram analysis for rectangular reinforced concrete columns under biaxial bending by rotating the cross-section to an arbitrary angle θ. Unlike conventional tools that only check X- and Y- direction independently (or use approximations like the Bresler load contour method), this tool computes the true single-axis P-M diagram in the actual seismic loading direction.
Calculations follow the standard strain compatibility method consistent with ACI 318 and the Taiwan Concrete Structure Design Code (TW NSC).
Key features
✅ Arbitrary seismic angle θ (0°–360°) — true biaxial bending capability
✅ Polygon clipping by Sutherland-Hodgman algorithm — handles rotated stress block correctly (triangular, trapezoidal, or pentagonal compression zones)
✅ Up to 20 × 20 = 400 main rebars with arbitrary coordinates (asymmetric layouts supported via the `bar` sheet)
✅ Whitney equivalent stress block (ACI 22.2) with β₁ auto-computed from f'c
✅ φ factor with smooth transition (0.65 → 0.90) per ACI 21.2 / TW NSC §3
✅ 21-point P-M curve per angle, with linear interpolation between adjacent points to find design intersection
✅ No VBA / no macros — pure Excel formulas, safe to open on any machine
✅ Includes geometric output sheets (R0–R20) showing actual clipped compression polygon for visual verification
Quick start
Open `RC-Column-PM-Checker-EN.xlsx`
Go to the `check` sheet and fill in the yellow input cells:
Column dimensions B (X-dir), D (Y-dir) in cm
Concrete strength f'c, steel yield strength fy in kgf/cm²
Main rebar count (bars per side, parallel to width and depth)
Cover to bar centroid d'
Confinement type (1 = spiral, 2 = rectangular tie)
Seismic angle θ (degrees, clockwise positive) ← key input for biaxial bending
Demand axial Pu (tf) and bending moment Mu (tf-m)
The P-M curve and check result update automatically
Verify the rebar layout in the `bar` sheet matches your intended configuration
Calculation reference
Concrete model: Whitney equivalent stress block, 0.85 f'c × a, where a = β₁ × c
β₁: linearly interpolated from 0.85 (f'c ≤ 280 kgf/cm²) to 0.65 (f'c ≥ 560 kgf/cm²)
Steel model: elastic-perfectly plastic, fs = E·ε bounded by ±fy (E = 2,040,000 kgf/cm²)
Ultimate concrete strain: ε_cu = 0.003
Compression zone: computed by rotating the rectangular cross-section by (90° + θ) and clipping by a horizontal line at the neutral axis position, using Sutherland-Hodgman algorithm
Pure compression cap: 0.80 × P₀ per ACI 22.4.2.1
φ factor:
φ = 0.65 if εt ≤ 0.002 (compression-controlled)
φ = 0.90 if εt ≥ 0.005 (tension-controlled)
Linear transition between the two
Limitations / Known constraints
Item	Constraint
Cross-section shape	Rectangular only (L, T, circular, hollow not supported)
Maximum rebar count	20 × 20 = 400 bars
Bar layout	Symmetric grid (custom positions possible via `bar` sheet, but with caveats)
Bar buckling	Not considered (assumes adequate tie spacing)
Confined concrete	Standard 0.85 f'c only — Mander or similar confined models not implemented
P-Δ slenderness	Not considered — handle separately per code provisions
Number of P-M points	21 fixed (sheets R0 to R20)
Single-angle analysis	One θ at a time — sweep manually for envelope
File structure
```
RC-Column-PM-Checker-EN.xlsx
├── check         ← main input/output sheet
├── rebarps       ← rebar property table
├── rectangle     ← cross-section geometric helper
├── bar           ← rebar coordinates (X, Y, As per bar)
└── R0 ... R20    ← 21 P-M points (one neutral-axis position each)
```

Contributing
Found a bug? Have a suggestion? Please open an issue on GitHub or contact the author.
Disclaimer
This tool is provided as-is for educational and engineering reference. The author makes no warranty regarding the accuracy or fitness for any particular purpose. Users are responsible for verifying results against established codes and engineering judgment before applying to any actual project.
---
中文
工具概述
本 Excel 工具是針對矩形 RC 柱進行 P-M 互交曲線分析的試算表，支援任意地震角度 θ 的雙向彎曲。不同於一般工具只能各別檢核 X、Y 方向（或用 Bresler 載重輪廓法近似），本工具透過把斷面旋轉到實際地震方向，直接算出該方向上真實的單軸 P-M 曲線。
計算依據符合 ACI 318 與 中華民國混凝土結構設計規範的應變相容法。
主要功能
✅ 任意地震角度 θ（0°–360°）—— 真實雙向彎曲檢核能力
✅ Sutherland-Hodgman 多邊形裁切演算法 —— 正確處理旋轉後的壓力塊（三角形、梯形或五邊形壓力區皆可）
✅ 最多 20 × 20 = 400 支主筋，支援不對稱配置（透過 `bar` 分頁）
✅ Whitney 等值應力塊（ACI 22.2），β₁ 依 f'c 自動計算
✅ φ 過渡平滑（0.65 → 0.90），依 ACI 21.2 / 規範 §3
✅ 每個角度產生 21 點 P-M 曲線，相鄰點線性內插求設計交點
✅ 無 VBA、無巨集 —— 純 Excel 公式，任何電腦開啟都安全
✅ 含幾何輸出分頁（R0–R20）顯示實際裁切後的壓力多邊形，方便視覺驗證
快速開始
開啟 `RC-Column-PM-Checker-EN.xlsx`（英文版）或 `矩形柱檢核p-m-轉角度-final1150527.xlsx`（中文原版）
進入 `check` 分頁，填寫黃色輸入欄位：
柱尺寸 B（X 向）、D（Y 向），單位 cm
混凝土強度 f'c、鋼筋降伏強度 fy，單位 kgf/cm²
主筋支數（平行柱寬、平行柱深）
主筋形心至外緣 d'
圍束型式（1 = 螺箍，2 = 橫箍）
地震力轉角 θ（度，順時針為正） ← 雙向彎曲的關鍵輸入
需求軸力 Pu (tf) 與需求彎矩 Mu (tf-m)
PM 曲線與檢核結果會自動更新
至 `bar` 分頁確認鋼筋配置與預期一致
計算依據
混凝土模型：Whitney 等值應力塊，0.85 f'c × a，其中 a = β₁ × c
β₁：從 0.85（f'c ≤ 280 kgf/cm²）線性內插到 0.65（f'c ≥ 560 kgf/cm²）
鋼筋模型：彈完全塑性，fs = E·ε，截斷在 ±fy（E = 2,040,000 kgf/cm²）
混凝土極限應變：ε_cu = 0.003
壓力塊：將矩形斷面旋轉 (90° + θ) 度後，以中性軸位置的水平線做 Sutherland-Hodgman 裁切，求出實際壓力多邊形
純壓上限：0.80 × P₀，依 ACI 22.4.2.1
φ 強度折減因子：
εt ≤ 0.002 時 φ = 0.65（壓力控制）
εt ≥ 0.005 時 φ = 0.90（拉力控制）
兩者之間線性過渡
限制與已知條件
項目	限制
斷面型式	僅限矩形（L 型、T 型、圓形、空心柱不支援）
最大鋼筋數	20 × 20 = 400 支
鋼筋配置	對稱矩陣（透過 `bar` 分頁可設定自訂座標，但需注意對稱性）
鋼筋挫屈	未考慮（假設箍筋間距足夠）
圍束混凝土	僅使用 0.85 f'c——未實作 Mander 等圍束模型
P-Δ 二階效應	未考慮——請依規範另行檢核長細比
P-M 點數	固定 21 點（R0 至 R20 分頁）
單一角度	每次僅能計算一個 θ，包絡圖需手動掃描
檔案結構
```
RC-Column-PM-Checker-EN.xlsx
├── check         ← 主要輸入/輸出分頁
├── rebarps       ← 鋼筋規格表
├── rectangle     ← 斷面幾何輔助
├── bar           ← 鋼筋座標 (X, Y, 每根 As)
└── R0 ... R20    ← 21 個 P-M 點 (每個對應一個中性軸位置)
```

回饋
發現 bug？有建議？請在 GitHub 開 issue，或直接聯絡作者。
免責聲明
本工具現狀提供（as-is），僅供教育與工程參考。作者不對結果的正確性或適用性提供任何保證。使用者於正式專案應用前，有責任自行依規範與工程判斷驗證結果。
---

---
Contact / 聯絡資訊
Author / 作者:dihung chen
Email: cow1216.2@gmail.com
---
Last updated: 2026-05-27
