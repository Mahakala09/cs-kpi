import { useState, useRef } from "react";

// --- Types ---
interface CriteriaItem {
  name: string;
  weight: number;
}
interface CriteriaCategory {
  category: string;
  items: CriteriaItem[];
}
interface Level {
  label: string;
  value: number;
  color: string;
}
interface Info {
  name: string;
  dept: string;
  id: string;
  evaluator: string;
  period: string;
  date: string;
}
interface Comments {
  achievements: string;
  improvements: string;
  plan: string;
  overall: string;
}
interface Grade {
  text: string;
  color: string;
}

// --- Data ---
const criteria: CriteriaCategory[] = [
  {
    category: "服务质量",
    items: [
      { name: "客户满意度评分（CSAT）", weight: 15 },
      { name: "服务态度与礼貌用语", weight: 10 },
      { name: "问题解决率（首次解决率）", weight: 10 },
    ],
  },
  {
    category: "工作效率",
    items: [
      { name: "平均响应时间", weight: 10 },
      { name: "日均处理工单/通话量", weight: 10 },
      { name: "工单处理及时率", weight: 5 },
    ],
  },
  {
    category: "专业能力",
    items: [
      { name: "产品/业务知识掌握", weight: 10 },
      { name: "沟通表达能力", weight: 5 },
      { name: "情绪管理与抗压能力", weight: 5 },
    ],
  },
  {
    category: "工作态度",
    items: [
      { name: "出勤与纪律", weight: 5 },
      { name: "主动性与责任心", weight: 5 },
      { name: "团队协作", weight: 5 },
    ],
  },
  {
    category: "学习与成长",
    items: [
      { name: "培训参与及考核成绩", weight: 3 },
      { name: "改进与自我提升", weight: 2 },
    ],
  },
];

const levels: Level[] = [
  { label: "优秀", value: 5, color: "#10b981" },
  { label: "良好", value: 4, color: "#3b82f6" },
  { label: "合格", value: 3, color: "#f59e0b" },
  { label: "待改进", value: 2, color: "#f97316" },
  { label: "不合格", value: 1, color: "#ef4444" },
];

const levelText: Record<number, string> = {
  5: "优秀",
  4: "良好",
  3: "合格",
  2: "待改进",
  1: "不合格",
};

// --- Excel Export ---
function exportToExcel(
  info: Info,
  scores: Record<string, number>,
  comments: Comments,
  finalScore: string,
  grade: Grade,
  criteriaData: CriteriaCategory[]
) {
  const getScore = (cat: string, item: string) => scores[`${cat}-${item}`] || 0;
  const rows: string[] = [];

  rows.push(`<Row ss:Height="30"><Cell ss:MergeAcross="4" ss:StyleID="title"><Data ss:Type="String">客服人员工作评估表</Data></Cell></Row>`);
  rows.push(`<Row><Cell ss:MergeAcross="4"><Data ss:Type="String"></Data></Cell></Row>`);

  const infoRows: string[][] = [
    [`姓名：${info.name}`, `工号：${info.id}`, `部门：${info.dept}`],
    [`评估人：${info.evaluator}`, `评估周期：${info.period}`, `评估日期：${info.date}`],
  ];
  infoRows.forEach(r => {
    rows.push(`<Row>${r.map((c, i) => {
      const merge = i === 2 ? ' ss:MergeAcross="2"' : i === 0 ? ' ss:MergeAcross="1"' : '';
      return `<Cell${merge} ss:StyleID="info"><Data ss:Type="String">${c}</Data></Cell>`;
    }).join("")}</Row>`);
  });

  rows.push(`<Row><Cell><Data ss:Type="String"></Data></Cell></Row>`);
  rows.push(`<Row ss:Height="24">
    <Cell ss:StyleID="header"><Data ss:Type="String">评估维度</Data></Cell>
    <Cell ss:StyleID="header"><Data ss:Type="String">评估指标</Data></Cell>
    <Cell ss:StyleID="header"><Data ss:Type="String">权重</Data></Cell>
    <Cell ss:StyleID="header"><Data ss:Type="String">评分</Data></Cell>
    <Cell ss:StyleID="header"><Data ss:Type="String">加权得分</Data></Cell>
  </Row>`);

  criteriaData.forEach(cat => {
    cat.items.forEach((item, ii) => {
      const sc = getScore(cat.category, item.name);
      const weighted = sc > 0 ? ((sc / 5) * item.weight).toFixed(1) : "0";
      const scoreLabel = sc > 0 ? `${sc}分（${levelText[sc]}）` : "未评分";
      rows.push(`<Row>
        <Cell ss:StyleID="catCell"><Data ss:Type="String">${ii === 0 ? cat.category : ""}</Data></Cell>
        <Cell ss:StyleID="cell"><Data ss:Type="String">${item.name}</Data></Cell>
        <Cell ss:StyleID="centerCell"><Data ss:Type="String">${item.weight}%</Data></Cell>
        <Cell ss:StyleID="centerCell"><Data ss:Type="String">${scoreLabel}</Data></Cell>
        <Cell ss:StyleID="centerCell"><Data ss:Type="Number">${weighted}</Data></Cell>
      </Row>`);
    });
  });

  rows.push(`<Row><Cell><Data ss:Type="String"></Data></Cell></Row>`);
  rows.push(`<Row>
    <Cell ss:MergeAcross="2" ss:StyleID="totalLabel"><Data ss:Type="String">加权总分</Data></Cell>
    <Cell ss:StyleID="totalValue"><Data ss:Type="Number">${finalScore}</Data></Cell>
    <Cell ss:StyleID="totalValue"><Data ss:Type="String">等级：${grade.text}</Data></Cell>
  </Row>`);

  rows.push(`<Row><Cell><Data ss:Type="String"></Data></Cell></Row>`);

  const commentSections: [string, string][] = [
    ["主要成绩与贡献", comments.achievements],
    ["不足与改进建议", comments.improvements],
    ["下一周期发展计划", comments.plan],
    ["综合评语", comments.overall],
  ];
  commentSections.forEach(([label, val]) => {
    rows.push(`<Row><Cell ss:StyleID="commentLabel"><Data ss:Type="String">${label}</Data></Cell><Cell ss:MergeAcross="3" ss:StyleID="commentVal"><Data ss:Type="String">${val || ""}</Data></Cell></Row>`);
  });

  rows.push(`<Row><Cell><Data ss:Type="String"></Data></Cell></Row>`);
  rows.push(`<Row>
    <Cell ss:StyleID="info"><Data ss:Type="String">被评估人签字：</Data></Cell>
    <Cell ss:StyleID="info"><Data ss:Type="String"></Data></Cell>
    <Cell ss:StyleID="info"><Data ss:Type="String">直属上级签字：</Data></Cell>
    <Cell ss:StyleID="info"><Data ss:Type="String"></Data></Cell>
    <Cell ss:StyleID="info"><Data ss:Type="String">部门负责人签字：</Data></Cell>
  </Row>`);

  const xml = `<?xml version="1.0" encoding="UTF-8"?>
<?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">
 <Styles>
  <Style ss:ID="Default"><Alignment ss:Vertical="Center" ss:WrapText="1"/><Font ss:FontName="微软雅黑" ss:Size="10"/></Style>
  <Style ss:ID="title"><Alignment ss:Horizontal="Center" ss:Vertical="Center"/><Font ss:FontName="微软雅黑" ss:Size="16" ss:Bold="1" ss:Color="#1e40af"/></Style>
  <Style ss:ID="info"><Font ss:FontName="微软雅黑" ss:Size="10"/><Alignment ss:Vertical="Center"/></Style>
  <Style ss:ID="header"><Interior ss:Color="#1e40af" ss:Pattern="Solid"/><Font ss:FontName="微软雅黑" ss:Size="10" ss:Bold="1" ss:Color="#FFFFFF"/><Alignment ss:Horizontal="Center" ss:Vertical="Center"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/></Borders></Style>
  <Style ss:ID="cell"><Alignment ss:Vertical="Center" ss:WrapText="1"/><Font ss:FontName="微软雅黑" ss:Size="10"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#e2e8f0"/></Borders></Style>
  <Style ss:ID="catCell"><Alignment ss:Vertical="Center"/><Font ss:FontName="微软雅黑" ss:Size="10" ss:Bold="1"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#e2e8f0"/></Borders></Style>
  <Style ss:ID="centerCell"><Alignment ss:Horizontal="Center" ss:Vertical="Center"/><Font ss:FontName="微软雅黑" ss:Size="10"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#e2e8f0"/></Borders></Style>
  <Style ss:ID="totalLabel"><Alignment ss:Horizontal="Right" ss:Vertical="Center"/><Font ss:FontName="微软雅黑" ss:Size="11" ss:Bold="1"/></Style>
  <Style ss:ID="totalValue"><Alignment ss:Horizontal="Center" ss:Vertical="Center"/><Font ss:FontName="微软雅黑" ss:Size="14" ss:Bold="1" ss:Color="#1e40af"/></Style>
  <Style ss:ID="commentLabel"><Font ss:FontName="微软雅黑" ss:Size="10" ss:Bold="1"/><Alignment ss:Vertical="Top"/><Interior ss:Color="#f1f5f9" ss:Pattern="Solid"/></Style>
  <Style ss:ID="commentVal"><Font ss:FontName="微软雅黑" ss:Size="10"/><Alignment ss:Vertical="Top" ss:WrapText="1"/></Style>
 </Styles>
 <Worksheet ss:Name="客服评估表">
  <Table ss:DefaultRowHeight="22">
   <Column ss:Width="100"/>
   <Column ss:Width="200"/>
   <Column ss:Width="70"/>
   <Column ss:Width="120"/>
   <Column ss:Width="100"/>
   ${rows.join("\n   ")}
  </Table>
 </Worksheet>
</Workbook>`;

  const blob = new Blob([xml], { type: "application/vnd.ms-excel" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  const empName = info.name || "员工";
  const dt = info.date || new Date().toISOString().slice(0, 10);
  a.download = `客服评估表_${empName}_${dt}.xls`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

// --- PDF Export (html2canvas + jsPDF loaded from CDN at runtime) ---
async function exportToPDF(element: HTMLElement, info: Info) {
  const loadScript = (src: string): Promise<void> =>
    new Promise((resolve, reject) => {
      if (document.querySelector(`script[src="${src}"]`)) { resolve(); return; }
      const s = document.createElement("script");
      s.src = src;
      s.onload = () => resolve();
      s.onerror = () => reject(new Error(`Failed to load ${src}`));
      document.head.appendChild(s);
    });

  await loadScript("https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js");
  await loadScript("https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js");

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const html2canvas = (window as any).html2canvas;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const { jsPDF } = (window as any).jspdf;

  const canvas = await html2canvas(element, {
    scale: 2,
    useCORS: true,
    backgroundColor: "#ffffff",
    logging: false,
  });

  const imgData = canvas.toDataURL("image/png");
  const pdf = new jsPDF({ orientation: "portrait", unit: "mm", format: "a4" });

  const pageWidth = pdf.internal.pageSize.getWidth();   // 210mm
  const pageHeight = pdf.internal.pageSize.getHeight(); // 297mm
  const margin = 10;
  const printWidth = pageWidth - margin * 2;
  const printHeight = (canvas.height / canvas.width) * printWidth;

  if (printHeight <= pageHeight - margin * 2) {
    // Fits on one page
    pdf.addImage(imgData, "PNG", margin, margin, printWidth, printHeight);
  } else {
    // Multi-page split
    const ratio = canvas.width / printWidth;
    const pageHeightPx = (pageHeight - margin * 2) * ratio;
    let offsetY = 0;
    while (offsetY < canvas.height) {
      const sliceHeight = Math.min(pageHeightPx, canvas.height - offsetY);
      const sliceCanvas = document.createElement("canvas");
      sliceCanvas.width = canvas.width;
      sliceCanvas.height = sliceHeight;
      const ctx = sliceCanvas.getContext("2d")!;
      ctx.drawImage(canvas, 0, -offsetY);
      const sliceData = sliceCanvas.toDataURL("image/png");
      const slicePrintHeight = sliceHeight / ratio;
      pdf.addImage(sliceData, "PNG", margin, margin, printWidth, slicePrintHeight);
      offsetY += sliceHeight;
      if (offsetY < canvas.height) pdf.addPage();
    }
  }

  const empName = info.name || "员工";
  const dt = info.date || new Date().toISOString().slice(0, 10);
  pdf.save(`客服评估表_${empName}_${dt}.pdf`);
}

// --- Main Component ---
export default function CSEvaluation() {
  const [info, setInfo] = useState<Info>({ name: "", dept: "", id: "", evaluator: "", period: "", date: "" });
  const [scores, setScores] = useState<Record<string, number>>({});
  const [comments, setComments] = useState<Comments>({ achievements: "", improvements: "", plan: "", overall: "" });
  const [exportedXls, setExportedXls] = useState(false);
  const [exportedPdf, setExportedPdf] = useState(false);
  const [pdfLoading, setPdfLoading] = useState(false);
  const printRef = useRef<HTMLDivElement>(null);

  const setScore = (cat: string, item: string, val: number) => {
    setScores((p) => ({ ...p, [`${cat}-${item}`]: val }));
  };

  const getScore = (cat: string, item: string): number => scores[`${cat}-${item}`] || 0;

  const totalWeight = criteria.reduce((s, c) => s + c.items.reduce((a, i) => a + i.weight, 0), 0);
  const weightedTotal = criteria.reduce(
    (s, c) => s + c.items.reduce((a, i) => a + (getScore(c.category, i.name) / 5) * i.weight, 0), 0
  );
  const finalScore = totalWeight > 0 ? ((weightedTotal / totalWeight) * 100).toFixed(1) : "0";

  const getGrade = (s: number): Grade => {
    if (s >= 90) return { text: "优秀", color: "#10b981" };
    if (s >= 80) return { text: "良好", color: "#3b82f6" };
    if (s >= 60) return { text: "合格", color: "#f59e0b" };
    if (s >= 40) return { text: "待改进", color: "#f97316" };
    return { text: "不合格", color: "#ef4444" };
  };

  const grade = getGrade(parseFloat(finalScore));
  const allScored = criteria.every((c) => c.items.every((i) => getScore(c.category, i.name) > 0));

  const handleExportXls = () => {
    exportToExcel(info, scores, comments, finalScore, grade, criteria);
    setExportedXls(true);
    setTimeout(() => setExportedXls(false), 2500);
  };

  const handleExportPdf = async () => {
    if (!printRef.current) return;
    setPdfLoading(true);
    try {
      await exportToPDF(printRef.current, info);
      setExportedPdf(true);
      setTimeout(() => setExportedPdf(false), 2500);
    } catch (e) {
      console.error("PDF export failed:", e);
      alert("PDF 导出失败，请检查网络连接后重试。");
    } finally {
      setPdfLoading(false);
    }
  };

  return (
    <div style={{ fontFamily: "system-ui, sans-serif", maxWidth: 800, margin: "0 auto", padding: 20, color: "#1e293b" }}>

      {/* ===== PRINTABLE AREA — export buttons are OUTSIDE this div ===== */}
      <div ref={printRef} style={{ background: "#fff", padding: 4 }}>

        {/* Header */}
        <div style={{ textAlign: "center", marginBottom: 24, borderBottom: "3px solid #1e40af", paddingBottom: 16 }}>
          <h1 style={{ fontSize: 22, fontWeight: 700, color: "#1e40af", margin: 0 }}>客服人员工作评估表</h1>
          <p style={{ color: "#64748b", fontSize: 13, margin: "6px 0 0" }}>Customer Service Performance Evaluation</p>
        </div>

        {/* Basic Info */}
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 10, marginBottom: 20 }}>
          {(
            [
              ["姓名", "name"], ["工号", "id"], ["部门", "dept"],
              ["评估人", "evaluator"], ["评估周期", "period"], ["评估日期", "date"],
            ] as [string, keyof Info][]
          ).map(([label, key]) => (
            <div key={key} style={{ display: "flex", alignItems: "center", gap: 6 }}>
              <label style={{ fontSize: 13, fontWeight: 600, whiteSpace: "nowrap", minWidth: 56 }}>{label}</label>
              <input
                value={info[key]}
                onChange={(e) => setInfo((p) => ({ ...p, [key]: e.target.value }))}
                placeholder={`请输入${label}`}
                style={{ flex: 1, border: "1px solid #cbd5e1", borderRadius: 4, padding: "5px 8px", fontSize: 13, outline: "none", minWidth: 0 }}
              />
            </div>
          ))}
        </div>

        {/* Rating Legend */}
        <div style={{ display: "flex", gap: 12, marginBottom: 16, flexWrap: "wrap", fontSize: 12 }}>
          {levels.map((l) => (
            <span key={l.value} style={{ display: "flex", alignItems: "center", gap: 4 }}>
              <span style={{ width: 10, height: 10, borderRadius: "50%", background: l.color, display: "inline-block" }} />
              {l.value}分 = {l.label}
            </span>
          ))}
        </div>

        {/* Scoring Table */}
        <table style={{ width: "100%", borderCollapse: "collapse", marginBottom: 20, fontSize: 13 }}>
          <thead>
            <tr style={{ background: "#1e40af", color: "#fff" }}>
              <th style={{ padding: "8px 10px", textAlign: "left", width: "15%" }}>评估维度</th>
              <th style={{ padding: "8px 10px", textAlign: "left", width: "35%" }}>评估指标</th>
              <th style={{ padding: "8px 10px", textAlign: "center", width: "10%" }}>权重</th>
              <th style={{ padding: "8px 10px", textAlign: "center", width: "30%" }}>评分（1-5）</th>
              <th style={{ padding: "8px 10px", textAlign: "center", width: "10%" }}>得分</th>
            </tr>
          </thead>
          <tbody>
            {criteria.map((cat, ci) =>
              cat.items.map((item, ii) => {
                const sc = getScore(cat.category, item.name);
                const weighted = sc > 0 ? ((sc / 5) * item.weight).toFixed(1) : "-";
                return (
                  <tr key={`${ci}-${ii}`} style={{ background: ci % 2 === 0 ? "#f8fafc" : "#fff", borderBottom: "1px solid #e2e8f0" }}>
                    {ii === 0 && (
                      <td rowSpan={cat.items.length} style={{ padding: "8px 10px", fontWeight: 600, background: ci % 2 === 0 ? "#eff6ff" : "#f0f9ff", verticalAlign: "middle", borderRight: "1px solid #e2e8f0" }}>
                        {cat.category}
                      </td>
                    )}
                    <td style={{ padding: "8px 10px" }}>{item.name}</td>
                    <td style={{ padding: "8px 10px", textAlign: "center", fontWeight: 600 }}>{item.weight}%</td>
                    <td style={{ padding: "4px 10px", textAlign: "center" }}>
                      <div style={{ display: "flex", gap: 4, justifyContent: "center" }}>
                        {levels.map((l) => (
                          <button
                            key={l.value}
                            onClick={() => setScore(cat.category, item.name, l.value)}
                            style={{
                              width: 32, height: 28, borderRadius: 4, border: "1px solid",
                              borderColor: sc === l.value ? l.color : "#d1d5db",
                              background: sc === l.value ? l.color : "#fff",
                              color: sc === l.value ? "#fff" : "#64748b",
                              fontWeight: 600, fontSize: 12, cursor: "pointer",
                              transition: "all 0.15s",
                            }}
                          >
                            {l.value}
                          </button>
                        ))}
                      </div>
                    </td>
                    <td style={{ padding: "8px 10px", textAlign: "center", fontWeight: 600, color: sc > 0 ? "#1e40af" : "#94a3b8" }}>
                      {weighted}
                    </td>
                  </tr>
                );
              })
            )}
          </tbody>
        </table>

        {/* Score Summary */}
        <div style={{ display: "flex", gap: 16, marginBottom: 20, padding: 16, background: "#f8fafc", borderRadius: 8, border: "1px solid #e2e8f0", alignItems: "center", justifyContent: "center" }}>
          <div style={{ textAlign: "center" }}>
            <div style={{ fontSize: 12, color: "#64748b", marginBottom: 4 }}>加权总分</div>
            <div style={{ fontSize: 28, fontWeight: 700, color: "#1e40af" }}>{allScored ? finalScore : "--"}</div>
          </div>
          <div style={{ width: 1, height: 40, background: "#cbd5e1" }} />
          <div style={{ textAlign: "center" }}>
            <div style={{ fontSize: 12, color: "#64748b", marginBottom: 4 }}>评定等级</div>
            <div style={{ fontSize: 20, fontWeight: 700, color: allScored ? grade.color : "#94a3b8", background: allScored ? `${grade.color}18` : "#f1f5f9", padding: "2px 16px", borderRadius: 6 }}>
              {allScored ? grade.text : "--"}
            </div>
          </div>
        </div>

        {/* Comments */}
        <div style={{ display: "grid", gap: 12, marginBottom: 20 }}>
          {(
            [
              ["主要成绩与贡献", "achievements", "请列出该员工在本评估周期内的突出表现和主要贡献..."],
              ["不足与改进建议", "improvements", "请指出需要改进的方面及具体建议..."],
              ["下一周期发展计划", "plan", "请填写下一评估周期的目标和培训发展计划..."],
              ["综合评语", "overall", "请填写总体评价..."],
            ] as [string, keyof Comments, string][]
          ).map(([label, key, ph]) => (
            <div key={key}>
              <label style={{ fontSize: 13, fontWeight: 600, display: "block", marginBottom: 4 }}>{label}</label>
              <textarea
                value={comments[key]}
                onChange={(e) => setComments((p) => ({ ...p, [key]: e.target.value }))}
                placeholder={ph}
                rows={3}
                style={{ width: "100%", border: "1px solid #cbd5e1", borderRadius: 6, padding: 8, fontSize: 13, resize: "vertical", outline: "none", boxSizing: "border-box" }}
              />
            </div>
          ))}
        </div>

        {/* Signatures */}
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 16, padding: 16, borderTop: "2px solid #e2e8f0" }}>
          {["被评估人签字", "直属上级签字", "部门负责人签字"].map((label) => (
            <div key={label} style={{ textAlign: "center" }}>
              <div style={{ fontSize: 13, fontWeight: 600, marginBottom: 8 }}>{label}</div>
              <div style={{ borderBottom: "1px solid #94a3b8", height: 40 }} />
              <div style={{ fontSize: 11, color: "#94a3b8", marginTop: 4 }}>日期：____年____月____日</div>
            </div>
          ))}
        </div>

      </div>
      {/* ===== END PRINTABLE AREA ===== */}

      {/* Export Buttons */}
      <div style={{ textAlign: "center", marginTop: 24, display: "flex", gap: 12, justifyContent: "center", flexWrap: "wrap" }}>
        <button
          onClick={handleExportXls}
          style={{
            padding: "12px 32px", fontSize: 14, fontWeight: 700, color: "#fff",
            background: exportedXls ? "#10b981" : "linear-gradient(135deg, #1e40af, #3b82f6)",
            border: "none", borderRadius: 8, cursor: "pointer",
            boxShadow: "0 2px 8px rgba(30,64,175,0.3)",
            transition: "all 0.2s",
          }}
        >
          {exportedXls ? "✓ 导出成功！" : "📊 导出为 Excel"}
        </button>

        <button
          onClick={handleExportPdf}
          disabled={pdfLoading}
          style={{
            padding: "12px 32px", fontSize: 14, fontWeight: 700, color: "#fff",
            background: exportedPdf ? "#10b981" : pdfLoading ? "#94a3b8" : "linear-gradient(135deg, #dc2626, #f97316)",
            border: "none", borderRadius: 8, cursor: pdfLoading ? "not-allowed" : "pointer",
            boxShadow: "0 2px 8px rgba(220,38,38,0.3)",
            transition: "all 0.2s",
          }}
        >
          {exportedPdf ? "✓ 导出成功！" : pdfLoading ? "⏳ 生成中..." : "📄 导出为 PDF"}
        </button>
      </div>

      <p style={{ textAlign: "center", fontSize: 11, color: "#94a3b8", marginTop: 8 }}>
        Excel 格式可用 Excel / WPS 打开 · PDF 格式适合打印归档
      </p>
    </div>
  );
}
