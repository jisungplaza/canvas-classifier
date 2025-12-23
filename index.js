// index.js — 캔버스 자동 분류 서버
// 멀티 시트 + 파일명 인코딩
// 헤더 위치 자동 탐색 + RESULT / ITEM NO / DESCRIPTION / QUANTITY 4컬럼 출력
// ITEM NO 컬럼 폭 15 고정 + 나머지 자동폭
// ITEM NO, QUANTITY 둘 다 없는 요약줄(Grimso 등)은 RESULT 비우기
// 추가: Serial / Our Item No. / Description / Quantity 요약 시트 생성

const express = require("express");
const multer = require("multer");
const cors = require("cors");
const xlsx = require("xlsx");
const path = require("path");

const {
  sizeMap,
  typeRules,
  roundConfig,
  thicknessMap,
  excludedForThickness,
  manualOverrides,
  sizeTolerance,
  thicknessTolerance,
  defaultTypeLabel,
} = require("./config/canvasConfig");

const app = express();
const PORT = 3100;

app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, "public")));

const upload = multer({ storage: multer.memoryStorage() });

/* ---------- 공통 유틸 ---------- */

function normalize(text) {
  return (text || "").toString().trim().toLowerCase();
}

/* ---------- triangle 인식 ---------- */

function detectTriangle(text) {
  const t = normalize(text);
  if (!t.includes("triangle")) return null;
  const m = t.match(/triangle[^0-9]*([0-9]{1,3}(\.[0-9]+)?)/i);
  if (!m) return null;
  return parseFloat(m[1]);
}

/* ---------- 원형 인식 ---------- */

function detectRound(text) {
  const t = normalize(text);

  const mDia = t.match(/dia[^0-9]*([0-9]{1,3}(\.[0-9]+)?)/i);
  if (mDia) {
    const d = parseFloat(mDia[1]);
    return { isRound: true, diameter: d };
  }

  const found = roundConfig.keywords.some((kw) => t.includes(kw));
  if (!found) return null;

  const num = t.match(/(\d+(\.\d+)?)/);
  const d = num ? parseFloat(num[1]) : null;
  return { isRound: true, diameter: d };
}

/* ---------- 타입 인식 ---------- */

function detectType(text) {
  const t = normalize(text);
  for (const rule of typeRules) {
    if (rule.keywords.some((kw) => t.includes(kw))) {
      return rule.labelKo;
    }
  }
  return defaultTypeLabel;
}

/* ---------- 크기 파싱 ---------- */

function parseSize(text) {
  const t = normalize(text);
  const hasDia = roundConfig.keywords.some((kw) => t.includes(kw));
  if (hasDia) return { shape: "round" };

  const re = /(\d+(\.\d+)?)\s*[x×*]\s*(\d+(\.\d+)?)/g;
  let match;
  let best = null;
  let bestArea = 0;

  while ((match = re.exec(t)) !== null) {
    const w = parseFloat(match[1]);
    const h = parseFloat(match[3]);
    const area = w * h;
    if (area > bestArea) {
      bestArea = area;
      best = { w, h };
    }
  }

  if (!best) return null;
  return { shape: "rect", w: best.w, h: best.h };
}

/* ---------- 두께 파싱 ---------- */

function detectThickness(text, hoNumber) {
  if (hoNumber == null) return null;
  if (!Array.isArray(thicknessMap) || thicknessMap.length === 0) return null;

  const t = normalize(text);
  const re = /(\d+(\.\d+)?)\s*[x×*]\s*(\d+(\.\d+)?)/g;

  let match;
  let smallest = null;
  let smallestArea = Infinity;

  while ((match = re.exec(t)) !== null) {
    const w = parseFloat(match[1]);
    const h = parseFloat(match[3]);
    const area = w * h;
    if (area < smallestArea) {
      smallestArea = area;
      smallest = { w, h };
    }
  }
  if (!smallest) return null;

  let best = null;
  let bestScore = Infinity;

  for (const rule of thicknessMap) {
    if (!rule.nos || !rule.nos.includes(hoNumber)) continue;

    const score =
      Math.abs(rule.w - smallest.w) + Math.abs(rule.h - smallest.h);
    if (score < bestScore && score <= thicknessTolerance) {
      bestScore = score;
      best = rule;
    }
  }
  return best ? best.labelKo : null;
}

/* ---------- 호수 매칭 ---------- */

function findSizeCode(w, h) {
  let best = null;
  let bestScore = 999;

  for (const s of sizeMap) {
    const score1 = Math.abs(s.w - w) + Math.abs(s.h - h);
    const score2 = Math.abs(s.w - h) + Math.abs(s.h - w);
    const score = Math.min(score1, score2);

    const dw = score === score1 ? Math.abs(s.w - w) : Math.abs(s.w - h);
    const dh = score === score1 ? Math.abs(s.h - h) : Math.abs(s.h - w);

    if (dw <= sizeTolerance && dh <= sizeTolerance && score < bestScore) {
      best = s;
      bestScore = score;
    }
  }
  return best;
}

/* ---------- 한 줄 분류 (텍스트 기반) ---------- */

function classifyText(fullText) {
  const text = (fullText || "").toString();
  const lower = normalize(text);

  if (!/canvas|panel/.test(lower)) {
    return "";
  }

  const baseType = detectType(text);

  const tri = detectTriangle(text);
  if (tri !== null) {
    return baseType === "캔버스보드"
      ? `캔버스보드 삼각형 ${tri}`
      : `삼각형 ${tri}`;
  }

  const round = detectRound(text);
  if (round) {
    return baseType === "캔버스보드"
      ? `캔버스보드 원형 ${round.diameter || ""}`
      : `원형캔버스 지름 ${round.diameter || ""}`;
  }

  const parsed = parseSize(text);
  let code = "";
  let sizeLabel = "";
  let hoNumber = null;

  if (parsed && parsed.shape === "rect") {
    const matched = findSizeCode(parsed.w, parsed.h);
    if (matched) {
      code = matched.code;
      const num = parseInt(matched.code.replace(/[^0-9]/g, ""), 10);
      if (!Number.isNaN(num)) hoNumber = num;
    } else {
      sizeLabel = `${parsed.w}x${parsed.h}`;
    }
  }

  let thicknessType = null;
  const isExcluded = excludedForThickness.some((x) => lower.includes(x));
  if (!isExcluded && hoNumber !== null) {
    thicknessType = detectThickness(text, hoNumber);
  }

  const finalType = thicknessType || baseType;

  if (code) return `${finalType} ${code}`.trim();
  if (sizeLabel) return `${finalType} ${sizeLabel}`.trim();
  return finalType;
}

/* ---------- 헤더 행/컬럼 위치 찾기 ---------- */

function normCell(v) {
  return v.toString().trim().toUpperCase().replace(/\s+/g, "");
}

function findHeaderInfo(rows) {
  let headerRowIndex = -1;
  let itemCol = -1;
  let descCol = -1;
  let qtyCol = -1;

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i] || [];
    for (let j = 0; j < row.length; j++) {
      const n = normCell(row[j]);
      if (!n) continue;

      if (n.includes("ITEMNO")) itemCol = j; // "Our Item No." 포함
      if (n.startsWith("DESC") || n.startsWith("DESCRIPT")) descCol = j;
      if (n.startsWith("QTY") || n.startsWith("QUANTITY")) qtyCol = j;
    }

    if (itemCol !== -1 || descCol !== -1 || qtyCol !== -1) {
      headerRowIndex = i;
      break;
    }
  }

  return { headerRowIndex, itemCol, descCol, qtyCol };
}

/* ---------- 컬럼 폭 자동 + ITEM NO 고정 ---------- */

function autoFitCols(header, rows, fixedItemNoWidth = 15) {
  const colCount = header.length;
  const maxLen = new Array(colCount).fill(0);

  header.forEach((h, idx) => {
    maxLen[idx] = h.toString().length;
  });

  rows.forEach((r) => {
    header.forEach((h, idx) => {
      const v = r[h] == null ? "" : r[h].toString();
      if (v.length > maxLen[idx]) maxLen[idx] = v.length;
    });
  });

  const cols = maxLen.map((len) => ({ wch: len + 2 }));

  // ITEM NO 또는 Our Item No. 폭 고정
  let itemNoIndex = header.indexOf("ITEM NO");
  if (itemNoIndex === -1) {
    itemNoIndex = header.indexOf("Our Item No.");
  }
  if (itemNoIndex !== -1) {
    cols[itemNoIndex] = { wch: fixedItemNoWidth };
  }

  return cols;
}

/* ---------- 업로드 API ---------- */

app.post("/upload", upload.single("file"), (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: "no file" });

    const originalName = req.file.originalname || "result";
    const baseName = originalName.replace(/\.[^/.]+$/, "");
    const downloadName = `${baseName}_분류결과.xlsx`;

    const wb = xlsx.read(req.file.buffer, { type: "buffer" });
    const outWB = xlsx.utils.book_new();

    wb.SheetNames.forEach((sheetName) => {
      const ws = wb.Sheets[sheetName];
      if (!ws) return;

      const rows = xlsx.utils.sheet_to_json(ws, { header: 1, defval: "" });
      if (!rows.length) return;

      const { headerRowIndex, itemCol, descCol, qtyCol } = findHeaderInfo(rows);

      if (headerRowIndex === -1) return;

      const filteredRows = [];

      // ----- 본 데이터 행 생성 -----
      for (let i = headerRowIndex + 1; i < rows.length; i++) {
        const rowArr = rows[i] || [];

        const itemNo =
          itemCol >= 0 ? (rowArr[itemCol] || "") : "";
        const desc =
          descCol >= 0 ? (rowArr[descCol] || "") : "";
        const qty =
          qtyCol >= 0 ? (rowArr[qtyCol] || "") : "";

        let result = "";

        // ITEM NO 또는 QTY 있을 때만 분류
        if (itemNo || qty) {
          if (itemNo && manualOverrides && manualOverrides[itemNo]) {
            const ov = manualOverrides[itemNo];
            result = `${ov.kind || defaultTypeLabel} ${ov.code || ""}`.trim();
          } else {
            const text = rowArr.join(" ");
            result = classifyText(text);
          }
        } else {
          result = "";
        }

        if (!result && !itemNo && !desc && !qty) continue;

        filteredRows.push({
          RESULT: result,
          "ITEM NO": itemNo,
          DESCRIPTION: desc,
          QUANTITY: qty,
        });
      }

      // ----- 결과 시트 (RESULT / ITEM NO / DESCRIPTION / QUANTITY) -----
      const header = ["RESULT", "ITEM NO", "DESCRIPTION", "QUANTITY"];
      const outWS = xlsx.utils.json_to_sheet(filteredRows, { header });
      outWS["!cols"] = autoFitCols(header, filteredRows, 15);

      xlsx.utils.book_append_sheet(
        outWB,
        outWS,
        `${sheetName}_result`.slice(0, 31)
      );

      // ----- 요약 시트 (Serial / Our Item No. / Description / Quantity) -----
      const summaryRows = filteredRows.map((r, idx) => ({
        Serial: idx + 1,
        "Our Item No.": r["ITEM NO"],
        Description: r.DESCRIPTION,
        Quantity: r.QUANTITY,
      }));

      const summaryHeader = [
        "Serial",
        "Our Item No.",
        "Description",
        "Quantity",
      ];
      const summaryWS = xlsx.utils.json_to_sheet(summaryRows, {
        header: summaryHeader,
      });
      summaryWS["!cols"] = autoFitCols(summaryHeader, summaryRows, 15);

      xlsx.utils.book_append_sheet(
        outWB,
        summaryWS,
        `${sheetName}_summary`.slice(0, 31)
      );
    });

    const buf = xlsx.write(outWB, { type: "buffer", bookType: "xlsx" });

    res.setHeader(
      "Content-Disposition",
      `attachment; filename*=UTF-8''${encodeURIComponent(downloadName)}`
    );
    res.send(buf);
  } catch (e) {
    console.error("UPLOAD ERROR:", e);
    res.status(500).json({ error: e.message });
  }
});

app.listen(PORT, () => {
  console.log(
    `RUNNING http://localhost:${PORT} (result + summary, ITEMNO=15, auto-cols)`
  );
});
