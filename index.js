// index.js — 캔버스 자동 분류 서버 (스타일 포함 exceljs 버전)
// - 입력 엑셀 읽기: xlsx
// - 출력 엑셀 생성/스타일: exceljs
// - 멀티 시트 처리
// - 상품명 / 아이템 코드 / 설명 / 수량 / 박스수(CTN) 출력
// - 아이템 코드 컬럼 폭 15 고정
// - Serial / Our Item No. / Description / Quantity / Boxes 요약 시트 생성
// - ITEM NO, QUANTITY, CTN(박스수) 모두 비면(Grimso 등) RESULT 비움
// - 스타일: 헤더 볼드+가운데, 전체 테두리, 기본 정렬

const express = require("express");
const multer = require("multer");
const cors = require("cors");
const xlsx = require("xlsx");
const ExcelJS = require("exceljs");
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
const PORT = process.env.PORT || 3100;

app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, "public")));

const upload = multer({ storage: multer.memoryStorage() });

/* ---------- 공통 유틸 ---------- */

// ✅ 숫자 오타 보정 (예: 3..8 -> 3.8)
function sanitizeNumericTypos(text) {
  return (text || "")
    .toString()
    // "3..8" 같은 케이스를 "3.8"로 교정
    .replace(/(\d)\.\.(\d)/g, "$1.$2")
    // 혹시 "3...8" 처럼 더 많은 점이 들어오면 1개로 축소
    .replace(/(\d)\.{2,}(\d)/g, "$1.$2");
}

function normalize(text) {
  // ✅ 여기서 먼저 오타를 고치고 나서 소문자화
  const fixed = sanitizeNumericTypos(text);
  return fixed.toString().trim().toLowerCase();
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
    if (rule.keywords.some((kw) => t.includes(kw))) return rule.labelKo;
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
    const score = Math.abs(rule.w - smallest.w) + Math.abs(rule.h - smallest.h);
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

/* ---------- 한 줄 분류 ---------- */

function classifyText(fullText) {
  const text = (fullText || "").toString();
  const lower = normalize(text);

  if (!/canvas|panel/.test(lower)) return "";

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

/* ---------- 헤더 탐색 ---------- */

function normCell(v) {
  return (v ?? "").toString().trim().toUpperCase().replace(/\s+/g, "");
}

function findHeaderInfo(rows) {
  let headerRowIndex = -1;
  let itemCol = -1;
  let descCol = -1;
  let qtyCol = -1;
  let ctnCol = -1;

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i] || [];
    for (let j = 0; j < row.length; j++) {
      const n = normCell(row[j]);
      if (!n) continue;

      if (n.includes("ITEMNO")) itemCol = j;
      if (n.startsWith("DESC") || n.startsWith("DESCRIPT")) descCol = j;
      if (n.startsWith("QTY") || n.startsWith("QUANTITY")) qtyCol = j;

      if (
        ctnCol === -1 &&
        (n === "CTN" ||
          n.includes("CTN(CTN)") ||
          n.startsWith("CTN") ||
          n.includes("CTNS") ||
          n.includes("CARTON"))
      ) {
        ctnCol = j;
      }
    }

    if (itemCol !== -1 || descCol !== -1 || qtyCol !== -1 || ctnCol !== -1) {
      headerRowIndex = i;
      break;
    }
  }

  return { headerRowIndex, itemCol, descCol, qtyCol, ctnCol };
}

/* ---------- ExcelJS 스타일 유틸 ---------- */

const THIN_BORDER = {
  top: { style: "thin" },
  left: { style: "thin" },
  bottom: { style: "thin" },
  right: { style: "thin" },
};

function setWorksheetColumns(ws, columns) {
  ws.columns = columns;

  const headerRow = ws.getRow(1);
  headerRow.font = { bold: true };
  headerRow.alignment = { vertical: "middle", horizontal: "center" };
  headerRow.height = 18;

  headerRow.eachCell((cell) => {
    cell.border = THIN_BORDER;
  });
}

function applyTableBorders(ws, lastRowNumber, lastColNumber) {
  for (let r = 2; r <= lastRowNumber; r++) {
    const row = ws.getRow(r);
    row.height = 16;
    for (let c = 1; c <= lastColNumber; c++) {
      const cell = row.getCell(c);
      cell.border = THIN_BORDER;
      cell.alignment = {
        vertical: "middle",
        horizontal: "left",
        wrapText: true,
      };
    }
  }
}

function autoFitColumns(ws, fixedMap = {}) {
  ws.columns.forEach((col) => {
    const key = col.key || col.header;
    if (fixedMap[key] != null) {
      col.width = fixedMap[key];
      return;
    }

    let max = col.header ? col.header.toString().length : 10;
    col.eachCell({ includeEmpty: false }, (cell) => {
      const v = cell.value == null ? "" : cell.value.toString();
      if (v.length > max) max = v.length;
    });

    col.width = Math.min(max + 2, 60);
  });
}

/* ---------- 업로드 API ---------- */

app.post("/upload", upload.single("file"), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: "no file" });

    const originalNameRaw = req.file.originalname || "result";

    // multer가 filename을 latin1로 줄 때가 있어 UTF-8로 복구
    let originalName = originalNameRaw;
    try {
      originalName = Buffer.from(originalNameRaw, "latin1").toString("utf8");
    } catch (e) {
      originalName = originalNameRaw;
    }

    const baseName = originalName.replace(/\.[^/.]+$/, "");
    const downloadName = `${baseName}_분류결과.xlsx`;

    const inWB = xlsx.read(req.file.buffer, { type: "buffer" });
    const outWB = new ExcelJS.Workbook();
    outWB.creator = "canvas-classifier";
    outWB.created = new Date();

    for (const sheetName of inWB.SheetNames) {
      const ws = inWB.Sheets[sheetName];
      if (!ws) continue;

      const rows = xlsx.utils.sheet_to_json(ws, { header: 1, defval: "" });
      if (!rows.length) continue;

      const { headerRowIndex, itemCol, descCol, qtyCol, ctnCol } =
        findHeaderInfo(rows);
      if (headerRowIndex === -1) continue;

      const filteredRows = [];

      for (let i = headerRowIndex + 1; i < rows.length; i++) {
        const rowArr = rows[i] || [];

        const itemNo = itemCol >= 0 ? rowArr[itemCol] || "" : "";
        const desc = descCol >= 0 ? rowArr[descCol] || "" : "";
        const qty = qtyCol >= 0 ? rowArr[qtyCol] || "" : "";
        const ctn = ctnCol >= 0 ? rowArr[ctnCol] || "" : "";

        let result = "";

        if (itemNo || qty || ctn) {
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

        if (!result && !itemNo && !desc && !qty && !ctn) continue;

        filteredRows.push({
          PRODUCT_NAME: result, // 상품명(=기존 RESULT)
          ITEM_CODE: itemNo, // 아이템 코드(=기존 ITEM NO)
          DESC_KO: desc, // 설명(=기존 DESCRIPTION)
          QTY_KO: qty, // 수량(=기존 QUANTITY)
          BOXES: ctn, // 박스수
        });
      }

      /* -------- result 시트 -------- */
      const resultSheetName = `${sheetName}_result`.slice(0, 31);
      const outWS = outWB.addWorksheet(resultSheetName);

      // ✅ 수량/박스수 폭 동일(12)
      setWorksheetColumns(outWS, [
        { header: "상품명", key: "PRODUCT_NAME", width: 20 },
        { header: "아이템 코드", key: "ITEM_CODE", width: 15 }, // 고정 15
        { header: "설명", key: "DESC_KO", width: 40 },
        { header: "수량", key: "QTY_KO", width: 12 },
        { header: "박스수", key: "BOXES", width: 12 }, // 수량과 동일
      ]);

      filteredRows.forEach((r) => outWS.addRow(r));

      const lastRow = outWS.rowCount;
      const lastCol = outWS.columnCount;

      applyTableBorders(outWS, lastRow, lastCol);

      // 수량/박스수 가운데 정렬
      outWS.getColumn("QTY_KO").alignment = {
        vertical: "middle",
        horizontal: "center",
      };
      outWS.getColumn("BOXES").alignment = {
        vertical: "middle",
        horizontal: "center",
      };

      // 컬럼 자동폭(아이템 코드는 15 고정)
      autoFitColumns(outWS, { ITEM_CODE: 15, QTY_KO: 12, BOXES: 12 });

      /* -------- summary 시트 -------- */
      const summarySheetName = `${sheetName}_summary`.slice(0, 31);
      const sumWS = outWB.addWorksheet(summarySheetName);

      setWorksheetColumns(sumWS, [
        { header: "Serial", key: "SERIAL", width: 8 },
        { header: "Our Item No.", key: "OUR_ITEM_NO", width: 15 },
        { header: "Description", key: "DESC", width: 40 },
        { header: "Quantity", key: "QTY", width: 12 },
        { header: "Boxes", key: "BOXES", width: 12 }, // Quantity와 동일
      ]);

      filteredRows.forEach((r, idx) => {
        sumWS.addRow({
          SERIAL: idx + 1,
          OUR_ITEM_NO: r.ITEM_CODE,
          DESC: r.DESC_KO,
          QTY: r.QTY_KO,
          BOXES: r.BOXES,
        });
      });

      const sLastRow = sumWS.rowCount;
      const sLastCol = sumWS.columnCount;

      applyTableBorders(sumWS, sLastRow, sLastCol);

      sumWS.getColumn("SERIAL").alignment = {
        vertical: "middle",
        horizontal: "center",
      };
      sumWS.getColumn("QTY").alignment = {
        vertical: "middle",
        horizontal: "center",
      };
      sumWS.getColumn("BOXES").alignment = {
        vertical: "middle",
        horizontal: "center",
      };

      autoFitColumns(sumWS, { OUR_ITEM_NO: 15, QTY: 12, BOXES: 12 });
    }

    const buf = await outWB.xlsx.writeBuffer();

    res.setHeader(
      "Content-Disposition",
      `attachment; filename*=UTF-8''${encodeURIComponent(downloadName)}`
    );
    res.send(Buffer.from(buf));
  } catch (e) {
    console.error("UPLOAD ERROR:", e);
    res.status(500).json({ error: e.message });
  }
});

app.listen(PORT, () => {
  console.log(`RUNNING http://0.0.0.0:${PORT}`);
});
