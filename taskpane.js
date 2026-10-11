/******/ (() => { // webpackBootstrap
/******/ 	"use strict";
/******/ 	var __webpack_modules__ = ({

/***/ "./src/taskpane/pdf-sign.mjs"
/*!***********************************!*\
  !*** ./src/taskpane/pdf-sign.mjs ***!
  \***********************************/
(__unused_webpack___webpack_module__, __webpack_exports__, __webpack_require__) {

__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   SIGN_FIELD_SIZE: () => (/* binding */ SIGN_FIELD_SIZE),
/* harmony export */   SIGN_NAME_RE: () => (/* binding */ SIGN_NAME_RE),
/* harmony export */   addSignatureFields: () => (/* binding */ addSignatureFields),
/* harmony export */   findMarkers: () => (/* binding */ findMarkers),
/* harmony export */   signMarker: () => (/* binding */ signMarker)
/* harmony export */ });
// Chèn các trường chữ ký số (PDF signature field, /FT /Sig) vào file PDF tại vị trí các chuỗi đánh
// dấu "#KY#<tên ô>#" mà add-in đã đặt sẵn trong văn bản Word.
// Hàm thuần — thư viện pdf.js (pdfjsLib) và pdf-lib (PDFLib) được truyền vào, nên chạy được cả trong
// add-in (tải từ CDN khi cần) lẫn khi test bằng Node.

// Tên ô ký: chữ không dấu, số, dấu gạch dưới (VD ct_kt, ct_tckt) — dùng làm tên trường trong PDF.
const SIGN_NAME_RE = /^[A-Za-z0-9_]{1,40}$/;

function signMarker(name) {
  return `#KY#${name}#`;
}

const MARKER_RE = /#KY#([A-Za-z0-9_]{1,40})#/g;

// Kích thước mỗi ô ký: ~4,9 x 2cm (đủ chỗ cho hình chữ ký + thông tin ký số).
const SIGN_FIELD_SIZE = { width: 140, height: 56 };

// Tìm mọi chuỗi đánh dấu trong PDF. Trả về [{ name, pageIndex, x, y }] theo toạ độ PDF (gốc ở góc
// dưới trái trang, đơn vị point).
async function findMarkers(pdfjsLib, bytes) {
  // pdf.js có thể chiếm (detach) buffer truyền vào nên đưa bản sao.
  const doc = await pdfjsLib.getDocument({ data: bytes.slice(), isEvalSupported: false }).promise;
  const found = [];
  try {
    for (let p = 1; p <= doc.numPages; p++) {
      const page = await doc.getPage(p);
      const { items } = await page.getTextContent();

      // Ghép chữ cả trang để tìm được cả khi chuỗi đánh dấu bị tách thành nhiều đoạn (item).
      let text = "";
      const starts = items.map((item) => {
        const start = text.length;
        text += item.str || "";
        return start;
      });

      for (const match of text.matchAll(MARKER_RE)) {
        const at = match.index;
        const i = starts.findIndex((start, k) => at >= start && at < start + (items[k].str || "").length);
        const item = items[i];
        const charWidth = item.str.length ? item.width / item.str.length : 0;
        found.push({
          name: match[1],
          pageIndex: p - 1,
          x: item.transform[4] + charWidth * (at - starts[i]),
          y: item.transform[5]
        });
      }
    }
    return found;
  } finally {
    await doc.destroy();
  }
}

// Thêm các trường ký (chưa ký). Mỗi trường căn giữa theo chiều ngang tại điểm đánh dấu (vị trí con trỏ)
// và chạy từ dòng đó xuống dưới — nằm gọn trong khoảng trống giữa chức vụ và họ tên người ký, không
// đè lên dòng chữ phía trên. Trả về bytes PDF mới.
// Đỉnh ô cao hơn chân dòng chữ (baseline) của điểm đánh dấu một khoảng ~ chiều cao chữ cỡ 13-14.
const LINE_TOP_OFFSET = 12;

async function addSignatureFields(PDFLib, bytes, positions, size = SIGN_FIELD_SIZE) {
  const { PDFDocument, PDFString } = PDFLib;
  const pdf = await PDFDocument.load(bytes);
  const ctx = pdf.context;
  const form = pdf.catalog.getOrCreateAcroForm();
  const { width: w, height: h } = size;

  positions.forEach((pos) => {
    const page = pdf.getPage(pos.pageIndex);
    const { width: pageW, height: pageH } = page.getSize();
    const left = Math.min(Math.max(pos.x - w / 2, 0), pageW - w);
    const bottom = Math.min(Math.max(pos.y + LINE_TOP_OFFSET - h, 0), pageH - h);

    const appearance = ctx.register(
      ctx.stream("", { Type: "XObject", Subtype: "Form", BBox: [0, 0, w, h], Resources: {} })
    );
    const widget = ctx.register(
      ctx.obj({
        Type: "Annot",
        Subtype: "Widget",
        FT: "Sig",
        T: PDFString.of(pos.name),
        F: 4, // Print
        Rect: [left, bottom, left + w, bottom + h],
        P: page.ref,
        AP: { N: appearance }
      })
    );
    page.node.addAnnot(widget);
    form.addField(widget);
  });

  return pdf.save();
}


/***/ },

/***/ "./src/taskpane/vn-text.mjs"
/*!**********************************!*\
  !*** ./src/taskpane/vn-text.mjs ***!
  \**********************************/
(__unused_webpack___webpack_module__, __webpack_exports__, __webpack_require__) {

__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   numberToVietnameseWords: () => (/* binding */ numberToVietnameseWords)
/* harmony export */ });
// Đổi số thành chữ tiếng Việt — hàm thuần (không dùng Word API), tách riêng để chạy unit test
// bằng `npm test` mà không cần mở Word.
const VN_ONES = ["không", "một", "hai", "ba", "bốn", "năm", "sáu", "bảy", "tám", "chín"];
// Trong 1 khối 9 chữ số chỉ có đơn vị / nghìn / triệu; từ hàng tỷ trở lên đọc đệ quy
// ("một nghìn năm trăm tỷ", "một tỷ tỷ") thay vì dùng các tên ghép như "nghìn tỷ".
const VN_SCALES = ["", "nghìn", "triệu"];

function threeDigitGroupToWords(n, isLeadingGroup) {
  const hundreds = Math.floor(n / 100);
  const tens = Math.floor((n % 100) / 10);
  const ones = n % 10;
  const parts = [];

  if (hundreds > 0) {
    parts.push(VN_ONES[hundreds], "trăm");
  } else if (!isLeadingGroup) {
    parts.push("không trăm");
  }

  if (tens === 0) {
    if (ones > 0) {
      parts.push(hundreds > 0 || !isLeadingGroup ? "linh" : null, VN_ONES[ones]);
    }
  } else if (tens === 1) {
    parts.push("mười");
    if (ones === 5) {
      parts.push("lăm");
    } else if (ones > 0) {
      parts.push(VN_ONES[ones]);
    }
  } else {
    parts.push(VN_ONES[tens], "mươi");
    if (ones === 1) {
      parts.push("mốt");
    } else if (ones === 5) {
      parts.push("lăm");
    } else if (ones > 0) {
      parts.push(VN_ONES[ones]);
    }
  }

  return parts.filter(Boolean).join(" ");
}

// Đọc khối tối đa 9 chữ số (dưới 1 tỷ). Trả về "" nếu toàn số 0.
function readUnderBillion(digits, isLeading) {
  const groups = [];
  let rest = digits;
  while (rest.length > 0) {
    groups.unshift(rest.slice(-3));
    rest = rest.slice(0, -3);
  }

  const words = [];
  groups.forEach((group, i) => {
    const value = parseInt(group, 10);
    if (value === 0) {
      return;
    }
    const scale = VN_SCALES[groups.length - 1 - i];
    const groupWords = threeDigitGroupToWords(value, isLeading && i === 0);
    words.push(scale ? `${groupWords} ${scale}` : groupWords);
  });

  return words.join(" ");
}

// Đọc chuỗi chữ số nguyên dương (không giới hạn độ dài): tách 9 chữ số cuối, phần phía trước
// đọc đệ quy rồi thêm "tỷ".
function readInteger(digits, isLeading = true) {
  if (/^0*$/.test(digits)) {
    return isLeading ? "không" : "";
  }
  if (digits.length <= 9) {
    return readUnderBillion(digits, isLeading);
  }

  const high = readInteger(digits.slice(0, -9), isLeading);
  const low = readUnderBillion(digits.slice(-9), false);
  return low ? `${high} tỷ ${low}` : `${high} tỷ`;
}

// Phần thập phân: các số 0 ở đầu đọc riêng từng chữ "không", phần còn lại đọc như số nguyên
// (VD: 2,05 -> "hai phẩy không năm"; 2,25 -> "hai phẩy hai mươi lăm").
function readDecimal(digits) {
  const leadingZeros = digits.match(/^0*/)[0].length;
  const words = Array(leadingZeros).fill("không");
  const rest = digits.slice(leadingZeros);
  if (rest) {
    words.push(readInteger(rest));
  }
  return words.join(" ");
}

// Chuyển chuỗi số thành chữ tiếng Việt theo quy ước Việt Nam: dấu chấm phân cách hàng nghìn,
// dấu phẩy phân cách thập phân (VD: "5.000.000", "2,5", "-1.234,05").
function numberToVietnameseWords(rawInput) {
  const trimmed = String(rawInput).trim().replace(/\s/g, "");
  if (!trimmed) {
    throw new Error("Vui lòng nhập một số.");
  }

  const isNegative = trimmed.startsWith("-");
  const unsigned = trimmed.replace(/^-/, "");

  if (!/^\d{1,3}(\.\d{3})*(,\d+)?$/.test(unsigned) && !/^\d+(,\d+)?$/.test(unsigned)) {
    throw new Error(
      "Số không hợp lệ. Dùng dấu chấm phân cách hàng nghìn và dấu phẩy cho phần thập phân (VD: 1.234.567,5)."
    );
  }

  const [intPart, decPart] = unsigned.split(",");
  let words = readInteger(intPart.replace(/\./g, ""));
  if (decPart) {
    words += ` phẩy ${readDecimal(decPart)}`;
  }

  if (isNegative && /[1-9]/.test(unsigned)) {
    return `Âm ${words}`;
  }
  return words.charAt(0).toUpperCase() + words.slice(1);
}


/***/ }

/******/ 	});
/************************************************************************/
/******/ 	// The module cache
/******/ 	const __webpack_module_cache__ = {};
/******/ 	
/******/ 	// The require function
/******/ 	function __webpack_require__(moduleId) {
/******/ 		// Check if module is in cache
/******/ 		const cachedModule = __webpack_module_cache__[moduleId];
/******/ 		if (cachedModule !== undefined) {
/******/ 			return cachedModule.exports;
/******/ 		}
/******/ 		// Create a new module (and put it into the cache)
/******/ 		const module = __webpack_module_cache__[moduleId] = {
/******/ 			// no module.id needed
/******/ 			// no module.loaded needed
/******/ 			exports: {}
/******/ 		};
/******/ 	
/******/ 		// Execute the module function
/******/ 		if (!(moduleId in __webpack_modules__)) {
/******/ 			delete __webpack_module_cache__[moduleId];
/******/ 			const e = new Error("Cannot find module '" + moduleId + "'");
/******/ 			e.code = 'MODULE_NOT_FOUND';
/******/ 			throw e;
/******/ 		}
/******/ 		__webpack_modules__[moduleId](module, module.exports, __webpack_require__);
/******/ 	
/******/ 		// Return the exports of the module
/******/ 		return module.exports;
/******/ 	}
/******/ 	
/************************************************************************/
/******/ 	/* webpack/runtime/define property getters */
/******/ 	// define getter/value functions for harmony exports
/******/ 	__webpack_require__.d = (exports, definition) => {
/******/ 		for(var key in definition) {
/******/ 			if(__webpack_require__.o(definition, key) && !__webpack_require__.o(exports, key)) {
/******/ 				Object.defineProperty(exports, key, { enumerable: true, get: definition[key] });
/******/ 			}
/******/ 		}
/******/ 	};
/******/ 	
/******/ 	/* webpack/runtime/hasOwnProperty shorthand */
/******/ 	__webpack_require__.o = (obj, prop) => (Object.prototype.hasOwnProperty.call(obj, prop));
/******/ 	
/******/ 	/* webpack/runtime/make namespace object */
/******/ 	// define __esModule on exports
/******/ 	__webpack_require__.r = (exports) => {
/******/ 		Object.defineProperty(exports, Symbol.toStringTag, { value: 'Module' });
/******/ 		Object.defineProperty(exports, '__esModule', { value: true });
/******/ 	};
/******/ 	
/************************************************************************/
let __webpack_exports__ = {};
// This entry needs to be wrapped in an IIFE because it needs to be isolated against other modules in the chunk.
(() => {
/*!**********************************!*\
  !*** ./src/taskpane/taskpane.js ***!
  \**********************************/
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _vn_text_mjs__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./vn-text.mjs */ "./src/taskpane/vn-text.mjs");
/* harmony import */ var _pdf_sign_mjs__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./pdf-sign.mjs */ "./src/taskpane/pdf-sign.mjs");
/* global Office, Word, document, console, window, fetch, Blob, URL, setTimeout */



Office.onReady((info) => {
  if (info.host !== Office.HostType.Word) {
    setStatus("Add-in này chỉ chạy trong Microsoft Word.", "error");
    document.querySelectorAll("button").forEach((b) => (b.disabled = true));
    return;
  }

  bind("btn-quick-normalize", quickNormalize, formatQuickNormalizeMessage);
  bind("btn-page-setup", applyPageSetup, "Đã chuẩn hoá lề trang: Trên 20mm • Dưới 20mm • Trái 30mm • Phải 15mm.");
  bind("btn-unify-font", unifyFont, "Đã đồng bộ toàn bộ font chữ trong văn bản thành Times New Roman.");
  bind("btn-page-number", insertPageNumbers, "Đã đánh số trang: căn giữa phía trên, cỡ 14, không hiện số ở trang đầu.");
  bind("btn-place-sign", placeSignField, (name) => `Đã đặt ô ký "${name}" tại vị trí con trỏ.`);
  bind("btn-save-pdf-sign", saveAsPdfWithSignatures, formatSignedPdfMessage);
  document.getElementById("select-sign-name").onchange = toggleCustomSignName;
  renderSignNameOptions([]);
  refreshSignList();
  bind("btn-save-pdf", saveAsPdf, (name) => `Đã xuất PDF: ${name} (xem trong thư mục Downloads).`);
  bind("btn-insert-landscape", insertLandscapePage, "Đã chèn thêm 1 trang ngang ở cuối văn bản.");
  bind("btn-insert-portrait", insertPortraitPage, "Đã chèn thêm 1 trang dọc ở cuối văn bản.");
  bind(
    "btn-insert-quochieu",
    insertQuocHieu,
    "Đã chèn khối Quốc hiệu – Tiêu ngữ ở đầu văn bản. Hãy kiểm tra lại số hiệu văn bản (Số: .../ĐLTM-KHKT)."
  );
  bind(
    "btn-insert-noinhan",
    insertNoiNhanChuKy,
    "Đã chèn khối \"Nơi nhận\" và \"Chữ ký\" ở cuối văn bản. Hãy sửa lại họ tên người ký cho đúng."
  );
  bind("btn-insert-doc-title", insertDocTitle, (t) => `Đã chèn tiêu đề "${t}" tại vị trí con trỏ.`);
  bind("btn-repeat-header", repeatTableHeaders, (n) => `Đã đặt lặp lại tiêu đề (hàng 1) cho ${n} bảng.`);
  bind("btn-center-table", centerTableContent, (n) => `Đã căn giữa nội dung cho ${n} bảng.`);
  bind("btn-first-line-indent", applyFirstLineIndent, (n) => `Đã thụt lề đầu dòng (1cm) cho ${n} đoạn đang chọn.`);
  bind(
    "btn-line-spacing",
    normalizeLineSpacing,
    (n) => `Đã chuẩn hoá giãn dòng (Multiple 1,17 ≈ 14pt; cách đoạn trên 3pt, dưới 6pt) cho ${n} đoạn đang chọn.`
  );
  bind("btn-cleanup", cleanupExtraSpaces, (n) => `Đã xoá ${n} chỗ dấu cách thừa.`);
  bind("btn-justify-selection", justifySelection, (n) => `Đã căn đều hai bên ${n} đoạn đang chọn.`);
  bind("btn-justify-all", justifyAll, (n) => `Đã căn đều hai bên ${n} đoạn trong toàn văn bản.`);
  bind("btn-number-to-words", insertNumberAsWords, (w) => `Đã chèn: "${w}".`);
});

function bind(buttonId, action, message) {
  const el = document.getElementById(buttonId);
  el.onclick = () => runAction(action, message);
}

// Khoá toàn bộ nút trong lúc đang xử lý để tránh bấm 2 lần (VD: chèn Quốc hiệu 2 lần).
function setButtonsDisabled(disabled) {
  document.querySelectorAll("button").forEach((b) => (b.disabled = disabled));
}

async function runAction(action, message) {
  setStatus("Đang xử lý...", "");
  setButtonsDisabled(true);
  try {
    let result;
    await Word.run(async (context) => {
      result = await action(context);
    });
    const text = typeof message === "function" ? message(result) : message;
    setStatus(text, "success");
  } catch (error) {
    console.error(error);
    const detail = error && error.message ? error.message : String(error);
    setStatus("Lỗi: " + detail, "error");
  } finally {
    setButtonsDisabled(false);
    refreshSignList();
  }
}

function setStatus(text, cssClass) {
  const el = document.getElementById("status");
  el.textContent = text;
  el.className = cssClass || "";
}

function mmToPoints(mm) {
  return (mm / 25.4) * 72;
}

// Lề trang chuẩn (mm) và khổ A4 — dùng chung cho "Chuẩn hoá lề trang" và để tính độ rộng
// vùng chữ khi chèn bảng Quốc hiệu.
const PAGE_MARGINS_MM = { top: 20, bottom: 20, left: 30, right: 15 };
const A4_WIDTH_MM = 210;
const TEXT_WIDTH_MM = A4_WIDTH_MM - PAGE_MARGINS_MM.left - PAGE_MARGINS_MM.right; // 165mm
// Độ rộng 2 cột bảng Quốc hiệu (tổng 175mm = vùng chữ + 10mm lấn sang lề phải).
const QUOC_HIEU_COLUMNS_MM = { left: 70, right: TEXT_WIDTH_MM + 10 - 70 };

function escapeXml(str) {
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function buildRunProps(font) {
  const f = font || {};
  const parts = [];
  if (f.name) {
    parts.push(`<w:rFonts w:ascii="${f.name}" w:hAnsi="${f.name}" w:cs="${f.name}"/>`);
  }
  if (f.bold) {
    parts.push("<w:b/>");
  }
  if (f.italic) {
    parts.push("<w:i/>");
  }
  if (f.underline) {
    parts.push('<w:u w:val="single"/>');
  }
  if (f.size) {
    const halfPoints = Math.round(f.size * 2);
    parts.push(`<w:sz w:val="${halfPoints}"/>`, `<w:szCs w:val="${halfPoints}"/>`);
  }
  return parts.join("");
}

function buildRunXml(text, font) {
  const rPr = buildRunProps(font);
  return `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ""}<w:t xml:space="preserve">${escapeXml(text || "")}</w:t></w:r>`;
}

// Một dòng có thể là { text, font } (1 run) hoặc { runs: [{ text, font }, ...] }
// khi cần nhiều kiểu định dạng khác nhau trong cùng một đoạn (vd. chỉ gạch chân một phần chữ).
// Luôn đặt cứng giãn dòng đơn, không cách đoạn (trừ spaceBefore), không thụt lề — để khối
// trong ô không bị ăn theo style Normal của văn bản (VD Normal có thụt đầu dòng 1cm sẽ đẩy
// chữ căn giữa lệch sang phải, hoặc cách đoạn 6-8pt làm hai cột cao thấp khác nhau).
// Tuỳ chọn: indentMm — thụt đều 2 bên; rule — kẻ đường ngang (viền trên của đoạn), dùng cùng
// indentMm để đường kẻ chỉ dài bằng phần giữa.
function buildParagraphXml(line) {
  const runsData = line.runs || [{ text: line.text, font: line.font }];
  const before = Math.round((line.spaceBefore || 0) * 20); // pt -> twip
  const spacing = `<w:spacing w:before="${before}" w:after="0" w:line="240" w:lineRule="auto"/>`;
  const indent = Math.round(mmToPoints(line.indentMm || 0) * 20);
  const ind = `<w:ind w:left="${indent}" w:right="${indent}" w:firstLine="0"/>`;
  const border = line.rule ? '<w:pBdr><w:top w:val="single" w:sz="6" w:space="1" w:color="auto"/></w:pBdr>' : "";
  const jc = line.alignment === Word.Alignment.centered ? '<w:jc w:val="center"/>' : '<w:jc w:val="left"/>';
  const runs = runsData.map((r) => buildRunXml(r.text, r.font)).join("");
  return `<w:p><w:pPr>${border}${spacing}${ind}${jc}</w:pPr>${runs}</w:p>`;
}

// Dựng gói OOXML tối thiểu (pkg:package) theo đúng mẫu chính thức của Microsoft để chèn
// nhiều paragraph cùng lúc trong MỘT lệnh insertOoxml duy nhất — tránh hoàn toàn việc phải
// tạo từng paragraph rời rạc bằng insertParagraph (cách này đã nhiều lần gây lỗi lệch/mất
// nội dung khi thao tác trên ô bảng).
function buildCellOoxml(lines) {
  return buildPackageXml(lines.map(buildParagraphXml).join(""));
}

function buildPackageXml(body) {
  return (
    '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">' +
    '<pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml" pkg:padding="512">' +
    '<pkg:xmlData><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
    "</Relationships></pkg:xmlData></pkg:part>" +
    '<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">' +
    '<pkg:xmlData><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' +
    body +
    "</w:body></w:document></pkg:xmlData></pkg:part></pkg:package>"
  );
}

// Bỏ viền + tắt định dạng "cột đầu tiên/băng màu" mặc định của table style — nếu không,
// Word có thể tự áp định dạng khác cho cột trái so với cột phải, ghi đè lên căn giữa/định
// dạng mình đặt.
function makeTablePlain(table) {
  table.getBorder("All").type = "None";
  table.styleFirstColumn = false;
  table.styleBandedColumns = false;
  table.styleBandedRows = false;
}

function tnr(size, extra) {
  return { name: "Times New Roman", size, ...extra };
}

// Đặt lại alignment và khoảng cách trực tiếp qua API (không qua OOXML) để đảm bảo chắc chắn
// áp dụng, vì <w:jc>/<w:spacing>/<w:ind> trong OOXML không phải lúc nào cũng được Word áp dụng.
async function applyLineFormats(context, paragraphs, lines) {
  paragraphs.load("items");
  await context.sync();

  paragraphs.items.forEach((p, i) => {
    const line = lines[i];
    if (!line) {
      return;
    }
    if (line.alignment) {
      p.alignment = line.alignment;
    }
    const indent = mmToPoints(line.indentMm || 0);
    p.set({
      spaceBefore: line.spaceBefore || 0,
      spaceAfter: 0,
      lineSpacing: 12, // dòng đơn
      firstLineIndent: 0,
      leftIndent: indent,
      rightIndent: indent
    });
  });

  await context.sync();
}

async function fillCell(context, cell, lines) {
  cell.body.insertOoxml(buildCellOoxml(lines), Word.InsertLocation.replace);
  await context.sync();
  await applyLineFormats(context, cell.body.paragraphs, lines);
}

// 1. Chuẩn hoá page setup: Trên 20 / Dưới 20 / Trái 30 / Phải 15 (mm)
async function applyPageSetup(context) {
  if (!Office.context.requirements.isSetSupported("WordApiDesktop", "1.3")) {
    throw new Error("Chức năng này chỉ dùng được trên Word desktop (Windows/Mac), không hỗ trợ Word trên trình duyệt.");
  }

  const sections = context.document.sections;
  sections.load("items");
  await context.sync();

  sections.items.forEach((section) => {
    section.pageSetup.set({
      topMargin: mmToPoints(PAGE_MARGINS_MM.top),
      bottomMargin: mmToPoints(PAGE_MARGINS_MM.bottom),
      leftMargin: mmToPoints(PAGE_MARGINS_MM.left),
      rightMargin: mmToPoints(PAGE_MARGINS_MM.right)
    });
  });

  await context.sync();
}

// 1c. Chèn thêm 1 trang mới (ngắt sang trang bằng Section Break) ở cuối văn bản, với
// hướng giấy riêng (ngang/dọc) — dùng khi cần 1 trang khác hướng với phần còn lại
// (VD: 1 trang ngang để chứa bảng rộng). Trang mới không ảnh hưởng hướng giấy các trang trước.
async function insertPageWithOrientation(context, orientation) {
  if (!Office.context.requirements.isSetSupported("WordApiDesktop", "1.3")) {
    throw new Error("Chức năng này chỉ dùng được trên Word desktop (Windows/Mac), không hỗ trợ Word trên trình duyệt.");
  }

  context.document.body.insertBreak(Word.BreakType.sectionNext, Word.InsertLocation.end);
  await context.sync();

  const sections = context.document.sections;
  sections.load("items");
  await context.sync();

  sections.items[sections.items.length - 1].pageSetup.orientation = orientation;
  await context.sync();
}

async function insertLandscapePage(context) {
  await insertPageWithOrientation(context, Word.PageOrientation.landscape);
}

async function insertPortraitPage(context) {
  await insertPageWithOrientation(context, Word.PageOrientation.portrait);
}

// 1b. Đồng bộ toàn bộ font chữ trong văn bản (kể cả trong bảng, header, footer) thành Times
// New Roman. Chỉ đổi tên font, giữ nguyên cỡ chữ/đậm/nghiêng/màu đang có.
// Bỏ qua chữ đang dùng font ký hiệu (Symbol, Wingdings...): đổi sang Times New Roman sẽ biến
// ký hiệu (dấu tick, mũi tên, ô vuông...) thành chữ cái vô nghĩa.
const SYMBOL_FONT_RE = /^(symbol|wingdings.*|webdings|marlett|mt extra|zapf ?dingbats)$/i;

async function unifyFont(context) {
  const bodies = [context.document.body];
  const sections = context.document.sections;
  sections.load("items");
  await context.sync();
  sections.items.forEach((section) => {
    ["Primary", "FirstPage", "EvenPages"].forEach((type) => {
      bodies.push(section.getHeader(type), section.getFooter(type));
    });
  });

  const paragraphLists = bodies.map((b) => b.paragraphs);
  paragraphLists.forEach((list) => list.load("items/font/name"));
  await context.sync();
  const paragraphs = paragraphLists.flatMap((list) => list.items);

  // Đoạn chỉ có 1 font: đổi luôn cả đoạn. Đoạn trộn nhiều font (font.name = null): tách theo
  // từ rồi đổi từng từ không phải font ký hiệu.
  const mixed = [];
  paragraphs.forEach((p) => {
    const name = p.font.name;
    if (name === null || name === "") {
      mixed.push(p.getTextRanges([" "], false));
    } else if (!SYMBOL_FONT_RE.test(name)) {
      p.font.name = "Times New Roman";
    }
  });
  mixed.forEach((ranges) => ranges.load("items/font/name"));
  await context.sync();

  mixed.forEach((ranges) => {
    ranges.items.forEach((r) => {
      if (r.font.name && !SYMBOL_FONT_RE.test(r.font.name)) {
        r.font.name = "Times New Roman";
      }
    });
  });
  await context.sync();
}

// 1d. Đánh số trang theo NĐ 30/2020: chữ số Ả Rập, cỡ 14, chữ đứng, căn giữa trong phần lề trên
// (header), KHÔNG hiển thị số ở trang thứ nhất.
// Không dùng tuỳ chọn "Different First Page" (Word JS chưa có API ổn định để bật) mà dùng mã
// trường { IF { = { SECTION } + { PAGE } } > 2 "{ PAGE }" "" }: chỉ trang 1 của section 1 có
// SECTION + PAGE = 2 nên bị ẩn; mọi trang khác đều hiện số — kể cả trang đầu của section mới
// tạo bằng ngắt "Next Page" có cài đánh số lại từ 1 (điều kiện PAGE > 1 sẽ ẩn nhầm trang đó).
// Ghi đè nội dung header hiện có.

// Một trường Word. Mỗi phần của mã trường là chuỗi (chữ thường) hoặc { xml } (trường lồng bên trong).
function fieldXml(parts, cachedText, rPr, dirty) {
  const r = (inner) => `<w:r>${rPr}${inner}</w:r>`;
  const code = parts
    .map((p) => (typeof p === "string" ? r(`<w:instrText xml:space="preserve">${escapeXml(p)}</w:instrText>`) : p.xml))
    .join("");
  return (
    r(`<w:fldChar w:fldCharType="begin"${dirty ? ' w:dirty="true"' : ""}/>`) +
    code +
    r('<w:fldChar w:fldCharType="separate"/>') +
    (cachedText ? r(`<w:t>${cachedText}</w:t>`) : "") +
    r('<w:fldChar w:fldCharType="end"/>')
  );
}

function buildPageNumberOoxml() {
  const rPr = `<w:rPr>${buildRunProps(tnr(14))}</w:rPr>`;
  const field = (parts, cached, dirty) => ({ xml: fieldXml(parts, cached, rPr, dirty) });
  const page = () => field([" PAGE "], "2");
  const sum = field([" = ", field([" SECTION "], "1"), " + ", page(), " "], "3");
  const ifField = fieldXml([" IF ", sum, ' > 2 "', page(), '" "" '], "", rPr, true);
  const pPr =
    '<w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>' +
    '<w:ind w:left="0" w:right="0" w:firstLine="0"/><w:jc w:val="center"/></w:pPr>';
  return buildPackageXml(`<w:p>${pPr}${ifField}</w:p>`);
}

async function insertPageNumbers(context) {
  const sections = context.document.sections;
  sections.load("items");
  await context.sync();

  const headers = [];
  sections.items.forEach((section) => {
    ["Primary", "FirstPage", "EvenPages"].forEach((type) => headers.push(section.getHeader(type)));
  });

  headers.forEach((h) => h.insertOoxml(buildPageNumberOoxml(), Word.InsertLocation.replace));
  await context.sync();

  // insertOoxml hay để thừa đoạn trống (trước hoặc sau đoạn số trang) — dòng trống làm header cao
  // thêm và đẩy chữ xuống. Nhận diện đoạn số trang bằng việc nó chứa trường (không dựa vào chữ:
  // ở trang 1 trường cho kết quả rỗng), xoá các đoạn trống còn lại.
  const paragraphLists = headers.map((h) => h.paragraphs);
  paragraphLists.forEach((list) => list.load("items/text"));
  await context.sync();
  const fieldLists = paragraphLists.map((list) => list.items.map((p) => p.fields));
  fieldLists.flat().forEach((fields) => fields.load("items"));
  await context.sync();

  paragraphLists.forEach((list, i) => {
    const items = list.items;
    const fieldIndex = fieldLists[i].findIndex((fields) => fields.items.length > 0);
    if (fieldIndex < 0) {
      return;
    }
    const pageParagraph = items[fieldIndex];

    // Đặt lại định dạng qua API — style "Header" của văn bản có thể ghi đè <w:jc>/cỡ chữ trong OOXML.
    pageParagraph.set({
      alignment: Word.Alignment.centered,
      leftIndent: 0,
      rightIndent: 0,
      firstLineIndent: 0,
      spaceBefore: 0,
      spaceAfter: 0
    });
    pageParagraph.getRange("Whole").font.set({ name: "Times New Roman", size: 14, bold: false, italic: false });

    items.slice(0, fieldIndex).forEach((p) => {
      if (p.text.trim() === "") {
        p.delete();
      }
    });
    if (items.slice(fieldIndex + 1).every((p) => p.text.trim() === "") && fieldIndex < items.length - 1) {
      pageParagraph.getRange("End").expandTo(headers[i].getRange("End")).delete();
    }
  });
  await context.sync();

  return sections.items.length;
}

// 1e. Lưu văn bản thành PDF: Word tự xuất PDF (getFileAsync), add-in ghép các phần (slice) lại
// rồi tải xuống với tên trùng tên file Word. Không cần thư viện ngoài.
function officeAsync(fn) {
  return new Promise((resolve, reject) => {
    fn((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        reject(new Error(result.error ? result.error.message : "Lỗi không xác định khi xuất PDF."));
      }
    });
  });
}

function pdfFileName() {
  const url = Office.context.document.url || "";
  let last = url.split(/[\\/]/).pop() || "";
  try {
    last = decodeURIComponent(last); // URL OneDrive/SharePoint mã hoá tiếng Việt dạng %xx
  } catch (e) {
    // đường dẫn cục bộ có ký tự "%" lẻ — giữ nguyên
  }
  last = last.replace(/\.[^.]+$/, "");
  return `${last || "van-ban"}.pdf`;
}

function ensurePdfSupported() {
  if (!Office.context.requirements.isSetSupported("PdfFile", "1.1")) {
    throw new Error("Phiên bản Word này không hỗ trợ xuất PDF từ add-in. Dùng File → Save As → PDF.");
  }
}

async function exportPdfBytes() {
  const file = await officeAsync((cb) =>
    Office.context.document.getFileAsync(Office.FileType.Pdf, { sliceSize: 4 * 1024 * 1024 }, cb)
  );
  try {
    const chunks = [];
    let total = 0;
    for (let i = 0; i < file.sliceCount; i++) {
      const slice = await officeAsync((cb) => file.getSliceAsync(i, cb));
      chunks.push(new Uint8Array(slice.data));
      total += slice.data.length;
    }
    const bytes = new Uint8Array(total);
    let offset = 0;
    chunks.forEach((c) => {
      bytes.set(c, offset);
      offset += c.length;
    });
    return bytes;
  } finally {
    file.closeAsync();
  }
}

function downloadPdf(bytes, name) {
  const url = URL.createObjectURL(new Blob([bytes], { type: "application/pdf" }));
  const link = document.createElement("a");
  link.href = url;
  link.download = name;
  document.body.appendChild(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 60000);
}

async function saveAsPdf() {
  ensurePdfSupported();
  const name = pdfFileName();
  downloadPdf(await exportPdfBytes(), name);
  return name;
}

// 1f. Ô ký số trong PDF: đặt nhiều ô ký trong Word (mỗi ô một tên, VD ct_kt, ct_tckt) rồi xuất PDF —
// mỗi ô thành 1 trường chữ ký số (ô trống để ký bằng USB token trong PDF-XChange/VNPT...).
// Mỗi ô ký là chuỗi đánh dấu "#KY#<tên>#" (chữ trắng, cỡ 1pt — không nhìn thấy) bọc trong content
// control (tag "ccvb-ky", title = tên ô, hiện nhãn tên ô trên màn hình) để liệt kê/nhảy tới/xoá được và lưu cùng file Word.
// Khi xuất: Word xuất PDF, add-in tìm toạ độ các chuỗi đánh dấu trong PDF và thêm trường ký đúng chỗ.
// Thư viện đọc PDF (pdf.js) và sửa PDF (pdf-lib) ~1,9MB chỉ tải từ CDN khi xuất lần đầu.
const SIGN_TAG = "ccvb-ky";
// Tên ô ký chọn nhanh trên panel; chọn "Tên khác…" để gõ tên tuỳ ý.
const SIGN_NAME_PRESETS = ["chuky", "ct_khvt", "ct_kt", "ct_tckt", "ct_qlda", "ct_chuky"];
const CUSTOM_SIGN_NAME = "__khac";

const PDF_LIB_URL = "https://cdn.jsdelivr.net/npm/pdf-lib@1.17.1/dist/pdf-lib.min.js";
const PDFJS_URL = "https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.min.js";
const PDFJS_WORKER_URL = "https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.worker.min.js";

function loadScript(url) {
  return new Promise((resolve, reject) => {
    const script = document.createElement("script");
    script.src = url;
    script.onload = resolve;
    script.onerror = () => reject(new Error("Không tải được thư viện PDF (cần kết nối mạng): " + url));
    document.head.appendChild(script);
  });
}

let pdfLibsPromise = null;
function loadPdfLibs() {
  if (!pdfLibsPromise) {
    pdfLibsPromise = (async () => {
      await Promise.all([loadScript(PDF_LIB_URL), loadScript(PDFJS_URL)]);
      // Worker khác domain không chạy trực tiếp được — tải về rồi chạy qua blob URL.
      const workerCode = await (await fetch(PDFJS_WORKER_URL)).text();
      window.pdfjsLib.GlobalWorkerOptions.workerSrc = URL.createObjectURL(
        new Blob([workerCode], { type: "text/javascript" })
      );
      return { PDFLib: window.PDFLib, pdfjsLib: window.pdfjsLib };
    })().catch((error) => {
      pdfLibsPromise = null; // cho phép thử lại lần sau
      throw error;
    });
  }
  return pdfLibsPromise;
}

function getSignFields(context) {
  const controls = context.document.contentControls.getByTag(SIGN_TAG);
  controls.load("items/title,items/appearance");
  return controls;
}

// Hiện ô ký trên màn hình Word bằng nhãn đầu/cuối của content control (chế độ "Tags") mang tên ô,
// màu cam. Nhãn chỉ hiện khi soạn thảo — không in ra, không có trong PDF.
const SIGN_FIELD_LOOK = { appearance: "Tags", color: "#CC785C" };

// Danh sách chọn tên ô ký: tên đã đặt trong văn bản bị khoá "(đã đặt)"; tự chọn tên kế tiếp chưa dùng.
function renderSignNameOptions(placedNames) {
  const select = document.getElementById("select-sign-name");
  const current = select.value;
  const options = SIGN_NAME_PRESETS.map((name) => {
    const option = document.createElement("option");
    option.value = name;
    option.disabled = placedNames.includes(name);
    option.textContent = option.disabled ? `${name} (đã đặt)` : name;
    return option;
  });
  const custom = document.createElement("option");
  custom.value = CUSTOM_SIGN_NAME;
  custom.textContent = "Tên khác…";
  select.replaceChildren(...options, custom);

  const keep = options.find((o) => o.value === current && !o.disabled);
  const firstFree = options.find((o) => !o.disabled);
  select.value = current === CUSTOM_SIGN_NAME ? CUSTOM_SIGN_NAME : (keep || firstFree || custom).value;
  toggleCustomSignName();
}

function toggleCustomSignName() {
  const isCustom = document.getElementById("select-sign-name").value === CUSTOM_SIGN_NAME;
  const input = document.getElementById("input-sign-name");
  input.hidden = !isCustom;
  if (isCustom) {
    input.focus();
  }
}

function selectedSignName() {
  const value = document.getElementById("select-sign-name").value;
  return value === CUSTOM_SIGN_NAME ? document.getElementById("input-sign-name").value.trim() : value;
}

async function placeSignField(context) {
  const name = selectedSignName();
  if (!_pdf_sign_mjs__WEBPACK_IMPORTED_MODULE_1__.SIGN_NAME_RE.test(name)) {
    throw new Error("Tên ô ký chỉ gồm chữ không dấu, số và dấu gạch dưới, tối đa 40 ký tự (VD: ct_kt).");
  }

  const existing = getSignFields(context);
  await context.sync();
  if (existing.items.some((c) => c.title === name)) {
    throw new Error(`Đã có ô ký "${name}" trong văn bản. Mỗi ô ký phải có tên khác nhau.`);
  }

  const marker = context.document.getSelection().insertText((0,_pdf_sign_mjs__WEBPACK_IMPORTED_MODULE_1__.signMarker)(name), Word.InsertLocation.start);
  marker.font.set({ color: "#FFFFFF", size: 1, bold: false, italic: false, underline: "None", highlightColor: null });
  const control = marker.insertContentControl();
  control.set({ tag: SIGN_TAG, title: name, ...SIGN_FIELD_LOOK });
  await context.sync();

  document.getElementById("input-sign-name").value = "";
  return name;
}

// Danh sách ô ký trên panel: bấm tên để nhảy tới chỗ đặt ô, bấm × để xoá ô.
async function refreshSignList() {
  const list = document.getElementById("sign-list");
  try {
    let names = [];
    await Word.run(async (context) => {
      const controls = getSignFields(context);
      await context.sync();
      names = controls.items.map((c) => c.title);
      // Ô ký đặt bằng bản cũ (ẩn hoàn toàn) -> chuyển sang hiện nhãn.
      const hidden = controls.items.filter((c) => c.appearance !== SIGN_FIELD_LOOK.appearance);
      if (hidden.length) {
        hidden.forEach((c) => c.set(SIGN_FIELD_LOOK));
        await context.sync();
      }
    });
    list.replaceChildren(
      ...names.map((name) => {
        const chip = document.createElement("span");
        chip.className = "chip";
        const go = document.createElement("button");
        go.textContent = name;
        go.title = "Đi tới vị trí ô ký này";
        go.onclick = () => runSignControlAction(name, (c) => c.select(), `Đang ở vị trí ô ký "${name}".`);
        const remove = document.createElement("button");
        remove.textContent = "×";
        remove.title = "Xoá ô ký này";
        remove.onclick = () => runSignControlAction(name, (c) => c.delete(false), `Đã xoá ô ký "${name}".`);
        chip.append(go, remove);
        return chip;
      })
    );
    list.hidden = names.length === 0;
    renderSignNameOptions(names);
  } catch (e) {
    list.hidden = true;
  }
}

function runSignControlAction(name, fn, message) {
  return runAction(async (context) => {
    const controls = getSignFields(context);
    await context.sync();
    const control = controls.items.find((c) => c.title === name);
    if (!control) {
      throw new Error(`Không tìm thấy ô ký "${name}" (có thể đã bị xoá).`);
    }
    fn(control);
    await context.sync();
  }, message);
}

async function saveAsPdfWithSignatures(context) {
  ensurePdfSupported();
  const controls = getSignFields(context);
  await context.sync();
  const names = controls.items.map((c) => c.title);
  if (names.length === 0) {
    throw new Error("Chưa đặt ô ký nào. Đặt con trỏ vào chỗ cần ký, nhập tên ô (VD ct_kt) rồi bấm \"Đặt ô ký\".");
  }

  const libsPromise = loadPdfLibs(); // tải thư viện song song trong lúc Word xuất PDF
  libsPromise.catch(() => {}); // lỗi được báo ở chỗ await bên dưới

  // Bản PDF xuất từ add-in có vẽ cả nhãn content control: nhãn vừa lọt vào PDF vừa đẩy chuỗi đánh
  // dấu lệch khỏi vị trí con trỏ. Tạm ẩn nhãn trong lúc xuất rồi bật lại.
  controls.items.forEach((c) => (c.appearance = "Hidden"));
  await context.sync();
  let bytes;
  try {
    bytes = await exportPdfBytes();
  } finally {
    controls.items.forEach((c) => c.set(SIGN_FIELD_LOOK));
    await context.sync();
  }
  const { PDFLib, pdfjsLib } = await libsPromise;

  // Mỗi tên chỉ lấy vị trí đầu tiên (phòng khi chuỗi đánh dấu bị copy trùng sang chỗ khác).
  const positions = [];
  (await (0,_pdf_sign_mjs__WEBPACK_IMPORTED_MODULE_1__.findMarkers)(pdfjsLib, bytes)).forEach((m) => {
    if (names.includes(m.name) && !positions.some((p) => p.name === m.name)) {
      positions.push(m);
    }
  });
  if (positions.length === 0) {
    throw new Error("Không tìm thấy vị trí ô ký nào trong PDF. Ô ký phải nằm trong phần nội dung văn bản (không ở header/footer, hộp văn bản).");
  }

  const fileName = pdfFileName();
  downloadPdf(await (0,_pdf_sign_mjs__WEBPACK_IMPORTED_MODULE_1__.addSignatureFields)(PDFLib, bytes, positions), fileName);
  return {
    fileName,
    placed: positions.map((p) => `${p.name} (tr.${p.pageIndex + 1})`),
    missing: names.filter((n) => !positions.some((p) => p.name === n))
  };
}

function formatSignedPdfMessage(r) {
  const missing = r.missing.length ? ` Không tìm thấy vị trí: ${r.missing.join(", ")}.` : "";
  return `Đã xuất ${r.fileName} với ${r.placed.length} ô ký: ${r.placed.join(", ")} (xem thư mục Downloads).${missing}`;
}

// 2. Chèn khối Quốc hiệu - Tiêu ngữ (dạng bảng 2 hàng x 2 cột, không viền) ở đầu văn bản
// Hàng 1: tên cơ quan / quốc hiệu - tiêu ngữ. Hàng 2: số hiệu văn bản / địa danh - ngày tháng năm.
// Cỡ chữ theo Nghị định 30/2020/NĐ-CP: quốc hiệu/tên cơ quan 13, tiêu ngữ/địa danh-ngày tháng 14.
async function insertQuocHieu(context) {
  const body = context.document.body;
  const table = body.insertTable(2, 2, Word.InsertLocation.start, [
    ["", ""],
    ["", ""]
  ]);
  await context.sync();

  makeTablePlain(table);

  // Bỏ lề trong ô, chia 2 cột: trái 70mm cho "CÔNG TY ĐIỆN LỰC ĐÀ NẴNG" (~65mm ở cỡ 13),
  // phải 105mm cho "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM" (~96mm ở cỡ 13 đậm) — mỗi bên còn
  // dư vài mm nên hai khối cách nhau ~7mm, không dính vào nhau. Tổng 175mm, rộng hơn vùng chữ
  // (165mm) nên bảng lấn 10mm sang lề phải (vẫn cách mép giấy 5mm), đẩy khối Quốc hiệu ra
  // phía phải trang như cách trình bày thông dụng.
  table.setCellPadding(Word.CellPaddingLocation.left, 0);
  table.setCellPadding(Word.CellPaddingLocation.right, 0);
  table.getCell(1, 0).column.width = mmToPoints(QUOC_HIEU_COLUMNS_MM.left);
  table.getCell(1, 1).column.width = mmToPoints(QUOC_HIEU_COLUMNS_MM.right);
  table.verticalAlignment = Word.VerticalAlignment.top;
  await context.sync();

  const font13 = tnr(13);
  const font14 = tnr(14);
  // NĐ 30/2020: ngày nhỏ hơn 10 và tháng 1, 2 phải ghi thêm số 0 ở trước (VD: ngày 05 tháng 01).
  const today = new Date();
  const day = String(today.getDate()).padStart(2, "0");
  const monthNumber = today.getMonth() + 1;
  const month = monthNumber <= 2 ? String(monthNumber).padStart(2, "0") : String(monthNumber);
  const year = today.getFullYear();

  await fillCell(context, table.getCell(0, 0), [
    { text: "CÔNG TY ĐIỆN LỰC ĐÀ NẴNG", alignment: Word.Alignment.centered, font: font13 },
    {
      alignment: Word.Alignment.centered,
      runs: [
        { text: "ĐIỆN ", font: { ...font13, bold: true } },
        { text: "LỰC TRÀ", font: { ...font13, bold: true, underline: true } },
        { text: " MY", font: { ...font13, bold: true } }
      ]
    }
  ]);

  await fillCell(context, table.getCell(0, 1), [
    { text: "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", alignment: Word.Alignment.centered, font: { ...font13, bold: true } },
    { text: "Độc lập - Tự do - Hạnh phúc", alignment: Word.Alignment.centered, font: { ...font14, bold: true, underline: true } }
  ]);

  // Hàng 2 của cả hai cột cùng cách hàng trên 6pt để "Số:" và dòng ngày tháng thẳng hàng.
  await fillCell(context, table.getCell(1, 0), [
    // Chỗ trống ghi số dùng dấu cách không ngắt (NBSP) để "Xoá dấu cách thừa" không gộp mất.
    {
      text: `Số:${String.fromCharCode(0xa0).repeat(8)}/ĐLTM-KHKT`,
      alignment: Word.Alignment.centered,
      font: font13,
      spaceBefore: 6
    }
  ]);

  await fillCell(context, table.getCell(1, 1), [
    {
      text: `Xã Trà My, ngày ${day} tháng ${month} năm ${year}`,
      alignment: Word.Alignment.centered,
      font: { ...font14, italic: true },
      spaceBefore: 6
    }
  ]);

  await context.sync();
}

// 3. Chèn khối "Nơi nhận" (trái) và "Chữ ký" (phải) ở cuối văn bản
// Theo Nghị định 30/2020/NĐ-CP: "Nơi nhận:" cỡ 12, nghiêng, đậm; các dòng liệt kê nơi nhận
// cỡ 11, chữ đứng; quyền hạn/chức vụ người ký cỡ 14, đậm; họ tên người ký cỡ 14, đứng, đậm.
async function insertNoiNhanChuKy(context) {
  const body = context.document.body;
  const table = body.insertTable(1, 2, Word.InsertLocation.end, [["", ""]]);
  await context.sync();

  makeTablePlain(table);
  table.autoFitWindow();
  const centered = Word.Alignment.centered;

  await fillCell(context, table.getCell(0, 0), [
    { text: "Nơi nhận:", font: tnr(12, { italic: true, bold: true }) },
    { text: "- Như trên;", font: tnr(11) },
    { text: "- Lưu: VT, KHKT.", font: tnr(11) }
  ]);

  await fillCell(context, table.getCell(0, 1), [
    { text: "KT. TRƯỞNG ĐIỆN LỰC", alignment: centered, font: tnr(14, { bold: true }) },
    { text: "PHÓ TRƯỞNG ĐIỆN LỰC", alignment: centered, font: tnr(14, { bold: true }) },
    { text: "(Ký, đóng dấu, ghi rõ họ tên)", alignment: centered, font: tnr(14, { italic: true }) },
    { text: "", font: tnr(14) },
    { text: "", font: tnr(14) },
    { text: "Nguyễn Văn A", alignment: centered, font: tnr(14, { bold: true }) }
  ]);

  await context.sync();
}

// 3b. Chèn tên loại văn bản + trích yếu tại vị trí con trỏ (Tờ trình, Báo cáo, Kế hoạch...).
// Theo NĐ 30/2020: tên loại văn bản in hoa, cỡ 14, đứng, đậm, căn giữa; trích yếu ngay bên dưới,
// cỡ 14, đứng, đậm, căn giữa; dưới trích yếu có đường kẻ ngang nét liền dài 1/3 - 1/2 dòng chữ.
// Công văn không có tên loại (trích yếu "V/v ..." nằm dưới số ký hiệu) nên không có trong danh sách.
async function insertDocTitle(context) {
  const docType = document.getElementById("select-doc-type").value;
  const summary = document.getElementById("input-doc-summary").value.trim();

  const selection = context.document.getSelection();
  const parentTable = selection.parentTableOrNullObject;
  parentTable.load("isNullObject");
  await context.sync();
  if (!parentTable.isNullObject) {
    throw new Error("Con trỏ đang nằm trong bảng. Hãy đặt con trỏ ở một dòng trống bên dưới khối Quốc hiệu.");
  }

  // Đường kẻ dài ~40% dòng trích yếu (ước lượng ~2,6mm/ký tự ở cỡ 14 đậm), trong khoảng 25-60mm.
  const textMm = Math.min((summary || docType).length * 2.6, TEXT_WIDTH_MM);
  const ruleMm = Math.max(25, Math.min(60, textMm * 0.4));
  const centered = Word.Alignment.centered;
  const lines = [
    { text: docType.toUpperCase(), alignment: centered, font: tnr(14, { bold: true }), spaceBefore: 12 },
    summary ? { text: summary, alignment: centered, font: tnr(14, { bold: true }) } : null,
    { text: "", alignment: centered, font: tnr(6), rule: true, indentMm: (TEXT_WIDTH_MM - ruleMm) / 2, spaceBefore: 2 },
    // Đoạn trống thường ở cuối: nếu Word gộp đoạn cuối với đoạn đang có con trỏ thì chỉ gộp
    // đoạn này, không làm chữ của văn bản dính đường kẻ / căn giữa.
    { text: "", font: tnr(14) }
  ].filter(Boolean);

  // Tạo 1 đoạn trống mới ngay trước đoạn có con trỏ (bằng API đoạn văn, luôn nằm ngoài bảng —
  // kể cả khi con trỏ ở dòng sát ngay sau bảng Quốc hiệu/Nơi nhận), rồi thay đoạn trống đó bằng
  // khối tiêu đề. Chèn OOXML thẳng "trước" đoạn sát sau bảng có thể bị Word đưa vào ô cuối bảng.
  const anchor = selection.paragraphs.getFirst().insertParagraph("", Word.InsertLocation.before);
  const inserted = anchor.insertOoxml(buildCellOoxml(lines), Word.InsertLocation.replace);
  await context.sync();
  await applyLineFormats(context, inserted.paragraphs, lines.slice(0, -1));

  return docType.toUpperCase();
}

// Bảng dàn trang Quốc hiệu – Tiêu ngữ và Nơi nhận – Chữ ký (do add-in chèn hoặc tự gõ) không
// phải bảng số liệu: không được căn giữa / lặp tiêu đề, nếu không sẽ làm hỏng bố cục văn bản.
function isLayoutTable(table) {
  const text = table.values.flat().join("\n").normalize("NFC");
  return /CỘNG H(ÒA|OÀ) XÃ HỘI CHỦ NGHĨA/i.test(text) || /^\s*Nơi nhận\s*:/im.test(text);
}

// Áp dụng "fn" cho mọi bảng số liệu trong thân văn bản (bỏ qua bảng dàn trang ở trên). Dùng
// chung cho các thao tác bảng bên dưới.
async function forEachTable(context, fn) {
  const tables = context.document.body.tables;
  tables.load("items/values");
  await context.sync();

  const dataTables = tables.items.filter((t) => !isLayoutTable(t));
  if (dataTables.length === 0) {
    throw new Error("Không tìm thấy bảng số liệu nào trong văn bản (đã bỏ qua bảng Quốc hiệu, Nơi nhận – Chữ ký).");
  }

  dataTables.forEach(fn);

  await context.sync();
  return dataTables.length;
}

// 4. Lặp lại tiêu đề bảng (hàng 1) cho mọi bảng trong văn bản
function repeatTableHeaders(context) {
  return forEachTable(context, (table) => {
    table.headerRowCount = 1;
  });
}

// 5. Căn giữa nội dung (ngang + dọc) trong mọi bảng
function centerTableContent(context) {
  return forEachTable(context, (table) => {
    table.horizontalAlignment = Word.Alignment.centered;
    table.verticalAlignment = Word.VerticalAlignment.center;
  });
}

// Áp dụng "fn" cho từng paragraph trong phạm vi đã chọn ("selection") hoặc toàn thân văn
// bản ("body", không tính nội dung trong bảng). Dùng chung cho các thao tác đoạn văn bên dưới.
// "properties": các thuộc tính đoạn cần đọc trong fn (VD "alignment,isListItem").
async function forEachParagraph(context, scope, fn, properties) {
  const paragraphs =
    scope === "selection" ? context.document.getSelection().paragraphs : context.document.body.paragraphs;
  paragraphs.load(properties ? properties.split(",").map((p) => `items/${p.trim()}`).join(",") : "items");
  await context.sync();

  let items = paragraphs.items;

  if (scope === "body") {
    // body.paragraphs bao gồm cả đoạn văn nằm trong bảng (từ WordApi 1.3 trở lên), nên phải
    // lọc bỏ chúng ở đây — nếu không, thao tác trên "toàn văn bản" sẽ ghi đè định dạng đã đặt
    // riêng cho nội dung trong bảng (vd. căn giữa bảng rồi justify toàn văn bản).
    const parentTables = items.map((p) => p.parentTableOrNullObject);
    parentTables.forEach((t) => t.load("isNullObject"));
    await context.sync();
    items = items.filter((p, i) => parentTables[i].isNullObject);
  }

  items.forEach(fn);

  await context.sync();
  return items.length;
}

// 6a. Thụt lề đầu dòng 1cm cho các đoạn văn đang được chọn (bôi đen)
function setFirstLineIndent(p) {
  p.firstLineIndent = mmToPoints(10); // 1cm = 10mm
}

function applyFirstLineIndent(context) {
  return forEachParagraph(context, "selection", setFirstLineIndent);
}

// 6b. Chuẩn hoá giãn dòng cho các đoạn văn đang được chọn (bôi đen): giãn dòng 1,17 dòng +
// cách đoạn trên 3pt, dưới 6pt — nằm trong khung của Nghị định 30/2020/NĐ-CP (giãn dòng từ
// dòng đơn đến 1,5 dòng, cách đoạn tối thiểu 6pt).
// Lưu ý: lineSpacing dùng đơn vị "1/12 dòng" của Word — 12 = dòng đơn, 14 ≈ 1,17 dòng,
// 18 = 1,5 dòng, 24 = dòng đôi.
function setLineSpacing(p) {
  p.lineSpacing = 14;
  p.spaceBefore = 3;
  p.spaceAfter = 6;
}

function normalizeLineSpacing(context) {
  return forEachParagraph(context, "selection", setLineSpacing);
}

// 6c. Xoá dấu cách thừa (2+ dấu cách liên tiếp -> 1 dấu cách, toàn văn bản kể cả trong bảng)
// Gộp bằng cách tìm-thay chuỗi thường "  " (2 dấu cách liền) -> " " (1 dấu cách), lặp nhiều
// lượt cho đến khi không còn khớp nào. Không dùng wildcard search (" @") vì cách đó từng bị
// Word hiểu sai, chèn thêm hàng loạt dấu cách thay vì gộp lại và gây lag.
async function cleanupExtraSpaces(context) {
  let removedSpaces = 0;
  for (let pass = 0; pass < 10; pass++) {
    const matches = context.document.body.search("  ");
    matches.load("items");
    await context.sync();

    if (matches.items.length === 0) {
      break;
    }

    matches.items.forEach((range) => {
      range.insertText(" ", Word.InsertLocation.replace);
    });
    removedSpaces += matches.items.length;
    await context.sync();
  }

  return removedSpaces;
}

function justifyParagraphs(p) {
  p.alignment = Word.Alignment.justified;
}

// 6d. Justify đoạn đang chọn
function justifySelection(context) {
  return forEachParagraph(context, "selection", justifyParagraphs);
}

// 6e. Justify toàn văn bản (không tính nội dung trong bảng)
function justifyAll(context) {
  return forEachParagraph(context, "body", justifyParagraphs);
}

// 7. Đổi số thành chữ tiếng Việt (offline, không gọi API) và chèn vào vị trí con trỏ.
// Logic đọc số nằm trong vn-text.mjs (có unit test).
async function insertNumberAsWords(context) {
  const input = document.getElementById("input-number");
  const words = (0,_vn_text_mjs__WEBPACK_IMPORTED_MODULE_0__.numberToVietnameseWords)(input.value);

  context.document.getSelection().insertText(words, Word.InsertLocation.replace);
  await context.sync();

  return words;
}

// 0. Chuẩn hoá nhanh toàn văn bản: chạy liền các thao tác ở trên trong 1 lần bấm.
//  - Lề trang (chỉ trên Word desktop; Word Online thì bỏ qua bước này, không báo lỗi).
//  - Font Times New Roman, xoá dấu cách thừa (toàn văn bản, kể cả trong bảng).
//  - Justify + giãn dòng + thụt đầu dòng 1cm cho đoạn văn ngoài bảng (bảng Quốc hiệu, Nơi
//    nhận – Chữ ký và bảng số liệu giữ nguyên). Bỏ qua đoạn đang căn giữa/căn phải (tên loại
//    văn bản, trích yếu, tiêu đề...) và không thụt đầu dòng cho đoạn gạch đầu dòng/đánh số.
async function quickNormalize(context) {
  const result = { pageSetup: false, spaces: 0, paragraphs: 0, skipped: 0 };

  if (Office.context.requirements.isSetSupported("WordApiDesktop", "1.3")) {
    await applyPageSetup(context);
    result.pageSetup = true;
  }

  await unifyFont(context);
  result.spaces = await cleanupExtraSpaces(context);

  await forEachParagraph(
    context,
    "body",
    (p) => {
      if (p.alignment === Word.Alignment.centered || p.alignment === Word.Alignment.right) {
        result.skipped++;
        return;
      }
      justifyParagraphs(p);
      setLineSpacing(p);
      if (!p.isListItem) {
        setFirstLineIndent(p);
      }
      result.paragraphs++;
    },
    "alignment,isListItem"
  );

  return result;
}

function formatQuickNormalizeMessage(r) {
  const parts = [
    r.pageSetup ? "lề trang" : null,
    "font Times New Roman",
    `xoá ${r.spaces} dấu cách thừa`,
    `căn đều, giãn dòng, thụt đầu dòng ${r.paragraphs} đoạn`
  ].filter(Boolean);
  const notes = [];
  if (r.skipped > 0) {
    notes.push(`giữ nguyên ${r.skipped} đoạn căn giữa/căn phải`);
  }
  if (!r.pageSetup) {
    notes.push("bỏ qua lề trang vì không phải Word desktop");
  }
  return `Đã chuẩn hoá: ${parts.join("; ")}.` + (notes.length ? ` (${notes.join("; ")})` : "");
}

})();

/******/ })()
;
//# sourceMappingURL=taskpane.js.map