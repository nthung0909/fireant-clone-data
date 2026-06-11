/*
  bulk-export.js
  - Cập nhật FIREANT_TOKEN với token của bạn.
  - Mặc định thử với ["VNM", "VCB"].
  - Mỗi cổ phiếu cách nhau 5 giây.
  - Mỗi cổ phiếu xuất ra một file Excel riêng.
  - File Excel có các sheet và format theo logic của popup.js.
*/

const STOCK_IDS = ["VNM", "VCB"];
const FIREANT_TOKEN = ""; // Điền token ở đây
const STOCK_DELAY_MS = 5000;

function delay(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function getExcelSections(isBank) {
  const basicRows = [
    { label: "EPS", getValue: (f) => f.BasicEPS },
    { label: "BVPS", getValue: (f) => f.BookValuePerShare },
    { label: "P/E", getValue: (f) => f.PE },
    { label: "P/B", getValue: (f) => f.PB },
    { label: "ROA", getValue: (f) => f.ROA, isPercent: true },
    { label: "ROE", getValue: (f) => f.ROE, isPercent: true },
    { label: "ROIC", getValue: (f) => f.ROIC, isPercent: true },
    {
      label: "Nợ vay/VCSH",
      getValue: (f) => {
        if (!f) return null;
        if (f.TotalDebtOverEquity !== undefined && f.TotalDebtOverEquity !== null) {
          return f.TotalDebtOverEquity;
        }
        const debt = f.TotalDebt || (f.ShortTermInterestBearingDebt || 0) + (f.LongTermInterestBearingDebt || 0) || 0;
        const equity = f.TotalEquity || f.TotalStockHolderEquity || 1;
        return debt / (equity || 1);
      },
    },
    {
      label: "Tổng nợ/Tổng tài sản",
      getValue: (f) => {
        if (!f) return null;
        if (f.TotalDebtOverAsset !== undefined && f.TotalDebtOverAsset !== null) {
          return f.TotalDebtOverAsset;
        }
        const debt = f.TotalDebt || 0;
        const asset = f.TotalAsset || 1;
        return debt / (asset || 1);
      },
    },
  ];

  if (isBank) {
    basicRows.push(
      { label: "NII (ngàn tỷ)", getValue: (f) => f.NetInterestIncome, scale: 1000000000 },
      { label: "NIM", getValue: (f) => f.NIM, isPercent: true },
      { label: "CIR", getValue: (f) => f.CIR, isPercent: true },
      { label: "NPL", getValue: (f) => f.NPLToLoan, isPercent: true },
      { label: "CAR", getValue: (f) => f.CAR, isPercent: true },
      { label: "LDR", getValue: (f) => f.LDR },
      { label: "LAR", getValue: (f) => f.LAR },
      { label: "LLR", getValue: (f) => f.LoanlossReservesToNPL }
    );
  } else {
    basicRows.push(
      { label: "Thanh toán nhanh", getValue: (f) => f.QuickRatio },
      { label: "Thanh toán hiện hành", getValue: (f) => f.CurrentRatio },
      { label: "Tỷ lệ tiền mặt", getValue: (f) => f.CashRatio }
    );
  }

  return [
    {
      title: "Chỉ số cơ bản",
      rows: basicRows,
    },
    {
      title: "Kết quả kinh doanh",
      rows: [
        { label: "Doanh thu thuần", getValue: (f) => f.NetSale },
        { label: "Giá vốn hàng bán", getValue: (f) => f.CostOfGoodSold, hideForBank: true },
        { label: "Lợi nhuận gộp", getValue: (f) => f.GrossProfit },
        {
          label: "LN HĐTC & Cty LDLK",
          getValue: (f) => (f?.ProfitFromFinancialActivity || 0) + (f?.ProfitFromAssociate || 0),
        },
        { label: "LN khác", getValue: (f) => f.OtherProfit },
        { label: "LN trước thuế", getValue: (f) => f.ProfitBeforeTax },
        { label: "LN sau thuế", getValue: (f) => f.ProfitAfterTax },
        { label: "Lợi nhuận sau thuế công ty mẹ", getValue: (f) => f.ParentCompanyShareholderProfitAfterTax },
        {
          label: "Lợi ích CĐ không kiểm soát",
          getValue: (f) => (f?.ProfitAfterTax || 0) - (f?.ParentCompanyShareholderProfitAfterTax || 0),
        },
        { label: "Lợi nhuận ròng", getValue: (f) => f.ParentCompanyShareholderProfitAfterTax },
      ],
    },
    {
      title: "Tài sản và VCSH",
      rows: [
        { label: "Tổng nợ", getValue: (f) => f.TotalDebt },
        { label: "Nợ ngắn hạn", getValue: (f) => f.TotalShortTermDebt, hideForBank: true },
        { label: "Nợ dài hạn", getValue: (f) => f.TotalLongTermDebt, hideForBank: true },
        { label: "Nợ nhóm 1", getValue: (f) => f.StandardDebt },
        { label: "Nợ nhóm 2", getValue: (f) => f.WatchlistDebt },
        { label: "Nợ nhóm 3", getValue: (f) => f.SubstandardDebt },
        { label: "Nợ nhóm 4", getValue: (f) => f.DoubtfulDebt },
        { label: "Nợ nhóm 5", getValue: (f) => f.BadDebt },
        { label: "Vốn chủ sở hữu", getValue: (f) => f.TotalStockHolderEquity },
        { label: "Tổng giá trị tồn kho", getValue: (f) => f.TotalInventory, hideForBank: true },
        { label: "Tồn kho nguyên vật liệu", getValue: () => null, hideForBank: true },
        { label: "Công cụ, dụng cụ", getValue: () => null, hideForBank: true },
        { label: "Chi phí SXKD dở dang", getValue: () => null, hideForBank: true },
        { label: "Thành phẩm", getValue: () => null, hideForBank: true },
        { label: "Hàng hoá", getValue: () => null, hideForBank: true },
        { label: "Hàng gửi bán", getValue: () => null, hideForBank: true },
      ],
    },
    {
      title: "Lưu chuyển tiền tệ",
      rows: [
        { label: "Lưu chuyển tiền thuần từ HĐ Kinh doanh", getValue: (f) => f.CashflowFromOperatingActivity },
        { label: "Lưu chuyển tiền thuần từ HĐ Tài chính", getValue: (f) => f.CashflowFromFinancingActivity },
        { label: "Lưu chuyển tiền thuần từ HĐ Đầu tư", getValue: (f) => f.CashflowFromInvestingActivity },
        {
          label: "Tiền mặt",
          getValue: (f, isBank) => (isBank ? f?.CashGoldJewelry : f?.Cash),
        },
        { label: "Tiền và tương đương tiền cuối kỳ", getValue: (f) => f.CashAndCashEquivalentAtTheEndOfPeriod },
      ],
    },
  ];
}

function buildSheet(workbook, section, sortedData, isQuarter, isBank) {
  const sheet = workbook.addWorksheet(section.title);
  const defaultNumFmt = section.defaultNumFmt || "#,##0.00";
  const defaultPercentFmt = section.percentNumFmt || "0.00%";

  const headers = sortedData.map((item) => {
    return item.quarter != null ? `Q${item.quarter}/${item.year}` : `${item.year}`;
  });

  sheet.addRow(["Chỉ số", ...headers]);

  section.rows.forEach((rowConfig) => {
    if (rowConfig.hideForBank && isBank) return;
    const rowValues = [rowConfig.label];

    sortedData.forEach((item) => {
      const rawValue = rowConfig.getValue(item.financialValues, isBank);
      let value = rawValue;
      const numericValue = Number(rawValue);
      if (Number.isFinite(numericValue)) {
        value = numericValue / (rowConfig.scale || 1);
        if (rowConfig.isPercent) {
          value = Math.abs(value) <= 1 ? value : value / 100;
        }
      }
      rowValues.push(value);
    });

    const addedRow = sheet.addRow(rowValues);
    if (rowConfig.isPercent) {
      addedRow.eachCell((cell, colNumber) => {
        if (colNumber === 1) return;
        if (typeof cell.value === "number") {
          cell.numFmt = rowConfig.numFmt || defaultPercentFmt;
        }
      });
    }
  });

  sheet.columns.forEach((column, idx) => {
    column.width = idx === 0 ? 40 : 20;
  });

  sheet.getRow(1).font = { bold: true };
  sheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    row.eachCell((cell, colNumber) => {
      if (colNumber === 1) return;
      if (cell.numFmt) return;
      if (typeof cell.value === "number") {
        cell.numFmt = defaultNumFmt;
      }
    });
  });
}

function sortSheetData(items) {
  return [...items].sort((a, b) => {
    if (a.year !== b.year) {
      return b.year - a.year;
    }
    const qa = a.quarter != null ? a.quarter : -1;
    const qb = b.quarter != null ? b.quarter : -1;
    return qb - qa;
  });
}

async function fetchFinancialData(stockID, token, type, count) {
  const apiType = type === "quarter" ? "Q" : "Y";
  const url = `https://restv2.fireant.vn/symbols/${encodeURIComponent(stockID)}/financial-data?type=${apiType}&count=${count}`;

  const response = await fetch(url, {
    headers: {
      accept: "application/json, text/plain, */*",
      authorization: `Bearer ${token}`,
    },
  });

  if (!response.ok) {
    throw new Error(`API Error ${response.status} for ${stockID} ${type}`);
  }

  const data = await response.json();
  if (!Array.isArray(data)) {
    throw new Error(`Dữ liệu trả về không phải mảng cho ${stockID} ${type}`);
  }

  return data;
}

function getStockCompanyType(yearData, quarterData) {
  return yearData?.[0]?.companyType || quarterData?.[0]?.companyType || "";
}

async function writeWorkbookToFile(workbook, stockID) {
  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const filename = `[${stockID}] Phân tích cơ bản.xlsx`;

  await new Promise((resolve, reject) => {
    chrome.downloads.download(
      {
        url,
        filename,
        saveAs: true,
      },
      (downloadId) => {
        if (chrome.runtime.lastError) {
          reject(new Error(chrome.runtime.lastError.message));
          return;
        }
        resolve(downloadId);
      }
    );
  });
}

async function exportStock(stockID, token) {
  console.info(`Bắt đầu lấy dữ liệu cho ${stockID}`);
  const [yearData, quarterData] = await Promise.all([
    fetchFinancialData(stockID, token, "year", 6),
    fetchFinancialData(stockID, token, "quarter", 1),
  ]);

  if (!yearData.length && !quarterData.length) {
    throw new Error(`Không có dữ liệu financial cho ${stockID}`);
  }

  const isBank = getStockCompanyType(yearData, quarterData) === "Bank";
  const combinedData = [...quarterData, ...yearData];
  const sortedData = sortSheetData(combinedData);

  const workbook = new ExcelJS.Workbook();
  const sections = getExcelSections(isBank);
  sections.forEach((section) => buildSheet(workbook, section, sortedData, quarterData.length > 0, isBank));

  await writeWorkbookToFile(workbook, stockID);
  console.info(`Hoàn tất xuất file cho ${stockID}`);
}

async function exportFireantStocks(stockIDs, token) {
  if (!token || !token.trim()) {
    console.error("Vui lòng nhập FIREANT_TOKEN vào bulk-export.js trước khi chạy.");
    return;
  }

  if (!Array.isArray(stockIDs) || stockIDs.length === 0) {
    console.error("Danh sách cổ phiếu rỗng. Vui lòng cấu hình STOCK_IDS.");
    return;
  }

  for (let index = 0; index < stockIDs.length; index += 1) {
    const stockID = String(stockIDs[index]).trim().toUpperCase();
    if (!stockID) {
      continue;
    }

    try {
      await exportStock(stockID, token);
    } catch (error) {
      console.error(`Lỗi với ${stockID}:`, error);
    }

    if (index < stockIDs.length - 1) {
      console.info(`Đợi ${STOCK_DELAY_MS / 1000}s trước khi tiếp tục...`);
      await delay(STOCK_DELAY_MS);
    }
  }
}

exportFireantStocks(STOCK_IDS, FIREANT_TOKEN);
