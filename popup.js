const stockIDInput = document.getElementById("stockID");
const typeInput = document.getElementById("type");
const countInput = document.getElementById("count");
const submitBtn = document.getElementById("submitBtn");
const downloadBtn = document.getElementById("downloadBtn");
const statusEl = document.getElementById("status");
const previewNav = document.getElementById("previewNav");
const previewContainer = document.getElementById("preview");

let previewData = null;
let previewStockID = null;
let previewType = null;
let previewIsBank = false;

downloadBtn.disabled = true;

submitBtn.addEventListener("click", async () => {
  try {
    const stockID = stockIDInput.value.trim().toUpperCase();
    const type = typeInput.value;
    const count = Number(countInput.value);

    if (!stockID) {
      throw new Error("Thiếu StockID");
    }

    previewData = null;
    previewStockID = null;
    previewType = null;
    downloadBtn.disabled = true;
    previewContainer.innerHTML = "";
    previewNav.innerHTML = "";
    previewNav.hidden = true;

    setStatus("Đang lấy token từ trang Fireant...");
    const token = await getAccessTokenFromPage();

    setStatus("Đang tải dữ liệu...");
    const data = await fetchFinancialData({
      stockID,
      token,
      type,
      count,
    });


    if (!Array.isArray(data) || data.length === 0) {
      throw new Error("Không có dữ liệu");
    }

    previewData = data;
    previewStockID = stockID;
    previewType = type;
    previewIsBank = data[0]?.companyType === "Bank";

    renderPreviewTable(data, type, previewIsBank);
    downloadBtn.disabled = false;
    previewNav.hidden = false;
    setStatus("Dữ liệu đã sẵn sàng. Kiểm tra bảng và nhấn Download excel để tải về.");
  } catch (err) {
    console.error(err);
    downloadBtn.disabled = true;
    setStatus(err.message);
  }
});

downloadBtn.addEventListener("click", async () => {
  try {
    if (!previewData || !previewStockID || !previewType) {
      throw new Error("Không có dữ liệu để tải xuống. Vui lòng nhấn Submit trước.");
    }

    setStatus("Đang tạo Excel và tải về...");
    await generateExcel(previewData, previewStockID, previewType, previewIsBank);
    setStatus("Hoàn tất tải file Excel.");
  } catch (err) {
    console.error(err);
    setStatus(err.message);
  }
});

function getPreviewSections(isBank) {
  const basicRows = [
    { label: "EPS", key: "BasicEPS" },
    { label: "BVPS", key: "BookValuePerShare" },
    { label: "P/E", key: "PE" },
    { label: "P/B", key: "PB" },
    { label: "ROA", key: "ROA", isPercent: true },
    { label: "ROE", key: "ROE", isPercent: true },
    { label: "ROIC", key: "ROIC", isPercent: true },
    {
      label: "Nợ vay/VCSH",
      getValue: (item) => {
        const f = item.financialValues || {};
        if (f.TotalDebtOverEquity !== undefined && f.TotalDebtOverEquity !== null) {
          return f.TotalDebtOverEquity;
        }
        const debt = f.TotalDebt || (f.ShortTermInterestBearingDebt || 0) + (f.LongTermInterestBearingDebt || 0) || 0;
        const equity = f.TotalEquity || f.TotalStockHolderEquity || 1;
        return debt / (equity || 1);
      },
    },
    { label: "CAGR Doanh thu 3 năm", getValue: (item) => (item.financialValues || {}).SaleGrowth_03Yr },
    { label: "CAGR LNST 3 năm", getValue: (item) => {
      const f = item.financialValues || {};
      return f.ProfitAfterTaxGrowth_03Yr ?? f.ProfitGrowth_03Yr;
    }},
    { label: "CAGR EPS 3 năm", getValue: (item) => (item.financialValues || {}).BasicEPSGrowth_03Yr },
    { label: "CAGR Tài sản 3 năm", getValue: (item) => (item.financialValues || {}).TotalAssetGrowth_03Yr },
    { label: "CAGR Vốn chủ 3 năm", getValue: (item) => (item.financialValues || {}).EquityGrowth_03Yr },
    { label: "Tăng trưởng tín dụng", getValue: (item, idx, all) => {
      const f = item.financialValues || {};
      const currentLoan = Number(f.CustomerLoan ?? f.CustomerLoans ?? 0);
      const prev = all && all[idx + 1];
      const prevLoan = Number(prev?.financialValues?.CustomerLoan ?? prev?.financialValues?.CustomerLoans ?? 0);
      if (!prevLoan) return null;
      return (currentLoan - prevLoan) / prevLoan;
    }},
    {
      label: "Tổng nợ/Tổng tài sản",
      getValue: (item) => {
        const f = item.financialValues || {};
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
      { label: "NII (ngàn tỷ)", key: "NetInterestIncome", scale: 1000000000 },
      { label: "NIM", key: "NIM", isPercent: true },
      { label: "CIR", key: "CIR", isPercent: true },
      { label: "NPL", key: "NPLToLoan", isPercent: true },
      { label: "CAR", key: "CAR", isPercent: true },
      { label: "LDR", key: "LDR" },
      { label: "LAR", key: "LAR" },
      { label: "LLR", key: "LoanlossReservesToNPL" }
    );
  }
  else {
    // For non-bank companies add liquidity ratios
    basicRows.push(
      { label: "Thanh toán nhanh", key: "QuickRatio" },
      { label: "Thanh toán hiện hành", key: "CurrentRatio" },
      { label: "Tỷ lệ tiền mặt", key: "CashRatio" }
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
        { label: "Doanh thu thuần", key: "NetSale" },
        // CostOfGoodSold will be hidden for Banks
        { label: "Giá vốn hàng bán", key: "CostOfGoodSold", hideForBank: true },
        { label: "Lợi nhuận gộp", key: "GrossProfit" },
        {
          label: "LN HĐTC & Cty LDLK",
          getValue: (item) => {
            const f = item.financialValues || {};
            return (f.ProfitFromFinancialActivity || 0) + (f.ProfitFromAssociate || 0);
          },
        },
        { label: "LN khác", key: "OtherProfit" },
        { label: "LN trước thuế", key: "ProfitBeforeTax" },
        { label: "LN sau thuế", key: "ProfitAfterTax" },
        { label: "Lợi nhuận sau thuế công ty mẹ", key: "ParentCompanyShareholderProfitAfterTax" },
        {
          label: "Lợi ích CĐ không kiểm soát",
          getValue: (item) => {
            const f = item.financialValues || {};
            return (f.ProfitAfterTax || 0) - (f.ParentCompanyShareholderProfitAfterTax || 0);
          },
        },
        { label: "Lợi nhuận ròng", key: "ParentCompanyShareholderProfitAfterTax" },
      ],
    },
    {
      title: "Tài sản và VCSH",
      rows: [
        { label: "Tổng nợ", key: "TotalDebt" },
        { label: "Nợ ngắn hạn", key: "TotalShortTermDebt", hideForBank: true },
        { label: "Nợ dài hạn", key: "TotalLongTermDebt", hideForBank: true },
        { label: "Nợ nhóm 1", key: "StandardDebt" },
        { label: "Nợ nhóm 2", key: "WatchlistDebt" },
        { label: "Nợ nhóm 3", key: "SubstandardDebt" },
        { label: "Nợ nhóm 4", key: "DoubtfulDebt" },
        { label: "Nợ nhóm 5", key: "BadDebt" },
        { label: "Vốn chủ sở hữu", key: "TotalStockHolderEquity" },
        { label: "Tổng giá trị tồn kho", key: "TotalInventory", hideForBank: true },
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
        { label: "Lưu chuyển tiền thuần từ HĐ Kinh doanh", key: "CashflowFromOperatingActivity" },
        { label: "Lưu chuyển tiền thuần từ HĐ Tài chính", key: "CashflowFromFinancingActivity" },
        { label: "Lưu chuyển tiền thuần từ HĐ Đầu tư", key: "CashflowFromInvestingActivity" },
        {
          label: "Tiền mặt",
          getValue: (item) => {
            const f = item.financialValues || {};
            return isBank ? f.CashGoldJewelry : f.Cash;
          },
        },
        { label: "Tiền và tương đương tiền cuối kỳ", key: "CashAndCashEquivalentAtTheEndOfPeriod" },
      ],
    },
  ];
}

function renderPreviewTable(data, type, isBank) {
  const sortedData = [...data].sort((a, b) => {
    if (a.year !== b.year) {
      return b.year - a.year;
    }
    return b.quarter - a.quarter;
  });

  const headers = sortedData.map((item) => {
    return type === "quarter" ? `Q${item.quarter}/${item.year}` : `${item.year}`;
  });

  const sections = getPreviewSections(isBank);
  let html = "";

  previewNav.innerHTML = sections
    .map((section) => {
      const sectionId = `preview-${section.title.toLowerCase().replace(/[^a-z0-9]+/gi, "-")}`;
      return `<button type="button" class="button-tertiary" data-target="${sectionId}">${section.title}</button>`;
    })
    .join("");

  for (const section of sections) {
    const sectionId = `preview-${section.title.toLowerCase().replace(/[^a-z0-9]+/gi, "-")}`;
    html += `<div class="preview-section" id="${sectionId}"><div class="preview-title">${section.title}</div><table><thead><tr><th>Chỉ số</th>`;
    for (const header of headers) {
      html += `<th>${header}</th>`;
    }
    html += `</tr></thead><tbody>`;

    for (const metric of section.rows) {
      if (metric.hideForBank && isBank) {
        continue;
      }
      html += `<tr><td>${metric.label}</td>`;
      for (let i = 0; i < sortedData.length; i++) {
        const item = sortedData[i];
        let value = metric.getValue
          ? metric.getValue(item, i, sortedData)
          : item.financialValues?.[metric.key];
      if (metric.scale && Number.isFinite(Number(value))) {
        value = Number(value) / metric.scale;
      }
      const precision = section.title === "Chỉ số cơ bản" ? 3 : 2;
      html += `<td>${formatPreviewValue(value, metric.isPercent, precision)}</td>`;
      }
      html += `</tr>`;
    }

    html += `</tbody></table></div>`;
  }

  previewContainer.innerHTML = html;
}

function formatPreviewValue(value, isPercent, precision = 2) {
  if (value === null || value === undefined) {
    return "";
  }

  const num = Number(value);
  if (Number.isFinite(num)) {
    if (isPercent) {
      const percentValue = Math.abs(num) <= 1 ? num * 100 : num;
      return `${percentValue.toLocaleString("en-US", {
        maximumFractionDigits: precision,
        minimumFractionDigits: precision,
      })}%`;
    }

    return num.toLocaleString("en-US", {
      maximumFractionDigits: precision,
      minimumFractionDigits: precision,
    });
  }

  return String(value);
}

function setStatus(message) {
  statusEl.textContent = message;
}

previewNav.addEventListener("click", (event) => {
  const button = event.target.closest("button[data-target]");
  if (!button) {
    return;
  }

  const targetId = button.dataset.target;
  const targetSection = document.getElementById(targetId);
  if (targetSection) {
    targetSection.scrollIntoView({ behavior: "smooth", block: "start" });
  }
});

async function getAccessTokenFromPage() {
  const tabs = await chrome.tabs.query({ url: ["https://fireant.vn/*", "https://www.fireant.vn/*"] });
  const tab = tabs && tabs[0];

  if (!tab || !tab.id) {
    throw new Error("Vui lòng mở trang https://fireant.vn trong một tab khác trước khi sử dụng extension.");
  }

  const results = await chrome.scripting.executeScript({
    target: { tabId: tab.id },
    func: () => {
      const key = "oidc.user:https://accounts.fireant.vn:fireant.tradestation";
      const raw = localStorage.getItem(key);
      if (!raw) {
        return { error: "Không tìm thấy dữ liệu token trong localStorage." };
      }

      try {
        const parsed = JSON.parse(raw);
        return { token: parsed.access_token || null };
      } catch (err) {
        return { error: "Dữ liệu token không hợp lệ." };
      }
    },
  });

  const result = results?.[0]?.result;
  if (!result) {
    throw new Error("Không thể lấy token từ trang.");
  }
  if (result.error) {
    throw new Error(result.error);
  }
  if (!result.token) {
    throw new Error("Token lấy được rỗng.");
  }

  return result.token;
}

async function fetchFinancialData({ stockID, token, type, count }) {
  const apiType = type === "quarter" ? "Q" : "Y";

  const response = await fetch(
    `https://restv2.fireant.vn/symbols/${stockID}/financial-data?type=${apiType}&count=${count}`,
    {
      headers: {
        accept: "application/json, text/plain, */*",
        authorization: `Bearer ${token}`,
      },
    }
  );

  if (!response.ok) {
    if (response.status === 401) {
      await chrome.tabs.create({ url: "https://fireant.vn/" });
      throw new Error(
        "API trả về 401. Mở tab https://fireant.vn/ để đăng nhập lại và tạo token mới."
      );
    }
    throw new Error(`API Error: ${response.status}`);
  }

  return response.json();
}

function createSheet(workbook, name, sortedData, rowsConfig, isQuarter, options = {}) {
  const sheet = workbook.addWorksheet(name);
  const defaultNumFmt = options.defaultNumFmt || "#,##0.00";
  const defaultPercentFmt = options.percentNumFmt || "0.00%";

  const headers = sortedData.map((item) => {
    if (isQuarter) {
      return `Q${item.quarter}/${item.year}`;
    }

    return item.year;
  });

  sheet.addRow(["Chỉ số", ...headers]);

  rowsConfig.forEach((rowConfig) => {
    const row = [rowConfig.label];

    sortedData.forEach((item, idx) => {
      const rawValue = rowConfig.getValue
        ? rowConfig.getValue(item.financialValues, idx, sortedData)
        : item.financialValues?.[rowConfig.key];
      let value = rawValue;
      const numericValue = Number(rawValue);
      if (Number.isFinite(numericValue)) {
        value = numericValue / (rowConfig.scale || 1);
        if (rowConfig.isPercent) {
          value = Math.abs(value) <= 1 ? value : value / 100;
        }
      }
      row.push(value);
    });

    const addedRow = sheet.addRow(row);
    if (rowConfig.isPercent) {
      addedRow.eachCell((cell, colNumber) => {
        if (colNumber === 1) return;
        if (typeof cell.value === "number") {
          cell.numFmt = rowConfig.numFmt || defaultPercentFmt;
        }
      });
    }
  });

  sheet.columns.forEach((col, index) => {
    col.width = index === 0 ? 40 : 20;
  });

  sheet.getRow(1).font = {
    bold: true,
  };

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

async function generateExcel(data, stockID, type, isBank) {
  const workbook = new ExcelJS.Workbook();

  const sortedData = [...data].sort((a, b) => {
    if (a.year !== b.year) {
      return b.year - a.year;
    }

    return b.quarter - a.quarter;
  });

  const isQuarter = type === "quarter";

  const basicRows = [
    {
      label: "EPS",
      getValue: (f) => f.BasicEPS,
    },
    {
      label: "BVPS",
      getValue: (f) => f.BookValuePerShare,
    },
    {
      label: "P/E",
      getValue: (f) => f.PE,
    },
    {
      label: "P/B",
      getValue: (f) => f.PB,
    },
    {
      label: "ROA",
      getValue: (f) => f.ROA,
      isPercent: true,
    },
    {
      label: "ROE",
      getValue: (f) => f.ROE,
      isPercent: true,
    },
    {
      label: "ROIC",
      getValue: (f) => f.ROIC,
      isPercent: true,
    },
    { label: "CAGR Doanh thu 3 năm", getValue: (f) => f.SaleGrowth_03Yr },
    {
      label: "CAGR LNST 3 năm",
      getValue: (f) => f.ProfitAfterTaxGrowth_03Yr ?? f.ProfitGrowth_03Yr,
    },
    { label: "CAGR EPS 3 năm", getValue: (f) => f.BasicEPSGrowth_03Yr },
    { label: "CAGR Tài sản 3 năm", getValue: (f) => f.TotalAssetGrowth_03Yr },
    { label: "CAGR Vốn chủ 3 năm", getValue: (f) => f.EquityGrowth_03Yr },
    {
      label: "Tăng trưởng tín dụng",
      getValue: (f, idx, all) => {
        const currentLoan = Number(f?.CustomerLoan ?? f?.CustomerLoans ?? 0);
        const prev = all && all[idx + 1];
        const prevLoan = Number(prev?.financialValues?.CustomerLoan ?? prev?.financialValues?.CustomerLoans ?? 0);
        if (!prevLoan) return null;
        return (currentLoan - prevLoan) / prevLoan;
      },
    },
    {
      label: "Nợ vay/VCSH",
      getValue: (f) => {
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
  }
  else {
    basicRows.push(
      { label: "Thanh toán nhanh", getValue: (f) => f.QuickRatio },
      { label: "Thanh toán hiện hành", getValue: (f) => f.CurrentRatio },
      { label: "Tỷ lệ tiền mặt", getValue: (f) => f.CashRatio }
    );
  }

  createSheet(
    workbook,
    "Chỉ số cơ bản",
    sortedData,
    basicRows,
    isQuarter,
    {
      defaultNumFmt: "#,##0.000",
      percentNumFmt: "0.000%",
    }
  );

  // Build Kết quả kinh doanh rows, hiding CostOfGoodSold for banks
  const kqRows = [
    { label: "Doanh thu thuần", getValue: (f) => f.NetSale },
  ];
  if (!isBank) {
    kqRows.push({ label: "Giá vốn hàng bán", getValue: (f) => f.CostOfGoodSold });
  }
  kqRows.push(
    { label: "Lợi nhuận gộp", getValue: (f) => f.GrossProfit },
    {
      label: "LN HĐTC & Cty LDLK",
      getValue: (f) => (f.ProfitFromFinancialActivity || 0) + (f.ProfitFromAssociate || 0),
    },
    { label: "LN khác", getValue: (f) => f.OtherProfit },
    { label: "LN trước thuế", getValue: (f) => f.ProfitBeforeTax },
    { label: "LN sau thuế", getValue: (f) => f.ProfitAfterTax },
    { label: "Lợi nhuận sau thuế công ty mẹ", getValue: (f) => f.ParentCompanyShareholderProfitAfterTax },
    {
      label: "Lợi ích CĐ không kiểm soát",
      getValue: (f) => (f.ProfitAfterTax || 0) - (f.ParentCompanyShareholderProfitAfterTax || 0),
    },
    { label: "Lợi nhuận ròng", getValue: (f) => f.ParentCompanyShareholderProfitAfterTax }
  );

  createSheet(workbook, "Kết quả kinh doanh", sortedData, kqRows, isQuarter);

  // Asset rows; exclude inventory and short/long term debt for banks
  const assetRows = [];
  assetRows.push({ label: "Tổng nợ", getValue: (f) => f.TotalDebt });
  if (!isBank) {
    assetRows.push({ label: "Nợ ngắn hạn", getValue: (f) => f.TotalShortTermDebt });
    assetRows.push({ label: "Nợ dài hạn", getValue: (f) => f.TotalLongTermDebt });
  }
  assetRows.push(
    { label: "Nợ nhóm 1", getValue: (f) => f.StandardDebt },
    { label: "Nợ nhóm 2", getValue: (f) => f.WatchlistDebt },
    { label: "Nợ nhóm 3", getValue: (f) => f.SubstandardDebt },
    { label: "Nợ nhóm 4", getValue: (f) => f.DoubtfulDebt },
    { label: "Nợ nhóm 5", getValue: (f) => f.BadDebt }
  );

  assetRows.push({ label: "Vốn chủ sở hữu", getValue: (f) => f.TotalStockHolderEquity });
  if (!isBank) {
    assetRows.push({ label: "Tổng giá trị tồn kho", getValue: (f) => f.TotalInventory });
    assetRows.push({ label: "Tồn kho nguyên vật liệu", getValue: () => null });
    assetRows.push({ label: "Công cụ, dụng cụ", getValue: () => null });
    assetRows.push({ label: "Chi phí SXKD dở dang", getValue: () => null });
    assetRows.push({ label: "Thành phẩm", getValue: () => null });
    assetRows.push({ label: "Hàng hoá", getValue: () => null });
    assetRows.push({ label: "Hàng gửi bán", getValue: () => null });
  }

  createSheet(workbook, "Tài sản và VCSH", sortedData, assetRows, isQuarter);

  createSheet(
    workbook,
    "Lưu chuyển tiền tệ",
    sortedData,
    [
      { label: "Lưu chuyển tiền thuần từ HĐ Kinh doanh", getValue: (f) => f.CashflowFromOperatingActivity },
      { label: "Lưu chuyển tiền thuần từ HĐ Tài chính", getValue: (f) => f.CashflowFromFinancingActivity },
      { label: "Lưu chuyển tiền thuần từ HĐ Đầu tư", getValue: (f) => f.CashflowFromInvestingActivity },
      { label: "Tiền mặt", getValue: (f) => (isBank ? f.CashGoldJewelry : f.Cash) },
      { label: "Tiền và tương đương tiền cuối kỳ", getValue: (f) => f.CashAndCashEquivalentAtTheEndOfPeriod },
    ],
    isQuarter
  );

  const buffer = await workbook.xlsx.writeBuffer();

  const blob = new Blob([
    buffer,
  ], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });

  const url = URL.createObjectURL(blob);

  const fileName = `[${stockID}] Phân tích cơ bản.xlsx`;

  chrome.downloads.download({
    url,
    filename: fileName,
    saveAs: true,
  });
}
