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
    { label: "ROA", key: "ROA" },
    { label: "ROE", key: "ROE" },
    { label: "ROIC", key: "ROIC" },
    {
      label: "Nợ vay/VCSH",
      getValue: (item) => {
        const f = item.financialValues || {};
        const debt = (f.ShortTermInterestBearingDebt || 0) + (f.LongTermInterestBearingDebt || 0);
        const equity = f.StockHolderEquity || 1;
        return debt / equity;
      },
    },
  ]; 

  if (isBank) {
    basicRows.push(
      { label: "NII", key: "NetInterestIncome" },
      { label: "NIM", key: "LoanlossReservesToNPL" },
      { label: "CIR", key: "CIR" },
      { label: "NPL", key: "NPLToLoan" }
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
        { label: "Giá vốn hàng bán", key: "CostOfGoodSold" },
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
        { label: "Nợ ngắn hạn", key: "TotalShortTermDebt" },
        { label: "Nợ dài hạn", key: "TotalLongTermDebt" },
        { label: "Vốn chủ sở hữu", key: "StockHolderEquity" },
        { label: "Tổng giá trị tồn kho", key: "TotalInventory" },
        { label: "Tồn kho nguyên vật liệu", getValue: () => null },
        { label: "Công cụ, dụng cụ", getValue: () => null },
        { label: "Chi phí SXKD dở dang", getValue: () => null },
        { label: "Thành phẩm", getValue: () => null },
        { label: "Hàng hoá", getValue: () => null },
        { label: "Hàng gửi bán", getValue: () => null },
      ],
    },
    {
      title: "Lưu chuyển tiền tệ",
      rows: [
        { label: "Lưu chuyển tiền thuần từ HĐ Kinh doanh", key: "CashflowFromOperatingActivity" },
        { label: "Lưu chuyển tiền thuần từ HĐ Tài chính", key: "CashflowFromFinancingActivity" },
        { label: "Lưu chuyển tiền thuần từ HĐ Đầu tư", key: "CashflowFromInvestingActivity" },
        { label: "Tiền mặt", key: "Cash" },
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
      html += `<tr><td>${metric.label}</td>`;
      for (const item of sortedData) {
        const value = metric.getValue
          ? metric.getValue(item)
          : item.financialValues?.[metric.key];
        html += `<td>${formatPreviewValue(value)}</td>`;
      }
      html += `</tr>`;
    }

    html += `</tbody></table></div>`;
  }

  previewContainer.innerHTML = html;
}

function formatPreviewValue(value) {
  if (value === null || value === undefined) {
    return "";
  }

  const num = Number(value);
  if (Number.isFinite(num)) {
    return num.toLocaleString("en-US", {
      maximumFractionDigits: 2,
      minimumFractionDigits: 0,
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

function createSheet(workbook, name, sortedData, rowsConfig, isQuarter) {
  const sheet = workbook.addWorksheet(name);

  const headers = sortedData.map((item) => {
    if (isQuarter) {
      return `Q${item.quarter}/${item.year}`;
    }

    return item.year;
  });

  sheet.addRow(["Chỉ số", ...headers]);

  rowsConfig.forEach((rowConfig) => {
    const row = [rowConfig.label];

    sortedData.forEach((item) => {
      row.push(rowConfig.getValue(item.financialValues));
    });

    sheet.addRow(row);
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

      if (typeof cell.value === "number") {
        cell.numFmt = "#,##0.0#";
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
    },
    {
      label: "ROE",
      getValue: (f) => f.ROE,
    },
    {
      label: "ROIC",
      getValue: (f) => f.ROIC,
    },
    {
      label: "Nợ vay/VCSH",
      getValue: (f) => {
        const debt =
          (f.ShortTermInterestBearingDebt || 0) +
          (f.LongTermInterestBearingDebt || 0);

        return debt / (f.StockHolderEquity || 1);
      },
    },
  ]; 

  if (isBank) {
    basicRows.push(
      { label: "NII", getValue: (f) => f.NetInterestIncome },
      { label: "NIM", getValue: (f) => f.LoanlossReservesToNPL },
      { label: "CIR", getValue: (f) => f.CIR },
      { label: "NPL", getValue: (f) => f.NPLToLoan }
    );
  }

  createSheet(
    workbook,
    "Chỉ số cơ bản",
    sortedData,
    basicRows,
    isQuarter
  );

  createSheet(
    workbook,
    "Kết quả kinh doanh",
    sortedData,
    [
      {
        label: "Doanh thu thuần",
        getValue: (f) => f.NetSale,
      },
      {
        label: "Giá vốn hàng bán",
        getValue: (f) => f.CostOfGoodSold,
      },
      {
        label: "Lợi nhuận gộp",
        getValue: (f) => f.GrossProfit,
      },
      {
        label: "LN HĐTC & Cty LDLK",
        getValue: (f) =>
          (f.ProfitFromFinancialActivity || 0) +
          (f.ProfitFromAssociate || 0),
      },
      {
        label: "LN khác",
        getValue: (f) => f.OtherProfit,
      },
      {
        label: "LN trước thuế",
        getValue: (f) => f.ProfitBeforeTax,
      },
      {
        label: "LN sau thuế",
        getValue: (f) => f.ProfitAfterTax,
      },
      {
        label: "Lợi nhuận sau thuế công ty mẹ",
        getValue: (f) => f.ParentCompanyShareholderProfitAfterTax,
      },
      {
        label: "Lợi ích CĐ không kiểm soát",
        getValue: (f) =>
          (f.ProfitAfterTax || 0) -
          (f.ParentCompanyShareholderProfitAfterTax || 0),
      },
      {
        label: "Lợi nhuận ròng",
        getValue: (f) => f.ParentCompanyShareholderProfitAfterTax,
      },
    ],
    isQuarter
  );

  createSheet(
    workbook,
    "Tài sản và VCSH",
    sortedData,
    [
      {
        label: "Tổng nợ",
        getValue: (f) => f.TotalDebt,
      },
      {
        label: "Nợ ngắn hạn",
        getValue: (f) => f.TotalShortTermDebt,
      },
      {
        label: "Nợ dài hạn",
        getValue: (f) => f.TotalLongTermDebt,
      },
      {
        label: "Vốn chủ sở hữu",
        getValue: (f) => f.StockHolderEquity,
      },
      {
        label: "Tổng giá trị tồn kho",
        getValue: (f) => f.TotalInventory,
      },
      {
        label: "Tồn kho nguyên vật liệu",
        getValue: () => null,
      },
      {
        label: "Công cụ, dụng cụ",
        getValue: () => null,
      },
      {
        label: "Chi phí SXKD dở dang",
        getValue: () => null,
      },
      {
        label: "Thành phẩm",
        getValue: () => null,
      },
      {
        label: "Hàng hoá",
        getValue: () => null,
      },
      {
        label: "Hàng gửi bán",
        getValue: () => null,
      },
    ],
    isQuarter
  );

  createSheet(
    workbook,
    "Lưu chuyển tiền tệ",
    sortedData,
    [
      {
        label: "Lưu chuyển tiền thuần từ HĐ Kinh doanh",
        getValue: (f) => f.CashflowFromOperatingActivity,
      },
      {
        label: "Lưu chuyển tiền thuần từ HĐ Tài chính",
        getValue: (f) => f.CashflowFromFinancingActivity,
      },
      {
        label: "Lưu chuyển tiền thuần từ HĐ Đầu tư",
        getValue: (f) => f.CashflowFromInvestingActivity,
      },
      {
        label: "Tiền mặt",
        getValue: (f) => f.Cash,
      },
      {
        label: "Tiền và tương đương tiền cuối kỳ",
        getValue: (f) => f.CashAndCashEquivalentAtTheEndOfPeriod,
      },
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
