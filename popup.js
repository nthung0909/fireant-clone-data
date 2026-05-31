const stockIDInput = document.getElementById("stockID");
const typeInput = document.getElementById("type");
const countInput = document.getElementById("count");
const submitBtn = document.getElementById("submitBtn");
const statusEl = document.getElementById("status");

submitBtn.addEventListener("click", async () => {
  try {
    const stockID = stockIDInput.value.trim().toUpperCase();
    const type = typeInput.value;
    const count = Number(countInput.value);

    if (!stockID) {
      throw new Error("Thiếu StockID");
    }

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

    setStatus("Đang tạo excel...");

    await generateExcel(data, stockID, type);

    setStatus("Hoàn tất");
  } catch (err) {
    console.error(err);
    setStatus(err.message);
  }
});

function setStatus(message) {
  statusEl.textContent = message;
}

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
        cell.numFmt = "#,##0.00";
      }
    });
  });
}

async function generateExcel(data, stockID, type) {
  const workbook = new ExcelJS.Workbook();

  const sortedData = [...data].sort((a, b) => {
    if (a.year !== b.year) {
      return a.year - b.year;
    }

    return a.quarter - b.quarter;
  });

  const isQuarter = type === "quarter";

  createSheet(
    workbook,
    "Chỉ số cơ bản",
    sortedData,
    [
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
    ],
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

  const fileName = `${stockID}-${type}.xlsx`;

  chrome.downloads.download({
    url,
    filename: fileName,
    saveAs: true,
  });
}
