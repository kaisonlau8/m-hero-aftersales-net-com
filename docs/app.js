const EXCEL_URL = new URL(
  document.currentScript?.dataset.excelUrl || "./售后网络.xlsx",
  window.location.href,
);

function clean(value) {
  return String(value ?? "").trim();
}

function htmlEscape(value) {
  return clean(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

function groupBy(rows, selector) {
  const map = new Map();
  for (const row of rows) {
    const key = selector(row) || "未填写";
    if (!map.has(key)) map.set(key, []);
    map.get(key).push(row);
  }
  return map;
}

function sortZhEntries(entries) {
  return [...entries].sort((a, b) => String(a[0]).localeCompare(String(b[0]), "zh-Hans-CN"));
}

function parseRows(workbook) {
  const firstSheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[firstSheetName];
  const table = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
  const header = (table[0] || []).map(clean);
  const regionIndex = header.indexOf("区域");
  const storeIndex = header.indexOf("门店名称");
  const provinceIndex = header.indexOf("省份");
  const cityIndex = header.indexOf("城市");
  const districtIndex = header.indexOf("区");

  if ([regionIndex, storeIndex, provinceIndex, cityIndex, districtIndex].some((index) => index < 0)) {
    throw new Error("Excel 表头必须包含：区域、门店名称、省份、城市、区");
  }

  return table.slice(1).map((row, index) => ({
    rowNo: index + 2,
    region: clean(row[regionIndex]),
    store: clean(row[storeIndex]),
    province: clean(row[provinceIndex]),
    city: clean(row[cityIndex]),
    district: clean(row[districtIndex]),
  })).filter((row) => row.store);
}

function navLinks(regions) {
  return regions.map(([region, regionRows]) => `
    <a href="#${encodeURIComponent(region)}">
      <span>${htmlEscape(region)}</span>
      <strong>${regionRows.length}</strong>
    </a>
  `).join("");
}

function regionSections(regions) {
  return regions.map(([region, regionRows]) => {
    const provinceMap = groupBy(regionRows, (row) => row.province);
    const provinceHtml = sortZhEntries(provinceMap.entries()).map(([province, provinceRows]) => {
      const cityMap = groupBy(provinceRows, (row) => row.city);
      const cityHtml = sortZhEntries(cityMap.entries()).map(([city, cityRows]) => {
        const districtMap = groupBy(cityRows, (row) => row.district);
        const districts = sortZhEntries(districtMap.entries()).map(([district, districtRows]) => `
          <div class="district-pill">
            <span>${htmlEscape(district)}</span>
            <strong>${districtRows.length}</strong>
          </div>
        `).join("");
        const stores = cityRows
          .slice()
          .sort((a, b) => a.district.localeCompare(b.district, "zh-Hans-CN") || a.store.localeCompare(b.store, "zh-Hans-CN"))
          .map((row) => `<li><span>${htmlEscape(row.district)}</span>${htmlEscape(row.store)}</li>`)
          .join("");
        return `
          <article class="city-block">
            <header>
              <div>
                <h4>${htmlEscape(city)}</h4>
                <p>${new Set(cityRows.map((row) => row.district)).size} 个区县</p>
              </div>
              <strong>${cityRows.length}</strong>
            </header>
            <div class="district-grid">${districts}</div>
            <ul class="store-list">${stores}</ul>
          </article>
        `;
      }).join("");
      return `
        <section class="province-block">
          <div class="province-title">
            <h3>${htmlEscape(province)}</h3>
            <span>${provinceRows.length} 个网点</span>
          </div>
          <div class="city-grid">${cityHtml}</div>
        </section>
      `;
    }).join("");
    return `
      <section class="region-section" id="${encodeURIComponent(region)}">
        <div class="region-heading">
          <div>
            <p>区域</p>
            <h2>${htmlEscape(region)}</h2>
          </div>
          <strong>${regionRows.length}</strong>
        </div>
        ${provinceHtml}
      </section>
    `;
  }).join("");
}

function render(rows, lastModified) {
  const regions = [...groupBy(rows, (row) => row.region).entries()];

  document.querySelector("#sideNav").innerHTML = navLinks(regions);
  document.querySelector("#regions").innerHTML = regionSections(regions);
}

async function loadWorkbook() {
  const target = new URL(EXCEL_URL.href);
  target.searchParams.set("v", Date.now().toString());
  const response = await fetch(target, { cache: "no-store" });
  if (!response.ok) {
    throw new Error(`读取 Excel 失败：HTTP ${response.status}`);
  }
  const lastModified = response.headers.get("last-modified");
  const buffer = await response.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array" });
  return { rows: parseRows(workbook), lastModified };
}

async function init() {
  try {
    const { rows, lastModified } = await loadWorkbook();
    render(rows, lastModified);
  } catch (error) {
    document.querySelector("#regions").innerHTML = `
      <div class="error-panel">
        <strong>无法自动读取 Excel。</strong><br>
        请通过本地 HTTP 服务打开本页面，并确认 <code>售后网络.xlsx</code> 位于项目根目录。<br>
        错误信息：${htmlEscape(error.message)}
      </div>
    `;
  }
}

init();
