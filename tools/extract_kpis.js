// Prints the platform's built-in KPI list (after the layout, renames and standing targets are applied) and the
// template's master data as JSON. Used by build_excel_template.py. Usage: node tools/extract_kpis.js [year]
const fs = require("fs"), path = require("path"), vm = require("vm");
const html = fs.readFileSync(path.join(__dirname, "..", "index.html"), "utf8");
const start = html.indexOf("<script>\nconst STORAGE");
const end = html.lastIndexOf("init();");
const code = html.slice(start + "<script>".length, end);
const el = () => new Proxy(function(){}, { get: (t, k) => k === Symbol.toPrimitive ? () => "" : el(), apply: () => el(), set: () => true });
const sandbox = { window: { addEventListener(){} }, document: el(), localStorage: { getItem: () => null, setItem(){} }, navigator: {}, console, setTimeout, clearTimeout, setInterval, Intl, Date, Math, JSON };
vm.createContext(sandbox);
vm.runInContext(code, sandbox, { filename: "index.html" });
const year = Number(process.argv[2]) || 2026;
// Everything the template generator needs comes from index.html, so the workbook, the import screen and setup.sql's
// master data seed never disagree: KPIs (with their calculation type, allowed range and standing target), master data,
// calculation types, direction codes and validation codes.
const out = vm.runInContext(`(() => {
    const list = defaultKpis.map(k => ({ ...k })); addMissingKpis(list); applyDirectionRules(list); applyKpiLayout(list, []);
    list.forEach(k => { k.calcType = calcTypeOf(k); k.range = validRange(k); k.templateTarget = effectiveTarget(k, ${year}); });
    return JSON.stringify({
        kpis: list, categories: CATEGORY_TREE, derived: Object.keys(DERIVED_KPIS),
        templateVersion: TEMPLATE_VERSION, templateFile: TEMPLATE_FILE, dataStartDate: DATA_START_DATE,
        facilities: MASTER_FACILITIES, departments: MASTER_DEPARTMENTS, masterCategories: MASTER_CATEGORIES,
        calcTypes: CALC_TYPES, arabicLabels: KPI_ARABIC_LABELS, periodTypes: PERIOD_TYPES, directionCodes: DIRECTION_CODES, importCodes: IMPORT_CODES
    });
})()`, sandbox);
process.stdout.write(out);
