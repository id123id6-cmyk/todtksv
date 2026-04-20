(function () {
  "use strict";

  const SHEET_TARGET = "발주서 입력";
  const SHEET_PAST_ORDER = "지난발주서";
  const SHEET_STOCK = "재고";
  /** 한 품목당 구분 행: 지난발주(지난발주서 시트) · 차이수량 · 발주량(발주서 입력) · 부족품 */
  const SUB_ROWS = ["지난발주", "차이수량", "발주량", "부족품"];

  const DATE_KEYS = ["날짜", "일자", "date", "발주일", "납기일", "주문일", "생산일", "작업일자", "작업일"];
  /** '품목' 제외: '품목코드'가 제품명으로 오인됨. '제품'은 아래 guessColumnIndex에서 정확히 일치할 때만 사용 */
  const NAME_KEYS = ["제품명", "품명", "품목명", "상품명", "product", "item", "자재명"];
  /** 재고_단가 전용: `item` 키가 "Item Code"·itemcode 열에 잘못 붙는 것을 막음 */
  const STOCK_UNIT_PRICE_NAME_KEYS = ["품목명", "제품명", "품명", "상품명", "자재명"];
  const CODE_KEYS = ["품목코드", "코드", "자재코드", "itemcode", "item_code", "품번", "상품코드"];
  const TYPE_KEYS = ["타입", "type", "category", "종류", "라인", "공정"];
  const QTY_KEYS = ["발주량", "수량", "qty", "quantity", "주문량", "요청량", "생산량", "계획량"];
  const STOCK_KEYS = ["재고"];
  const EXPORT_KEYS = ["수출발주량", "수출"];
  const LACK_KEYS = ["부족량", "부족"];
  const GUBUN_KEYS = ["구분"];
  /** 재고 전용 엑셀 — 카테고리·분류 열 (일별 통합표 「타입」과 별도) */
  const STOCK_CATEGORY_KEYS = ["카테고리", "category", "분류", "대분류", "제품군", "상품군", "구분", "종류"];

  /** 드로잉팀 — 타입 열 값 기준(부분 일치). 이중관·화승이중관 구분을 위해 긴 키워드 우선 */
  const DRAWING_TEAM_SEGMENTS = ["드로잉", "페럴", "머플러", "이중관", "화승이중관"];

  /**
   * 참조 시트(팀별 색 구간·개당단가) 기준 기본 개당 단가(원).
   * 재고_단가 엑셀에 해당 품목(코드+명 또는 명만) 행이 없을 때만 일별 통합표 재고 금액에 사용합니다.
   * 값은 구간별 대표·가중 평균(드로잉 150·130 혼합, 페럴 120·100 혼합 등)에 맞춘 정수입니다.
   */
  const BUILTIN_STOCK_UNIT_PRICE_WON_BY_TEAM = Object.freeze({
    드로잉: 136,
    페럴: 108,
    이중관: 250,
    화승이중관: 290,
    머플러: 325,
  });
  /**
   * 고정 단가 마스터(`단가.xlsx`)에서 추출한 품목명→개당단가.
   * index.html에서 `assets/default-unit-prices.js`를 먼저 로드하면 window.DEFAULT_UNIT_PRICE_ROWS에 주입된다.
   */
  const DEFAULT_UNIT_PRICE_BY_NORM_NAME = (() => {
    /** @type {Map<string, number>} */
    const m = new Map();
    const rows =
      typeof window !== "undefined" && Array.isArray(window.DEFAULT_UNIT_PRICE_ROWS) ? window.DEFAULT_UNIT_PRICE_ROWS : [];
    for (const r of rows) {
      const name = String(r?.name ?? "").trim();
      const unit = parseUnitPrice(r?.unit);
      if (!name || !Number.isFinite(unit)) continue;
      const k = norm(name);
      if (!k || m.has(k)) continue;
      m.set(k, unit);
    }
    return m;
  })();

  /** 사내 드로잉 양식: 엑셀 Q~T열(0-based 16~19) 정지 — 교환/수리/소재/계획정지, U열(20) 공장불량 */
  const DRAWING_STOP_COL_QRST = [16, 17, 18, 19];
  const DRAWING_STOP_KIND_LABELS = ["교환", "수리", "소재", "계획정지"];
  const DRAWING_FACTORY_DEFECT_COL_U = 20;

  /** 페럴 시트: 정지 T~W(0-based 19~22), 불량 X+AB(23,27) */
  const PARALLEL_STOP_COL_TUVW = [19, 20, 21, 22];
  const PARALLEL_DEFECT_COL_X_AB = [23, 27];

  /** 이중관 시트: 비가동 S~V열, 불량 W+AJ열, 생산 N열(0-based 13) */
  const DOUBLE_PIPE_STOP_COL_STUV = [18, 19, 20, 21];
  const DOUBLE_PIPE_DEFECT_COL_W_AJ = [22, 35];
  const DOUBLE_PIPE_PROD_COL_N = 13;
  const DOUBLE_PIPE_STOP_LABELS = ["비가동(S)", "비가동(T)", "비가동(U)", "비가동(V)"];

  /** 화승이중관: 생산 L열, 비가동 R~V열, 소재불량 X+Y+Z(합=W), 공정불량 AK+AL+AR(합=AJ) */
  const HWASEUNG_PROD_COL_L = 11;
  const HWASEUNG_STOP_COL_RSTUV = [17, 18, 19, 20, 21];
  const HWASEUNG_DEFECT_MATERIAL_XYZ = [23, 24, 25];
  const HWASEUNG_DEFECT_PROCESS_AK_AL_AR = [36, 37, 43];
  const HWASEUNG_STOP_LABELS = ["비가동(R)", "비가동(S)", "비가동(T)", "비가동(U)", "비가동(V)"];
  /** 화승이중관 공정명(집계·필터 기준, 엑셀 공정 열 값을 이 이름으로 정규화) */
  const HWASEUNG_DOUBLE_PIPE_PROCESS_KEYS = [
    "전조",
    "스파이럴",
    "피어싱, 확관",
    "수동 롤링",
    "자동롤링",
    "압착",
  ];
  const HWASEUNG_DOUBLE_PIPE_ALL_PROC_KEYS = [...HWASEUNG_DOUBLE_PIPE_PROCESS_KEYS, "기타"];

  /** 머플러·단조머플러 공통: M열 생산, R~U 비가동(교환·수리·소재·계획정지), 불량=V+AH~AL(소재+공정 세부 합) */
  const MUFFLER_PROD_COL_M = 12;
  const MUFFLER_STOP_COL_RSTU = [17, 18, 19, 20];
  const MUFFLER_DEFECT_COL_V = 21;
  const MUFFLER_DEFECT_PROC_DETAIL_COLS_AH_AL = [33, 34, 35, 36, 37];

  const fileInput = document.getElementById("fileInput");
  const dropzone = document.getElementById("dropzone");
  const fileNameEl = document.getElementById("fileName");
  const sheetInfoEl = document.getElementById("sheetInfo");
  const stockFileInput = document.getElementById("stockFileInput");
  const stockFileNameEl = document.getElementById("stockFileName");
  const stockSheetInfoEl = document.getElementById("stockSheetInfo");
  const mappingBlock = document.getElementById("mappingBlock");
  const colDate = document.getElementById("colDate");
  const colName = document.getElementById("colName");
  const colCode = document.getElementById("colCode");
  const colType = document.getElementById("colType");
  const colQty = document.getElementById("colQty");
  const colStock = document.getElementById("colStock");
  const btnApply = document.getElementById("btnApply");
  const tableWrap = document.getElementById("tableWrap");
  const emptyState = document.getElementById("emptyState");
  const rowCountEl = document.getElementById("rowCount");
  const btnExport = document.getElementById("btnExport");
  const simpleTableWrap = document.getElementById("simpleTableWrap");
  const emptyStateSimple = document.getElementById("emptyStateSimple");
  const summaryContent = document.getElementById("summaryContent");
  const mainWrap = document.querySelector(".main-wrap");
  const kmStripEl = document.querySelector(".km-strip");
  const viewHome = document.getElementById("viewHome");
  const viewUpload = document.getElementById("viewUpload");
  const viewTeamProductionUpload = document.getElementById("viewTeamProductionUpload");
  const teamProdDrawingDropzone = document.getElementById("teamProdDrawingDropzone");
  const teamProdDrawingFileInput = document.getElementById("teamProdDrawingFileInput");
  const teamProdParallelDropzone = document.getElementById("teamProdParallelDropzone");
  const teamProdParallelFileInput = document.getElementById("teamProdParallelFileInput");
  const teamProdMufflerDropzone = document.getElementById("teamProdMufflerDropzone");
  const teamProdMufflerFileInput = document.getElementById("teamProdMufflerFileInput");
  const teamProdDoublePipeDropzone = document.getElementById("teamProdDoublePipeDropzone");
  const teamProdDoublePipeFileInput = document.getElementById("teamProdDoublePipeFileInput");
  const teamProdHwaseungDropzone = document.getElementById("teamProdHwaseungDropzone");
  const teamProdHwaseungFileInput = document.getElementById("teamProdHwaseungFileInput");
  const viewDaily = document.getElementById("viewDaily");
  const viewOrderCalendar = document.getElementById("viewOrderCalendar");
  const viewStockTable = document.getElementById("viewStockTable");
  const stockTableWrap = document.getElementById("stockTableWrap");
  const stockTableEmpty = document.getElementById("stockTableEmpty");
  const stockTablePanel = document.getElementById("stockTablePanel");
  const stockTableFilterBar = document.getElementById("stockTableFilterBar");
  const stockFilterFieldCategory = document.getElementById("stockFilterFieldCategory");
  const btnFilterStockCategory = document.getElementById("btnFilterStockCategory");
  const panelFilterStockCategory = document.getElementById("panelFilterStockCategory");
  const searchFilterStockCategory = document.getElementById("searchFilterStockCategory");
  const listFilterStockCategory = document.getElementById("listFilterStockCategory");
  const btnFilterStockCode = document.getElementById("btnFilterStockCode");
  const panelFilterStockCode = document.getElementById("panelFilterStockCode");
  const searchFilterStockCode = document.getElementById("searchFilterStockCode");
  const listFilterStockCode = document.getElementById("listFilterStockCode");
  const btnStockCodeSelectBySearch = document.getElementById("btnStockCodeSelectBySearch");
  const btnStockCodeClearBySearch = document.getElementById("btnStockCodeClearBySearch");
  const btnFilterStockName = document.getElementById("btnFilterStockName");
  const panelFilterStockName = document.getElementById("panelFilterStockName");
  const searchFilterStockName = document.getElementById("searchFilterStockName");
  const listFilterStockName = document.getElementById("listFilterStockName");
  const btnStockNameSelectBySearch = document.getElementById("btnStockNameSelectBySearch");
  const btnStockNameClearBySearch = document.getElementById("btnStockNameClearBySearch");
  const btnFilterStockQty = document.getElementById("btnFilterStockQty");
  const panelFilterStockQty = document.getElementById("panelFilterStockQty");
  const searchFilterStockQty = document.getElementById("searchFilterStockQty");
  const listFilterStockQty = document.getElementById("listFilterStockQty");
  const stockTableFilterReset = document.getElementById("stockTableFilterReset");
  const dailyFiltersSlotDaily = document.getElementById("dailyFiltersSlotDaily");
  const dailyFiltersSlotCalendar = document.getElementById("dailyFiltersSlotCalendar");
  const orderCalTeamHint = document.getElementById("orderCalTeamHint");
  const viewSimple = document.getElementById("viewSimple");
  const viewSummary = document.getElementById("viewSummary");
  const viewDrawingLog = document.getElementById("viewDrawingLog");
  const viewParallelLog = document.getElementById("viewParallelLog");
  const viewDoublePipeLog = document.getElementById("viewDoublePipeLog");
  const viewHwaseungDoublePipeLog = document.getElementById("viewHwaseungDoublePipeLog");
  const parallelLogEmpty = document.getElementById("parallelLogEmpty");
  const parallelLogContent = document.getElementById("parallelLogContent");
  const parallelLogMeta = document.getElementById("parallelLogMeta");
  const kpiUtilParallel = document.getElementById("kpiUtilParallel");
  const kpiOeeParallel = document.getElementById("kpiOeeParallel");
  const kpiDefectParallel = document.getElementById("kpiDefectParallel");
  const kpiDefectLabelParallel = document.getElementById("kpiDefectLabelParallel");
  const parallelLogTablesWrap = document.getElementById("parallelLogTablesWrap");
  const parallelLogProdWrap = document.getElementById("parallelLogProdWrap");
  const parallelLogMaintWrap = document.getElementById("parallelLogMaintWrap");
  const parallelLogOps = document.getElementById("parallelLogOps");
  const parallelLogTimeline = document.getElementById("parallelLogTimeline");
  const parallelLogSlicerClear = document.getElementById("parallelLogSlicerClear");
  const parallelLogGranularityEl = document.getElementById("parallelLogGranularity");
  const parallelLogSlicerSelection = document.getElementById("parallelLogSlicerSelection");
  const parallelLogYearLabel = document.getElementById("parallelLogYearLabel");
  const parallelLogYearPrev = document.getElementById("parallelLogYearPrev");
  const parallelLogYearNext = document.getElementById("parallelLogYearNext");
  const parallelLogMonthStrip = document.getElementById("parallelLogMonthStrip");
  const parallelLogDayStrip = document.getElementById("parallelLogDayStrip");
  const parallelLogFilterCutEl = document.getElementById("parallelLogFilterCut");
  const parallelLogFilterBendEl = document.getElementById("parallelLogFilterBend");
  const parallelLogKpiScopeEl = document.getElementById("parallelLogKpiScope");
  const doublePipeLogEmpty = document.getElementById("doublePipeLogEmpty");
  const doublePipeLogContent = document.getElementById("doublePipeLogContent");
  const doublePipeLogMeta = document.getElementById("doublePipeLogMeta");
  const kpiUtilDoublePipe = document.getElementById("kpiUtilDoublePipe");
  const kpiOeeDoublePipe = document.getElementById("kpiOeeDoublePipe");
  const kpiDefectDoublePipe = document.getElementById("kpiDefectDoublePipe");
  const kpiDefectLabelDoublePipe = document.getElementById("kpiDefectLabelDoublePipe");
  const doublePipeLogTablesWrap = document.getElementById("doublePipeLogTablesWrap");
  const doublePipeLogProdWrap = document.getElementById("doublePipeLogProdWrap");
  const doublePipeLogMaintWrap = document.getElementById("doublePipeLogMaintWrap");
  const doublePipeLogOps = document.getElementById("doublePipeLogOps");
  const doublePipeLogTimeline = document.getElementById("doublePipeLogTimeline");
  const doublePipeLogSlicerClear = document.getElementById("doublePipeLogSlicerClear");
  const doublePipeLogGranularityEl = document.getElementById("doublePipeLogGranularity");
  const doublePipeLogSlicerSelection = document.getElementById("doublePipeLogSlicerSelection");
  const doublePipeLogYearLabel = document.getElementById("doublePipeLogYearLabel");
  const doublePipeLogYearPrev = document.getElementById("doublePipeLogYearPrev");
  const doublePipeLogYearNext = document.getElementById("doublePipeLogYearNext");
  const doublePipeLogMonthStrip = document.getElementById("doublePipeLogMonthStrip");
  const doublePipeLogDayStrip = document.getElementById("doublePipeLogDayStrip");
  const doublePipeLogFilterMachEl = document.getElementById("doublePipeLogFilterMach");
  const doublePipeLogFilterFormEl = document.getElementById("doublePipeLogFilterForm");
  const doublePipeLogKpiScopeEl = document.getElementById("doublePipeLogKpiScope");
  const hwaseungDoublePipeLogEmpty = document.getElementById("hwaseungDoublePipeLogEmpty");
  const hwaseungDoublePipeLogContent = document.getElementById("hwaseungDoublePipeLogContent");
  const hwaseungDoublePipeLogMeta = document.getElementById("hwaseungDoublePipeLogMeta");
  const kpiUtilHwaseungDoublePipe = document.getElementById("kpiUtilHwaseungDoublePipe");
  const kpiOeeHwaseungDoublePipe = document.getElementById("kpiOeeHwaseungDoublePipe");
  const kpiDefectHwaseungDoublePipe = document.getElementById("kpiDefectHwaseungDoublePipe");
  const kpiDefectLabelHwaseungDoublePipe = document.getElementById("kpiDefectLabelHwaseungDoublePipe");
  const hwaseungDoublePipeLogTablesWrap = document.getElementById("hwaseungDoublePipeLogTablesWrap");
  const hwaseungDoublePipeLogProdWrap = document.getElementById("hwaseungDoublePipeLogProdWrap");
  const hwaseungDoublePipeLogMaintWrap = document.getElementById("hwaseungDoublePipeLogMaintWrap");
  const hwaseungDoublePipeLogOps = document.getElementById("hwaseungDoublePipeLogOps");
  const hwaseungDoublePipeLogTimeline = document.getElementById("hwaseungDoublePipeLogTimeline");
  const hwaseungDoublePipeLogSlicerClear = document.getElementById("hwaseungDoublePipeLogSlicerClear");
  const hwaseungDoublePipeLogGranularityEl = document.getElementById("hwaseungDoublePipeLogGranularity");
  const hwaseungDoublePipeLogSlicerSelection = document.getElementById("hwaseungDoublePipeLogSlicerSelection");
  const hwaseungDoublePipeLogYearLabel = document.getElementById("hwaseungDoublePipeLogYearLabel");
  const hwaseungDoublePipeLogYearPrev = document.getElementById("hwaseungDoublePipeLogYearPrev");
  const hwaseungDoublePipeLogYearNext = document.getElementById("hwaseungDoublePipeLogYearNext");
  const hwaseungDoublePipeLogMonthStrip = document.getElementById("hwaseungDoublePipeLogMonthStrip");
  const hwaseungDoublePipeLogDayStrip = document.getElementById("hwaseungDoublePipeLogDayStrip");
  const hwaseungDoublePipeLogProcessFilters = document.getElementById("hwaseungDoublePipeLogProcessFilters");
  const hwaseungDoublePipeLogKpiScopeEl = document.getElementById("hwaseungDoublePipeLogKpiScope");
  const viewMufflerLog = document.getElementById("viewMufflerLog");
  const viewStockUnitPrice = document.getElementById("viewStockUnitPrice");
  const stockUnitPriceFileInput = document.getElementById("stockUnitPriceFileInput");
  const stockUnitPriceDropzone = document.getElementById("stockUnitPriceDropzone");
  const stockUnitPriceFileName = document.getElementById("stockUnitPriceFileName");
  const stockUnitPriceSheetInfo = document.getElementById("stockUnitPriceSheetInfo");
  const stockUnitPriceTableWrap = document.getElementById("stockUnitPriceTableWrap");
  const stockUnitPriceError = document.getElementById("stockUnitPriceError");
  const stockUnitPriceFilterBar = document.getElementById("stockUnitPriceFilterBar");
  const btnStockUnitPriceFilterCode = document.getElementById("btnStockUnitPriceFilterCode");
  const panelStockUnitPriceFilterCode = document.getElementById("panelStockUnitPriceFilterCode");
  const searchStockUnitPriceFilterCode = document.getElementById("searchStockUnitPriceFilterCode");
  const listStockUnitPriceFilterCode = document.getElementById("listStockUnitPriceFilterCode");
  const btnStockUnitPriceFilterReset = document.getElementById("btnStockUnitPriceFilterReset");
  /** 재고 단가 미리보기 — 품목코드 필터(일별 통합표 picker와 동일 UX) */
  const stockUnitPriceCodeFilter = {
    button: btnStockUnitPriceFilterCode,
    panel: panelStockUnitPriceFilterCode,
    search: searchStockUnitPriceFilterCode,
    list: listStockUnitPriceFilterCode,
    /** @type {string[]} */
    options: [],
    /** @type {Set<string>} */
    selected: new Set(),
  };
  const mufflerLogEmpty = document.getElementById("mufflerLogEmpty");
  const mufflerLogContent = document.getElementById("mufflerLogContent");
  const mufflerLogMeta = document.getElementById("mufflerLogMeta");
  const mufflerLogSegmentsHost = document.getElementById("mufflerLogSegmentsHost");
  const mufflerLogMaintWrap = document.getElementById("mufflerLogMaintWrap");
  const drawingLogEmpty = document.getElementById("drawingLogEmpty");
  const drawingLogContent = document.getElementById("drawingLogContent");
  const drawingLogMeta = document.getElementById("drawingLogMeta");
  const kpiUtil = document.getElementById("kpiUtil");
  const kpiOee = document.getElementById("kpiOee");
  const kpiDefect = document.getElementById("kpiDefect");
  const drawingLogTablesWrap = document.getElementById("drawingLogTablesWrap");
  const drawingLogProdWrap = document.getElementById("drawingLogProdWrap");
  const drawingLogMaintWrap = document.getElementById("drawingLogMaintWrap");
  const drawingLogOps = document.getElementById("drawingLogOps");
  const kpiDefectLabel = document.getElementById("kpiDefectLabel");
  const drawingLogTimeline = document.getElementById("drawingLogTimeline");
  const drawingLogSlicerClear = document.getElementById("drawingLogSlicerClear");
  const drawingLogGranularityEl = document.getElementById("drawingLogGranularity");
  const drawingLogSlicerSelection = document.getElementById("drawingLogSlicerSelection");
  const drawingLogYearLabel = document.getElementById("drawingLogYearLabel");
  const drawingLogYearPrev = document.getElementById("drawingLogYearPrev");
  const drawingLogYearNext = document.getElementById("drawingLogYearNext");
  const drawingLogMonthStrip = document.getElementById("drawingLogMonthStrip");
  const drawingLogDayStrip = document.getElementById("drawingLogDayStrip");
  const navItems = document.querySelectorAll(".nav-item");
  const sidebarBrandHome = document.getElementById("sidebarBrandHome");
  const dailyFilters = document.getElementById("dailyFilters");
  const pickerGubun = document.getElementById("pickerGubun");
  const btnFilterGubun = document.getElementById("btnFilterGubun");
  const panelFilterGubun = document.getElementById("panelFilterGubun");
  const searchFilterGubun = document.getElementById("searchFilterGubun");
  const listFilterGubun = document.getElementById("listFilterGubun");
  const pickerCode = document.getElementById("pickerCode");
  const btnFilterCode = document.getElementById("btnFilterCode");
  const panelFilterCode = document.getElementById("panelFilterCode");
  const searchFilterCode = document.getElementById("searchFilterCode");
  const listFilterCode = document.getElementById("listFilterCode");
  const pickerName = document.getElementById("pickerName");
  const btnFilterName = document.getElementById("btnFilterName");
  const panelFilterName = document.getElementById("panelFilterName");
  const searchFilterName = document.getElementById("searchFilterName");
  const listFilterName = document.getElementById("listFilterName");
  const pickerType = document.getElementById("pickerType");
  const btnFilterType = document.getElementById("btnFilterType");
  const panelFilterType = document.getElementById("panelFilterType");
  const searchFilterType = document.getElementById("searchFilterType");
  const listFilterType = document.getElementById("listFilterType");
  const pickerExport = document.getElementById("pickerExport");
  const btnFilterExport = document.getElementById("btnFilterExport");
  const panelFilterExport = document.getElementById("panelFilterExport");
  const searchFilterExport = document.getElementById("searchFilterExport");
  const listFilterExport = document.getElementById("listFilterExport");
  const filterReset = document.getElementById("filterReset");
  const dailyTeamHint = document.getElementById("dailyTeamHint");
  const orderCalendarPanel = document.getElementById("orderCalendarPanel");
  const orderCalendarGrid = document.getElementById("orderCalendarGrid");
  const orderCalendarMonthSummary = document.getElementById("orderCalendarMonthSummary");
  const orderCalMonthLabel = document.getElementById("orderCalMonthLabel");
  const orderCalPrev = document.getElementById("orderCalPrev");
  const orderCalNext = document.getElementById("orderCalNext");
  const orderCalendarTypeSelect = document.getElementById("orderCalendarTypeSelect");
  const orderCalendarTypeActive = document.getElementById("orderCalendarTypeActive");

  /** 발주 달력에 표시할 월(1일 기준). 이전/다음 달 버튼으로만 이동합니다. */
  let orderCalendarCursor = new Date();
  orderCalendarCursor.setDate(1);
  orderCalendarCursor.setHours(0, 0, 0, 0);
  /** 새 파일·초기화 후 첫 렌더에서 데이터 시작 월로 맞춤 */
  let orderCalendarNeedsSync = true;
  /** 달력에서 클릭해 강조한 날짜 (YYYY-MM-DD), 없으면 "" */
  let orderCalendarSelectedYmd = "";

  /** @type {string} 드로잉팀 사이드 메뉴: "" 이면 타입 필터 미적용 */
  let drawingTeamSegment = "";
  /** @type {'home' | 'upload' | 'teamProductionUpload' | 'daily' | 'orderCalendar' | 'stockTable' | 'simple' | 'summary' | 'drawingLog' | 'parallelLog' | 'doublePipeLog' | 'hwaseungDoublePipeLog' | 'mufflerLog' | 'stockUnitPrice'} */
  let currentView = "home";

  /** @type {null | ReturnType<typeof parseDrawingLogFromMatrix> & { sheetName: string; fileLabel: string; maintenance?: { fileLabel: string; failures: any[]; failSheets: string[]; pmRows: any[]; pmSheets: string[]; equipmentStats: any[] } | null }} */
  let lastDrawingLog = null;

  /** @type {null | ReturnType<typeof parseParallelLogFromMatrix> & { sheetName: string; fileLabel: string; maintenance?: { fileLabel: string; failures: any[]; failSheets: string[]; pmRows: any[]; pmSheets: string[]; equipmentStats: any[] } | null }} */
  let lastParallelLog = null;

  /** @type {null | ReturnType<typeof parseParallelLogFromMatrix> & { sheetName: string; fileLabel: string; maintenance?: { fileLabel: string; failures: any[]; failSheets: string[]; pmRows: any[]; pmSheets: string[]; equipmentStats: any[] } | null }} */
  let lastDoublePipeLog = null;

  /** @type {null | ReturnType<typeof parseHwaseungDoublePipeLogFromMatrix> & { sheetName: string; fileLabel: string; maintenance?: { fileLabel: string; failures: any[]; failSheets: string[]; pmRows: any[]; pmSheets: string[]; equipmentStats: any[] } | null }} */
  let lastHwaseungDoublePipeLog = null;

  /**
   * @typedef {{ granularity: 'month' | 'day'; selectedMonths: Set<string>; selectedDays: Set<string>; timelineYear: number; processCut: boolean; processForming: boolean; processMachine: boolean; kpiScope: 'overall' | 'process' }} MufflerSegUiState
   */
  /** @type {null | { fileLabel: string; maintenance?: { fileLabel: string; failures: any[]; failSheets: string[]; pmRows: any[]; pmSheets: string[]; equipmentStats: any[] } | null; segments: (ReturnType<typeof parseMufflerLogFromMatrix> & { sheetName: string })[] }} */
  let lastMufflerLog = null;
  /** @type {MufflerSegUiState[]} 시트별 타임라인·공정 필터 상태 */
  let mufflerLogSegStates = [];

  /** @type {'month' | 'day'} */
  let drawingLogGranularity = "month";
  /** @type {Set<string>} YYYY-MM (월 단위 다중 선택) */
  let drawingLogSelectedMonths = new Set();
  /** @type {Set<string>} YYYY-MM-DD 또는 일자미상 (일 단위 다중 선택) */
  let drawingLogSelectedDays = new Set();
  /** 타임라인에 표시 중인 연도 */
  let drawingLogTimelineYear = new Date().getFullYear();

  /** @type {'month' | 'day'} */
  let parallelLogGranularity = "month";
  /** @type {Set<string>} */
  let parallelLogSelectedMonths = new Set();
  /** @type {Set<string>} */
  let parallelLogSelectedDays = new Set();
  let parallelLogTimelineYear = new Date().getFullYear();
  /** 체크된 공정 행만 집계 (절단·벤딩; 둘 다 해제 시 전체로 복귀) */
  let parallelLogProcessCut = true;
  let parallelLogProcessBend = true;

  /** @type {'month' | 'day'} */
  let doublePipeLogGranularity = "month";
  /** @type {Set<string>} */
  let doublePipeLogSelectedMonths = new Set();
  /** @type {Set<string>} */
  let doublePipeLogSelectedDays = new Set();
  let doublePipeLogTimelineYear = new Date().getFullYear();
  let doublePipeLogProcessMach = true;
  let doublePipeLogProcessForm = true;

  /** @type {'month' | 'day'} */
  let hwaseungDoublePipeLogGranularity = "month";
  /** @type {Set<string>} */
  let hwaseungDoublePipeLogSelectedMonths = new Set();
  /** @type {Set<string>} */
  let hwaseungDoublePipeLogSelectedDays = new Set();
  let hwaseungDoublePipeLogTimelineYear = new Date().getFullYear();
  /** @type {Set<string>} 체크된 공정만 집계(전부면 필터 없음) */
  let hwaseungDoublePipeLogProcessSelected = new Set(HWASEUNG_DOUBLE_PIPE_ALL_PROC_KEYS);

  /**
   * 상단 KPI 카드 집계: overall=공정 체크는 표에만 적용·KPI는 전 공정(선택 월·일은 반영),
   * process=표와 동일하게 공정 필터 반영
   * @type {'overall' | 'process'}
   */
  let parallelLogKpiScope = "process";
  let doublePipeLogKpiScope = "process";
  let hwaseungDoublePipeLogKpiScope = "process";

  /** @type {string[][] | null} */
  let rawRows = null;
  /** @type {string[]} */
  let headers = [];
  let lastFileName = "";
  /** @type {'wide' | 'long'} */
  let dataMode = "long";
  /** @type {ReturnType<typeof buildBoardFromWide> | ReturnType<typeof buildBoardFromLong> | null} */
  let lastBoard = null;
  /** @type {any} */
  let lastWorkbook = null;
  /** @type {{ stockByKey: Map<string, number>, stockByName: Map<string, number> } | null} */
  let lastStockData = null;
  /**
   * @type {{
   *   byKey: Map<string, number>,
   *   byName: Map<string, number>,
   *   byNameRows?: Map<string, { code: string, unit: number }[]>,
   *   byNormName?: Map<string, number>,
   *   fileLabel: string,
   *   sheetName: string,
   *   rowCount: number,
   *   priceMode: string,
   *   uploadDateLabel: string,
   *   previewRows: { code: string, name: string, stock: number, unit: number, amount: number }[],
   * } | null}
   */
  let lastStockUnitPriceData = null;
  /** @type {Record<string, { key:string, label:string, picker:HTMLElement, button:HTMLButtonElement, panel:HTMLElement, search:HTMLInputElement, list:HTMLElement, options:string[], selected:Set<string> }>} */
  const filterState = {};
  /**
   * 재고 표 필터 — 옵션 값: category(원문·""), code·name(트림·""), stock(숫자 문자열 키)
   * @type {{ category: { options: string[], selected: Set<string>, enabled: boolean }, code: { options: string[], selected: Set<string> }, name: { options: string[], selected: Set<string> }, stock: { options: string[], selected: Set<string> } }}
   */
  const stockTableFilterState = {
    category: { options: [], selected: new Set(), enabled: false },
    code: { options: [], selected: new Set() },
    name: { options: [], selected: new Set() },
    stock: { options: [], selected: new Set() },
  };

  function updateNavActiveState() {
    navItems.forEach((btn) => {
      const v = btn.dataset.view;
      let active = false;
      if (v === currentView) {
        if (currentView === "daily") {
          const btnSeg = btn.dataset.teamSegment !== undefined ? String(btn.dataset.teamSegment) : "";
          const cur = drawingTeamSegment || "";
          active = btnSeg === cur;
        } else {
          active = true;
        }
      }
      btn.classList.toggle("is-active", active);
    });
    if (sidebarBrandHome) {
      sidebarBrandHome.classList.toggle("is-active", currentView === "home");
    }
  }

  function updateDailyTeamHint() {
    const segText = drawingTeamSegment ? ` · 드로잉팀「${drawingTeamSegment}」` : "";
    if (dailyTeamHint) {
      dailyTeamHint.textContent = segText;
      dailyTeamHint.hidden = !drawingTeamSegment;
    }
    if (orderCalTeamHint) {
      orderCalTeamHint.textContent = segText;
      orderCalTeamHint.hidden = !drawingTeamSegment;
    }
  }

  /**
   * 일별 통합표 ↔ 날짜별 발주 달력에서 동일한 필터 바를 공유합니다.
   * @param {'home' | 'upload' | 'teamProductionUpload' | 'daily' | 'orderCalendar' | 'stockTable' | 'simple' | 'summary' | 'drawingLog' | 'parallelLog' | 'doublePipeLog' | 'hwaseungDoublePipeLog' | 'mufflerLog' | 'stockUnitPrice'} viewKey
   */
  function mountDailyFiltersForView(viewKey) {
    if (!dailyFilters || !dailyFiltersSlotDaily || !dailyFiltersSlotCalendar) return;
    if (viewKey === "orderCalendar") {
      dailyFiltersSlotCalendar.appendChild(dailyFilters);
    } else {
      dailyFiltersSlotDaily.appendChild(dailyFilters);
    }
  }

  /**
   * @param {'home' | 'upload' | 'teamProductionUpload' | 'daily' | 'orderCalendar' | 'stockTable' | 'simple' | 'summary' | 'drawingLog' | 'parallelLog' | 'doublePipeLog' | 'hwaseungDoublePipeLog' | 'mufflerLog' | 'stockUnitPrice'} viewKey
   * @param {{ teamSegment?: string }} [opts] 일별 통합표·날짜별 달력에서만 `teamSegment`로 드로잉팀 구간 적용
   */
  function setView(viewKey, opts) {
    const o = opts || {};
    currentView = viewKey;
    if ((viewKey === "daily" || viewKey === "orderCalendar") && o.teamSegment !== undefined) {
      setDrawingTeamSegment(o.teamSegment || "");
    } else if (viewKey !== "daily" && viewKey !== "orderCalendar") {
      setDrawingTeamSegment("");
    }
    if (viewHome) viewHome.hidden = viewKey !== "home";
    viewUpload.hidden = viewKey !== "upload";
    if (viewTeamProductionUpload) viewTeamProductionUpload.hidden = viewKey !== "teamProductionUpload";
    viewDaily.hidden = viewKey !== "daily";
    if (viewOrderCalendar) viewOrderCalendar.hidden = viewKey !== "orderCalendar";
    if (viewStockTable) viewStockTable.hidden = viewKey !== "stockTable";
    viewSimple.hidden = viewKey !== "simple";
    viewSummary.hidden = viewKey !== "summary";
    if (viewDrawingLog) viewDrawingLog.hidden = viewKey !== "drawingLog";
    if (viewParallelLog) viewParallelLog.hidden = viewKey !== "parallelLog";
    if (viewDoublePipeLog) viewDoublePipeLog.hidden = viewKey !== "doublePipeLog";
    if (viewHwaseungDoublePipeLog) viewHwaseungDoublePipeLog.hidden = viewKey !== "hwaseungDoublePipeLog";
    if (viewMufflerLog) viewMufflerLog.hidden = viewKey !== "mufflerLog";
    if (viewStockUnitPrice) viewStockUnitPrice.hidden = viewKey !== "stockUnitPrice";
    if (mainWrap) mainWrap.classList.toggle("main-wrap--home", viewKey === "home");
    if (kmStripEl) kmStripEl.hidden = viewKey === "home";
    document.body.classList.toggle("app-body--print-calendar-only", viewKey === "orderCalendar");
    document.body.classList.toggle("app-body--print-stock-only", viewKey === "stockTable");
    mountDailyFiltersForView(viewKey);
    if (viewKey !== "orderCalendar") clearOrderCalendarPrintFit();
    if (viewKey !== "stockTable") clearStockTablePrintFit();
    updateNavActiveState();
    updateDailyTeamHint();
  }

  /**
   * @param {HTMLElement} el `data-view`가 있는 네비·메인 카드 버튼
   */
  function applyMainViewFromElement(el) {
    const v = el.dataset.view;
    if (!v) return;
    if (v === "daily" && el.dataset.teamSegment !== undefined) {
      setView("daily", { teamSegment: el.dataset.teamSegment });
    } else {
      setView(v);
    }
    if (v === "drawingLog") renderDrawingLogPanel();
    else if (v === "parallelLog") renderParallelLogPanel();
    else if (v === "doublePipeLog") renderDoublePipeLogPanel();
    else if (v === "hwaseungDoublePipeLog") renderHwaseungDoublePipeLogPanel();
    else if (v === "mufflerLog") renderMufflerLogPanel();
    else if (v === "teamProductionUpload") {
      /* 허브: 별도 렌더 없음 */
    } else if (v === "stockUnitPrice") renderStockUnitPricePanel();
    else if (v === "stockTable") renderStockTableView();
    else if (lastBoard && (v === "daily" || v === "simple")) {
      renderBoard(lastBoard);
      renderSimpleTable(lastBoard);
    } else if (lastBoard && v === "orderCalendar") {
      renderOrderCalendar(lastBoard);
    }
  }

  function clearOutputs() {
    lastBoard = null;
    lastWorkbook = null;
    lastStockData = null;
    lastStockUnitPriceData = null;
    const bt = tableWrap.querySelector("table.board-table");
    if (bt) bt.remove();
    emptyState.hidden = false;
    const st = simpleTableWrap.querySelector("table.simple-table");
    if (st) st.remove();
    emptyStateSimple.hidden = false;
    dailyFilters.classList.add("filter-bar--hidden");
    clearFilterSelections();
    Object.keys(filterState).forEach((k) => {
      filterState[k].options = [];
      filterState[k].selected.clear();
      filterState[k].list.innerHTML = "";
      filterState[k].button.textContent = "전체";
    });
    renderSummary(null);
    rowCountEl.textContent = "0건";
    btnExport.disabled = true;
    stockFileNameEl.textContent = "";
    stockSheetInfoEl.textContent = "";
    if (stockUnitPriceFileName) stockUnitPriceFileName.textContent = "";
    if (stockUnitPriceSheetInfo) stockUnitPriceSheetInfo.textContent = "";
    hideStockUnitPriceError();
    clearStockUnitPricePreviewTable();
    resetStockUnitPriceCodeFilter();
    lastDrawingLog = null;
    clearDrawingLogUi();
    lastParallelLog = null;
    clearParallelLogUi();
    lastDoublePipeLog = null;
    clearDoublePipeLogUi();
    lastHwaseungDoublePipeLog = null;
    clearHwaseungDoublePipeLogUi();
    lastMufflerLog = null;
    clearMufflerLogUi();
    orderCalendarNeedsSync = true;
    orderCalendarSelectedYmd = "";
    renderOrderCalendar(null);
    clearStockTableView();
    resetStockTableFilters();
    if (mappingBlock) mappingBlock.hidden = true;
  }

  function onDataReady() {
    if (!lastBoard) return;
    dailyFilters.classList.remove("filter-bar--hidden");
    populateAllFilters(lastBoard, true);
    renderBoard(lastBoard);
    renderSimpleTable(lastBoard);
    renderSummary(lastBoard);
    if (currentView === "orderCalendar") renderOrderCalendar(lastBoard);
  }

  function initializePickers() {
    filterState.gubun = {
      key: "gubun",
      label: "구분",
      picker: pickerGubun,
      button: btnFilterGubun,
      panel: panelFilterGubun,
      search: searchFilterGubun,
      list: listFilterGubun,
      options: [],
      selected: new Set(),
    };
    filterState.code = {
      key: "code",
      label: "품목코드",
      picker: pickerCode,
      button: btnFilterCode,
      panel: panelFilterCode,
      search: searchFilterCode,
      list: listFilterCode,
      options: [],
      selected: new Set(),
    };
    filterState.name = {
      key: "name",
      label: "제품",
      picker: pickerName,
      button: btnFilterName,
      panel: panelFilterName,
      search: searchFilterName,
      list: listFilterName,
      options: [],
      selected: new Set(),
    };
    filterState.type = {
      key: "type",
      label: "타입",
      picker: pickerType,
      button: btnFilterType,
      panel: panelFilterType,
      search: searchFilterType,
      list: listFilterType,
      options: [],
      selected: new Set(),
    };
    filterState.export = {
      key: "export",
      label: "수출발주량",
      picker: pickerExport,
      button: btnFilterExport,
      panel: panelFilterExport,
      search: searchFilterExport,
      list: listFilterExport,
      options: [],
      selected: new Set(),
    };
  }

  function updatePickerButton(key) {
    const st = filterState[key];
    const total = st.options.length;
    const sel = st.selected.size;
    if (total === 0 || sel === 0 || sel === total) {
      st.button.textContent = "전체";
      return;
    }
    if (sel === 1) {
      st.button.textContent = [...st.selected][0];
      return;
    }
    st.button.textContent = `${sel}개 선택`;
  }

  function renderPickerList(key) {
    const st = filterState[key];
    const q = norm(st.search.value || "");
    const filtered = st.options.filter((v) => norm(v).includes(q));
    st.list.innerHTML = "";

    const allRow = document.createElement("label");
    allRow.className = "picker-option";
    const allCb = document.createElement("input");
    allCb.type = "checkbox";
    allCb.checked = st.selected.size === st.options.length && st.options.length > 0;
    allCb.indeterminate = st.selected.size > 0 && st.selected.size < st.options.length;
    allCb.addEventListener("change", () => {
      if (allCb.checked) st.selected = new Set(st.options);
      else st.selected.clear();
      renderPickerList(key);
      updatePickerButton(key);
      if (lastBoard) {
        renderBoard(lastBoard);
        renderSimpleTable(lastBoard);
      }
    });
    const allTx = document.createElement("span");
    allTx.textContent = "전체";
    allRow.appendChild(allCb);
    allRow.appendChild(allTx);
    st.list.appendChild(allRow);

    if (filtered.length === 0) {
      const em = document.createElement("div");
      em.className = "picker-empty";
      em.textContent = "검색 결과 없음";
      st.list.appendChild(em);
      return;
    }

    filtered.forEach((v) => {
      const row = document.createElement("label");
      row.className = "picker-option";
      const cb = document.createElement("input");
      cb.type = "checkbox";
      cb.checked = st.selected.has(v);
      cb.addEventListener("change", () => {
        if (cb.checked) st.selected.add(v);
        else st.selected.delete(v);
        updatePickerButton(key);
        renderPickerList(key);
        if (lastBoard) {
          renderBoard(lastBoard);
          renderSimpleTable(lastBoard);
        }
      });
      const tx = document.createElement("span");
      tx.textContent = v || "—";
      row.appendChild(cb);
      row.appendChild(tx);
      st.list.appendChild(row);
    });
  }

  function setPickerOptions(key, values, preserveSelection) {
    const st = filterState[key];
    const prev = preserveSelection ? new Set(st.selected) : new Set();
    st.options = values.slice();
    st.selected = new Set(st.options.filter((v) => preserveSelection && prev.has(v)));
    if (st.selected.size === 0) st.selected = new Set(st.options);
    renderPickerList(key);
    updatePickerButton(key);
  }

  function clearFilterSelections() {
    setDrawingTeamSegment("");
    Object.keys(filterState).forEach((k) => {
      const st = filterState[k];
      st.search.value = "";
      st.selected = new Set(st.options);
      renderPickerList(k);
      updatePickerButton(k);
    });
  }

  function setDrawingTeamSegment(seg) {
    const s = typeof seg === "string" ? seg.trim() : "";
    drawingTeamSegment = !s ? "" : DRAWING_TEAM_SEGMENTS.includes(s) ? s : "";
    updateNavActiveState();
    updateDailyTeamHint();
  }

  /**
   * @param {string} typeStr
   * @param {string} segment DRAWING_TEAM_SEGMENTS 중 하나
   */
  function productTypeMatchesDrawingSegment(typeStr, segment) {
    const t = String(typeStr ?? "").trim();
    if (!segment) return true;
    if (segment === "화승이중관") return t.includes("화승이중관");
    if (segment === "이중관") return t.includes("이중관") && !t.includes("화승이중관");
    if (segment === "머플러") return t.includes("머플러");
    if (segment === "페럴") return t.includes("페럴");
    if (segment === "드로잉") return t.includes("드로잉");
    return true;
  }

  /** 일별 통합표 `타입`·제품명에서 팀 구간을 추정해 기본 개당 단가에 매핑 */
  function inferDrawingTeamSegmentForPricing(typeStr, nameStr) {
    const type = String(typeStr ?? "").trim();
    const name = String(nameStr ?? "").trim();
    const ordered = ["화승이중관", "이중관", "머플러", "페럴", "드로잉"];
    for (const seg of ordered) {
      if (productTypeMatchesDrawingSegment(type, seg)) return seg;
    }
    for (const seg of ordered) {
      if (productTypeMatchesDrawingSegment(name, seg)) return seg;
    }
    return "";
  }

  function populateAllFilters(board, preserveSelection) {
    setPickerOptions("gubun", [...SUB_ROWS], preserveSelection);
    const codes = [...new Set(board.products.map((p) => p.code || "").filter(Boolean))].sort((a, b) =>
      a.localeCompare(b, "ko")
    );
    const names = [...new Set(board.products.map((p) => p.name || "").filter(Boolean))].sort((a, b) =>
      a.localeCompare(b, "ko")
    );
    const types = [...new Set(board.products.map((p) => p.type || "").filter(Boolean))].sort((a, b) =>
      a.localeCompare(b, "ko")
    );
    setPickerOptions("code", codes, preserveSelection);
    setPickerOptions("name", names, preserveSelection);
    setPickerOptions("type", types, preserveSelection);

    const exportVals = [];
    for (const pack of board.products) {
      for (const sub of SUB_ROWS) {
        const line = pack.rows.get(sub);
        if (!line) continue;
        exportVals.push(String(Number(line.exportCol || 0)));
      }
    }
    const uniqExport = [...new Set(exportVals)].sort((a, b) => Number(a) - Number(b));
    setPickerOptions("export", uniqExport, preserveSelection);
    syncOrderCalendarTypeUi();
  }

  /**
   * 날짜별 달력 상단 타입 셀렉트·안내 문구를 일별 통합표 「타입」 필터와 맞춥니다.
   */
  function syncOrderCalendarTypeUi() {
    const sel = orderCalendarTypeSelect;
    const hint = orderCalendarTypeActive;
    if (!sel || !filterState.type) return;
    const opts = filterState.type.options || [];
    const selected = filterState.type.selected;
    const allOn = opts.length > 0 && selected.size === opts.length;

    sel.innerHTML = "";
    const optAll = document.createElement("option");
    optAll.value = "";
    optAll.textContent = "전체";
    sel.appendChild(optAll);
    for (const t of opts) {
      const o = document.createElement("option");
      o.value = t;
      o.textContent = t || "(빈값)";
      sel.appendChild(o);
    }

    if (allOn) {
      sel.value = "";
    } else if (selected.size === 1) {
      const v = [...selected][0];
      sel.value = opts.includes(v) ? v : "";
    } else {
      sel.value = "";
    }

    if (!hint) return;
    if (opts.length === 0) {
      hint.textContent = "타입 열이 없거나 비어 있습니다.";
    } else if (allOn) {
      hint.textContent = "지금 보는 타입: 전체";
    } else if (selected.size === 1) {
      const v = [...selected][0];
      hint.textContent = `지금 보는 타입: ${v || "(빈값)"}`;
    } else if (selected.size === 0) {
      hint.textContent = "선택된 타입이 없습니다. 위 필터에서 선택해 주세요.";
    } else {
      const list = [...selected].sort((a, b) => a.localeCompare(b, "ko")).join(", ");
      hint.textContent = `지금 보는 타입: ${list} (${selected.size}개)`;
    }
  }

  function applyOrderCalendarTypeFromSelect() {
    if (!orderCalendarTypeSelect || !filterState.type || !lastBoard) return;
    const v = orderCalendarTypeSelect.value;
    const opts = filterState.type.options || [];
    if (v === "") {
      filterState.type.selected = new Set(opts);
    } else {
      filterState.type.selected = new Set([v]);
    }
    updatePickerButton("type");
    renderPickerList("type");
    renderBoard(lastBoard);
    renderSimpleTable(lastBoard);
  }

  function getFilteredProductPacks(board) {
    const codeSet = filterState.code.selected;
    const codeAll = codeSet.size === filterState.code.options.length;
    const nameSet = filterState.name.selected;
    const nameAll = nameSet.size === filterState.name.options.length;
    const typeSet = filterState.type.selected;
    const typeAll = typeSet.size === filterState.type.options.length;
    return board.products.filter((pack) => {
      if (!codeAll && !codeSet.has(pack.code || "")) return false;
      if (!nameAll && !nameSet.has(pack.name || "")) return false;
      if (!typeAll && !typeSet.has(pack.type || "")) return false;
      if (drawingTeamSegment && !productTypeMatchesDrawingSegment(pack.type || "", drawingTeamSegment)) return false;
      return true;
    });
  }

  function getFilteredSubs() {
    const set = filterState.gubun.selected;
    if (set.size === filterState.gubun.options.length) return SUB_ROWS;
    return SUB_ROWS.filter((s) => set.has(s));
  }

  function norm(s) {
    return String(s ?? "")
      .trim()
      .toLowerCase()
      .replace(/\s+/g, "");
  }

  function matchHeader(header, keys) {
    const h = norm(header);
    for (const k of keys) {
      if (h === norm(k) || h.includes(norm(k))) return true;
    }
    return false;
  }

  function findSheetName(workbook) {
    const names = workbook.SheetNames || [];
    const exact = names.find((n) => String(n).trim() === SHEET_TARGET);
    if (exact) return exact;
    const partial = names.find((n) => {
      const t = norm(n);
      return t.includes("발주서") && t.includes("입력");
    });
    return partial || null;
  }

  /** @param {any} workbook */
  function findPastOrderSheetName(workbook) {
    const names = workbook.SheetNames || [];
    const exact = names.find((n) => String(n).trim() === SHEET_PAST_ORDER);
    if (exact) return exact;
    const loose = names.find((n) => {
      const t = norm(n);
      return t.includes("지난발주");
    });
    return loose || null;
  }

  /** @param {any} workbook */
  function findStockSheetName(workbook) {
    const names = workbook.SheetNames || [];
    const exact = names.find((n) => String(n).trim() === SHEET_STOCK);
    if (exact) return exact;
    const loose = names.find((n) => norm(n).includes("재고"));
    return loose || names[0] || null;
  }

  function splitProductKey(key) {
    const i = String(key).indexOf("\t");
    if (i < 0) return { code: "", name: String(key) };
    return { code: key.slice(0, i), name: key.slice(i + 1) };
  }

  /**
   * @param {any[][]} matrix
   * @param {number} hi 헤더 행 인덱스
   */
  function parseStockRowsWithHeaderRow(matrix, hi) {
    /** @type {Map<string, number>} */
    const stockByKey = new Map();
    /** @type {Map<string, number>} */
    const stockByName = new Map();
    /** @type {{ code: string, name: string, stock: number, category: string }[]} */
    const rows = [];
    const empty = { stockByKey, stockByName, rows, rowCount: 0, headerRowIndex: -1, hasCategoryColumn: false };
    if (!matrix || matrix.length < 2 || hi < 0 || hi >= matrix.length) return empty;

    const header = matrix[hi] || [];
    const iCode = findHeaderIndex(header, CODE_KEYS);
    const iName = findProductNameColumnIndex(header);
    const iStock = findHeaderIndex(header, ["재고수량", ...STOCK_KEYS, "재고량", "현재고"]);
    const iCat = findHeaderIndex(header, STOCK_CATEGORY_KEYS);
    if (iName < 0 || iStock < 0) return empty;
    const hasCategoryColumn = iCat >= 0;

    for (let r = hi + 1; r < matrix.length; r++) {
      const row = matrix[r];
      if (!row || row.length === 0) continue;
      const code = iCode >= 0 ? String(row[iCode] ?? "").trim() : "";
      const name = String(row[iName] ?? "").trim();
      if (!code && !name) continue;
      const stock = parseNumber(row[iStock]);
      const category = hasCategoryColumn ? String(row[iCat] ?? "").trim() : "";
      rows.push({ code, name, stock, category });
      const key = `${code}\t${name}`;
      stockByKey.set(key, (stockByKey.get(key) || 0) + stock);
      if (name) stockByName.set(name, (stockByName.get(name) || 0) + stock);
    }
    return {
      stockByKey,
      stockByName,
      rows,
      rowCount: rows.length,
      headerRowIndex: hi,
      hasCategoryColumn,
    };
  }

  /**
   * 재고 파일(날짜 없음) -> 코드+제품명 매칭용 재고 마스터
   * @param {any[][]} matrix
   * @returns {{ stockByKey: Map<string, number>, stockByName: Map<string, number>, rows?: { code: string, name: string, stock: number }[], rowCount?: number, headerRowIndex?: number }}
   */
  function extractStockData(matrix) {
    const empty = {
      stockByKey: new Map(),
      stockByName: new Map(),
      rows: [],
      rowCount: 0,
      headerRowIndex: -1,
      hasCategoryColumn: false,
    };
    if (!matrix || matrix.length < 2) return empty;

    const scanLimit = Math.min(80, matrix.length);
    let bestHi = -1;
    let bestScore = -1;
    for (let r = 0; r < scanLimit; r++) {
      const row = matrix[r] || [];
      if (!row.length) continue;
      const iCode = findHeaderIndex(row, CODE_KEYS);
      const iName = findProductNameColumnIndex(row);
      const iStock = findHeaderIndex(row, ["재고수량", ...STOCK_KEYS, "재고량", "현재고"]);
      const iCat = findHeaderIndex(row, STOCK_CATEGORY_KEYS);
      let sc = 0;
      if (iCode >= 0) sc += 2;
      if (iName >= 0) sc += 3;
      if (iStock >= 0) sc += 6;
      if (iCat >= 0) sc += 1;
      if (sc > bestScore) {
        bestScore = sc;
        bestHi = r;
      }
    }
    if (bestHi >= 0 && bestScore >= 6) {
      const out = parseStockRowsWithHeaderRow(matrix, bestHi);
      if (out.rowCount > 0) return out;
    }
    for (let hi = 0; hi < Math.min(50, matrix.length - 1); hi++) {
      const out = parseStockRowsWithHeaderRow(matrix, hi);
      if (out.rowCount > 0) return out;
    }
    return empty;
  }

  function hideStockUnitPriceError() {
    if (!stockUnitPriceError) return;
    stockUnitPriceError.hidden = true;
    stockUnitPriceError.textContent = "";
  }

  function closeAllStockUnitPricePicker() {
    if (stockUnitPriceCodeFilter.panel) stockUnitPriceCodeFilter.panel.hidden = true;
  }

  function hideStockUnitPriceFilterBar() {
    if (stockUnitPriceFilterBar) stockUnitPriceFilterBar.classList.add("filter-bar--hidden");
  }

  function showStockUnitPriceFilterBar() {
    if (stockUnitPriceFilterBar) stockUnitPriceFilterBar.classList.remove("filter-bar--hidden");
  }

  function resetStockUnitPriceCodeFilter() {
    if (!stockUnitPriceCodeFilter.button) return;
    stockUnitPriceCodeFilter.options = [];
    stockUnitPriceCodeFilter.selected.clear();
    if (stockUnitPriceCodeFilter.search) stockUnitPriceCodeFilter.search.value = "";
    if (stockUnitPriceCodeFilter.list) stockUnitPriceCodeFilter.list.innerHTML = "";
    stockUnitPriceCodeFilter.button.textContent = "전체";
    closeAllStockUnitPricePicker();
    hideStockUnitPriceFilterBar();
  }

  function stockUnitPriceCodeLabel(code) {
    const c = code != null ? String(code) : "";
    return c === "" ? "(코드 없음)" : c;
  }

  function updateStockUnitPriceCodeFilterButton() {
    const st = stockUnitPriceCodeFilter;
    if (!st.button) return;
    const total = st.options.length;
    const sel = st.selected.size;
    if (total === 0 || sel === 0 || sel === total) {
      st.button.textContent = "전체";
      return;
    }
    if (sel === 1) {
      st.button.textContent = stockUnitPriceCodeLabel([...st.selected][0]);
      return;
    }
    st.button.textContent = `${sel}개 선택`;
  }

  function renderStockUnitPriceCodePickerList() {
    const st = stockUnitPriceCodeFilter;
    if (!st.list || !st.search) return;
    const q = norm(st.search.value || "");
    const filtered = st.options.filter((v) => norm(stockUnitPriceCodeLabel(v)).includes(q) || norm(v).includes(q));
    st.list.innerHTML = "";

    const allRow = document.createElement("label");
    allRow.className = "picker-option";
    const allCb = document.createElement("input");
    allCb.type = "checkbox";
    allCb.checked = st.selected.size === st.options.length && st.options.length > 0;
    allCb.indeterminate = st.selected.size > 0 && st.selected.size < st.options.length;
    allCb.addEventListener("change", () => {
      if (allCb.checked) st.selected = new Set(st.options);
      else st.selected.clear();
      renderStockUnitPriceCodePickerList();
      updateStockUnitPriceCodeFilterButton();
      renderStockUnitPricePreviewTable();
    });
    const allTx = document.createElement("span");
    allTx.textContent = "전체";
    allRow.appendChild(allCb);
    allRow.appendChild(allTx);
    st.list.appendChild(allRow);

    if (filtered.length === 0) {
      const em = document.createElement("div");
      em.className = "picker-empty";
      em.textContent = "검색 결과 없음";
      st.list.appendChild(em);
      return;
    }

    filtered.forEach((v) => {
      const row = document.createElement("label");
      row.className = "picker-option";
      const cb = document.createElement("input");
      cb.type = "checkbox";
      cb.checked = st.selected.has(v);
      cb.addEventListener("change", () => {
        if (cb.checked) st.selected.add(v);
        else st.selected.delete(v);
        updateStockUnitPriceCodeFilterButton();
        renderStockUnitPriceCodePickerList();
        renderStockUnitPricePreviewTable();
      });
      const tx = document.createElement("span");
      tx.textContent = stockUnitPriceCodeLabel(v);
      row.appendChild(cb);
      row.appendChild(tx);
      st.list.appendChild(row);
    });
  }

  function populateStockUnitPriceCodeFilterFromPreview() {
    if (!lastStockUnitPriceData?.previewRows?.length || !stockUnitPriceCodeFilter.button) return;
    const seen = new Set();
    const codes = [];
    for (const pr of lastStockUnitPriceData.previewRows) {
      const c = pr.code != null ? String(pr.code) : "";
      if (seen.has(c)) continue;
      seen.add(c);
      codes.push(c);
    }
    codes.sort((a, b) => a.localeCompare(b, "ko", { numeric: true }));
    stockUnitPriceCodeFilter.options = codes;
    stockUnitPriceCodeFilter.selected = new Set(codes);
    if (stockUnitPriceCodeFilter.search) stockUnitPriceCodeFilter.search.value = "";
    renderStockUnitPriceCodePickerList();
    updateStockUnitPriceCodeFilterButton();
    showStockUnitPriceFilterBar();
  }

  function showStockUnitPriceError(msg) {
    if (stockUnitPriceError) {
      stockUnitPriceError.textContent = msg;
      stockUnitPriceError.hidden = false;
    } else {
      alert(msg);
    }
  }

  function clearStockUnitPricePreviewTable() {
    if (!stockUnitPriceTableWrap) return;
    stockUnitPriceTableWrap.innerHTML = "";
    stockUnitPriceTableWrap.hidden = true;
  }

  function formatKoreanUploadDate(d) {
    try {
      return d.toLocaleDateString("ko-KR", {
        year: "numeric",
        month: "long",
        day: "numeric",
        weekday: "short",
      });
    } catch {
      return "";
    }
  }

  function findSellBasisAmountColumn(headerRow) {
    let i = findExactHeaderIndex(headerRow, ["매출기준액"]);
    if (i >= 0) return i;
    return findHeaderIndex(headerRow, ["매출기준액"]);
  }

  /**
   * @param {any[][]} matrix
   * @param {number} hi 헤더 행 인덱스
   */
  function parseStockUnitPriceWithHeaderRow(matrix, hi) {
    /** @type {Map<string, number>} */
    const byKey = new Map();
    /** @type {Map<string, number>} */
    const byName = new Map();
    /** @type {Map<string, number>} */
    const byNormName = new Map();
    /** 동일 품목명·다른 품목코드 행 구분용(lookup 시 코드 매칭) */
    /** @type {Map<string, { code: string, unit: number }[]>} */
    const byNameRows = new Map();
    /** @type {{ code: string, name: string, stock: number, unit: number, amount: number }[]} */
    const previewRows = [];
    const empty = { byKey, byName, byNormName, byNameRows, rowCount: 0, priceMode: "", headerRowIndex: -1, previewRows };

    const header = matrix[hi] || [];
    const iCode = findHeaderIndex(header, CODE_KEYS);
    let iName = findHeaderIndex(header, STOCK_UNIT_PRICE_NAME_KEYS);
    if (iName < 0) iName = findProductNameColumnIndex(header);

    const iPriceDirect = findUnitPriceColumnIndex(header);
    const iBuyAmt = findBuyBasisAmountColumn(header);
    const iSellAmt = findSellBasisAmountColumn(header);
    const iCurSt = findCurrentStockColumnForPricing(header);
    if (iName < 0) {
      iName = inferProductNameColumnFromRows(matrix, hi, [iCode, iPriceDirect, iBuyAmt, iSellAmt, iCurSt]);
    }
    if (iName < 0 || iName === iCode) return empty;
    const canDirect = iPriceDirect >= 0;
    const canDerivedBuy = iBuyAmt >= 0 && iCurSt >= 0;
    const canDerivedSell = iSellAmt >= 0 && iCurSt >= 0;
    if (!canDirect && !canDerivedBuy && !canDerivedSell) return empty;

    let priceMode = "direct";
    if (!canDirect) priceMode = canDerivedBuy ? "derived-buy" : "derived-sell";

    let rowCount = 0;
    const maxPreview = 4000;

    for (let r = hi + 1; r < matrix.length; r++) {
      const row = matrix[r];
      if (!row || row.length === 0) continue;
      const code = iCode >= 0 ? String(row[iCode] ?? "").trim() : "";
      const name = String(row[iName] ?? "").trim();
      if (!name && !code) continue;

      const stockV = iCurSt >= 0 ? parseNumber(row[iCurSt]) : 0;

      let unit = NaN;
      if (canDirect) unit = parseUnitPrice(row[iPriceDirect]);
      if (!Number.isFinite(unit) && canDerivedBuy) {
        const q = parseNumber(row[iCurSt]);
        const a = parseNumber(row[iBuyAmt]);
        if (q > 0 && a > 0) unit = a / q;
      }
      if (!Number.isFinite(unit) && canDerivedSell) {
        const q = parseNumber(row[iCurSt]);
        const a = parseNumber(row[iSellAmt]);
        if (q > 0 && a > 0) unit = a / q;
      }
      if (!Number.isFinite(unit) || unit < 0) continue;

      const key = `${code}\t${name}`;
      byKey.set(key, unit);
      if (name) {
        if (!byNameRows.has(name)) byNameRows.set(name, []);
        byNameRows.get(name).push({ code, unit });
      }
      const amount = Math.round(stockV * unit);
      if (previewRows.length < maxPreview) {
        previewRows.push({ code, name, stock: stockV, unit, amount });
      }
      rowCount++;
    }

    for (const [name, rows] of byNameRows) {
      if (rows.length === 1) byName.set(name, rows[0].unit);
      const nk = norm(name);
      if (nk && rows.length > 0 && Number.isFinite(rows[0].unit) && !byNormName.has(nk)) {
        byNormName.set(nk, rows[0].unit);
      }
    }

    return { byKey, byName, byNormName, byNameRows, rowCount, priceMode, headerRowIndex: hi, previewRows };
  }

  /**
   * 재고_단가 등 마스터 엑셀 → 품목코드+품목명으로 단가 조회 + 미리보기 행
   * @param {any[][]} matrix
   */
  function extractStockUnitPriceData(matrix) {
    const empty = {
      byKey: new Map(),
      byName: new Map(),
      byNormName: new Map(),
      byNameRows: new Map(),
      rowCount: 0,
      priceMode: "",
      headerRowIndex: -1,
      previewRows: [],
    };
    if (!matrix || matrix.length < 2) return empty;

    const scanLimit = Math.min(80, matrix.length);
    let bestHi = -1;
    let bestScore = -1;

    function scoreHeaderRow(row) {
      if (!row || !row.length) return -1;
      const iCode = findHeaderIndex(row, CODE_KEYS);
      let iName = findHeaderIndex(row, STOCK_UNIT_PRICE_NAME_KEYS);
      if (iName < 0) iName = findProductNameColumnIndex(row);
      if (iName === iCode) return -1;
      const iPrice = findUnitPriceColumnIndex(row);
      const iBuy = findBuyBasisAmountColumn(row);
      const iSell = findSellBasisAmountColumn(row);
      const iCur = findCurrentStockColumnForPricing(row);
      let sc = 0;
      if (iCode >= 0) sc += 2;
      if (iName >= 0) sc += 3;
      if (iPrice >= 0) sc += 6;
      else if (iBuy >= 0 && iCur >= 0) sc += 4;
      else if (iSell >= 0 && iCur >= 0) sc += 4;
      return sc;
    }

    for (let r = 0; r < scanLimit; r++) {
      const row = matrix[r] || [];
      const sc = scoreHeaderRow(row);
      if (sc > bestScore) {
        bestScore = sc;
        bestHi = r;
      }
    }

    if (bestHi >= 0 && bestScore >= 5) {
      const out = parseStockUnitPriceWithHeaderRow(matrix, bestHi);
      if (out.rowCount > 0) return out;
    }

    for (let hi = 0; hi < Math.min(50, matrix.length - 1); hi++) {
      const out = parseStockUnitPriceWithHeaderRow(matrix, hi);
      if (out.rowCount > 0) return out;
    }

    return empty;
  }

  /**
   * 재고 단가 통합 엑셀(예: `연간 주별 재고금액.xlsx`)에서 실제 단가 표가 있는 시트 우선순위.
   * 시트명이 `단가 (2)`인 경우가 많아, 행 수와 관계없이 이 시트를 다른 시트보다 먼저 씁니다.
   */
  function stockUnitPriceSheetPreferenceRank(sheetName) {
    const raw = String(sheetName ?? "");
    const h = norm(raw);
    if (!h) return 0;
    if (h === "단가(2)" || h.includes("단가(2)")) return 100;
    if (h.startsWith("단가") && h.includes("(2)")) return 100;
    // 시트명이 깨져도 "(2)"는 보존되는 경우가 많아 우선권 부여
    if (h.includes("(2)")) return 95;
    if (h === "단가") return 80;
    if (h.includes("단가")) return 50;
    return 0;
  }

  /** @param {any} wb */
  function pickStockUnitPriceSheetAndMatrix(wb) {
    const names = wb.SheetNames || [];
    /** @type {{ name: string, matrix: any[][], parsed: ReturnType<typeof extractStockUnitPriceData> }[]} */
    const candidates = [];
    for (const n of names) {
      const matrix = sheetToMatrix(wb, n);
      const parsed = extractStockUnitPriceData(matrix);
      if (parsed.rowCount > 0) candidates.push({ name: n, matrix, parsed });
    }
    if (candidates.length === 0) {
      return { sheetName: names[0] || null, matrix: null, parsed: null };
    }
    candidates.sort((a, b) => {
      const ra = stockUnitPriceSheetPreferenceRank(a.name);
      const rb = stockUnitPriceSheetPreferenceRank(b.name);
      if (rb !== ra) return rb - ra;
      return b.parsed.rowCount - a.parsed.rowCount;
    });
    const best = candidates[0];
    return { sheetName: best.name, matrix: best.matrix, parsed: best.parsed };
  }

  /** 재고_단가 업로드 파일에서 재고 시트만 읽고, 단가는 고정 단가.xlsx(품목명 기준)로 매핑 */
  function buildStockUnitPriceDataFromWorkbook(wb) {
    const names = wb.SheetNames || [];
    /** @type {{ name: string, parsed: ReturnType<typeof extractStockData> } | null} */
    let stockCandidate = null;
    let bestScore = -Infinity;
    for (const n of names) {
      const m = sheetToMatrix(wb, n);
      const st = extractStockData(m);
      const score = st.rowCount + (norm(n).includes("재고") ? 100000 : 0);
      if (score > bestScore) {
        bestScore = score;
        stockCandidate = { name: n, parsed: st };
      }
    }
    if (!stockCandidate || !stockCandidate.parsed || stockCandidate.parsed.rowCount <= 0) {
      return { sheetName: names[0] || null, parsed: null };
    }

    /** @type {Map<string, number>} */
    const byKey = new Map();
    /** @type {Map<string, number>} */
    const byName = new Map();
    /** @type {Map<string, number>} */
    const byNormName = new Map();
    /** @type {Map<string, { code: string, unit: number }[]>} */
    const byNameRows = new Map();
    /** @type {{ code: string, name: string, stock: number, unit: number, amount: number }[]} */
    const previewRows = [];
    let matchedCount = 0;

    for (const sr of stockCandidate.parsed.rows || []) {
      const code = String(sr.code ?? "").trim();
      const name = String(sr.name ?? "").trim();
      const unit = lookupDefaultUnitPriceByName(name);

      const stock = parseNumber(sr.stock);
      const amount = Number.isFinite(unit) ? Math.round(stock * unit) : NaN;
      previewRows.push({ code, name, stock, unit, amount });

      if (Number.isFinite(unit)) {
        matchedCount++;
        const key = `${code}\t${name}`;
        byKey.set(key, unit);
        if (!byNameRows.has(name)) byNameRows.set(name, []);
        byNameRows.get(name).push({ code, unit });
        const nk = norm(name);
        if (nk && !byNormName.has(nk)) byNormName.set(nk, unit);
      }
    }

    for (const [name, rows] of byNameRows) {
      if (rows.length === 1) byName.set(name, rows[0].unit);
    }

    const mergedParsed = {
      byKey,
      byName,
      byNormName,
      byNameRows,
      rowCount: previewRows.length,
      priceMode: "default-master-name-match",
      headerRowIndex: stockCandidate.parsed.headerRowIndex ?? -1,
      previewRows,
      matchedCount,
    };
    const mergedSheetName = `${stockCandidate.name} + 단가.xlsx`;
    return { sheetName: mergedSheetName, parsed: mergedParsed };
  }

  function renderStockUnitPricePreviewTable() {
    if (!stockUnitPriceTableWrap) return;
    if (!lastStockUnitPriceData || !lastStockUnitPriceData.previewRows || lastStockUnitPriceData.previewRows.length === 0) {
      stockUnitPriceTableWrap.innerHTML = "";
      stockUnitPriceTableWrap.hidden = true;
      return;
    }
    const allRows = lastStockUnitPriceData.previewRows;
    const upload = lastStockUnitPriceData.uploadDateLabel || "—";
    const maxShow = 800;
    const fc = stockUnitPriceCodeFilter;
    let dataRows = allRows.slice();
    if (fc.options.length > 0) {
      const totalCodes = fc.options.length;
      const sel = fc.selected.size;
      if (sel > 0 && sel < totalCodes) {
        dataRows = dataRows.filter((pr) => fc.selected.has(pr.code != null ? String(pr.code) : ""));
      }
    }

    const table = document.createElement("table");
    table.className = "stock-unit-price-table";
    const thead = document.createElement("thead");
    const trh = document.createElement("tr");
    ["날짜", "품목코드", "품목명", "재고", "재고 금액", "개당 단가"].forEach((lab) => {
      const th = document.createElement("th");
      th.textContent = lab;
      if (lab === "재고 금액") th.classList.add("col-stock-amount-head");
      trh.appendChild(th);
    });
    thead.appendChild(trh);
    table.appendChild(thead);

    const tbody = document.createElement("tbody");
    const slice = dataRows.slice(0, maxShow);
    if (dataRows.length === 0) {
      const tr = document.createElement("tr");
      const td = document.createElement("td");
      td.colSpan = 6;
      td.className = "stock-unit-price-empty-filter";
      td.textContent = "선택한 품목코드에 해당하는 행이 없습니다. 필터를 바꿔 보세요.";
      tr.appendChild(td);
      tbody.appendChild(tr);
    } else {
      for (const pr of slice) {
        const tr = document.createElement("tr");
        const td0 = document.createElement("td");
        td0.textContent = upload;
        const td1 = document.createElement("td");
        td1.textContent = pr.code;
        const td2 = document.createElement("td");
        td2.textContent = pr.name;
        td2.className = "cell-name";
        const td3 = document.createElement("td");
        td3.textContent = Number(pr.stock).toLocaleString("ko-KR");
        td3.className = "num";
        const td4 = document.createElement("td");
        td4.textContent = Number.isFinite(pr.amount) ? pr.amount.toLocaleString("ko-KR") : "—";
        td4.className = "num col-stock-amount";
        const td5 = document.createElement("td");
        td5.textContent = Number.isFinite(pr.unit)
          ? Number(pr.unit).toLocaleString("ko-KR", { maximumFractionDigits: 6 })
          : "—";
        td5.className = "num";
        tr.append(td0, td1, td2, td3, td4, td5);
        tbody.appendChild(tr);
      }
    }
    table.appendChild(tbody);

    if (dataRows.length > 0) {
      let sumStock = 0;
      let sumAmount = 0;
      for (const pr of dataRows) {
        sumStock += parseNumber(pr.stock);
        if (Number.isFinite(pr.amount)) sumAmount += pr.amount;
      }
      const tfoot = document.createElement("tfoot");
      const trf = document.createElement("tr");
      trf.className = "stock-unit-price-tfoot-row";
      const tdSum0 = document.createElement("td");
      tdSum0.textContent = "합계";
      const tdSum1 = document.createElement("td");
      tdSum1.textContent = "—";
      const tdSum2 = document.createElement("td");
      tdSum2.className = "cell-name stock-unit-price-tfoot-label";
      const nShow = Math.min(dataRows.length, maxShow);
      tdSum2.textContent =
        dataRows.length > maxShow
          ? `총합 (${dataRows.length.toLocaleString("ko-KR")}건 중 표시 ${nShow.toLocaleString("ko-KR")}건 반영)`
          : `총합 (${dataRows.length.toLocaleString("ko-KR")}건)`;
      const tdSum3 = document.createElement("td");
      tdSum3.textContent = sumStock.toLocaleString("ko-KR");
      tdSum3.className = "num";
      const tdSum4 = document.createElement("td");
      tdSum4.textContent = sumAmount.toLocaleString("ko-KR");
      tdSum4.className = "num col-stock-amount col-stock-amount-total";
      const tdSum5 = document.createElement("td");
      tdSum5.textContent = "—";
      tdSum5.className = "num";
      trf.append(tdSum0, tdSum1, tdSum2, tdSum3, tdSum4, tdSum5);
      tfoot.appendChild(trf);
      table.appendChild(tfoot);
    }

    stockUnitPriceTableWrap.innerHTML = "";
    stockUnitPriceTableWrap.appendChild(table);
    const totalAll = allRows.length;
    const nFiltered = dataRows.length;
    if (nFiltered > maxShow) {
      const meta = document.createElement("p");
      meta.className = "stock-unit-price-table-meta";
      meta.textContent = `미리보기는 필터 결과 중 처음 ${maxShow}건만 표시합니다. (표시 대상 ${nFiltered}건 / 파일 전체 ${totalAll}건)`;
      stockUnitPriceTableWrap.appendChild(meta);
    } else if (nFiltered < totalAll) {
      const meta = document.createElement("p");
      meta.className = "stock-unit-price-table-meta";
      meta.textContent = `필터 적용: ${nFiltered}건 표시 (파일 전체 ${totalAll}건 — 일별 통합표에는 전체가 반영됩니다)`;
      stockUnitPriceTableWrap.appendChild(meta);
    } else if (totalAll > maxShow) {
      const meta = document.createElement("p");
      meta.className = "stock-unit-price-table-meta";
      meta.textContent = `미리보기는 처음 ${maxShow}건만 표시합니다. (전체 ${totalAll}건 — 일별 통합표에는 모두 반영)`;
      stockUnitPriceTableWrap.appendChild(meta);
    }
    stockUnitPriceTableWrap.hidden = false;
  }

  function renderStockUnitPricePanel() {
    if (stockUnitPriceSheetInfo) {
      if (!lastStockUnitPriceData) {
        stockUnitPriceSheetInfo.textContent = "";
        resetStockUnitPriceCodeFilter();
      } else {
        const pm = lastStockUnitPriceData.priceMode;
        const isNameMatchMode =
          (pm && pm.startsWith("stock-name-match")) || pm === "default-master-name-match";
        const modeLabel =
          isNameMatchMode
            ? "재고_단가 품목명 ↔ 단가.xlsx 품목명 매칭"
            : pm === "direct"
            ? "단가 열"
            : pm === "derived-buy"
              ? "매입기준액÷현재고(역산)"
              : pm === "derived-sell"
                ? "매출기준액÷현재고(역산)"
                : "역산 단가";
        const fn = lastStockUnitPriceData.fileLabel ? `${lastStockUnitPriceData.fileLabel} · ` : "";
        stockUnitPriceSheetInfo.textContent = `${fn}시트: ${lastStockUnitPriceData.sheetName} · ${lastStockUnitPriceData.rowCount}행 · ${modeLabel}`;
      }
    }
    renderStockUnitPricePreviewTable();
  }

  /**
   * 지난발주서 시트 → 품목키별 일자→수량
   * @param {any[][]} matrix
   */
  function extractPastOrderData(matrix) {
    /** @type {Map<string, Map<string, number>>} */
    const byKey = new Map();
    if (!matrix || matrix.length < 2) return { byKey, dates: [] };

    const spec = analyzeWideHeader(matrix[0]);
    if (spec.wide) {
      const dateAcc = new Set();
      for (let r = 1; r < matrix.length; r++) {
        const row = matrix[r];
        if (!row || row.length === 0) continue;
        const rawG = String(row[spec.iGubun] ?? "").trim();
        const g = normalizeGubun(row[spec.iGubun]);
        if (rawG !== "" && g !== "지난발주") continue;

        const code = spec.iCode >= 0 ? String(row[spec.iCode] ?? "").trim() : "";
        const name = String(row[spec.iName] ?? "").trim();
        if (!code && !name) continue;

        const pkey = `${code}\t${name}`;
        if (!byKey.has(pkey)) byKey.set(pkey, new Map());
        const dm = byKey.get(pkey);
        for (const { idx, ymd } of spec.dateCols) {
          const v = parseNumber(row[idx]);
          dm.set(ymd, (dm.get(ymd) || 0) + v);
          dateAcc.add(ymd);
        }
      }
      return { byKey, dates: [...dateAcc].sort() };
    }

    const idx = guessColumnIndex(matrix[0]);
    if (idx.date < 0 || idx.name < 0 || idx.qty < 0) return { byKey, dates: [] };
    const pivot = buildPivot(matrix, idx);
    if (!pivot) return { byKey, dates: [] };
    for (const pname of pivot.products) {
      const meta = pivot.metaByProduct.get(pname) || { code: "" };
      const pkey = `${meta.code}\t${pname}`;
      const dm = new Map();
      for (const d of pivot.dates) {
        dm.set(d, pivot.byDate.get(d).get(pname) || 0);
      }
      byKey.set(pkey, dm);
    }
    return { byKey, dates: pivot.dates };
  }

  function ensureProductSubRows(pack, dateList) {
    for (const sub of SUB_ROWS) {
      if (!pack.rows.has(sub)) {
        pack.rows.set(sub, {
          dates: new Map(dateList.map((d) => [d, 0])),
          stock: 0,
          exportCol: 0,
          lackCol: 0,
        });
      } else {
        const line = pack.rows.get(sub);
        for (const d of dateList) {
          if (!line.dates.has(d)) line.dates.set(d, 0);
        }
      }
    }
  }

  function recalcExportCols(pack) {
    const orderLine = pack.rows.get("발주량");
    let orderSum = 0;
    if (orderLine) {
      for (const v of orderLine.dates.values()) orderSum += v;
    }
    for (const key of SUB_ROWS) {
      const line = pack.rows.get(key);
      if (!line) continue;
      line.exportCol = orderSum;
    }
  }

  /**
   * 차이수량 = 발주량 - 지난발주 (일자별)
   * - 같은 날짜에 그대로 유지된 물량은 0
   * - 날짜 이동이 생긴 물량만 +/-
   * @param {any} pack
   * @param {string[]} dates
   */
  function recalcDiffFromPast(pack, dates) {
    const pastLine = pack.rows.get("지난발주");
    const orderLine = pack.rows.get("발주량");
    const diffLine = pack.rows.get("차이수량");
    if (!pastLine || !orderLine || !diffLine) return;
    for (const d of dates) {
      const past = parseNumber(pastLine.dates.get(d) ?? 0);
      const order = parseNumber(orderLine.dates.get(d) ?? 0);
      diffLine.dates.set(d, order - past);
    }
  }

  function alignBoardDates(board) {
    for (const pack of board.products) {
      ensureProductSubRows(pack, board.dates);
      recalcDiffFromPast(pack, board.dates);
      recalcExportCols(pack);
    }
  }

  /**
   * @param {any} wb
   * @param {NonNullable<typeof lastBoard>} board
   */
  function mergePastOrderSheetIntoBoard(wb, board) {
    const pastName = findPastOrderSheetName(wb);
    if (!pastName) return board;

    const matrix = sheetToMatrix(wb, pastName);
    const { byKey, dates: pastDates } = extractPastOrderData(matrix);
    if (byKey.size === 0) return board;

    const dateSet = new Set(board.dates);
    for (const d of pastDates) dateSet.add(d);
    board.dates = [...dateSet].sort();

    const existing = new Set(board.products.map((p) => `${p.code}\t${p.name}`));

    for (const pack of board.products) {
      const k = `${pack.code}\t${pack.name}`;
      const incoming = byKey.get(k);
      if (!incoming) continue;
      ensureProductSubRows(pack, board.dates);
      let line = pack.rows.get("지난발주");
      if (!line) {
        line = {
          dates: new Map(board.dates.map((d) => [d, 0])),
          stock: 0,
          exportCol: 0,
          lackCol: 0,
        };
        pack.rows.set("지난발주", line);
      }
      for (const d of board.dates) {
        if (!line.dates.has(d)) line.dates.set(d, 0);
      }
      for (const [ymd, v] of incoming) {
        line.dates.set(ymd, (line.dates.get(ymd) || 0) + v);
      }
    }

    let seq = board.products.length;
    for (const [k, incoming] of byKey) {
      if (existing.has(k)) continue;
      const { code, name } = splitProductKey(k);
      if (!name && !code) continue;
      const rows = new Map();
      for (const sub of SUB_ROWS) {
        rows.set(sub, {
          dates: new Map(board.dates.map((d) => [d, 0])),
          stock: 0,
          exportCol: 0,
          lackCol: 0,
        });
      }
      const linePast = rows.get("지난발주");
      for (const [ymd, v] of incoming) {
        linePast.dates.set(ymd, (linePast.dates.get(ymd) || 0) + v);
      }
      board.products.push({
        code: code || "",
        name: name || code || "",
        type: "",
        order: seq++,
        rows,
      });
    }

    for (const pack of board.products) ensureProductSubRows(pack, board.dates);
    return board;
  }

  /**
   * 재고파일 적용:
   * - 재고 열(모든 구분행) 갱신
   * - 차이수량: 일별 잔량(이월)
   * - 부족품: 일별 부족(음수)
   * @param {NonNullable<typeof lastBoard>} board
   * @param {{ stockByKey: Map<string, number>, stockByName: Map<string, number> }} stockData
   */
  function mergeStockDataIntoBoard(board, stockData) {
    if (!stockData) return board;

    const keyToPack = new Map();
    board.products.forEach((p) => {
      keyToPack.set(`${p.code}\t${p.name}`, p);
    });

    for (const pack of board.products) {
      const key = `${pack.code}\t${pack.name}`;
      let stockV = stockData.stockByKey.get(key);
      if (!Number.isFinite(stockV)) {
        stockV = stockData.stockByName.get(pack.name);
      }
      if (!Number.isFinite(stockV)) continue;

      ensureProductSubRows(pack, board.dates);

      for (const sub of SUB_ROWS) {
        const line = pack.rows.get(sub);
        if (line) line.stock = stockV;
      }

      const orderLine = pack.rows.get("발주량");
      const lackLine = pack.rows.get("부족품");
      if (!orderLine || !lackLine) continue;

      let running = stockV;
      for (const d of board.dates) {
        const order = parseNumber(orderLine.dates.get(d) ?? 0);
        // 차이수량은 지난발주 증감으로 계산하므로 재고 머지에서는 건드리지 않는다.
        running = running - order;
        lackLine.dates.set(d, running);
      }
    }
    return board;
  }

  function applyStockDataToCurrentBoard() {
    if (!lastBoard || !lastStockData) return;
    mergeStockDataIntoBoard(lastBoard, lastStockData);
    alignBoardDates(lastBoard);
  }

  function guessColumnIndex(headerRow) {
    const idx = { date: -1, name: -1, code: -1, type: -1, qty: -1, stock: -1 };
    headerRow.forEach((h, i) => {
      const hs = String(h ?? "");
      if (idx.date < 0 && matchHeader(hs, DATE_KEYS)) idx.date = i;
      else if (idx.code < 0 && matchHeader(hs, CODE_KEYS)) idx.code = i;
      else if (idx.name < 0 && matchHeader(hs, NAME_KEYS)) idx.name = i;
      else if (idx.type < 0 && matchHeader(hs, TYPE_KEYS)) idx.type = i;
      else if (idx.qty < 0 && matchHeader(hs, QTY_KEYS)) idx.qty = i;
      else if (idx.stock < 0 && matchHeader(hs, STOCK_KEYS)) idx.stock = i;
    });
    if (idx.name < 0) {
      headerRow.forEach((h, i) => {
        const hNorm = norm(String(h ?? ""));
        if (idx.name < 0 && hNorm === "제품") idx.name = i;
      });
    }
    if (idx.qty < 0) {
      headerRow.forEach((h, i) => {
        if (idx.qty < 0 && /수량|량$|qty/i.test(String(h))) idx.qty = i;
      });
    }
    return idx;
  }

  function findHeaderIndex(headerRow, keys) {
    for (let i = 0; i < headerRow.length; i++) {
      if (matchHeader(String(headerRow[i] ?? ""), keys)) return i;
    }
    return -1;
  }

  /** 헤더 셀 문자열이 라벨과 동일한 열(부분일치·장기재고 등 혼동 방지) */
  function findExactHeaderIndex(headerRow, labels) {
    for (let i = 0; i < headerRow.length; i++) {
      const h = norm(String(headerRow[i] ?? ""));
      if (!h) continue;
      for (const lab of labels) {
        if (h === norm(lab)) return i;
      }
    }
    return -1;
  }

  /**
   * 단가 열 — `matchHeader("단가")`가 `표준단가`에도 먼저 걸려 개당단가(27.45) 대신 표준단가(150)를 읽는 경우를 막기 위해
   * 「개당단가」류를 최우선으로 한 뒤, 표준/평균, 마지막에 일반 「단가」만 사용합니다.
   */
  function findUnitPriceColumnIndex(headerRow) {
    const tier1 = ["개당단가", "개당 단가", "개별단가"];
    const tier2 = ["재고단가", "매입단가", "매출단가"];
    const tier3 = ["표준단가", "평균단가"];
    const tier4 = ["단가", "unitprice", "unit price", "price"];

    function colExcluded(raw) {
      const h = norm(String(raw ?? ""));
      if (!h) return true;
      if (h.includes("매입기준") || h.includes("매출기준")) return true;
      if (h.includes("금액") && !h.includes("단가")) return true;
      return false;
    }

    /**
     * @param {string[]} keys
     * @param {boolean} looseTier — true면 표준/평균/개당 이 들어간 열은 제외(맨 마지막 일반 「단가」용)
     */
    function findInTier(keys, looseTier) {
      for (let i = 0; i < headerRow.length; i++) {
        const raw = String(headerRow[i] ?? "").trim();
        if (colExcluded(raw)) continue;
        const h = norm(raw);
        for (const k of keys) {
          const nk = norm(k);
          if (h === nk || h.includes(nk)) {
            if (looseTier && (h.includes("표준") || h.includes("평균") || h.includes("개당"))) continue;
            return i;
          }
        }
      }
      return -1;
    }

    let i = findInTier(tier1, false);
    if (i >= 0) return i;
    i = findInTier(tier2, false);
    if (i >= 0) return i;
    i = findInTier(tier3, false);
    if (i >= 0) return i;
    i = findInTier(tier4, true);
    return i;
  }

  function findCurrentStockColumnForPricing(headerRow) {
    let i = findExactHeaderIndex(headerRow, ["현재고"]);
    if (i >= 0) return i;
    i = findExactHeaderIndex(headerRow, ["재고수량", "재고량"]);
    if (i >= 0) return i;
    i = findExactHeaderIndex(headerRow, ["재고"]);
    if (i >= 0) return i;
    return -1;
  }

  function findBuyBasisAmountColumn(headerRow) {
    const i = findExactHeaderIndex(headerRow, ["매입기준액"]);
    if (i >= 0) return i;
    return findHeaderIndex(headerRow, ["매입기준액"]);
  }

  function parseUnitPrice(v) {
    if (v == null || v === "") return NaN;
    if (typeof v === "number" && Number.isFinite(v)) return v;
    const s = String(v).replace(/,/g, "").trim();
    if (!s) return NaN;
    const n = parseFloat(s);
    return Number.isFinite(n) ? n : NaN;
  }

  /** 고정 단가 마스터(단가.xlsx)에서 품목명 기준 단가 조회 */
  function lookupDefaultUnitPriceByName(name) {
    const n = String(name ?? "").trim();
    if (!n) return NaN;
    const k = norm(n);
    if (!k) return NaN;
    return DEFAULT_UNIT_PRICE_BY_NORM_NAME.has(k) ? DEFAULT_UNIT_PRICE_BY_NORM_NAME.get(k) : NaN;
  }

  function lookupStockUnitPrice(code, name, data) {
    if (!data) return NaN;
    const c = String(code ?? "").trim();
    const n = String(name ?? "").trim();
    if (
      String(data.priceMode || "").startsWith("stock-name-match") ||
      String(data.priceMode || "") === "default-master-name-match"
    ) {
      const nn = norm(n);
      if (nn && data.byNormName && data.byNormName.has(nn)) return data.byNormName.get(nn);
      if (n && data.byName.has(n)) return data.byName.get(n);
      if (nn && data.byNameRows) {
        for (const [nameKey, list] of data.byNameRows) {
          if (norm(nameKey) !== nn || !list || list.length === 0) continue;
          const u = list[0].unit;
          if (Number.isFinite(u)) return u;
        }
      }
      const fromDefault = lookupDefaultUnitPriceByName(n);
      if (Number.isFinite(fromDefault)) return fromDefault;
      return NaN;
    }
    const k = `${c}\t${n}`;
    if (data.byKey.has(k)) return data.byKey.get(k);
    if (n && data.byName.has(n)) return data.byName.get(n);
    const nn = norm(n);
    if (nn && data.byNormName && data.byNormName.has(nn)) return data.byNormName.get(nn);
    const rows = data.byNameRows && data.byNameRows.get(n);
    if (rows && rows.length > 0) {
      if (c) {
        const hit = rows.find((r) => String(r.code ?? "").trim() === c);
        if (hit && Number.isFinite(hit.unit)) return hit.unit;
      }
      if (rows.length === 1 && Number.isFinite(rows[0].unit)) return rows[0].unit;
    }
    if (n && data.byNameRows) {
      const nn = norm(n);
      for (const [nameKey, list] of data.byNameRows) {
        if (norm(nameKey) !== nn) continue;
        if (c) {
          const hit = list.find((r) => String(r.code ?? "").trim() === c);
          if (hit && Number.isFinite(hit.unit)) return hit.unit;
        }
        if (list.length === 1 && Number.isFinite(list[0].unit)) return list[0].unit;
      }
    }
    const fromDefault = lookupDefaultUnitPriceByName(n);
    if (Number.isFinite(fromDefault)) return fromDefault;
    return NaN;
  }

  /** 엑셀 단가가 없을 때 팀(타입)별 참조 개당 단가 */
  function lookupBuiltinStockUnitPrice(code, name, typeStr) {
    const seg = inferDrawingTeamSegmentForPricing(typeStr, name);
    if (!seg) return NaN;
    const v = BUILTIN_STOCK_UNIT_PRICE_WON_BY_TEAM[seg];
    return Number.isFinite(v) ? v : NaN;
  }

  /** 재고_단가 파일 → 없으면 내장 참조 단가(팀별) */
  function resolveStockUnitPriceWon(code, name, typeStr, data) {
    const fromFile = lookupStockUnitPrice(code, name, data);
    if (Number.isFinite(fromFile)) return fromFile;
    // 재고 시트→단가 시트 매칭 모드에서는 내장 단가로 덮어쓰지 않는다.
    if (
      data &&
      (String(data.priceMode || "").startsWith("stock-name-match") ||
        String(data.priceMode || "") === "default-master-name-match")
    )
      return NaN;
    return lookupBuiltinStockUnitPrice(code, name, typeStr);
  }

  function stockAmountFromLine(pack, line) {
    const unit = resolveStockUnitPriceWon(pack.code, pack.name, pack.type, lastStockUnitPriceData);
    const st = line.stock;
    if (!Number.isFinite(unit) || !Number.isFinite(st)) return NaN;
    return Math.round(st * unit);
  }

  /** 제품코드/품목코드 열이 '제품' 부분일치로 이름 열로 잡히지 않도록 */
  function findProductNameColumnIndex(headerRow) {
    for (let i = 0; i < headerRow.length; i++) {
      const h = norm(String(headerRow[i] ?? ""));
      if (!h) continue;
      if (h.includes("품목코드") || h.includes("자재코드") || h.includes("상품코드")) continue;
      if (/(코드|code)$/.test(h) || h.endsWith("코드")) continue;
      if (matchHeader(String(headerRow[i] ?? ""), NAME_KEYS)) return i;
      if (h === "제품") return i;
    }
    return -1;
  }

  /**
   * 단가 시트에서 제품명 헤더가 비어 있는 경우(예: 품번 다음 빈 헤더 열) 데이터 샘플로 제품명 열을 추정
   * @param {any[][]} matrix
   * @param {number} hi
   * @param {number[]} avoidCols
   */
  function inferProductNameColumnFromRows(matrix, hi, avoidCols) {
    const header = matrix[hi] || [];
    const avoid = new Set((avoidCols || []).filter((x) => Number.isInteger(x) && x >= 0));
    const scanEnd = Math.min(matrix.length, hi + 1 + 120);
    const maxCol = Math.max(header.length, ...matrix.slice(hi + 1, scanEnd).map((r) => (r ? r.length : 0)));
    let bestIdx = -1;
    let bestScore = -Infinity;

    for (let c = 0; c < maxCol; c++) {
      if (avoid.has(c)) continue;
      const hRaw = String(header[c] ?? "").trim();
      const h = norm(hRaw);
      if (h && (matchHeader(hRaw, CODE_KEYS) || matchHeader(hRaw, ["단가", "금액", "재고", "수량"]))) continue;

      let nonEmpty = 0;
      let textish = 0;
      let numericish = 0;
      for (let r = hi + 1; r < scanEnd; r++) {
        const row = matrix[r] || [];
        const raw = row[c];
        const s = String(raw ?? "").trim();
        if (!s) continue;
        nonEmpty++;
        const numericOnly = /^-?\d+(\.\d+)?$/.test(s.replace(/,/g, ""));
        if (numericOnly) numericish++;
        else textish++;
      }
      if (nonEmpty < 3) continue;
      const score = textish * 2 - numericish + (hRaw ? 0 : 1.2);
      if (score > bestScore) {
        bestScore = score;
        bestIdx = c;
      }
    }
    return bestIdx;
  }

  function fillSelect(sel, cols, selectedIndex) {
    sel.innerHTML = "";
    const optEmpty = document.createElement("option");
    optEmpty.value = "-1";
    optEmpty.textContent = "(없음)";
    sel.appendChild(optEmpty);
    cols.forEach((label, i) => {
      const o = document.createElement("option");
      o.value = String(i);
      o.textContent = label || `열 ${i + 1}`;
      sel.appendChild(o);
    });
    sel.value = String(selectedIndex >= 0 ? selectedIndex : -1);
  }

  function parseNumber(v) {
    if (v == null || v === "") return 0;
    if (typeof v === "number" && !Number.isNaN(v)) return v;
    const s = String(v).replace(/,/g, "").trim();
    const n = parseFloat(s);
    return Number.isFinite(n) ? n : 0;
  }

  function parseExcelDate(v) {
    if (v instanceof Date && !Number.isNaN(v.getTime())) {
      return formatYmd(v);
    }
    if (typeof v === "number" && v > 20000 && v < 65000) {
      const epoch = new Date(Date.UTC(1899, 11, 30));
      const d = new Date(epoch.getTime() + v * 86400000);
      return formatYmd(d);
    }
    const s = String(v ?? "").trim();
    if (!s) return "";
    if (/^\d{5}(\.\d+)?$/.test(s)) {
      const n = parseFloat(s);
      if (n > 20000 && n < 65000) {
        const epoch = new Date(Date.UTC(1899, 11, 30));
        const d = new Date(epoch.getTime() + n * 86400000);
        return formatYmd(d);
      }
    }
    const d = new Date(s);
    if (!Number.isNaN(d.getTime())) return formatYmd(d);
    return s;
  }

  function formatYmd(d) {
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, "0");
    const day = String(d.getDate()).padStart(2, "0");
    return `${y}-${m}-${day}`;
  }

  /** @param {string} ymd */
  function ymdToDisplay(ymd) {
    const p = String(ymd).split("-");
    if (p.length !== 3) return ymd;
    const y = parseInt(p[0], 10);
    const mo = parseInt(p[1], 10) - 1;
    const da = parseInt(p[2], 10);
    const d = new Date(y, mo, da);
    if (Number.isNaN(d.getTime())) return ymd;
    const w = ["일", "월", "화", "수", "목", "금", "토"][d.getDay()];
    return `${String(mo + 1).padStart(2, "0")}/${String(da).padStart(2, "0")} (${w})`;
  }

  /** @param {string|number|Date} cell */
  function headerCellToYmd(cell) {
    if (cell instanceof Date && !Number.isNaN(cell.getTime())) return formatYmd(cell);
    if (typeof cell === "number" && cell > 25000 && cell < 65000) return parseExcelDate(cell);
    const str = String(cell ?? "").trim();
    if (!str) return null;
    const mFull = str.match(/(\d{4})[./-](\d{1,2})[./-](\d{1,2})/);
    if (mFull) return `${mFull[1]}-${mFull[2].padStart(2, "0")}-${mFull[3].padStart(2, "0")}`;
    const mShort = str.match(/(\d{1,2})\s*[./]\s*(\d{1,2})/);
    if (mShort) {
      const y = 2026;
      return `${y}-${mShort[1].padStart(2, "0")}-${mShort[2].padStart(2, "0")}`;
    }
    const tryDate = new Date(str);
    if (!Number.isNaN(tryDate.getTime())) return formatYmd(tryDate);
    return null;
  }

  function isWeekendYmd(ymd) {
    const p = String(ymd).split("-");
    if (p.length !== 3) return false;
    const d = new Date(parseInt(p[0], 10), parseInt(p[1], 10) - 1, parseInt(p[2], 10));
    if (Number.isNaN(d.getTime())) return false;
    const day = d.getDay();
    return day === 0 || day === 6;
  }

  /**
   * 일별 통합표는 오늘~한 달 범위만 우선 노출
   * 범위 내 날짜가 있으면 연속 일자(토/일 포함)를 모두 만든다.
   * 범위 내 날짜가 없으면 기존 데이터 구간(min~max)의 연속 일자를 만든다.
   * @param {string[]} dates
   */
  function getDailyVisibleDates(dates) {
    if (!dates || dates.length === 0) return [];
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const end = new Date(today);
    end.setMonth(end.getMonth() + 1);

    /**
     * @param {Date} start
     * @param {Date} finish
     */
    function buildContinuousDates(start, finish) {
      const out = [];
      const cur = new Date(start);
      cur.setHours(0, 0, 0, 0);
      const fin = new Date(finish);
      fin.setHours(0, 0, 0, 0);
      while (cur <= fin) {
        out.push(formatYmd(cur));
        cur.setDate(cur.getDate() + 1);
      }
      return out;
    }

    const parsed = dates
      .map((ymd) => {
        const p = String(ymd).split("-");
        if (p.length !== 3) return null;
        const d = new Date(parseInt(p[0], 10), parseInt(p[1], 10) - 1, parseInt(p[2], 10));
        return Number.isNaN(d.getTime()) ? null : d;
      })
      .filter(Boolean)
      .sort((a, b) => a - b);

    const hasInRange = parsed.some((d) => d >= today && d <= end);
    if (hasInRange) {
      return buildContinuousDates(today, end);
    }

    const minD = parsed[0];
    const maxD = parsed[parsed.length - 1];
    if (!minD || !maxD) return dates;
    return buildContinuousDates(minD, maxD);
  }

  function normalizeGubun(raw) {
    const t = String(raw ?? "").trim();
    if (!t) return "";
    if (/지난/.test(t)) return "지난발주";
    if (/부족품/.test(t)) return "부족품";
    if (/차이/.test(t)) return "차이수량";
    if (t === "발주량" || /^발주량/.test(t)) return "발주량";
    if (t === "제품" || /^제품/.test(t)) return "발주량";
    return t;
  }

  function sheetToMatrix(workbook, sheetName) {
    const sheet = workbook.Sheets[sheetName];
    return XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: true });
  }

  function loadArrayBuffer(file) {
    return new Promise((resolve, reject) => {
      const r = new FileReader();
      r.onload = () => resolve(r.result);
      r.onerror = () => reject(r.error);
      r.readAsArrayBuffer(file);
    });
  }


  function formatPercent(v) {
    if (!Number.isFinite(v)) return "—";
    const a = Math.abs(v);
    if (a === 0) return "0.0%";
    if (a < 0.01) return `${v.toFixed(3)}%`;
    if (a < 1) return `${v.toFixed(2)}%`;
    return `${v.toFixed(1)}%`;
  }

  function detectHeaderRowIndex(matrix) {
    let bestIdx = 0;
    let bestScore = -1;
    const KEYSETS = [
      DATE_KEYS,
      ["가동율", "가동률"],
      ["설비가동율", "설비가동률"],
      ["설비효율", "oee", "효율"],
      ["불량율", "불량률", "불량수량", "불량수"],
      ["팀", "라인", "조"],
    ];
    const maxScan = Math.min(matrix.length, 8);
    for (let r = 0; r < maxScan; r++) {
      const row = matrix[r] || [];
      let score = 0;
      for (const ks of KEYSETS) {
        if (row.some((c) => matchHeader(String(c ?? ""), ks))) score++;
      }
      if (score > bestScore) {
        bestScore = score;
        bestIdx = r;
      }
    }
    return bestIdx;
  }

  /** 시트 행에서 값이 있는 최대 열(0-based 폭). 헤더가 희소 배열일 때 length만으로는 R~V·AK 등이 빠짐 */
  function matrixRowMaxUsedCol(row) {
    if (!row || !row.length) return 0;
    let m = row.length;
    for (let i = row.length - 1; i >= 0; i--) {
      const c = row[i];
      if (c !== undefined && c !== null && String(c).trim() !== "") {
        m = Math.max(m, i + 1);
        break;
      }
    }
    for (let i = 0; i < row.length; i++) {
      const c = row[i];
      if (c !== undefined && c !== null && String(c).trim() !== "") m = Math.max(m, i + 1);
    }
    return m;
  }

  /** 화승 이중관 본표: 헤더+데이터 샘플로 실제 열 폭 추정 */
  function inferHwaseungDoublePipeLayoutWidth(matrix, hi) {
    let w = 0;
    const scanEnd = Math.min(matrix.length, hi + 1 + 120);
    for (let r = 0; r < scanEnd; r++) w = Math.max(w, matrixRowMaxUsedCol(matrix[r]));
    return w;
  }

  /** 피벗/보조 시트가 아닌 본 작업일지 헤더 행을 고르기 위한 점수 */
  function findHwaseungDoublePipeHeaderRowIndex(matrix) {
    const maxScan = Math.min(matrix.length, 30);
    let bestRi = 0;
    let bestScore = -Infinity;
    for (let r = 0; r < maxScan; r++) {
      const row = matrix[r] || [];
      const width = matrixRowMaxUsedCol(row);
      const iWd = findHeaderIndex(row, ["작업일자", "작업일", ...DATE_KEYS]);
      let score = 0;
      if (iWd >= 0) score += 20;
      if (findHeaderIndex(row, ["규격"]) >= 0) score += 5;
      if (findHeaderIndex(row, ["공정"]) >= 0) score += 4;
      if (
        findHeaderIndex(row, [
          "생산량(ERP)",
          "생산량（ERP）",
          "생산량(erp)",
          "생산수량",
          "생산 수량",
          "생산량",
        ]) >= 0
      )
        score += 6;
      if (width >= 44) score += 8;
      else if (width >= 30) score += 4;
      else if (width >= 15) score += 1;
      const numSerialLike = row.slice(0, 4).filter((c) => typeof c === "number" && c > 40000 && c < 55000).length;
      if (numSerialLike >= 2) score -= 25;
      if (width <= 4 && iWd < 0) score -= 10;
      if (score > bestScore) {
        bestScore = score;
        bestRi = r;
      }
    }
    return bestRi;
  }

  function findCol(headerRow, keys) {
    for (let i = 0; i < headerRow.length; i++) {
      if (matchHeader(String(headerRow[i] ?? ""), keys)) return i;
    }
    return -1;
  }


  /**
   * @param {any[]} headerRow
   */
  function analyzeWideHeader(headerRow) {
    const iGubun = findHeaderIndex(headerRow, GUBUN_KEYS);
    const iCode = findHeaderIndex(headerRow, CODE_KEYS);
    const iName = findProductNameColumnIndex(headerRow);
    const iType = findHeaderIndex(headerRow, TYPE_KEYS);
    const iStock = findHeaderIndex(headerRow, STOCK_KEYS);
    const iExport = findHeaderIndex(headerRow, EXPORT_KEYS);
    const iLack = findHeaderIndex(headerRow, LACK_KEYS);

    /** @type {{ idx: number, ymd: string }[]} */
    const dateCols = [];
    for (let j = 0; j < headerRow.length; j++) {
      const ymd = headerCellToYmd(headerRow[j]);
      if (ymd) dateCols.push({ idx: j, ymd });
    }

    const wide = iGubun >= 0 && iName >= 0 && dateCols.length >= 2;

    return { wide, iGubun, iCode, iName, iType, iStock, iExport, iLack, dateCols };
  }

  /**
   * @param {any[][]} matrix
   * @param {ReturnType<typeof analyzeWideHeader>} spec
   */
  function buildBoardFromWide(matrix, spec) {
    /** @type {Map<string, { code: string, name: string, type: string, order: number, rows: Map<string, { dates: Map<string, number>, stock: number, exportCol: number, lackCol: number }> }>} */
    const products = new Map();
    let seq = 0;

    for (let r = 1; r < matrix.length; r++) {
      const row = matrix[r];
      if (!row || row.length === 0) continue;
      const code = spec.iCode >= 0 ? String(row[spec.iCode] ?? "").trim() : "";
      const name = String(row[spec.iName] ?? "").trim();
      const gubun = normalizeGubun(row[spec.iGubun]);
      if (!code && !name) continue;
      if (!SUB_ROWS.includes(gubun)) continue;

      const key = `${code}\t${name}`;
      if (!products.has(key)) {
        products.set(key, { code, name, type: "", order: seq++, rows: new Map() });
      }
      const pack = products.get(key);
      if (spec.iType >= 0) {
        const typ = String(row[spec.iType] ?? "").trim();
        if (typ && !pack.type) pack.type = typ;
      }
      if (!pack.rows.has(gubun)) {
        pack.rows.set(gubun, {
          dates: new Map(),
          stock: spec.iStock >= 0 ? parseNumber(row[spec.iStock]) : 0,
          exportCol: spec.iExport >= 0 ? parseNumber(row[spec.iExport]) : 0,
          lackCol: spec.iLack >= 0 ? parseNumber(row[spec.iLack]) : 0,
        });
      }
      const line = pack.rows.get(gubun);
      if (spec.iStock >= 0) line.stock = parseNumber(row[spec.iStock]);
      if (spec.iExport >= 0) line.exportCol = parseNumber(row[spec.iExport]);
      if (spec.iLack >= 0) line.lackCol = parseNumber(row[spec.iLack]);

      for (const { idx, ymd } of spec.dateCols) {
        const v = parseNumber(row[idx]);
        line.dates.set(ymd, (line.dates.get(ymd) || 0) + v);
      }
    }

    const dateSet = new Set(spec.dateCols.map((d) => d.ymd));
    const dates = [...dateSet].sort();

    const list = [...products.values()].sort((a, b) => a.order - b.order);
    for (const pack of list) ensureProductSubRows(pack, dates);
    return { mode: "wide", dates, products: list, spec };
  }

  function buildPivot(matrix, idx) {
    if (idx.date < 0 || idx.name < 0 || idx.qty < 0) return null;

    const byDate = new Map();
    const metaByProduct = new Map();

    for (let r = 1; r < matrix.length; r++) {
      const row = matrix[r];
      if (!row || row.length === 0) continue;
      const dateStr = parseExcelDate(row[idx.date]);
      const pname = String(row[idx.name] ?? "").trim();
      if (!dateStr || !pname) continue;
      const code = idx.code >= 0 ? String(row[idx.code] ?? "").trim() : "";
      const type = idx.type >= 0 ? String(row[idx.type] ?? "").trim() : "";
      const qty = parseNumber(row[idx.qty]);
      const stock = idx.stock >= 0 ? parseNumber(row[idx.stock]) : 0;

      if (!byDate.has(dateStr)) byDate.set(dateStr, new Map());
      const rowMap = byDate.get(dateStr);
      rowMap.set(pname, (rowMap.get(pname) || 0) + qty);

      if (!metaByProduct.has(pname)) {
        metaByProduct.set(pname, { code, type, stock });
      } else {
        const m = metaByProduct.get(pname);
        if (!m.code && code) m.code = code;
        if (!m.type && type) m.type = type;
        if (!m.stock && stock) m.stock = stock;
      }
    }

    const dates = [...byDate.keys()].sort();
    const products = [...metaByProduct.keys()].sort((a, b) => a.localeCompare(b, "ko"));
    return { dates, products, byDate, metaByProduct };
  }

  /**
   * @param {NonNullable<ReturnType<typeof buildPivot>>} pivot
   * @param {ReturnType<typeof readIndices>} idx
   */
  function buildBoardFromLong(pivot, idx) {
    /** @type {{ code: string, name: string, type: string, order: number, rows: Map<string, { dates: Map<string, number>, stock: number, exportCol: number, lackCol: number }> }[]} */
    const list = [];
    let order = 0;
    for (const name of pivot.products) {
      const meta = pivot.metaByProduct.get(name) || { code: "", type: "", stock: 0 };
      const rows = new Map();
      const 발주Dates = new Map();
      for (const d of pivot.dates) {
        const rowMap = pivot.byDate.get(d);
        const v = rowMap.get(name) || 0;
        발주Dates.set(d, v);
      }
      let exportSum = 0;
      for (const v of 발주Dates.values()) exportSum += v;
      rows.set("발주량", {
        dates: 발주Dates,
        stock: meta.stock || 0,
        exportCol: exportSum,
        lackCol: 0,
      });
      for (const sub of SUB_ROWS) {
        if (sub === "발주량") continue;
        rows.set(sub, {
          dates: new Map(pivot.dates.map((d) => [d, 0])),
          stock: meta.stock || 0,
          exportCol: 0,
          lackCol: 0,
        });
      }
      list.push({ code: meta.code, name, type: meta.type || "", order: order++, rows });
    }
    return { mode: "long", dates: pivot.dates, products: list, spec: null };
  }

  /**
   * @param {NonNullable<typeof lastBoard>} board
   */
  function syncOrderCalendarMonthFromBoard(board) {
    if (!orderCalendarNeedsSync || !board || !board.dates.length) return;
    const first = board.dates[0];
    const p = String(first).split("-");
    if (p.length === 3) {
      const y = parseInt(p[0], 10);
      const mo = parseInt(p[1], 10) - 1;
      orderCalendarCursor.setFullYear(y, mo, 1);
    }
    orderCalendarNeedsSync = false;
  }

  function hideOrderCalendarMonthSummary() {
    if (!orderCalendarMonthSummary) return;
    orderCalendarMonthSummary.hidden = true;
    orderCalendarMonthSummary.innerHTML = "";
  }

  /**
   * 현재 달력 월(y, mo) 기준, 필터된 품목의 발주량 일별 합 → 제품별 합계 블록
   * @param {number} y
   * @param {number} mo 0–11
   * @param {ReturnType<typeof getFilteredProductPacks>} packs
   */
  function renderOrderCalendarMonthSummary(y, mo, packs) {
    if (!orderCalendarMonthSummary) return;
    const lastDate = new Date(y, mo + 1, 0).getDate();
    /** @type {{ name: string, qty: number }[]} */
    const rows = [];
    for (const pack of packs) {
      const line = pack.rows.get("발주량");
      if (!line) continue;
      let sum = 0;
      for (let day = 1; day <= lastDate; day++) {
        const ymd = formatYmd(new Date(y, mo, day));
        sum += parseNumber(line.dates.get(ymd));
      }
      if (sum !== 0) rows.push({ name: pack.name, qty: sum });
    }
    rows.sort((a, b) => a.name.localeCompare(b.name, "ko"));

    orderCalendarMonthSummary.innerHTML = "";
    const title = document.createElement("div");
    title.className = "order-calendar-month-summary__title";
    title.textContent = `${y}년 ${mo + 1}월 제품별 발주량 합`;
    orderCalendarMonthSummary.appendChild(title);

    const list = document.createElement("div");
    list.className = "order-calendar-month-summary__list";
    if (rows.length === 0) {
      const empty = document.createElement("p");
      empty.className = "order-calendar-month-summary__empty";
      empty.textContent = "이 달에 표시 중인 발주량이 없습니다.";
      list.appendChild(empty);
    } else {
      for (const r of rows) {
        const row = document.createElement("div");
        row.className = "order-calendar-month-summary__row";
        const ns = document.createElement("span");
        ns.className = "order-calendar-month-summary__name";
        ns.textContent = r.name;
        const qs = document.createElement("span");
        qs.className = "order-calendar-month-summary__qty";
        qs.textContent = formatQty(r.qty);
        row.append(ns, qs);
        list.appendChild(row);
      }
    }
    orderCalendarMonthSummary.appendChild(list);
    orderCalendarMonthSummary.hidden = false;
  }

  /**
   * 일별 통합표와 동일 필터로 「발주량」만 월 달력 셀에 제품명·수량 표시
   * @param {typeof lastBoard} board
   */
  function renderOrderCalendar(board) {
    if (!orderCalendarPanel || !orderCalendarGrid || !orderCalMonthLabel) return;
    syncOrderCalendarTypeUi();

    if (!board || board.products.length === 0) {
      orderCalendarPanel.hidden = true;
      orderCalendarGrid.innerHTML = "";
      orderCalMonthLabel.textContent = "—";
      hideOrderCalendarMonthSummary();
      return;
    }

    orderCalendarPanel.hidden = false;
    syncOrderCalendarMonthFromBoard(board);

    const y = orderCalendarCursor.getFullYear();
    const mo = orderCalendarCursor.getMonth();
    orderCalMonthLabel.textContent = `${y}년 ${mo + 1}월`;

    const subs = getFilteredSubs();
    const exportSet = filterState.export.selected;
    const exportAll = exportSet.size === filterState.export.options.length;

    if (!subs.includes("발주량")) {
      hideOrderCalendarMonthSummary();
      orderCalendarGrid.innerHTML = "";
      orderCalendarGrid.className = "order-calendar-grid order-calendar-grid--message";
      const note = document.createElement("div");
      note.className = "order-calendar-note";
      note.textContent = "구분 필터에 「발주량」을 포함하면 날짜별로 제품명과 발주량이 표시됩니다.";
      orderCalendarGrid.appendChild(note);
      return;
    }

    const packs = getFilteredProductPacks(board).filter((pack) => {
      const line = pack.rows.get("발주량");
      if (!line) return false;
      if (!exportAll && !exportSet.has(String(Number(line.exportCol || 0)))) return false;
      return true;
    });

    const firstDow = new Date(y, mo, 1).getDay();
    const lastDate = new Date(y, mo + 1, 0).getDate();
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const todayYmd = formatYmd(today);

    renderOrderCalendarMonthSummary(y, mo, packs);

    orderCalendarGrid.innerHTML = "";
    orderCalendarGrid.className = "order-calendar-grid";

    for (let i = 0; i < firstDow; i++) {
      const pad = document.createElement("div");
      pad.className = "order-calendar-cell order-calendar-cell--pad";
      orderCalendarGrid.appendChild(pad);
    }

    for (let day = 1; day <= lastDate; day++) {
      const d = new Date(y, mo, day);
      const ymd = formatYmd(d);
      const cell = document.createElement("div");
      cell.className = "order-calendar-cell";
      cell.dataset.ymd = ymd;
      cell.setAttribute("role", "button");
      cell.tabIndex = 0;
      cell.setAttribute("aria-pressed", orderCalendarSelectedYmd === ymd ? "true" : "false");
      cell.setAttribute("aria-label", `${ymd} 발주`);
      if (isWeekendYmd(ymd)) cell.classList.add("order-calendar-cell--weekend");
      if (ymd === todayYmd) cell.classList.add("order-calendar-cell--today");
      if (orderCalendarSelectedYmd === ymd) cell.classList.add("order-calendar-cell--selected");

      const dayNum = document.createElement("div");
      dayNum.className = "order-calendar-daynum";
      dayNum.textContent = String(day);
      cell.appendChild(dayNum);

      /** @type {{ name: string, qty: number }[]} */
      const items = [];
      for (const pack of packs) {
        const line = pack.rows.get("발주량");
        if (!line) continue;
        const v = parseNumber(line.dates.get(ymd));
        if (v !== 0) items.push({ name: pack.name, qty: v });
      }
      items.sort((a, b) => a.name.localeCompare(b.name, "ko"));
      if (items.length === 0) cell.classList.add("order-calendar-cell--empty");

      const list = document.createElement("div");
      list.className = "order-calendar-items";
      for (const it of items) {
        const row = document.createElement("div");
        row.className = "order-calendar-item";
        const ns = document.createElement("span");
        ns.className = "order-calendar-item-name";
        ns.textContent = it.name;
        const qs = document.createElement("span");
        qs.className = "order-calendar-item-qty";
        qs.textContent = formatQty(it.qty);
        row.append(ns, qs);
        list.appendChild(row);
      }
      cell.appendChild(list);
      orderCalendarGrid.appendChild(cell);
    }
  }

  /**
   * @param {NonNullable<typeof lastBoard>} board
   */
  function renderBoard(board) {
    emptyState.hidden = true;
    const old = tableWrap.querySelector("table.board-table");
    if (old) old.remove();

    const visibleDates = getDailyVisibleDates(board.dates);
    const packs = getFilteredProductPacks(board);
    const subs = getFilteredSubs();
    const totalProd = board.products.length;
    const exportSet = filterState.export.selected;
    const exportAll = exportSet.size === filterState.export.options.length;
    const rowsData = [];
    const shownProductKeys = new Set();
    let groupCounter = 0;

    const table = document.createElement("table");
    table.className = "board-table";

    const thead = document.createElement("thead");
    const trh = document.createElement("tr");
    const fixedLabels = ["구분", "품목코드", "제품", "타입", "재고", "수출발주량", "부족량"];
    fixedLabels.forEach((label, i) => {
      const th = document.createElement("th");
      th.textContent = label;
      th.className = `sticky-left sticky-c${i + 1}` + (i >= 4 ? " col-summary" : "");
      trh.appendChild(th);
    });
    visibleDates.forEach((ymd) => {
      const th = document.createElement("th");
      th.textContent = ymdToDisplay(ymd);
      if (isWeekendYmd(ymd)) th.classList.add("col-weekend");
      trh.appendChild(th);
    });
    thead.appendChild(trh);
    table.appendChild(thead);

    const tbody = document.createElement("tbody");
    if (packs.length === 0) {
      const tr = document.createElement("tr");
      const td = document.createElement("td");
      td.colSpan = fixedLabels.length + visibleDates.length;
      td.className = "empty-filter-msg";
      td.textContent = "조건에 맞는 품목이 없습니다. 필터를 바꿔 보세요.";
      tr.appendChild(td);
      tbody.appendChild(tr);
    } else {
      for (const pack of packs) {
        const packRows = [];
        for (const sub of subs) {
          const line =
            pack.rows.get(sub) ||
            ({
              dates: new Map(),
              stock: 0,
              exportCol: 0,
              lackCol: 0,
            });
          if (!exportAll && !exportSet.has(String(Number(line.exportCol || 0)))) continue;
          packRows.push({ pack, sub, line });
        }
        if (packRows.length > 0) {
          const groupOdd = groupCounter % 2 === 1;
          packRows.forEach((row, idx) => {
            rowsData.push({
              ...row,
              groupStart: idx === 0,
              groupEnd: idx === packRows.length - 1,
              groupOdd,
            });
          });
          shownProductKeys.add(`${pack.code}\t${pack.name}`);
          groupCounter++;
        }
      }

      for (const row of rowsData) {
        const tr = document.createElement("tr");
        tr.classList.add("group-row");
        if (row.groupOdd) tr.classList.add("group-odd");
        if (row.groupStart) tr.classList.add("group-start");
        if (row.groupEnd) tr.classList.add("group-end");
        const pack = row.pack;
        const sub = row.sub;
        const line = row.line;

        const tdG = document.createElement("td");
        tdG.textContent = sub;
        tdG.className = "sticky-left sticky-c1";
        tr.appendChild(tdG);

        const tdC = document.createElement("td");
        tdC.textContent = pack.code;
        tdC.className = "sticky-left sticky-c2";
        tr.appendChild(tdC);

        const tdN = document.createElement("td");
        tdN.textContent = pack.name;
        tdN.className = "sticky-left sticky-c3 product-name";
        tr.appendChild(tdN);

        const tdT = document.createElement("td");
        tdT.textContent = pack.type || "—";
        tdT.className = "sticky-left sticky-c4 type-cell";
        tr.appendChild(tdT);

        const tdS = document.createElement("td");
        tdS.textContent = formatQty(line.stock);
        tdS.className = "sticky-left sticky-c5 col-summary num";
        tr.appendChild(tdS);

        const tdE = document.createElement("td");
        tdE.textContent = formatQty(line.exportCol);
        tdE.className = "sticky-left sticky-c6 col-summary num";
        tr.appendChild(tdE);

        const tdL = document.createElement("td");
        tdL.textContent = formatQty(line.lackCol);
        tdL.className = "sticky-left sticky-c7 col-summary num";
        if (line.lackCol < 0) tdL.classList.add("neg");
        tr.appendChild(tdL);

        visibleDates.forEach((ymd) => {
          const td = document.createElement("td");
          const v = line.dates.get(ymd) ?? 0;
          td.className = "num";
          if (sub === "차이수량" && v !== 0) {
            const mark = v > 0 ? "↑" : "↓";
            td.innerHTML = `<span class="diff-arrow ${v > 0 ? "up" : "down"}">${mark}</span> <span class="diff-value">${formatQty(
              Math.abs(v)
            )}</span>`;
            td.classList.add("cell-diff", v > 0 ? "cell-diff-up" : "cell-diff-down");
          } else {
            td.textContent = formatQty(v);
          }
          if (isWeekendYmd(ymd)) td.classList.add("col-weekend");
          if (v === 0) td.classList.add("zero");
          if ((sub === "지난발주" || sub === "발주량") && v > 0) td.classList.add("cell-order");
          if (sub === "부족품" && v < 0) td.classList.add("cell-short");
          tr.appendChild(td);
        });
        tbody.appendChild(tr);
      }
    }
    table.appendChild(tbody);
    tableWrap.appendChild(table);

    const nShow = shownProductKeys.size;
    const nRows = rowsData.length;
    rowCountEl.textContent =
      totalProd > 0
        ? `표시 ${nRows}행 · ${nShow}품목 (전체 ${totalProd}품목) · ${visibleDates.length}일`
        : `0건`;
    btnExport.disabled = totalProd === 0 || nRows === 0;
    if (currentView === "orderCalendar") renderOrderCalendar(board);
  }

  /**
   * @param {NonNullable<typeof lastBoard>} board
   */
  function renderSimpleTable(board) {
    const old = simpleTableWrap.querySelector("table.simple-table");
    if (old) old.remove();
    if (!board || board.products.length === 0) {
      emptyStateSimple.hidden = false;
      return;
    }
    emptyStateSimple.hidden = true;

    const table = document.createElement("table");
    table.className = "simple-table";
    const thead = document.createElement("thead");
    const trh = document.createElement("tr");
    ["품목코드", "제품"].forEach((label) => {
      const th = document.createElement("th");
      th.textContent = label;
      trh.appendChild(th);
    });
    board.dates.forEach((ymd) => {
      const th = document.createElement("th");
      th.textContent = ymdToDisplay(ymd);
      if (isWeekendYmd(ymd)) th.classList.add("col-weekend");
      trh.appendChild(th);
    });
    thead.appendChild(trh);
    table.appendChild(thead);

    const tbody = document.createElement("tbody");
    const packs = getFilteredProductPacks(board);
    for (const pack of packs) {
      const tr = document.createElement("tr");
      const line =
        pack.rows.get("발주량") ||
        ({ dates: new Map(), stock: 0, exportCol: 0, lackCol: 0 });
      const tdC = document.createElement("td");
      tdC.textContent = pack.code;
      tr.appendChild(tdC);
      const tdN = document.createElement("td");
      tdN.textContent = pack.name;
      tr.appendChild(tdN);
      board.dates.forEach((ymd) => {
        const td = document.createElement("td");
        const v = line.dates.get(ymd) ?? 0;
        td.textContent = formatQty(v);
        td.className = "num";
        if (isWeekendYmd(ymd)) td.classList.add("col-weekend");
        if (v === 0) td.classList.add("zero");
        if (v > 0) td.classList.add("cell-order");
        tr.appendChild(td);
      });
      tbody.appendChild(tr);
    }
    table.appendChild(tbody);
    simpleTableWrap.appendChild(table);
  }

  /**
   * @param {typeof lastBoard} board
   */
  function renderSummary(board) {
    summaryContent.innerHTML = "";
    if (!board || board.products.length === 0) {
      const d = document.createElement("div");
      d.className = "empty-state empty-state--flat";
      d.textContent = "발주서 엑셀 업로드에서 파일을 먼저 넣어 주세요.";
      summaryContent.appendChild(d);
      return;
    }

    let total발주 = 0;
    let total차이 = 0;
    let total지난 = 0;
    let total부족품 = 0;
    for (const pack of board.products) {
      for (const sub of ["발주량", "차이수량", "지난발주", "부족품"]) {
        const line = pack.rows.get(sub);
        if (!line) continue;
        let s = 0;
        for (const v of line.dates.values()) s += v;
        if (sub === "발주량") total발주 += s;
        if (sub === "차이수량") total차이 += s;
        if (sub === "지난발주") total지난 += s;
        if (sub === "부족품") total부족품 += s;
      }
    }

    const d0 = board.dates[0];
    const d1 = board.dates[board.dates.length - 1];
    const period =
      board.dates.length && d0 && d1 ? `${d0} ~ ${d1} (${board.dates.length}일)` : "—";

    /** @param {string} label @param {string} value @param {string} [sub] */
    function card(label, value, sub) {
      const el = document.createElement("div");
      el.className = "summary-card";
      el.innerHTML = `<div class="summary-card__label"></div><div class="summary-card__value"></div>`;
      el.querySelector(".summary-card__label").textContent = label;
      el.querySelector(".summary-card__value").textContent = value;
      if (sub) {
        const s = document.createElement("div");
        s.className = "summary-card__sub";
        s.textContent = sub;
        el.appendChild(s);
      }
      summaryContent.appendChild(el);
    }

    card("파일", lastFileName || "—", "");
    card("품목 수", String(board.products.length), "일별 통합표 기준 품목 그룹");
    card("기간", period, "");
    card("총 발주량", formatQty(total발주), "일자별 「발주량」 행 합계");
    card("총 차이수량", formatQty(total차이), "일자별 「차이수량」 행 합계");
    card("총 지난발주", formatQty(total지난), "일자별 「지난발주」 행 합계");
    card("총 부족품", formatQty(total부족품), "일자별 「부족품」 행 합계");
    card(
      "데이터 형식",
      board.mode === "wide" ? "넓은 표 (발주서)" : "세로 목록",
      board.mode === "long" ? "열 매핑으로 가공" : "시트 구조 그대로"
    );

    /** @type {Map<string, Map<string, number>>} */
    const monthlyByType = new Map();
    /** @type {Map<string, number>} */
    const rollingByType = new Map();

    const winStart = new Date();
    winStart.setHours(0, 0, 0, 0);
    const winEnd = new Date(winStart);
    winEnd.setMonth(winEnd.getMonth() + 1);

    for (const pack of board.products) {
      const type = (pack.type || "").trim() || "미분류";
      if (!monthlyByType.has(type)) monthlyByType.set(type, new Map());
      if (!rollingByType.has(type)) rollingByType.set(type, 0);
      const orderLine = pack.rows.get("발주량");
      if (!orderLine) continue;

      for (const [ymd, raw] of orderLine.dates) {
        const qty = parseNumber(raw);
        const monthKey = String(ymd).slice(0, 7);
        if (monthKey.length === 7) {
          const mm = monthlyByType.get(type);
          mm.set(monthKey, (mm.get(monthKey) || 0) + qty);
        }

        const p = String(ymd).split("-");
        if (p.length === 3) {
          const d = new Date(parseInt(p[0], 10), parseInt(p[1], 10) - 1, parseInt(p[2], 10));
          if (!Number.isNaN(d.getTime()) && d >= winStart && d <= winEnd) {
            rollingByType.set(type, (rollingByType.get(type) || 0) + qty);
          }
        }
      }
    }

    const winLabel = `${formatYmd(winStart)} ~ ${formatYmd(winEnd)}`;
    const blockMonthly = document.createElement("div");
    blockMonthly.className = "summary-block";
    const ttlMonthly = document.createElement("h3");
    ttlMonthly.className = "summary-block__title";
    ttlMonthly.textContent = "타입별 월별 발주량";
    blockMonthly.appendChild(ttlMonthly);

    const monthKeys = [...new Set([].concat(...[...monthlyByType.values()].map((m) => [...m.keys()])))].sort();
    const tableM = document.createElement("table");
    tableM.className = "summary-table";
    const theadM = document.createElement("thead");
    const trhM = document.createElement("tr");
    const thType = document.createElement("th");
    thType.textContent = "타입";
    trhM.appendChild(thType);
    monthKeys.forEach((mk) => {
      const th = document.createElement("th");
      th.textContent = mk;
      trhM.appendChild(th);
    });
    theadM.appendChild(trhM);
    tableM.appendChild(theadM);
    const tbodyM = document.createElement("tbody");
    [...monthlyByType.keys()].sort((a, b) => a.localeCompare(b, "ko")).forEach((type) => {
      const tr = document.createElement("tr");
      const tdType = document.createElement("td");
      tdType.textContent = type;
      tr.appendChild(tdType);
      monthKeys.forEach((mk) => {
        const td = document.createElement("td");
        td.className = "num";
        td.textContent = formatQty(monthlyByType.get(type).get(mk) || 0);
        tr.appendChild(td);
      });
      tbodyM.appendChild(tr);
    });
    tableM.appendChild(tbodyM);
    blockMonthly.appendChild(tableM);
    summaryContent.appendChild(blockMonthly);

    const blockRolling = document.createElement("div");
    blockRolling.className = "summary-block";
    const ttlRolling = document.createElement("h3");
    ttlRolling.className = "summary-block__title";
    ttlRolling.textContent = `타입별 발주량 (오늘~1개월: ${winLabel})`;
    blockRolling.appendChild(ttlRolling);
    const tableR = document.createElement("table");
    tableR.className = "summary-table";
    tableR.innerHTML = `<thead><tr><th>타입</th><th>발주량</th></tr></thead>`;
    const tbodyR = document.createElement("tbody");
    [...rollingByType.keys()].sort((a, b) => a.localeCompare(b, "ko")).forEach((type) => {
      const tr = document.createElement("tr");
      const tdT = document.createElement("td");
      tdT.textContent = type;
      const tdV = document.createElement("td");
      tdV.className = "num";
      tdV.textContent = formatQty(rollingByType.get(type) || 0);
      tr.appendChild(tdT);
      tr.appendChild(tdV);
      tbodyR.appendChild(tr);
    });
    tableR.appendChild(tbodyR);
    blockRolling.appendChild(tableR);
    summaryContent.appendChild(blockRolling);
  }

  function formatQty(n) {
    if (!Number.isFinite(n)) return "0";
    return Number(n).toLocaleString("ko-KR", { maximumFractionDigits: 2 });
  }

  function clearStockTableView() {
    if (!stockTableWrap) return;
    const t = stockTableWrap.querySelector("table.stock-table");
    if (t) t.remove();
    if (stockTableEmpty) {
      stockTableEmpty.hidden = false;
      stockTableEmpty.textContent = "발주서 엑셀 업로드에서 재고파일을 넣으면 목록이 표시됩니다.";
    }
  }

  function stockCategoryDisplayLabel(value) {
    return value === "" ? "(없음)" : String(value);
  }

  /** @param {"category"|"code"|"name"|"stock"} key */
  function stockDimDisplayLabel(key, value) {
    if (key === "category") return stockCategoryDisplayLabel(value);
    if (key === "code" || key === "name") return value === "" ? "(없음)" : String(value);
    if (key === "stock") return formatQty(Number(value));
    return String(value);
  }

  function closeStockTablePickerPanels() {
    if (panelFilterStockCategory) panelFilterStockCategory.hidden = true;
    if (panelFilterStockCode) panelFilterStockCode.hidden = true;
    if (panelFilterStockName) panelFilterStockName.hidden = true;
    if (panelFilterStockQty) panelFilterStockQty.hidden = true;
  }

  function resetStockTableFilters() {
    stockTableFilterState.category = { options: [], selected: new Set(), enabled: false };
    stockTableFilterState.code = { options: [], selected: new Set() };
    stockTableFilterState.name = { options: [], selected: new Set() };
    stockTableFilterState.stock = { options: [], selected: new Set() };
    if (btnFilterStockCategory) btnFilterStockCategory.textContent = "전체";
    if (btnFilterStockCode) btnFilterStockCode.textContent = "전체";
    if (btnFilterStockName) btnFilterStockName.textContent = "전체";
    if (btnFilterStockQty) btnFilterStockQty.textContent = "전체";
    if (searchFilterStockCategory) searchFilterStockCategory.value = "";
    if (searchFilterStockCode) searchFilterStockCode.value = "";
    if (searchFilterStockName) searchFilterStockName.value = "";
    if (searchFilterStockQty) searchFilterStockQty.value = "";
    if (listFilterStockCategory) listFilterStockCategory.innerHTML = "";
    if (listFilterStockCode) listFilterStockCode.innerHTML = "";
    if (listFilterStockName) listFilterStockName.innerHTML = "";
    if (listFilterStockQty) listFilterStockQty.innerHTML = "";
    if (stockTableFilterBar) stockTableFilterBar.classList.add("filter-bar--hidden");
    if (stockFilterFieldCategory) stockFilterFieldCategory.hidden = true;
    closeStockTablePickerPanels();
  }

  function populateStockTableFiltersFromData() {
    if (!lastStockData || !lastStockData.rows || lastStockData.rows.length === 0) {
      resetStockTableFilters();
      return;
    }
    const rows = lastStockData.rows;
    const hasCat = !!lastStockData.hasCategoryColumn;
    stockTableFilterState.category.enabled = hasCat;
    if (stockFilterFieldCategory) stockFilterFieldCategory.hidden = !hasCat;
    if (hasCat) {
      const u = new Set();
      for (const r of rows) u.add(r.category != null ? String(r.category) : "");
      stockTableFilterState.category.options = [...u].sort((a, b) => a.localeCompare(b, "ko"));
      stockTableFilterState.category.selected = new Set(stockTableFilterState.category.options);
    } else {
      stockTableFilterState.category.options = [];
      stockTableFilterState.category.selected = new Set();
    }

    const uCode = new Set();
    const uName = new Set();
    const uStock = new Set();
    for (const r of rows) {
      uCode.add(r.code != null ? String(r.code).trim() : "");
      uName.add(r.name != null ? String(r.name).trim() : "");
      uStock.add(String(parseNumber(r.stock)));
    }
    stockTableFilterState.code.options = [...uCode].sort((a, b) => a.localeCompare(b, "ko"));
    stockTableFilterState.code.selected = new Set(stockTableFilterState.code.options);
    stockTableFilterState.name.options = [...uName].sort((a, b) => a.localeCompare(b, "ko"));
    stockTableFilterState.name.selected = new Set(stockTableFilterState.name.options);
    stockTableFilterState.stock.options = [...uStock].sort((a, b) => parseNumber(a) - parseNumber(b));
    stockTableFilterState.stock.selected = new Set(stockTableFilterState.stock.options);

    if (stockTableFilterBar) stockTableFilterBar.classList.remove("filter-bar--hidden");
    ["category", "code", "name", "stock"].forEach((k) => {
      if (k === "category" && !stockTableFilterState.category.enabled) return;
      updateStockFilterButton(/** @type {"category"|"code"|"name"|"stock"} */ (k));
      renderStockFilterList(/** @type {"category"|"code"|"name"|"stock"} */ (k));
    });
  }

  /** @param {"category"|"code"|"name"|"stock"} key */
  function updateStockFilterButton(key) {
    const st = stockTableFilterState[key];
    const btn =
      key === "category"
        ? btnFilterStockCategory
        : key === "code"
          ? btnFilterStockCode
          : key === "name"
            ? btnFilterStockName
            : btnFilterStockQty;
    if (!btn) return;
    const total = st.options.length;
    const sel = st.selected.size;
    if (total === 0 || sel === 0 || sel === total) {
      btn.textContent = "전체";
      return;
    }
    if (sel === 1) {
      const v = [...st.selected][0];
      btn.textContent = stockDimDisplayLabel(key, v);
      return;
    }
    btn.textContent = `${sel}개 선택`;
  }

  /** @param {"category"|"code"|"name"|"stock"} key */
  function renderStockFilterList(key) {
    if (key === "category" && !stockTableFilterState.category.enabled) return;
    const st = stockTableFilterState[key];
    const searchEl =
      key === "category"
        ? searchFilterStockCategory
        : key === "code"
          ? searchFilterStockCode
          : key === "name"
            ? searchFilterStockName
            : searchFilterStockQty;
    const listEl =
      key === "category"
        ? listFilterStockCategory
        : key === "code"
          ? listFilterStockCode
          : key === "name"
            ? listFilterStockName
            : listFilterStockQty;
    if (!listEl || !searchEl) return;
    const q = norm(searchEl.value || "");
    const filtered = st.options.filter((v) => norm(stockDimDisplayLabel(key, v)).includes(q) || norm(v).includes(q));
    listEl.innerHTML = "";

    const allRow = document.createElement("label");
    allRow.className = "picker-option";
    const allCb = document.createElement("input");
    allCb.type = "checkbox";
    allCb.checked = st.selected.size === st.options.length && st.options.length > 0;
    allCb.indeterminate = st.selected.size > 0 && st.selected.size < st.options.length;
    allCb.addEventListener("change", () => {
      if (allCb.checked) st.selected = new Set(st.options);
      else st.selected.clear();
      renderStockFilterList(key);
      updateStockFilterButton(key);
      renderStockTableView();
    });
    const allTx = document.createElement("span");
    allTx.textContent = "전체";
    allRow.appendChild(allCb);
    allRow.appendChild(allTx);
    listEl.appendChild(allRow);

    if (filtered.length === 0) {
      const em = document.createElement("div");
      em.className = "picker-empty";
      em.textContent = "검색 결과 없음";
      listEl.appendChild(em);
      return;
    }

    filtered.forEach((v) => {
      const row = document.createElement("label");
      row.className = "picker-option";
      const cb = document.createElement("input");
      cb.type = "checkbox";
      cb.checked = st.selected.has(v);
      cb.addEventListener("change", () => {
        if (cb.checked) st.selected.add(v);
        else st.selected.delete(v);
        updateStockFilterButton(key);
        renderStockFilterList(key);
        renderStockTableView();
      });
      const tx = document.createElement("span");
      tx.textContent = stockDimDisplayLabel(key, v);
      row.appendChild(cb);
      row.appendChild(tx);
      listEl.appendChild(row);
    });
  }

  /**
   * 검색창 글자가 포함되는 옵션만 선택/해제 (품목코드·품목명)
   * @param {"code"|"name"} key
   * @param {"select"|"clear"} mode
   */
  function stockFilterBulkBySearchQuery(key, mode) {
    const st = stockTableFilterState[key];
    const searchEl = key === "code" ? searchFilterStockCode : searchFilterStockName;
    const q = norm(searchEl && searchEl.value ? searchEl.value : "");
    if (!q) return;
    for (const opt of st.options) {
      if (!norm(opt).includes(q)) continue;
      if (mode === "select") st.selected.add(opt);
      else st.selected.delete(opt);
    }
    updateStockFilterButton(key);
    renderStockFilterList(key);
    renderStockTableView();
  }

  /** 품목코드 검색창: 입력 시 일치 항목만 선택, 비우면 전체 선택 (목록 필터와 동일한 매칭) */
  function applyStockCodeSearchToSelection() {
    const st = stockTableFilterState.code;
    const q = norm(searchFilterStockCode && searchFilterStockCode.value ? searchFilterStockCode.value : "");
    if (!q) {
      st.selected = new Set(st.options);
    } else {
      st.selected = new Set(
        st.options.filter(
          (v) => norm(stockDimDisplayLabel("code", v)).includes(q) || norm(v).includes(q)
        )
      );
    }
    updateStockFilterButton("code");
    renderStockFilterList("code");
    renderStockTableView();
  }

  /** @param {"category"|"code"|"name"|"stock"} dimKey */
  function stockDimPasses(dimKey, rawValue) {
    if (dimKey === "category" && !stockTableFilterState.category.enabled) return true;
    const st = stockTableFilterState[dimKey];
    if (!st.options.length || st.selected.size === 0 || st.selected.size === st.options.length) return true;
    return st.selected.has(rawValue);
  }

  /** `lastStockData` + 필터 기준 재고 표 */
  function renderStockTableView() {
    if (!stockTableWrap || !stockTableEmpty) return;
    const old = stockTableWrap.querySelector("table.stock-table");
    if (old) old.remove();
    if (!lastStockData || !lastStockData.rows || lastStockData.rows.length === 0) {
      resetStockTableFilters();
      stockTableEmpty.hidden = false;
      stockTableEmpty.textContent =
        lastStockData && lastStockData.rowCount === 0
          ? "읽은 데이터 행이 없습니다. 품목코드·제품명·현재고(또는 재고) 열 헤더를 확인해 주세요."
          : "발주서 엑셀 업로드에서 재고파일을 넣으면 목록이 표시됩니다.";
      return;
    }
    const hasCat = !!lastStockData.hasCategoryColumn;
    const displayRows = lastStockData.rows.filter((r) => {
      const cat = r.category != null ? String(r.category) : "";
      const code = r.code != null ? String(r.code).trim() : "";
      const name = r.name != null ? String(r.name).trim() : "";
      const sk = String(parseNumber(r.stock));
      if (!stockDimPasses("category", cat)) return false;
      if (!stockDimPasses("code", code)) return false;
      if (!stockDimPasses("name", name)) return false;
      if (!stockDimPasses("stock", sk)) return false;
      return true;
    });
    if (displayRows.length === 0) {
      stockTableEmpty.hidden = false;
      stockTableEmpty.textContent = "선택한 필터에 해당하는 행이 없습니다. 필터를 조정해 주세요.";
      return;
    }
    stockTableEmpty.hidden = true;
    const table = document.createElement("table");
    table.className = "stock-table";
    table.setAttribute("role", "grid");
    const thead = document.createElement("thead");
    const trh = document.createElement("tr");
    const headers = hasCat ? ["카테고리", "품목코드", "품목명", "현재고"] : ["품목코드", "품목명", "현재고"];
    headers.forEach((label) => {
      const th = document.createElement("th");
      th.textContent = label;
      trh.appendChild(th);
    });
    thead.appendChild(trh);
    const tbody = document.createElement("tbody");
    for (const row of displayRows) {
      const tr = document.createElement("tr");
      if (hasCat) {
        const tdCat = document.createElement("td");
        tdCat.textContent = stockCategoryDisplayLabel(row.category != null ? String(row.category) : "");
        tr.appendChild(tdCat);
      }
      const tdCode = document.createElement("td");
      tdCode.textContent = row.code || "—";
      const tdName = document.createElement("td");
      tdName.className = "stock-table__name";
      tdName.textContent = row.name || "—";
      const tdStock = document.createElement("td");
      tdStock.className = "num";
      tdStock.textContent = formatQty(row.stock);
      tr.append(tdCode, tdName, tdStock);
      tbody.appendChild(tr);
    }
    table.appendChild(thead);
    table.appendChild(tbody);
    stockTableWrap.appendChild(table);
  }

  /**
   * @param {NonNullable<typeof lastBoard>} board
   */
  function exportBoard(board) {
    const packs = getFilteredProductPacks(board);
    const subs = getFilteredSubs();
    const exportSet = filterState.export.selected;
    const exportAll = exportSet.size === filterState.export.options.length;
    const head = [
      ...["구분", "품목코드", "제품", "타입", "재고", "수출발주량", "부족량"],
      ...board.dates.map(ymdToDisplay),
    ];
    const aoa = [head];
    for (const pack of packs) {
      for (const sub of subs) {
        const line =
          pack.rows.get(sub) ||
          ({ dates: new Map(), stock: 0, exportCol: 0, lackCol: 0 });
        if (!exportAll && !exportSet.has(String(Number(line.exportCol || 0)))) continue;
        const row = [
          sub,
          pack.code,
          pack.name,
          pack.type || "",
          line.stock,
          line.exportCol,
          line.lackCol,
        ];
        for (const d of board.dates) row.push(line.dates.get(d) ?? 0);
        aoa.push(row);
      }
    }
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "발주현황");
    XLSX.writeFile(wb, (lastFileName.replace(/\.[^.]+$/, "") || "발주현황") + "_보드.xlsx");
  }

  /** 가동율·가동률 우선, 없으면 설비가동율(동일 지표로 쓰는 파일 대응) */
  function findUtilizationRateColumnIndex(headerRow) {
    for (let i = 0; i < headerRow.length; i++) {
      const raw = String(headerRow[i] ?? "");
      const h = norm(raw);
      if (!h) continue;
      if (/설비가동|설비\s*가동/.test(h)) continue;
      if (matchHeader(raw, ["가동율", "가동률"])) return i;
    }
    return findHeaderIndex(headerRow, ["설비가동율", "설비가동률"]);
  }

  /** 설비효율·OEE (막연한 '효율' 단독 헤더는 맨 후순위) */
  function findOeeColumnIndex(headerRow) {
    const preferred = [
      ["설비효율", "설비 효율", "전체설비효율"],
      ["oee", "overall"],
      ["효율"],
    ];
    for (const keys of preferred) {
      const j = findHeaderIndex(headerRow, keys);
      if (j >= 0) return j;
    }
    return -1;
  }

  /** 불량율/률 (불량수량·불량수 단독은 제외) */
  function findDefectRateColumnIndex(headerRow) {
    for (let i = 0; i < headerRow.length; i++) {
      const raw = String(headerRow[i] ?? "");
      const h = norm(raw);
      if (!h) continue;
      if (/불량수량|^불량수$/.test(h) && !/율|률|%/.test(h)) continue;
      if (matchHeader(raw, ["불량율", "불량률"])) return i;
    }
    return -1;
  }

  /** 불량 건수·수량 (불량율/률 열 제외) */
  function findDefectQtyColumnIndex(headerRow) {
    for (let i = 0; i < headerRow.length; i++) {
      const raw = String(headerRow[i] ?? "");
      const h = norm(raw);
      if (!h) continue;
      if (/불량율|불량률/.test(h)) continue;
      if (
        h === "불량" ||
        matchHeader(raw, [
          "불량수량",
          "불량수",
          "불량 개수",
          "불량개수",
          "불량량",
          "불량건수",
          "불량ea",
          "ng수량",
          "ng건",
          "스크랩",
          "scrap",
        ])
      ) {
        return i;
      }
      if (/불량/.test(h) && /수량|개수|건수|ea|qty|ng|스크랩|scrap/.test(h) && !/율|률/.test(h)) return i;
    }
    return -1;
  }

  /** 생산 수량 열 (생산량·출하 등 엑셀 표기 차이 대응) */
  function findDrawingProductionQtyColumnIndex(headerRow) {
    const j = findHeaderIndex(headerRow, [
      "생산수량",
      "생산 수량",
      "생산량",
      "생산량(ERP)",
      "생산량（ERP）",
      "생산량(erp)",
      "총생산량",
      "양산수량",
      "출하수량",
      "실적수량",
      "생산량(개)",
      "생산수량(개)",
      "prdqty",
      "productionqty",
    ]);
    if (j >= 0) return j;
    for (let i = 0; i < headerRow.length; i++) {
      const h = norm(String(headerRow[i] ?? ""));
      if (!h) continue;
      if (/불량|defect|scrap|율|률|가동|oee|효율|투입시간|작업시간|생산성|규격|작업일|일자/.test(h)) continue;
      if (
        (/생산|양산|출하|실적|총생산/.test(h) && /수량|량$/i.test(h) && !/생산성$/.test(h)) ||
        /productionqty|prdqty/.test(h)
      ) {
        return i;
      }
    }
    return -1;
  }

  /** 정지·비가동 시간 열 (비가동률 등과 구분) */
  function findDrawingStopTimeColumnIndex(headerRow) {
    const j = findHeaderIndex(headerRow, [
      "정지시간",
      "정지 시간",
      "비가동시간",
      "비가동 시간",
      "다운타임",
      "downtime",
      "설비정지시간",
      "설비 정지시간",
      "정지합계",
    ]);
    if (j >= 0) return j;
    for (let i = 0; i < headerRow.length; i++) {
      const h = norm(String(headerRow[i] ?? ""));
      if (!h) continue;
      if (/비가동률|비가동율|가동률|가동율|정지율|정지률/.test(h)) continue;
      if (
        (/정지|비가동|다운타임|downtime|idle/.test(h) && /시간|분|hr|hour|min/i.test(h)) &&
        !/율|률|건수|횟수|횟$/.test(h)
      ) {
        return i;
      }
    }
    return -1;
  }

  /**
   * @param {any} wb
   * @returns {string | null} 시트 이름만 반환 (이름에 「드로잉」 포함, 정확히 「드로잉」 우선)
   */
  function findDrawingSheetName(wb) {
    const names = wb.SheetNames || [];
    for (const n of names) {
      if (String(n).trim() === "드로잉") return n;
    }
    for (const n of names) {
      if (norm(n).includes("드로잉")) return n;
    }
    return null;
  }

  /**
   * @param {any} wb
   * @returns {string | null} 시트 이름 — 정확히 「페럴」 우선, 이름에 페럴 포함
   */
  function findParallelSheetName(wb) {
    const names = wb.SheetNames || [];
    for (const n of names) {
      if (String(n).trim() === "페럴") return n;
    }
    for (const n of names) {
      if (norm(n).includes("페럴")) return n;
    }
    return null;
  }

  /**
   * @param {any} wb
   * @returns {string | null} 시트 이름 — 「이중관」 우선, 화승이중관 제외 후 이름에 이중관 포함
   */
  function findDoublePipeSheetName(wb) {
    const names = wb.SheetNames || [];
    for (const n of names) {
      if (String(n).trim() === "이중관") return n;
    }
    for (const n of names) {
      const nt = norm(String(n));
      if (nt.includes("화승") && nt.includes("이중관")) continue;
      if (nt.includes("이중관")) return n;
    }
    return null;
  }

  /**
   * @param {any} wb
   * @returns {string | null} 시트 이름 — 화승이중관(또는 이름에 화승+이중관)
   */
  function findHwaseungDoublePipeSheetName(wb) {
    const names = wb.SheetNames || [];
    const compact = (n) =>
      String(n ?? "")
        .trim()
        .replace(/[-_\s]+/g, "");
    const isPivotLike = (n) => /피벗|pivot/i.test(String(n));
    for (const n of names) {
      if (isPivotLike(n)) continue;
      if (compact(n) === "화승이중관") return n;
    }
    for (const n of names) {
      if (isPivotLike(n)) continue;
      const nt = norm(String(n));
      if (nt.includes("화승") && nt.includes("이중관")) return n;
    }
    for (const n of names) {
      const nt = norm(String(n));
      if (nt.includes("화승") && nt.includes("이중관")) return n;
    }
    return null;
  }

  /** @param {unknown} raw 공정 열 값 */
  function normalizeParallelProcess(raw) {
    const s = String(raw ?? "").trim();
    if (!s) return "기타";
    const n = norm(s);
    if (/절단/.test(n)) return "절단";
    if (/벤딩|bending/.test(n)) return "벤딩";
    return "기타";
  }

  /** @param {unknown} raw 공정 열 값 — 가공·성형 */
  function normalizeDoublePipeProcess(raw) {
    const s = String(raw ?? "").trim();
    if (!s) return "기타";
    const n = norm(s);
    if (/가공/.test(n)) return "가공";
    if (/성형|포밍|forming/i.test(n)) return "성형";
    return "기타";
  }

  /** @param {unknown} raw 공정 열 값 — 화승이중관 6공정 + 기타 */
  function normalizeHwaseungDoublePipeProcess(raw) {
    const s = String(raw ?? "").trim();
    if (!s) return "기타";
    const n = norm(s.replace(/\s+/g, " "));
    for (const k of HWASEUNG_DOUBLE_PIPE_PROCESS_KEYS) {
      const kn = norm(k.replace(/\s*,\s*/g, ","));
      if (n === kn || n === norm(k)) return k;
    }
    if (/압착/.test(n)) return "압착";
    if (/수동\s*롤링|수동롤링/.test(n)) return "수동 롤링";
    if (/자동\s*롤링|자동롤링/.test(n)) return "자동롤링";
    if (/스파이럴|spiral/i.test(n)) return "스파이럴";
    if (/전조/.test(n)) return "전조";
    if (/피어싱/.test(n) || /확관/.test(n)) return "피어싱, 확관";
    return "기타";
  }

  function sortHwaseungDoublePipeProcessKeys(a, b) {
    const rank = (k) => {
      const i = HWASEUNG_DOUBLE_PIPE_PROCESS_KEYS.indexOf(k);
      if (i >= 0) return i;
      if (k === "기타") return HWASEUNG_DOUBLE_PIPE_PROCESS_KEYS.length;
      return 999;
    };
    const ra = rank(a);
    const rb = rank(b);
    if (ra !== rb) return ra - rb;
    return String(a).localeCompare(String(b), "ko");
  }

  /** @param {unknown} raw 공정 열 값 — 절단·포밍·가공 */
  function normalizeMufflerProcess(raw) {
    const s = String(raw ?? "").trim();
    if (!s) return "기타";
    const n = norm(s);
    if (/절단/.test(n)) return "절단";
    if (/포밍|forming/.test(n)) return "포밍";
    if (/가공/.test(n)) return "가공";
    return "기타";
  }

  /**
   * @param {any} wb
   * @returns {{ mufflerSheet: string | null, forgedSheet: string | null }}
   */
  function findMufflerSheetNamesInWorkbook(wb) {
    const names = wb.SheetNames || [];
    let mufflerSheet = null;
    let forgedSheet = null;
    for (const n of names) {
      const t = String(n).trim();
      if (t === "머플러") mufflerSheet = mufflerSheet || n;
    }
    for (const n of names) {
      const t = String(n).trim();
      const nt = norm(t);
      if (t === "단조머플러" || (nt.includes("단조") && nt.includes("머플러"))) forgedSheet = forgedSheet || n;
    }
    if (!mufflerSheet) {
      for (const n of names) {
        const nt = norm(String(n));
        if (nt.includes("머플러") && !nt.includes("단조")) mufflerSheet = mufflerSheet || n;
      }
    }
    return { mufflerSheet, forgedSheet };
  }

  /**
   * @param {any} v
   * @param {{ ratioOk?: boolean }} [opts] true면 0~1 숫자를 %로 환산
   */
  function parseMetricCell(v, opts) {
    const ratioOk = opts && opts.ratioOk !== false;
    if (v == null || v === "") return NaN;
    if (typeof v === "number" && Number.isFinite(v)) {
      if (ratioOk && v > 0 && v <= 1) return v * 100;
      return v;
    }
    const s0 = String(v).trim();
    const hasPct = /%/.test(s0);
    const n = parseFloat(s0.replace(/,/g, "").replace(/%/g, ""));
    if (!Number.isFinite(n)) return NaN;
    if (hasPct) return n;
    if (ratioOk && n > 0 && n < 1) return n * 100;
    return n;
  }

  function averageFinite(values) {
    const ok = values.filter((x) => Number.isFinite(x));
    if (ok.length === 0) return null;
    return ok.reduce((a, b) => a + b, 0) / ok.length;
  }

  function sumFinite(values) {
    const ok = values.filter((x) => Number.isFinite(x));
    if (ok.length === 0) return null;
    return ok.reduce((a, b) => a + b, 0);
  }

  /**
   * 행 단위 가동율(%): 작업시간/(작업+정지) 우선, 없으면 작업시간/투입시간(상한 100%)
   * @param {number} workTime
   * @param {number} stopTime
   * @param {number} inputTime
   */
  function deriveRowUtilizationPct(workTime, stopTime, inputTime) {
    const w = Number.isFinite(workTime) ? workTime : NaN;
    const s = Number.isFinite(stopTime) ? stopTime : NaN;
    const inp = Number.isFinite(inputTime) ? inputTime : NaN;
    if (Number.isFinite(w) && Number.isFinite(s) && w + s > 0) return (w / (w + s)) * 100;
    if (Number.isFinite(inp) && inp > 0 && Number.isFinite(w) && w >= 0) return Math.min(100, (w / inp) * 100);
    return NaN;
  }

  /**
   * 행 단위 품질(%): 생산수량·불량수 우선, 없으면 불량율 열로 역산
   */
  function deriveRowQualityPct(prodQty, defectQty, defectRateCell) {
    const p = Number.isFinite(prodQty) ? prodQty : NaN;
    const d = Number.isFinite(defectQty) ? Math.max(0, defectQty) : NaN;
    if (Number.isFinite(p) && p > 0) {
      const q = ((p - (Number.isFinite(d) ? d : 0)) / p) * 100;
      return Math.max(0, Math.min(100, q));
    }
    const r = parseMetricCell(defectRateCell, { ratioOk: false });
    if (Number.isFinite(r) && r >= 0 && r <= 100) return Math.max(0, 100 - r);
    return NaN;
  }

  /** 생산성 열을 성능(%)으로 쓸 수 있을 때만 (0~150% 또는 0~1 비율) */
  function deriveRowPerformancePct(prodCell) {
    if (prodCell == null || prodCell === "") return NaN;
    const x = parseMetricCell(prodCell, { ratioOk: true });
    if (!Number.isFinite(x) || x < 0) return NaN;
    if (x > 150) return NaN;
    return Math.min(120, x);
  }

  function findEquipmentColumnIndex(headerRow) {
    return findHeaderIndex(headerRow, [
      "설비",
      "설비명",
      "호기",
      "설비호기",
      "기계",
      "기계명",
      "장비",
      "라인",
      "line",
      "equipment",
    ]);
  }

  function makeDrawingAgg() {
    return {
      dataRows: 0,
      utilVals: [],
      oeeVals: [],
      defRateVals: [],
      utilDerivedVals: [],
      oeeDerivedVals: [],
      prodQtyVals: [],
      workTimeVals: [],
      inputTimeVals: [],
      productivityVals: [],
      defectQtyVals: [],
      stopTimeVals: [],
      stopExchangeVals: [],
      stopRepairVals: [],
      stopMaterialVals: [],
      stopPlannedVals: [],
      stopFifthVals: [],
      workDates: [],
    };
  }

  /**
   * @param {Map<string, Map<string, ReturnType<typeof makeDrawingAgg>>>} outerMap
   */
  function ensureAggMapNested2(outerMap, outerKey, innerKey) {
    if (!outerMap.has(outerKey)) outerMap.set(outerKey, new Map());
    const inner = outerMap.get(outerKey);
    if (!inner.has(innerKey)) inner.set(innerKey, makeDrawingAgg());
    return inner.get(innerKey);
  }

  /**
   * 공정 → 일 → 설비
   * @param {Map<string, Map<string, Map<string, ReturnType<typeof makeDrawingAgg>>>>} procMap
   */
  function ensureAggMapNested3ProcDayEq(procMap, pk, dayKey, eqKey) {
    if (!procMap.has(pk)) procMap.set(pk, new Map());
    const dmap = procMap.get(pk);
    if (!dmap.has(dayKey)) dmap.set(dayKey, new Map());
    const emap = dmap.get(dayKey);
    if (!emap.has(eqKey)) emap.set(eqKey, makeDrawingAgg());
    return emap.get(eqKey);
  }

  /**
   * @param {ReturnType<typeof makeDrawingAgg>} agg
   * @param {any[]} row
   * @param {Record<string, number>} idx
   * @param {boolean} collectDates
   */
  function addDrawingRowToAgg(agg, row, idx, collectDates) {
    const {
      iUtil,
      iOee,
      iDefRate,
      iProdQty,
      iWorkTime,
      iInputTime,
      iProductivity,
      iDefectQty,
      iStopTime,
      iWorkDate,
    } = idx;
    const stopCols = Array.isArray(idx.iStopTimeCols) ? idx.iStopTimeCols : [];
    agg.dataRows++;
    if (iUtil >= 0) agg.utilVals.push(parseMetricCell(row[iUtil], { ratioOk: true }));
    if (iOee >= 0) agg.oeeVals.push(parseMetricCell(row[iOee], { ratioOk: true }));
    if (iDefRate >= 0) agg.defRateVals.push(parseMetricCell(row[iDefRate], { ratioOk: false }));
    if (iProdQty >= 0) agg.prodQtyVals.push(parseNumber(row[iProdQty]));
    if (iWorkTime >= 0) agg.workTimeVals.push(parseNumber(row[iWorkTime]));
    if (iInputTime >= 0) agg.inputTimeVals.push(parseNumber(row[iInputTime]));
    if (iProductivity >= 0) agg.productivityVals.push(parseMetricCell(row[iProductivity], { ratioOk: true }));
    const defectCols = Array.isArray(idx.iDefectQtyCols) ? idx.iDefectQtyCols : [];
    if (defectCols.length) {
      let dq = 0;
      for (const j of defectCols) {
        if (j >= 0) dq += parseNumber(row[j]);
      }
      agg.defectQtyVals.push(dq);
    } else if (iDefectQty >= 0) {
      agg.defectQtyVals.push(parseNumber(row[iDefectQty]));
    }
    let stopSumForDerive = NaN;
    if (stopCols.length === 4) {
      const v0 = parseNumber(row[stopCols[0]]);
      const v1 = parseNumber(row[stopCols[1]]);
      const v2 = parseNumber(row[stopCols[2]]);
      const v3 = parseNumber(row[stopCols[3]]);
      agg.stopExchangeVals.push(v0);
      agg.stopRepairVals.push(v1);
      agg.stopMaterialVals.push(v2);
      agg.stopPlannedVals.push(v3);
      stopSumForDerive = v0 + v1 + v2 + v3;
      agg.stopTimeVals.push(stopSumForDerive);
    } else if (stopCols.length === 5) {
      const vals = stopCols.map((j) => (j >= 0 ? parseNumber(row[j]) : 0));
      agg.stopExchangeVals.push(vals[0]);
      agg.stopRepairVals.push(vals[1]);
      agg.stopMaterialVals.push(vals[2]);
      agg.stopPlannedVals.push(vals[3]);
      agg.stopFifthVals.push(vals[4]);
      stopSumForDerive = vals[0] + vals[1] + vals[2] + vals[3] + vals[4];
      agg.stopTimeVals.push(stopSumForDerive);
    } else if (stopCols.length) {
      let st = 0;
      for (const j of stopCols) {
        if (j >= 0) st += parseNumber(row[j]);
      }
      stopSumForDerive = st;
      agg.stopTimeVals.push(st);
    } else if (iStopTime >= 0) {
      stopSumForDerive = parseNumber(row[iStopTime]);
      agg.stopTimeVals.push(stopSumForDerive);
    }
    if (collectDates && iWorkDate >= 0) {
      const d = parseExcelDate(row[iWorkDate]);
      if (d) agg.workDates.push(d);
    }

    const w = iWorkTime >= 0 ? parseNumber(row[iWorkTime]) : NaN;
    const s = stopCols.length ? stopSumForDerive : iStopTime >= 0 ? parseNumber(row[iStopTime]) : NaN;
    const inp = iInputTime >= 0 ? parseNumber(row[iInputTime]) : NaN;
    const p = iProdQty >= 0 ? parseNumber(row[iProdQty]) : NaN;
    let dqty = NaN;
    if (defectCols.length) {
      dqty = defectCols.reduce((acc, j) => acc + (j >= 0 ? parseNumber(row[j]) : 0), 0);
    } else if (iDefectQty >= 0) {
      dqty = parseNumber(row[iDefectQty]);
    }
    const defCell = iDefRate >= 0 ? row[iDefRate] : null;
    const prodCell = iProductivity >= 0 ? row[iProductivity] : null;

    const uDer = deriveRowUtilizationPct(w, s, inp);
    if (iUtil < 0 && Number.isFinite(uDer)) agg.utilDerivedVals.push(uDer);

    const uForOee = iUtil >= 0 ? parseMetricCell(row[iUtil], { ratioOk: true }) : uDer;
    if (Number.isFinite(uForOee)) {
      const qual = deriveRowQualityPct(p, dqty, defCell);
      const perf = deriveRowPerformancePct(prodCell);
      const pe = Number.isFinite(perf) ? perf : 100;
      const q = Number.isFinite(qual) ? qual : 100;
      agg.oeeDerivedVals.push((uForOee / 100) * (pe / 100) * (q / 100) * 100);
    }
  }

  /**
   * @param {ReturnType<typeof makeDrawingAgg>} agg
   * @param {Record<string, number>} idx
   */
  function finalizeDrawingAgg(agg, idx) {
    const { iUtil, iOee, iDefRate, iProdQty, iWorkTime, iInputTime, iProductivity, iDefectQty, iStopTime } = idx;
    const defectCols = Array.isArray(idx.iDefectQtyCols) ? idx.iDefectQtyCols : [];
    const hasDefectQtySource = defectCols.length > 0 || iDefectQty >= 0;
    const stopCols = Array.isArray(idx.iStopTimeCols) ? idx.iStopTimeCols : [];
    const hasStopSource = stopCols.length > 0 || iStopTime >= 0;
    const hasStopKindSplit = stopCols.length === 4 || stopCols.length === 5;
    const hasStopKindSplit5 = stopCols.length === 5;
    const utilFromCol = iUtil >= 0 ? averageFinite(agg.utilVals) : null;
    const utilFromDerived = averageFinite(agg.utilDerivedVals);
    const utilAvg = utilFromCol != null ? utilFromCol : utilFromDerived;
    const oeeFromCol = iOee >= 0 ? averageFinite(agg.oeeVals) : null;
    const oeeFromDerived = averageFinite(agg.oeeDerivedVals);
    const oeeAvg = oeeFromCol != null ? oeeFromCol : oeeFromDerived;
    const utilIsDerived = utilFromCol == null && utilFromDerived != null;
    const oeeIsDerived = oeeFromCol == null && oeeFromDerived != null;
    const sumProdQty = iProdQty >= 0 ? sumFinite(agg.prodQtyVals) : null;
    const sumDefectQty = hasDefectQtySource ? sumFinite(agg.defectQtyVals) : null;
    let defectAvg = iDefRate >= 0 ? averageFinite(agg.defRateVals) : null;
    const canQtyDefectRate =
      sumProdQty != null &&
      sumProdQty > 0 &&
      sumDefectQty != null &&
      Number.isFinite(sumDefectQty);
    if (canQtyDefectRate) {
      const fromQty = (sumDefectQty / sumProdQty) * 100;
      if (!Number.isFinite(defectAvg)) defectAvg = fromQty;
      else if (defectAvg === 0 && sumDefectQty > 0) defectAvg = fromQty;
    }
    return {
      dataRows: agg.dataRows,
      utilAvg,
      oeeAvg,
      utilIsDerived,
      oeeIsDerived,
      defectAvg,
      sumProdQty,
      sumWorkTime: iWorkTime >= 0 ? sumFinite(agg.workTimeVals) : null,
      sumInputTime: iInputTime >= 0 ? sumFinite(agg.inputTimeVals) : null,
      productivityAvg: iProductivity >= 0 ? averageFinite(agg.productivityVals) : null,
      sumDefectQty,
      sumStopTime: hasStopSource ? sumFinite(agg.stopTimeVals) : null,
      sumStopExchange: hasStopKindSplit ? sumFinite(agg.stopExchangeVals) : null,
      sumStopRepair: hasStopKindSplit ? sumFinite(agg.stopRepairVals) : null,
      sumStopMaterial: hasStopKindSplit ? sumFinite(agg.stopMaterialVals) : null,
      sumStopPlanned: hasStopKindSplit ? sumFinite(agg.stopPlannedVals) : null,
      sumStopFifth: hasStopKindSplit5 ? sumFinite(agg.stopFifthVals) : null,
    };
  }

  /** 설비 키에서 정렬용 숫자 추출 (1, 1호기, 호기 2 등) */
  function firstIntFromDrawingEqKey(s) {
    const t = String(s ?? "").trim();
    const m = t.match(/\d+/);
    return m ? parseInt(m[0], 10) : NaN;
  }

  /** 설비별 집계: 1·2·3… 호기 순, 미지정·전체는 맨 아래 */
  function compareDrawingEquipmentKeys(a, b) {
    const tail = new Set(["미지정", "전체"]);
    const aTail = tail.has(a);
    const bTail = tail.has(b);
    if (aTail !== bTail) return aTail ? 1 : -1;
    if (aTail && bTail) return a.localeCompare(b, "ko");
    const na = firstIntFromDrawingEqKey(a);
    const nb = firstIntFromDrawingEqKey(b);
    if (Number.isFinite(na) && Number.isFinite(nb) && na !== nb) return na - nb;
    if (Number.isFinite(na) && !Number.isFinite(nb)) return -1;
    if (!Number.isFinite(na) && Number.isFinite(nb)) return 1;
    return a.localeCompare(b, "ko");
  }

  /**
   * @param {any[][]} matrix
   * @returns {null | Record<string, any>}
   */
  function parseDrawingLogFromMatrix(matrix) {
    if (!matrix || matrix.length < 2) return null;
    const hi = detectHeaderRowIndex(matrix);
    const headerRow = matrix[hi] || [];

    const iWorkDate = findHeaderIndex(headerRow, ["작업일자", "작업일", ...DATE_KEYS]);
    const iSpec = findHeaderIndex(headerRow, ["규격"]);
    const iProdQty = findDrawingProductionQtyColumnIndex(headerRow);
    const iWorkTime = findHeaderIndex(headerRow, ["작업시간", "작업 시간"]);
    const iInputTime = findHeaderIndex(headerRow, ["투입시간", "투입 시간"]);
    const iProductivity = findHeaderIndex(headerRow, ["생산성"]);
    let iDefectQty = findHeaderIndex(headerRow, [
      "공장불량",
      "공장 불량",
      "공장불량수",
      "공장불량수량",
      "공장 불량수",
    ]);
    if (iDefectQty < 0 && headerRow.length > DRAWING_FACTORY_DEFECT_COL_U) {
      iDefectQty = DRAWING_FACTORY_DEFECT_COL_U;
    }
    if (iDefectQty < 0) iDefectQty = findDefectQtyColumnIndex(headerRow);

    const iStopTimeCols = headerRow.length > 19 ? [...DRAWING_STOP_COL_QRST] : null;
    const iStopTime = iStopTimeCols ? -1 : findDrawingStopTimeColumnIndex(headerRow);

    const iUtil = findUtilizationRateColumnIndex(headerRow);
    const iOee = findOeeColumnIndex(headerRow);
    const iDefRate = findDefectRateColumnIndex(headerRow);

    const hasCore =
      iWorkDate >= 0 ||
      iSpec >= 0 ||
      iProdQty >= 0 ||
      iWorkTime >= 0 ||
      iInputTime >= 0 ||
      iProductivity >= 0 ||
      iDefectQty >= 0 ||
      iStopTime >= 0 ||
      !!(iStopTimeCols && iStopTimeCols.length);
    const hasKpi = iUtil >= 0 || iOee >= 0 || iDefRate >= 0;
    if (!hasCore && !hasKpi) return null;

    const iEquipment = findEquipmentColumnIndex(headerRow);

    const idx = {
      iWorkDate,
      iUtil,
      iOee,
      iDefRate,
      iProdQty,
      iWorkTime,
      iInputTime,
      iProductivity,
      iDefectQty,
      iStopTime,
      iStopTimeCols: iStopTimeCols ? iStopTimeCols : [],
    };

    const aggGlobal = makeDrawingAgg();
    /** @type {Map<string, ReturnType<typeof makeDrawingAgg>>} */
    const byMonth = new Map();
    /** @type {Map<string, ReturnType<typeof makeDrawingAgg>>} */
    const byEq = new Map();
    /** @type {Map<string, ReturnType<typeof makeDrawingAgg>>} */
    const byDay = new Map();
    /** @type {Map<string, Map<string, ReturnType<typeof makeDrawingAgg>>>} */
    const byDayEq = new Map();

    for (let r = hi + 1; r < matrix.length; r++) {
      const row = matrix[r];
      if (!row || row.length === 0) continue;
      const nonempty = row.some((c) => String(c ?? "").trim() !== "");
      if (!nonempty) continue;

      addDrawingRowToAgg(aggGlobal, row, idx, true);

      const ymd = iWorkDate >= 0 ? parseExcelDate(row[iWorkDate]) : "";
      const ym = ymd && ymd.length >= 7 ? ymd.slice(0, 7) : "일자미상";
      if (!byMonth.has(ym)) byMonth.set(ym, makeDrawingAgg());
      addDrawingRowToAgg(byMonth.get(ym), row, idx, false);

      const eqKey = iEquipment >= 0 ? String(row[iEquipment] ?? "").trim() || "미지정" : "전체";
      if (!byEq.has(eqKey)) byEq.set(eqKey, makeDrawingAgg());
      addDrawingRowToAgg(byEq.get(eqKey), row, idx, false);

      const dayKey = ymd || "일자미상";
      if (!byDay.has(dayKey)) byDay.set(dayKey, makeDrawingAgg());
      addDrawingRowToAgg(byDay.get(dayKey), row, idx, false);
      if (!byDayEq.has(dayKey)) byDayEq.set(dayKey, new Map());
      const dayEqMap = byDayEq.get(dayKey);
      if (!dayEqMap.has(eqKey)) dayEqMap.set(eqKey, makeDrawingAgg());
      addDrawingRowToAgg(dayEqMap.get(eqKey), row, idx, false);
    }

    aggGlobal.workDates.sort();
    const g = finalizeDrawingAgg(aggGlobal, idx);
    const workDates = aggGlobal.workDates;

    const monthLabel = (ymKey) => {
      if (ymKey === "일자미상") return ymKey;
      const p = ymKey.split("-");
      if (p.length !== 2) return ymKey;
      return `${p[0]}년 ${parseInt(p[1], 10)}월`;
    };

    const monthlyStats = [...byMonth.entries()]
      .sort((a, b) => {
        if (a[0] === "일자미상") return 1;
        if (b[0] === "일자미상") return -1;
        return a[0].localeCompare(b[0]);
      })
      .map(([key, agg]) => ({ key, label: monthLabel(key), ...finalizeDrawingAgg(agg, idx) }));

    const equipmentStats = [...byEq.entries()]
      .sort((a, b) => compareDrawingEquipmentKeys(a[0], b[0]))
      .map(([key, agg]) => ({ key, label: key, ...finalizeDrawingAgg(agg, idx) }));

    const dailyStats = [...byDay.entries()]
      .sort((a, b) => {
        if (a[0] === "일자미상") return 1;
        if (b[0] === "일자미상") return -1;
        return a[0].localeCompare(b[0]);
      })
      .map(([key, agg]) => ({ key, label: key, ...finalizeDrawingAgg(agg, idx) }));

    /** @type {Record<string, any[]>} */
    const dayEquipmentStats = {};
    for (const [dKey, eqMap] of byDayEq.entries()) {
      dayEquipmentStats[dKey] = [...eqMap.entries()]
        .sort((a, b) => compareDrawingEquipmentKeys(a[0], b[0]))
        .map(([key, agg]) => ({ key, label: key, ...finalizeDrawingAgg(agg, idx) }));
    }

    return {
      hasAny: true,
      headerRowIndex: hi,
      iWorkDate,
      iEquipment,
      iSpec,
      iProdQty,
      iWorkTime,
      iInputTime,
      iProductivity,
      iDefectQty,
      iStopTime,
      iStopTimeCols: iStopTimeCols ? [...iStopTimeCols] : [],
      iUtil,
      iOee,
      iDef: iDefRate,
      ...g,
      dateFrom: workDates.length ? workDates[0] : null,
      dateTo: workDates.length ? workDates[workDates.length - 1] : null,
      monthlyStats,
      equipmentStats,
      dailyStats,
      dayEquipmentStats,
      hasEquipmentColumn: iEquipment >= 0,
      _filterMatrix: matrix,
      _filterHi: hi,
      _filterIdx: idx,
    };
  }

  /**
   * 「페럴」 시트만 사용. 생산량(ERP), 정지 T~W열, 불량 X+AB열, 공정(절단·벤딩)별 월·설비 집계.
   * @param {any[][]} matrix
   */
  function parseParallelLogFromMatrix(matrix) {
    if (!matrix || matrix.length < 2) return null;
    const hi = detectHeaderRowIndex(matrix);
    const headerRow = matrix[hi] || [];

    const iWorkDate = findHeaderIndex(headerRow, ["작업일자", "작업일", ...DATE_KEYS]);
    const iSpec = findHeaderIndex(headerRow, ["규격"]);
    let iProdQty = findHeaderIndex(headerRow, ["생산량(ERP)", "생산량（ERP）", "생산량(erp)"]);
    if (iProdQty < 0) iProdQty = findDrawingProductionQtyColumnIndex(headerRow);
    const iWorkTime = findHeaderIndex(headerRow, ["작업시간", "작업 시간"]);
    const iInputTime = findHeaderIndex(headerRow, ["투입시간", "투입 시간"]);
    const iProductivity = findHeaderIndex(headerRow, ["생산성"]);
    const iDefectQty = -1;
    const iDefectQtyCols = headerRow.length > 27 ? [...PARALLEL_DEFECT_COL_X_AB] : [];
    const iStopTimeCols = headerRow.length > 22 ? [...PARALLEL_STOP_COL_TUVW] : [];
    const iStopTime = iStopTimeCols.length >= 4 ? -1 : findDrawingStopTimeColumnIndex(headerRow);

    const iUtil = findUtilizationRateColumnIndex(headerRow);
    const iOee = findOeeColumnIndex(headerRow);
    const iDefRate = findDefectRateColumnIndex(headerRow);

    const iProcess = findHeaderIndex(headerRow, [
      "공정",
      "작업공정",
      "공정명",
      "세부공정",
      "공정구분",
      "작업구분",
    ]);

    const hasCore =
      iWorkDate >= 0 ||
      iSpec >= 0 ||
      iProdQty >= 0 ||
      iWorkTime >= 0 ||
      iInputTime >= 0 ||
      iProductivity >= 0 ||
      iDefectQtyCols.length > 0 ||
      iStopTime >= 0 ||
      iStopTimeCols.length > 0;
    const hasKpi = iUtil >= 0 || iOee >= 0 || iDefRate >= 0;
    if (!hasCore && !hasKpi) return null;

    const iEquipment = findEquipmentColumnIndex(headerRow);

    const idx = {
      iWorkDate,
      iUtil,
      iOee,
      iDefRate,
      iProdQty,
      iWorkTime,
      iInputTime,
      iProductivity,
      iDefectQty,
      iDefectQtyCols,
      iStopTime,
      iStopTimeCols: iStopTimeCols || [],
    };

    const aggGlobal = makeDrawingAgg();
    /** @type {Map<string, ReturnType<typeof makeDrawingAgg>>} */
    const byMonth = new Map();
    /** @type {Map<string, ReturnType<typeof makeDrawingAgg>>} */
    const byEq = new Map();
    /** @type {Map<string, ReturnType<typeof makeDrawingAgg>>} */
    const byDay = new Map();
    /** @type {Map<string, Map<string, ReturnType<typeof makeDrawingAgg>>>} */
    const byDayEq = new Map();

    /** @type {Map<string, Map<string, ReturnType<typeof makeDrawingAgg>>>} */
    const byProcMonth = new Map();
    /** @type {Map<string, Map<string, ReturnType<typeof makeDrawingAgg>>>} */
    const byProcEq = new Map();
    /** @type {Map<string, Map<string, ReturnType<typeof makeDrawingAgg>>>} */
    const byProcDay = new Map();
    /** @type {Map<string, Map<string, Map<string, ReturnType<typeof makeDrawingAgg>>>>} */
    const byProcDayEq = new Map();

    for (let r = hi + 1; r < matrix.length; r++) {
      const row = matrix[r];
      if (!row || row.length === 0) continue;
      const nonempty = row.some((c) => String(c ?? "").trim() !== "");
      if (!nonempty) continue;

      addDrawingRowToAgg(aggGlobal, row, idx, true);

      const ymd = iWorkDate >= 0 ? parseExcelDate(row[iWorkDate]) : "";
      const ym = ymd && ymd.length >= 7 ? ymd.slice(0, 7) : "일자미상";
      if (!byMonth.has(ym)) byMonth.set(ym, makeDrawingAgg());
      addDrawingRowToAgg(byMonth.get(ym), row, idx, false);

      const eqKey = iEquipment >= 0 ? String(row[iEquipment] ?? "").trim() || "미지정" : "전체";
      if (!byEq.has(eqKey)) byEq.set(eqKey, makeDrawingAgg());
      addDrawingRowToAgg(byEq.get(eqKey), row, idx, false);

      const dayKey = ymd || "일자미상";
      if (!byDay.has(dayKey)) byDay.set(dayKey, makeDrawingAgg());
      addDrawingRowToAgg(byDay.get(dayKey), row, idx, false);
      if (!byDayEq.has(dayKey)) byDayEq.set(dayKey, new Map());
      const dayEqMap = byDayEq.get(dayKey);
      if (!dayEqMap.has(eqKey)) dayEqMap.set(eqKey, makeDrawingAgg());
      addDrawingRowToAgg(dayEqMap.get(eqKey), row, idx, false);

      const procKey = normalizeParallelProcess(iProcess >= 0 ? row[iProcess] : "");
      addDrawingRowToAgg(ensureAggMapNested2(byProcMonth, procKey, ym), row, idx, false);
      addDrawingRowToAgg(ensureAggMapNested2(byProcEq, procKey, eqKey), row, idx, false);
      addDrawingRowToAgg(ensureAggMapNested2(byProcDay, procKey, dayKey), row, idx, false);
      addDrawingRowToAgg(ensureAggMapNested3ProcDayEq(byProcDayEq, procKey, dayKey, eqKey), row, idx, false);
    }

    aggGlobal.workDates.sort();
    const g = finalizeDrawingAgg(aggGlobal, idx);
    const workDates = aggGlobal.workDates;

    const monthLabel = (ymKey) => {
      if (ymKey === "일자미상") return ymKey;
      const p = ymKey.split("-");
      if (p.length !== 2) return ymKey;
      return `${p[0]}년 ${parseInt(p[1], 10)}월`;
    };

    const monthlyStats = [...byMonth.entries()]
      .sort((a, b) => {
        if (a[0] === "일자미상") return 1;
        if (b[0] === "일자미상") return -1;
        return a[0].localeCompare(b[0]);
      })
      .map(([key, agg]) => ({ key, label: monthLabel(key), ...finalizeDrawingAgg(agg, idx) }));

    const equipmentStats = [...byEq.entries()]
      .sort((a, b) => compareDrawingEquipmentKeys(a[0], b[0]))
      .map(([key, agg]) => ({ key, label: key, ...finalizeDrawingAgg(agg, idx) }));

    const dailyStats = [...byDay.entries()]
      .sort((a, b) => {
        if (a[0] === "일자미상") return 1;
        if (b[0] === "일자미상") return -1;
        return a[0].localeCompare(b[0]);
      })
      .map(([key, agg]) => ({ key, label: key, ...finalizeDrawingAgg(agg, idx) }));

    /** @type {Record<string, any[]>} */
    const dayEquipmentStats = {};
    for (const [dKey, eqMap] of byDayEq.entries()) {
      dayEquipmentStats[dKey] = [...eqMap.entries()]
        .sort((a, b) => compareDrawingEquipmentKeys(a[0], b[0]))
        .map(([key, agg]) => ({ key, label: key, ...finalizeDrawingAgg(agg, idx) }));
    }

    const sortProcKeys = (a, b) => {
      const rank = (k) => (k === "절단" ? 0 : k === "벤딩" ? 1 : k === "기타" ? 2 : 99);
      const ra = rank(a);
      const rb = rank(b);
      if (ra !== rb) return ra - rb;
      return String(a).localeCompare(String(b));
    };

    const procKeySet = new Set([...byProcMonth.keys(), ...byProcEq.keys()]);
    const processBlocks = [...procKeySet].sort(sortProcKeys).map((pk) => {
      const pm = byProcMonth.get(pk) || new Map();
      const peq = byProcEq.get(pk) || new Map();
      const pd = byProcDay.get(pk) || new Map();
      const pde = byProcDayEq.get(pk) || new Map();

      const monthlyStatsP = [...pm.entries()]
        .sort((a, b) => {
          if (a[0] === "일자미상") return 1;
          if (b[0] === "일자미상") return -1;
          return a[0].localeCompare(b[0]);
        })
        .map(([key, agg]) => ({ key, label: monthLabel(key), ...finalizeDrawingAgg(agg, idx) }));

      const equipmentStatsP = [...peq.entries()]
        .sort((a, b) => compareDrawingEquipmentKeys(a[0], b[0]))
        .map(([key, agg]) => ({ key, label: key, ...finalizeDrawingAgg(agg, idx) }));

      const dailyStatsP = [...pd.entries()]
        .sort((a, b) => {
          if (a[0] === "일자미상") return 1;
          if (b[0] === "일자미상") return -1;
          return a[0].localeCompare(b[0]);
        })
        .map(([key, agg]) => ({ key, label: key, ...finalizeDrawingAgg(agg, idx) }));

      /** @type {Record<string, any[]>} */
      const dayEquipmentStatsP = {};
      for (const [dKey, eqMap] of pde.entries()) {
        dayEquipmentStatsP[dKey] = [...eqMap.entries()]
          .sort((a, b) => compareDrawingEquipmentKeys(a[0], b[0]))
          .map(([key, agg]) => ({ key, label: key, ...finalizeDrawingAgg(agg, idx) }));
      }

      return {
        process: pk,
        monthlyStats: monthlyStatsP,
        equipmentStats: equipmentStatsP,
        dailyStats: dailyStatsP,
        dayEquipmentStats: dayEquipmentStatsP,
      };
    });

    return {
      hasAny: true,
      headerRowIndex: hi,
      iWorkDate,
      iEquipment,
      iSpec,
      iProdQty,
      iWorkTime,
      iInputTime,
      iProductivity,
      iDefectQty,
      iDefectQtyCols: [...iDefectQtyCols],
      iStopTime,
      iStopTimeCols: iStopTimeCols ? [...iStopTimeCols] : [],
      iProcess,
      iUtil,
      iOee,
      iDef: iDefRate,
      ...g,
      dateFrom: workDates.length ? workDates[0] : null,
      dateTo: workDates.length ? workDates[workDates.length - 1] : null,
      monthlyStats,
      equipmentStats,
      dailyStats,
      dayEquipmentStats,
      processBlocks,
      hasEquipmentColumn: iEquipment >= 0,
      _filterMatrix: matrix,
      _filterHi: hi,
      _filterIdx: idx,
    };
  }

  /**
   * 「이중관」 시트. 생산 N열, 비가동 S~V열, 불량 W+AJ열, 공정 가공·성형.
   * @param {any[][]} matrix
   */
  function parseDoublePipeLogFromMatrix(matrix) {
    if (!matrix || matrix.length < 2) return null;
    const hi = detectHeaderRowIndex(matrix);
    const headerRow = matrix[hi] || [];

    const iWorkDate = findHeaderIndex(headerRow, ["작업일자", "작업일", ...DATE_KEYS]);
    const iSpec = findHeaderIndex(headerRow, ["규격"]);
    let iProdQty = findHeaderIndex(headerRow, ["생산량(ERP)", "생산량（ERP）", "생산량(erp)", "생산수량", "생산 수량", "생산량"]);
    if (iProdQty < 0) iProdQty = findDrawingProductionQtyColumnIndex(headerRow);
    if (iProdQty < 0 && headerRow.length > DOUBLE_PIPE_PROD_COL_N) iProdQty = DOUBLE_PIPE_PROD_COL_N;
    const iWorkTime = findHeaderIndex(headerRow, ["작업시간", "작업 시간"]);
    const iInputTime = findHeaderIndex(headerRow, ["투입시간", "투입 시간"]);
    const iProductivity = findHeaderIndex(headerRow, ["생산성"]);
    const iDefectQty = -1;
    const iDefectQtyCols = [];
    if (headerRow.length > DOUBLE_PIPE_DEFECT_COL_W_AJ[1]) iDefectQtyCols.push(...DOUBLE_PIPE_DEFECT_COL_W_AJ);
    else if (headerRow.length > DOUBLE_PIPE_DEFECT_COL_W_AJ[0]) iDefectQtyCols.push(DOUBLE_PIPE_DEFECT_COL_W_AJ[0]);
    const iStopTimeCols = headerRow.length > 21 ? [...DOUBLE_PIPE_STOP_COL_STUV] : [];
    const iStopTime = iStopTimeCols.length >= 4 ? -1 : findDrawingStopTimeColumnIndex(headerRow);

    const iUtil = findUtilizationRateColumnIndex(headerRow);
    const iOee = findOeeColumnIndex(headerRow);
    const iDefRate = findDefectRateColumnIndex(headerRow);

    const iProcess = findHeaderIndex(headerRow, [
      "공정",
      "작업공정",
      "공정명",
      "세부공정",
      "공정구분",
      "작업구분",
    ]);

    const hasCore =
      iWorkDate >= 0 ||
      iSpec >= 0 ||
      iProdQty >= 0 ||
      iWorkTime >= 0 ||
      iInputTime >= 0 ||
      iProductivity >= 0 ||
      iDefectQtyCols.length > 0 ||
      iStopTime >= 0 ||
      iStopTimeCols.length > 0;
    const hasKpi = iUtil >= 0 || iOee >= 0 || iDefRate >= 0;
    if (!hasCore && !hasKpi) return null;

    const iEquipment = findEquipmentColumnIndex(headerRow);

    const idx = {
      iWorkDate,
      iUtil,
      iOee,
      iDefRate,
      iProdQty,
      iWorkTime,
      iInputTime,
      iProductivity,
      iDefectQty,
      iDefectQtyCols,
      iStopTime,
      iStopTimeCols: iStopTimeCols || [],
    };

    const aggGlobal = makeDrawingAgg();
    const byMonth = new Map();
    const byEq = new Map();
    const byDay = new Map();
    const byDayEq = new Map();
    const byProcMonth = new Map();
    const byProcEq = new Map();
    const byProcDay = new Map();
    const byProcDayEq = new Map();

    for (let r = hi + 1; r < matrix.length; r++) {
      const row = matrix[r];
      if (!row || row.length === 0) continue;
      const nonempty = row.some((c) => String(c ?? "").trim() !== "");
      if (!nonempty) continue;

      addDrawingRowToAgg(aggGlobal, row, idx, true);

      const ymd = iWorkDate >= 0 ? parseExcelDate(row[iWorkDate]) : "";
      const ym = ymd && ymd.length >= 7 ? ymd.slice(0, 7) : "일자미상";
      if (!byMonth.has(ym)) byMonth.set(ym, makeDrawingAgg());
      addDrawingRowToAgg(byMonth.get(ym), row, idx, false);

      const eqKey = iEquipment >= 0 ? String(row[iEquipment] ?? "").trim() || "미지정" : "전체";
      if (!byEq.has(eqKey)) byEq.set(eqKey, makeDrawingAgg());
      addDrawingRowToAgg(byEq.get(eqKey), row, idx, false);

      const dayKey = ymd || "일자미상";
      if (!byDay.has(dayKey)) byDay.set(dayKey, makeDrawingAgg());
      addDrawingRowToAgg(byDay.get(dayKey), row, idx, false);
      if (!byDayEq.has(dayKey)) byDayEq.set(dayKey, new Map());
      const dayEqMap = byDayEq.get(dayKey);
      if (!dayEqMap.has(eqKey)) dayEqMap.set(eqKey, makeDrawingAgg());
      addDrawingRowToAgg(dayEqMap.get(eqKey), row, idx, false);

      const procKey = normalizeDoublePipeProcess(iProcess >= 0 ? row[iProcess] : "");
      addDrawingRowToAgg(ensureAggMapNested2(byProcMonth, procKey, ym), row, idx, false);
      addDrawingRowToAgg(ensureAggMapNested2(byProcEq, procKey, eqKey), row, idx, false);
      addDrawingRowToAgg(ensureAggMapNested2(byProcDay, procKey, dayKey), row, idx, false);
      addDrawingRowToAgg(ensureAggMapNested3ProcDayEq(byProcDayEq, procKey, dayKey, eqKey), row, idx, false);
    }

    aggGlobal.workDates.sort();
    const g = finalizeDrawingAgg(aggGlobal, idx);
    const workDates = aggGlobal.workDates;

    const monthLabel = (ymKey) => {
      if (ymKey === "일자미상") return ymKey;
      const p = ymKey.split("-");
      if (p.length !== 2) return ymKey;
      return `${p[0]}년 ${parseInt(p[1], 10)}월`;
    };

    const monthlyStats = [...byMonth.entries()]
      .sort((a, b) => {
        if (a[0] === "일자미상") return 1;
        if (b[0] === "일자미상") return -1;
        return a[0].localeCompare(b[0]);
      })
      .map(([key, agg]) => ({ key, label: monthLabel(key), ...finalizeDrawingAgg(agg, idx) }));

    const equipmentStats = [...byEq.entries()]
      .sort((a, b) => compareDrawingEquipmentKeys(a[0], b[0]))
      .map(([key, agg]) => ({ key, label: key, ...finalizeDrawingAgg(agg, idx) }));

    const dailyStats = [...byDay.entries()]
      .sort((a, b) => {
        if (a[0] === "일자미상") return 1;
        if (b[0] === "일자미상") return -1;
        return a[0].localeCompare(b[0]);
      })
      .map(([key, agg]) => ({ key, label: key, ...finalizeDrawingAgg(agg, idx) }));

    const dayEquipmentStats = {};
    for (const [dKey, eqMap] of byDayEq.entries()) {
      dayEquipmentStats[dKey] = [...eqMap.entries()]
        .sort((a, b) => compareDrawingEquipmentKeys(a[0], b[0]))
        .map(([key, agg]) => ({ key, label: key, ...finalizeDrawingAgg(agg, idx) }));
    }

    const sortProcKeys = (a, b) => {
      const rank = (k) => (k === "가공" ? 0 : k === "성형" ? 1 : k === "기타" ? 2 : 99);
      const ra = rank(a);
      const rb = rank(b);
      if (ra !== rb) return ra - rb;
      return String(a).localeCompare(String(b));
    };

    const procKeySet = new Set([...byProcMonth.keys(), ...byProcEq.keys()]);
    const processBlocks = [...procKeySet].sort(sortProcKeys).map((pk) => {
      const pm = byProcMonth.get(pk) || new Map();
      const peq = byProcEq.get(pk) || new Map();
      const pd = byProcDay.get(pk) || new Map();
      const pde = byProcDayEq.get(pk) || new Map();

      const monthlyStatsP = [...pm.entries()]
        .sort((a, b) => {
          if (a[0] === "일자미상") return 1;
          if (b[0] === "일자미상") return -1;
          return a[0].localeCompare(b[0]);
        })
        .map(([key, agg]) => ({ key, label: monthLabel(key), ...finalizeDrawingAgg(agg, idx) }));

      const equipmentStatsP = [...peq.entries()]
        .sort((a, b) => compareDrawingEquipmentKeys(a[0], b[0]))
        .map(([key, agg]) => ({ key, label: key, ...finalizeDrawingAgg(agg, idx) }));

      const dailyStatsP = [...pd.entries()]
        .sort((a, b) => {
          if (a[0] === "일자미상") return 1;
          if (b[0] === "일자미상") return -1;
          return a[0].localeCompare(b[0]);
        })
        .map(([key, agg]) => ({ key, label: key, ...finalizeDrawingAgg(agg, idx) }));

      const dayEquipmentStatsP = {};
      for (const [dKey, eqMap] of pde.entries()) {
        dayEquipmentStatsP[dKey] = [...eqMap.entries()]
          .sort((a, b) => compareDrawingEquipmentKeys(a[0], b[0]))
          .map(([key, agg]) => ({ key, label: key, ...finalizeDrawingAgg(agg, idx) }));
      }

      return {
        process: pk,
        monthlyStats: monthlyStatsP,
        equipmentStats: equipmentStatsP,
        dailyStats: dailyStatsP,
        dayEquipmentStats: dayEquipmentStatsP,
      };
    });

    return {
      hasAny: true,
      headerRowIndex: hi,
      iWorkDate,
      iEquipment,
      iSpec,
      iProdQty,
      iWorkTime,
      iInputTime,
      iProductivity,
      iDefectQty,
      iDefectQtyCols: [...iDefectQtyCols],
      iStopTime,
      iStopTimeCols: iStopTimeCols ? [...iStopTimeCols] : [],
      iProcess,
      iUtil,
      iOee,
      iDef: iDefRate,
      ...g,
      dateFrom: workDates.length ? workDates[0] : null,
      dateTo: workDates.length ? workDates[workDates.length - 1] : null,
      monthlyStats,
      equipmentStats,
      dailyStats,
      dayEquipmentStats,
      processBlocks,
      hasEquipmentColumn: iEquipment >= 0,
      _filterMatrix: matrix,
      _filterHi: hi,
      _filterIdx: idx,
    };
  }

  /**
   * 「화승이중관」 시트. 생산 L열, 비가동 R~V열, 불량 소재(X+Y+Z)+공정(AK+AL+AR), 공정 전조·스파이럴 등 6종.
   * @param {any[][]} matrix
   */
  function parseHwaseungDoublePipeLogFromMatrix(matrix) {
    if (!matrix || matrix.length < 2) return null;
    const hi = findHwaseungDoublePipeHeaderRowIndex(matrix);
    const headerRow = matrix[hi] || [];
    const layoutWidth = Math.max(
      matrixRowMaxUsedCol(headerRow),
      inferHwaseungDoublePipeLayoutWidth(matrix, hi),
      HWASEUNG_DEFECT_PROCESS_AK_AL_AR[HWASEUNG_DEFECT_PROCESS_AK_AL_AR.length - 1] + 1
    );

    let iWorkDate = findHeaderIndex(headerRow, ["작업일자", "작업일", ...DATE_KEYS]);
    if (iWorkDate < 0 && layoutWidth > 5) {
      let ok = 0;
      let trials = 0;
      for (let r = hi + 1; r < Math.min(matrix.length, hi + 50); r++) {
        const row = matrix[r];
        if (!row || !row.length) continue;
        const v = row[0];
        if (v === "" || v === undefined || v === null) continue;
        trials++;
        const y = parseExcelDate(v);
        if (/^\d{4}-\d{2}-\d{2}$/.test(y)) ok++;
      }
      if (trials >= 4 && ok / trials >= 0.55) iWorkDate = 0;
    }
    const iSpec = findHeaderIndex(headerRow, ["규격"]);
    let iProdQty = findHeaderIndex(headerRow, ["생산량(ERP)", "생산량（ERP）", "생산량(erp)", "생산수량", "생산 수량", "생산량"]);
    if (iProdQty < 0) iProdQty = findDrawingProductionQtyColumnIndex(headerRow);
    if (iProdQty < 0 && layoutWidth > HWASEUNG_PROD_COL_L) iProdQty = HWASEUNG_PROD_COL_L;
    const iWorkTime = findHeaderIndex(headerRow, ["작업시간", "작업 시간"]);
    const iInputTime = findHeaderIndex(headerRow, ["투입시간", "투입 시간"]);
    const iProductivity = findHeaderIndex(headerRow, ["생산성"]);
    const iDefectQty = -1;
    const iDefectQtyCols = [];
    if (layoutWidth > 43) {
      iDefectQtyCols.push(...HWASEUNG_DEFECT_MATERIAL_XYZ, ...HWASEUNG_DEFECT_PROCESS_AK_AL_AR);
    } else if (layoutWidth > 25) {
      iDefectQtyCols.push(...HWASEUNG_DEFECT_MATERIAL_XYZ);
    }
    const iStopTimeCols = layoutWidth > 21 ? [...HWASEUNG_STOP_COL_RSTUV] : [];
    const iStopTime = iStopTimeCols.length >= 4 ? -1 : findDrawingStopTimeColumnIndex(headerRow);

    const iUtil = findUtilizationRateColumnIndex(headerRow);
    const iOee = findOeeColumnIndex(headerRow);
    const iDefRate = findDefectRateColumnIndex(headerRow);

    const iProcess = findHeaderIndex(headerRow, [
      "공정",
      "작업공정",
      "공정명",
      "세부공정",
      "공정구분",
      "작업구분",
    ]);

    const hasCore =
      iWorkDate >= 0 ||
      iSpec >= 0 ||
      iProdQty >= 0 ||
      iWorkTime >= 0 ||
      iInputTime >= 0 ||
      iProductivity >= 0 ||
      iDefectQtyCols.length > 0 ||
      iStopTime >= 0 ||
      iStopTimeCols.length > 0;
    const hasKpi = iUtil >= 0 || iOee >= 0 || iDefRate >= 0;
    if (!hasCore && !hasKpi) return null;

    const iEquipment = findEquipmentColumnIndex(headerRow);

    const idx = {
      iWorkDate,
      iUtil,
      iOee,
      iDefRate,
      iProdQty,
      iWorkTime,
      iInputTime,
      iProductivity,
      iDefectQty,
      iDefectQtyCols,
      iStopTime,
      iStopTimeCols: iStopTimeCols || [],
    };

    const aggGlobal = makeDrawingAgg();
    const byMonth = new Map();
    const byEq = new Map();
    const byDay = new Map();
    const byDayEq = new Map();
    const byProcMonth = new Map();
    const byProcEq = new Map();
    const byProcDay = new Map();
    const byProcDayEq = new Map();

    for (let r = hi + 1; r < matrix.length; r++) {
      const row = matrix[r];
      if (!row || row.length === 0) continue;
      const nonempty = row.some((c) => String(c ?? "").trim() !== "");
      if (!nonempty) continue;

      addDrawingRowToAgg(aggGlobal, row, idx, true);

      const ymd = iWorkDate >= 0 ? parseExcelDate(row[iWorkDate]) : "";
      const ym = ymd && ymd.length >= 7 ? ymd.slice(0, 7) : "일자미상";
      if (!byMonth.has(ym)) byMonth.set(ym, makeDrawingAgg());
      addDrawingRowToAgg(byMonth.get(ym), row, idx, false);

      const eqKey = iEquipment >= 0 ? String(row[iEquipment] ?? "").trim() || "미지정" : "전체";
      if (!byEq.has(eqKey)) byEq.set(eqKey, makeDrawingAgg());
      addDrawingRowToAgg(byEq.get(eqKey), row, idx, false);

      const dayKey = ymd || "일자미상";
      if (!byDay.has(dayKey)) byDay.set(dayKey, makeDrawingAgg());
      addDrawingRowToAgg(byDay.get(dayKey), row, idx, false);
      if (!byDayEq.has(dayKey)) byDayEq.set(dayKey, new Map());
      const dayEqMap = byDayEq.get(dayKey);
      if (!dayEqMap.has(eqKey)) dayEqMap.set(eqKey, makeDrawingAgg());
      addDrawingRowToAgg(dayEqMap.get(eqKey), row, idx, false);

      const procKey = normalizeHwaseungDoublePipeProcess(iProcess >= 0 ? row[iProcess] : "");
      addDrawingRowToAgg(ensureAggMapNested2(byProcMonth, procKey, ym), row, idx, false);
      addDrawingRowToAgg(ensureAggMapNested2(byProcEq, procKey, eqKey), row, idx, false);
      addDrawingRowToAgg(ensureAggMapNested2(byProcDay, procKey, dayKey), row, idx, false);
      addDrawingRowToAgg(ensureAggMapNested3ProcDayEq(byProcDayEq, procKey, dayKey, eqKey), row, idx, false);
    }

    aggGlobal.workDates.sort();
    const g = finalizeDrawingAgg(aggGlobal, idx);
    const workDates = aggGlobal.workDates;

    const monthLabel = (ymKey) => {
      if (ymKey === "일자미상") return ymKey;
      const p = ymKey.split("-");
      if (p.length !== 2) return ymKey;
      return `${p[0]}년 ${parseInt(p[1], 10)}월`;
    };

    const monthlyStats = [...byMonth.entries()]
      .sort((a, b) => {
        if (a[0] === "일자미상") return 1;
        if (b[0] === "일자미상") return -1;
        return a[0].localeCompare(b[0]);
      })
      .map(([key, agg]) => ({ key, label: monthLabel(key), ...finalizeDrawingAgg(agg, idx) }));

    const equipmentStats = [...byEq.entries()]
      .sort((a, b) => compareDrawingEquipmentKeys(a[0], b[0]))
      .map(([key, agg]) => ({ key, label: key, ...finalizeDrawingAgg(agg, idx) }));

    const dailyStats = [...byDay.entries()]
      .sort((a, b) => {
        if (a[0] === "일자미상") return 1;
        if (b[0] === "일자미상") return -1;
        return a[0].localeCompare(b[0]);
      })
      .map(([key, agg]) => ({ key, label: key, ...finalizeDrawingAgg(agg, idx) }));

    const dayEquipmentStats = {};
    for (const [dKey, eqMap] of byDayEq.entries()) {
      dayEquipmentStats[dKey] = [...eqMap.entries()]
        .sort((a, b) => compareDrawingEquipmentKeys(a[0], b[0]))
        .map(([key, agg]) => ({ key, label: key, ...finalizeDrawingAgg(agg, idx) }));
    }

    const procKeySet = new Set([...byProcMonth.keys(), ...byProcEq.keys()]);
    const processBlocks = [...procKeySet].sort(sortHwaseungDoublePipeProcessKeys).map((pk) => {
      const pm = byProcMonth.get(pk) || new Map();
      const peq = byProcEq.get(pk) || new Map();
      const pd = byProcDay.get(pk) || new Map();
      const pde = byProcDayEq.get(pk) || new Map();

      const monthlyStatsP = [...pm.entries()]
        .sort((a, b) => {
          if (a[0] === "일자미상") return 1;
          if (b[0] === "일자미상") return -1;
          return a[0].localeCompare(b[0]);
        })
        .map(([key, agg]) => ({ key, label: monthLabel(key), ...finalizeDrawingAgg(agg, idx) }));

      const equipmentStatsP = [...peq.entries()]
        .sort((a, b) => compareDrawingEquipmentKeys(a[0], b[0]))
        .map(([key, agg]) => ({ key, label: key, ...finalizeDrawingAgg(agg, idx) }));

      const dailyStatsP = [...pd.entries()]
        .sort((a, b) => {
          if (a[0] === "일자미상") return 1;
          if (b[0] === "일자미상") return -1;
          return a[0].localeCompare(b[0]);
        })
        .map(([key, agg]) => ({ key, label: key, ...finalizeDrawingAgg(agg, idx) }));

      const dayEquipmentStatsP = {};
      for (const [dKey, eqMap] of pde.entries()) {
        dayEquipmentStatsP[dKey] = [...eqMap.entries()]
          .sort((a, b) => compareDrawingEquipmentKeys(a[0], b[0]))
          .map(([key, agg]) => ({ key, label: key, ...finalizeDrawingAgg(agg, idx) }));
      }

      return {
        process: pk,
        monthlyStats: monthlyStatsP,
        equipmentStats: equipmentStatsP,
        dailyStats: dailyStatsP,
        dayEquipmentStats: dayEquipmentStatsP,
      };
    });

    return {
      hasAny: true,
      headerRowIndex: hi,
      iWorkDate,
      iEquipment,
      iSpec,
      iProdQty,
      iWorkTime,
      iInputTime,
      iProductivity,
      iDefectQty,
      iDefectQtyCols: [...iDefectQtyCols],
      iStopTime,
      iStopTimeCols: iStopTimeCols ? [...iStopTimeCols] : [],
      iProcess,
      iUtil,
      iOee,
      iDef: iDefRate,
      ...g,
      dateFrom: workDates.length ? workDates[0] : null,
      dateTo: workDates.length ? workDates[workDates.length - 1] : null,
      monthlyStats,
      equipmentStats,
      dailyStats,
      dayEquipmentStats,
      processBlocks,
      hasEquipmentColumn: iEquipment >= 0,
      _filterMatrix: matrix,
      _filterHi: hi,
      _filterIdx: idx,
    };
  }

  /**
   * 「머플러」「단조머플러」 시트 각각에 동일 적용. M열 생산, R~U 정지(4종), 불량=V+AH~AL 합, 공정 절단·포밍·가공.
   * @param {any[][]} matrix
   */
  function parseMufflerLogFromMatrix(matrix) {
    if (!matrix || matrix.length < 2) return null;
    const hi = detectHeaderRowIndex(matrix);
    const headerRow = matrix[hi] || [];

    const iWorkDate = findHeaderIndex(headerRow, ["작업일자", "작업일", ...DATE_KEYS]);
    const iSpec = findHeaderIndex(headerRow, ["규격"]);
    let iProdQty = findHeaderIndex(headerRow, [
      "생산수량",
      "생산 수량",
      "생산량",
      "생산량(ERP)",
      "생산량（ERP）",
      "생산량(erp)",
    ]);
    if (iProdQty < 0) iProdQty = findDrawingProductionQtyColumnIndex(headerRow);
    if (iProdQty < 0 && headerRow.length > MUFFLER_PROD_COL_M) iProdQty = MUFFLER_PROD_COL_M;

    const iWorkTime = findHeaderIndex(headerRow, ["작업시간", "작업 시간"]);
    const iInputTime = findHeaderIndex(headerRow, ["투입시간", "투입 시간"]);
    const iProductivity = findHeaderIndex(headerRow, ["생산성"]);
    const iDefectQty = -1;
    const iDefectQtyCols = [];
    if (headerRow.length > MUFFLER_DEFECT_COL_V) iDefectQtyCols.push(MUFFLER_DEFECT_COL_V);
    const lastDetail = MUFFLER_DEFECT_PROC_DETAIL_COLS_AH_AL[MUFFLER_DEFECT_PROC_DETAIL_COLS_AH_AL.length - 1];
    if (headerRow.length > lastDetail) {
      for (const j of MUFFLER_DEFECT_PROC_DETAIL_COLS_AH_AL) iDefectQtyCols.push(j);
    }
    const iStopTimeCols = headerRow.length > 20 ? [...MUFFLER_STOP_COL_RSTU] : [];
    const iStopTime = iStopTimeCols.length >= 4 ? -1 : findDrawingStopTimeColumnIndex(headerRow);

    const iUtil = findUtilizationRateColumnIndex(headerRow);
    const iOee = findOeeColumnIndex(headerRow);
    const iDefRate = findDefectRateColumnIndex(headerRow);

    const iProcess = findHeaderIndex(headerRow, [
      "공정",
      "작업공정",
      "공정명",
      "세부공정",
      "공정구분",
      "작업구분",
    ]);

    const hasCore =
      iWorkDate >= 0 ||
      iSpec >= 0 ||
      iProdQty >= 0 ||
      iWorkTime >= 0 ||
      iInputTime >= 0 ||
      iProductivity >= 0 ||
      iDefectQtyCols.length > 0 ||
      iStopTime >= 0 ||
      iStopTimeCols.length > 0;
    const hasKpi = iUtil >= 0 || iOee >= 0 || iDefRate >= 0;
    if (!hasCore && !hasKpi) return null;

    const iEquipment = findEquipmentColumnIndex(headerRow);

    const idx = {
      iWorkDate,
      iUtil,
      iOee,
      iDefRate,
      iProdQty,
      iWorkTime,
      iInputTime,
      iProductivity,
      iDefectQty,
      iDefectQtyCols: [...iDefectQtyCols],
      iStopTime,
      iStopTimeCols: iStopTimeCols || [],
    };

    const aggGlobal = makeDrawingAgg();
    const byMonth = new Map();
    const byEq = new Map();
    const byDay = new Map();
    const byDayEq = new Map();
    const byProcMonth = new Map();
    const byProcEq = new Map();
    const byProcDay = new Map();
    const byProcDayEq = new Map();

    for (let r = hi + 1; r < matrix.length; r++) {
      const row = matrix[r];
      if (!row || row.length === 0) continue;
      const nonempty = row.some((c) => String(c ?? "").trim() !== "");
      if (!nonempty) continue;

      addDrawingRowToAgg(aggGlobal, row, idx, true);

      const ymd = iWorkDate >= 0 ? parseExcelDate(row[iWorkDate]) : "";
      const ym = ymd && ymd.length >= 7 ? ymd.slice(0, 7) : "일자미상";
      if (!byMonth.has(ym)) byMonth.set(ym, makeDrawingAgg());
      addDrawingRowToAgg(byMonth.get(ym), row, idx, false);

      const eqKey = iEquipment >= 0 ? String(row[iEquipment] ?? "").trim() || "미지정" : "전체";
      if (!byEq.has(eqKey)) byEq.set(eqKey, makeDrawingAgg());
      addDrawingRowToAgg(byEq.get(eqKey), row, idx, false);

      const dayKey = ymd || "일자미상";
      if (!byDay.has(dayKey)) byDay.set(dayKey, makeDrawingAgg());
      addDrawingRowToAgg(byDay.get(dayKey), row, idx, false);
      if (!byDayEq.has(dayKey)) byDayEq.set(dayKey, new Map());
      const dayEqMap = byDayEq.get(dayKey);
      if (!dayEqMap.has(eqKey)) dayEqMap.set(eqKey, makeDrawingAgg());
      addDrawingRowToAgg(dayEqMap.get(eqKey), row, idx, false);

      const procKey = normalizeMufflerProcess(iProcess >= 0 ? row[iProcess] : "");
      addDrawingRowToAgg(ensureAggMapNested2(byProcMonth, procKey, ym), row, idx, false);
      addDrawingRowToAgg(ensureAggMapNested2(byProcEq, procKey, eqKey), row, idx, false);
      addDrawingRowToAgg(ensureAggMapNested2(byProcDay, procKey, dayKey), row, idx, false);
      addDrawingRowToAgg(ensureAggMapNested3ProcDayEq(byProcDayEq, procKey, dayKey, eqKey), row, idx, false);
    }

    aggGlobal.workDates.sort();
    const g = finalizeDrawingAgg(aggGlobal, idx);
    const workDates = aggGlobal.workDates;

    const monthLabel = (ymKey) => {
      if (ymKey === "일자미상") return ymKey;
      const p = ymKey.split("-");
      if (p.length !== 2) return ymKey;
      return `${p[0]}년 ${parseInt(p[1], 10)}월`;
    };

    const monthlyStats = [...byMonth.entries()]
      .sort((a, b) => {
        if (a[0] === "일자미상") return 1;
        if (b[0] === "일자미상") return -1;
        return a[0].localeCompare(b[0]);
      })
      .map(([key, agg]) => ({ key, label: monthLabel(key), ...finalizeDrawingAgg(agg, idx) }));

    const equipmentStats = [...byEq.entries()]
      .sort((a, b) => compareDrawingEquipmentKeys(a[0], b[0]))
      .map(([key, agg]) => ({ key, label: key, ...finalizeDrawingAgg(agg, idx) }));

    const dailyStats = [...byDay.entries()]
      .sort((a, b) => {
        if (a[0] === "일자미상") return 1;
        if (b[0] === "일자미상") return -1;
        return a[0].localeCompare(b[0]);
      })
      .map(([key, agg]) => ({ key, label: key, ...finalizeDrawingAgg(agg, idx) }));

    const dayEquipmentStats = {};
    for (const [dKey, eqMap] of byDayEq.entries()) {
      dayEquipmentStats[dKey] = [...eqMap.entries()]
        .sort((a, b) => compareDrawingEquipmentKeys(a[0], b[0]))
        .map(([key, agg]) => ({ key, label: key, ...finalizeDrawingAgg(agg, idx) }));
    }

    const sortProcKeys = (a, b) => {
      const rank = (k) => (k === "절단" ? 0 : k === "포밍" ? 1 : k === "가공" ? 2 : k === "기타" ? 3 : 99);
      const ra = rank(a);
      const rb = rank(b);
      if (ra !== rb) return ra - rb;
      return String(a).localeCompare(String(b));
    };

    const procKeySet = new Set([...byProcMonth.keys(), ...byProcEq.keys()]);
    const processBlocks = [...procKeySet].sort(sortProcKeys).map((pk) => {
      const pm = byProcMonth.get(pk) || new Map();
      const peq = byProcEq.get(pk) || new Map();
      const pd = byProcDay.get(pk) || new Map();
      const pde = byProcDayEq.get(pk) || new Map();

      const monthlyStatsP = [...pm.entries()]
        .sort((a, b) => {
          if (a[0] === "일자미상") return 1;
          if (b[0] === "일자미상") return -1;
          return a[0].localeCompare(b[0]);
        })
        .map(([key, agg]) => ({ key, label: monthLabel(key), ...finalizeDrawingAgg(agg, idx) }));

      const equipmentStatsP = [...peq.entries()]
        .sort((a, b) => compareDrawingEquipmentKeys(a[0], b[0]))
        .map(([key, agg]) => ({ key, label: key, ...finalizeDrawingAgg(agg, idx) }));

      const dailyStatsP = [...pd.entries()]
        .sort((a, b) => {
          if (a[0] === "일자미상") return 1;
          if (b[0] === "일자미상") return -1;
          return a[0].localeCompare(b[0]);
        })
        .map(([key, agg]) => ({ key, label: key, ...finalizeDrawingAgg(agg, idx) }));

      const dayEquipmentStatsP = {};
      for (const [dKey, eqMap] of pde.entries()) {
        dayEquipmentStatsP[dKey] = [...eqMap.entries()]
          .sort((a, b) => compareDrawingEquipmentKeys(a[0], b[0]))
          .map(([key, agg]) => ({ key, label: key, ...finalizeDrawingAgg(agg, idx) }));
      }

      return {
        process: pk,
        monthlyStats: monthlyStatsP,
        equipmentStats: equipmentStatsP,
        dailyStats: dailyStatsP,
        dayEquipmentStats: dayEquipmentStatsP,
      };
    });

    return {
      hasAny: true,
      headerRowIndex: hi,
      iWorkDate,
      iEquipment,
      iSpec,
      iProdQty,
      iWorkTime,
      iInputTime,
      iProductivity,
      iDefectQty,
      iDefectQtyCols: [...iDefectQtyCols],
      iStopTime,
      iStopTimeCols: iStopTimeCols ? [...iStopTimeCols] : [],
      iProcess,
      iUtil,
      iOee,
      iDef: iDefRate,
      ...g,
      dateFrom: workDates.length ? workDates[0] : null,
      dateTo: workDates.length ? workDates[workDates.length - 1] : null,
      monthlyStats,
      equipmentStats,
      dailyStats,
      dayEquipmentStats,
      processBlocks,
      hasEquipmentColumn: iEquipment >= 0,
      _filterMatrix: matrix,
      _filterHi: hi,
      _filterIdx: idx,
    };
  }

  /**
   * 「머플러」「단조머플러」 시트를 각각 파싱해 배열로 반환 (행을 이어 붙이지 않음).
   * @param {any} wb
   * @returns {null | { segments: { sheetName: string, parsed: ReturnType<typeof parseMufflerLogFromMatrix> }[] }}
   */
  function findMufflerLogSheetInWorkbook(wb) {
    const { mufflerSheet, forgedSheet } = findMufflerSheetNamesInWorkbook(wb);
    /** @type {{ sheetName: string, parsed: ReturnType<typeof parseMufflerLogFromMatrix> }[]} */
    const segments = [];
    if (mufflerSheet) {
      const parsed = parseMufflerLogFromMatrix(sheetToMatrix(wb, mufflerSheet));
      if (parsed && parsed.hasAny) segments.push({ sheetName: mufflerSheet, parsed });
    }
    if (forgedSheet) {
      const parsed = parseMufflerLogFromMatrix(sheetToMatrix(wb, forgedSheet));
      if (parsed && parsed.hasAny) segments.push({ sheetName: forgedSheet, parsed });
    }
    if (segments.length === 0) return null;
    return { segments };
  }

  function resetDrawingLogSlicerState() {
    drawingLogGranularity = "month";
    drawingLogSelectedMonths = new Set();
    drawingLogSelectedDays = new Set();
    drawingLogTimelineYear = new Date().getFullYear();
  }

  function drawingLogYearsFromStats(log) {
    if (!log) return [];
    const ys = new Set();
    (log.monthlyStats || []).forEach((m) => {
      if (!m || m.key === "일자미상") return;
      const y = parseInt(String(m.key).slice(0, 4), 10);
      if (Number.isFinite(y)) ys.add(y);
    });
    return [...ys].sort((a, b) => a - b);
  }

  function initDrawingLogSlicerFromData() {
    resetDrawingLogSlicerState();
    if (!lastDrawingLog) return;
    const ys = drawingLogYearsFromStats(lastDrawingLog);
    if (ys.length) drawingLogTimelineYear = ys[ys.length - 1];
    else if (lastDrawingLog.dateTo && String(lastDrawingLog.dateTo).length >= 4) {
      const y = parseInt(String(lastDrawingLog.dateTo).slice(0, 4), 10);
      if (Number.isFinite(y)) drawingLogTimelineYear = y;
    }
  }

  /**
   * 필터로 포함된 행만 월·일 단위로 묶어 「선택 월·생산수량」피벗에 사용
   * @param {Map<string, ReturnType<typeof makeDrawingAgg>>} byMonthPivot
   * @param {Map<string, ReturnType<typeof makeDrawingAgg>>} byDayPivot
   * @param {Record<string, any>} idx
   */
  function buildPivotMonthlyDailyStatsFromAggMaps(byMonthPivot, byDayPivot, idx) {
    const monthLabel = (ymKey) => {
      if (ymKey === "일자미상") return ymKey;
      const p = ymKey.split("-");
      if (p.length !== 2) return ymKey;
      return `${p[0]}년 ${parseInt(p[1], 10)}월`;
    };
    const pivotMonthlyStats = [...byMonthPivot.entries()]
      .sort((a, b) => {
        if (a[0] === "일자미상") return 1;
        if (b[0] === "일자미상") return -1;
        return a[0].localeCompare(b[0]);
      })
      .map(([key, agg]) => ({ key, label: monthLabel(key), ...finalizeDrawingAgg(agg, idx) }));
    const pivotDailyStats = [...byDayPivot.entries()]
      .sort((a, b) => {
        if (a[0] === "일자미상") return 1;
        if (b[0] === "일자미상") return -1;
        return a[0].localeCompare(b[0]);
      })
      .map(([key, agg]) => ({ key, label: key, ...finalizeDrawingAgg(agg, idx) }));
    return { pivotMonthlyStats, pivotDailyStats };
  }

  /**
   * 월·일 다중 선택 시 시트에서 다시 집계
   * @returns {null | { equipmentStats: any[] } & ReturnType<typeof finalizeDrawingAgg>}
   */
  function reaggregateDrawingLogFiltered() {
    if (!lastDrawingLog || !lastDrawingLog._filterMatrix) return null;
    const hasMonth =
      drawingLogGranularity === "month" && drawingLogSelectedMonths.size > 0;
    const hasDay = drawingLogGranularity === "day" && drawingLogSelectedDays.size > 0;
    if (!hasMonth && !hasDay) return null;

    const matrix = lastDrawingLog._filterMatrix;
    const hi = lastDrawingLog._filterHi;
    const idx = lastDrawingLog._filterIdx;
    const iWorkDate = idx.iWorkDate;
    const iEq = lastDrawingLog.iEquipment;

    const rowIncluded = (row) => {
      if (!row || !row.length) return false;
      if (!row.some((c) => String(c ?? "").trim() !== "")) return false;
      const ymd = iWorkDate >= 0 ? parseExcelDate(row[iWorkDate]) : "";
      const ym = ymd && ymd.length >= 7 ? ymd.slice(0, 7) : "일자미상";
      const dayKey = ymd || "일자미상";
      if (hasDay) return drawingLogSelectedDays.has(dayKey);
      if (hasMonth) return drawingLogSelectedMonths.has(ym);
      return false;
    };

    const aggGlobal = makeDrawingAgg();
    /** @type {Map<string, ReturnType<typeof makeDrawingAgg>>} */
    const byEq = new Map();
    /** @type {Map<string, ReturnType<typeof makeDrawingAgg>>} */
    const byMonthPivot = new Map();
    /** @type {Map<string, ReturnType<typeof makeDrawingAgg>>} */
    const byDayPivot = new Map();
    for (let r = hi + 1; r < matrix.length; r++) {
      const row = matrix[r];
      if (!rowIncluded(row)) continue;
      addDrawingRowToAgg(aggGlobal, row, idx, false);
      const ymdPV = iWorkDate >= 0 ? parseExcelDate(row[iWorkDate]) : "";
      const ymPV = ymdPV && ymdPV.length >= 7 ? ymdPV.slice(0, 7) : "일자미상";
      const dayKeyPV = ymdPV || "일자미상";
      if (!byMonthPivot.has(ymPV)) byMonthPivot.set(ymPV, makeDrawingAgg());
      addDrawingRowToAgg(byMonthPivot.get(ymPV), row, idx, false);
      if (!byDayPivot.has(dayKeyPV)) byDayPivot.set(dayKeyPV, makeDrawingAgg());
      addDrawingRowToAgg(byDayPivot.get(dayKeyPV), row, idx, false);
      const eqKey = iEq >= 0 ? String(row[iEq] ?? "").trim() || "미지정" : "전체";
      if (!byEq.has(eqKey)) byEq.set(eqKey, makeDrawingAgg());
      addDrawingRowToAgg(byEq.get(eqKey), row, idx, false);
    }
    const g = finalizeDrawingAgg(aggGlobal, idx);
    const equipmentStats = [...byEq.entries()]
      .sort((a, b) => compareDrawingEquipmentKeys(a[0], b[0]))
      .map(([key, agg]) => ({ key, label: key, ...finalizeDrawingAgg(agg, idx) }));
    const { pivotMonthlyStats, pivotDailyStats } = buildPivotMonthlyDailyStatsFromAggMaps(byMonthPivot, byDayPivot, idx);
    return { ...g, equipmentStats, dataRows: aggGlobal.dataRows, pivotMonthlyStats, pivotDailyStats };
  }

  function getDrawingLogEquipmentStatsForTables() {
    const sub = reaggregateDrawingLogFiltered();
    if (sub && Array.isArray(sub.equipmentStats)) return sub.equipmentStats;
    return lastDrawingLog && lastDrawingLog.equipmentStats ? lastDrawingLog.equipmentStats : [];
  }

  /**
   * @param {any} wb
   * @returns {null | { sheetName: string, parsed: ReturnType<typeof parseDrawingLogFromMatrix> }}
   */
  function findDrawingLogSheetInWorkbook(wb) {
    const sheetName = findDrawingSheetName(wb);
    if (!sheetName) return null;
    const matrix = sheetToMatrix(wb, sheetName);
    const parsed = parseDrawingLogFromMatrix(matrix);
    if (!parsed || !parsed.hasAny) return null;
    return { sheetName, parsed };
  }

  /**
   * @param {any} wb
   * @returns {null | { sheetName: string, parsed: ReturnType<typeof parseParallelLogFromMatrix> }}
   */
  function findParallelLogSheetInWorkbook(wb) {
    const sheetName = findParallelSheetName(wb);
    if (!sheetName) return null;
    const matrix = sheetToMatrix(wb, sheetName);
    const parsed = parseParallelLogFromMatrix(matrix);
    if (!parsed || !parsed.hasAny) return null;
    return { sheetName, parsed };
  }

  /**
   * @param {any} wb
   * @returns {null | { sheetName: string, parsed: ReturnType<typeof parseDoublePipeLogFromMatrix> }}
   */
  function findDoublePipeLogSheetInWorkbook(wb) {
    const sheetName = findDoublePipeSheetName(wb);
    if (!sheetName) return null;
    const matrix = sheetToMatrix(wb, sheetName);
    const parsed = parseDoublePipeLogFromMatrix(matrix);
    if (!parsed || !parsed.hasAny) return null;
    return { sheetName, parsed };
  }

  /**
   * @param {any} wb
   * @returns {null | { sheetName: string, parsed: ReturnType<typeof parseHwaseungDoublePipeLogFromMatrix> }}
   */
  function findHwaseungDoublePipeLogSheetInWorkbook(wb) {
    const sheetName = findHwaseungDoublePipeSheetName(wb);
    if (!sheetName) return null;
    const matrix = sheetToMatrix(wb, sheetName);
    const parsed = parseHwaseungDoublePipeLogFromMatrix(matrix);
    if (!parsed || !parsed.hasAny) return null;
    return { sheetName, parsed };
  }

  function clearDrawingLogUi() {
    if (drawingLogEmpty) drawingLogEmpty.hidden = false;
    if (drawingLogContent) drawingLogContent.hidden = true;
    if (kpiUtil) kpiUtil.textContent = "—";
    if (kpiOee) kpiOee.textContent = "—";
    if (kpiDefect) kpiDefect.textContent = "—";
    if (kpiDefectLabel) kpiDefectLabel.textContent = "불량율";
    if (drawingLogMeta) drawingLogMeta.textContent = "";
    if (drawingLogOps) {
      drawingLogOps.textContent = "";
      drawingLogOps.hidden = true;
    }
    if (drawingLogTablesWrap) {
      drawingLogTablesWrap.innerHTML = "";
      drawingLogTablesWrap.hidden = true;
    }
    if (drawingLogProdWrap) {
      drawingLogProdWrap.innerHTML = "";
      drawingLogProdWrap.hidden = true;
    }
    if (drawingLogMaintWrap) {
      drawingLogMaintWrap.innerHTML = "";
      drawingLogMaintWrap.hidden = true;
    }
    resetDrawingLogSlicerState();
    if (drawingLogTimeline) drawingLogTimeline.hidden = true;
    if (drawingLogMonthStrip) drawingLogMonthStrip.innerHTML = "";
    if (drawingLogDayStrip) {
      drawingLogDayStrip.innerHTML = "";
      drawingLogDayStrip.hidden = true;
    }
    if (drawingLogSlicerSelection) drawingLogSlicerSelection.textContent = "";
    if (drawingLogGranularityEl) drawingLogGranularityEl.value = "month";
  }

  function clearParallelLogUi() {
    if (parallelLogEmpty) parallelLogEmpty.hidden = false;
    if (parallelLogContent) parallelLogContent.hidden = true;
    if (kpiUtilParallel) kpiUtilParallel.textContent = "—";
    if (kpiOeeParallel) kpiOeeParallel.textContent = "—";
    if (kpiDefectParallel) kpiDefectParallel.textContent = "—";
    if (kpiDefectLabelParallel) kpiDefectLabelParallel.textContent = "불량율";
    if (parallelLogMeta) parallelLogMeta.textContent = "";
    if (parallelLogOps) {
      parallelLogOps.textContent = "";
      parallelLogOps.hidden = true;
    }
    if (parallelLogTablesWrap) {
      parallelLogTablesWrap.innerHTML = "";
      parallelLogTablesWrap.hidden = true;
    }
    if (parallelLogProdWrap) {
      parallelLogProdWrap.innerHTML = "";
      parallelLogProdWrap.hidden = true;
    }
    if (parallelLogMaintWrap) {
      parallelLogMaintWrap.innerHTML = "";
      parallelLogMaintWrap.hidden = true;
    }
    parallelLogGranularity = "month";
    parallelLogSelectedMonths = new Set();
    parallelLogSelectedDays = new Set();
    parallelLogTimelineYear = new Date().getFullYear();
    if (parallelLogTimeline) parallelLogTimeline.hidden = true;
    if (parallelLogMonthStrip) parallelLogMonthStrip.innerHTML = "";
    if (parallelLogDayStrip) {
      parallelLogDayStrip.innerHTML = "";
      parallelLogDayStrip.hidden = true;
    }
    if (parallelLogSlicerSelection) parallelLogSlicerSelection.textContent = "";
    if (parallelLogGranularityEl) parallelLogGranularityEl.value = "month";
    parallelLogProcessCut = true;
    parallelLogProcessBend = true;
    if (parallelLogFilterCutEl) parallelLogFilterCutEl.checked = true;
    if (parallelLogFilterBendEl) parallelLogFilterBendEl.checked = true;
    parallelLogKpiScope = "process";
    if (parallelLogKpiScopeEl) parallelLogKpiScopeEl.value = "process";
  }

  function clearDoublePipeLogUi() {
    if (doublePipeLogEmpty) doublePipeLogEmpty.hidden = false;
    if (doublePipeLogContent) doublePipeLogContent.hidden = true;
    if (kpiUtilDoublePipe) kpiUtilDoublePipe.textContent = "—";
    if (kpiOeeDoublePipe) kpiOeeDoublePipe.textContent = "—";
    if (kpiDefectDoublePipe) kpiDefectDoublePipe.textContent = "—";
    if (kpiDefectLabelDoublePipe) kpiDefectLabelDoublePipe.textContent = "불량율";
    if (doublePipeLogMeta) doublePipeLogMeta.textContent = "";
    if (doublePipeLogOps) {
      doublePipeLogOps.textContent = "";
      doublePipeLogOps.hidden = true;
    }
    if (doublePipeLogTablesWrap) {
      doublePipeLogTablesWrap.innerHTML = "";
      doublePipeLogTablesWrap.hidden = true;
    }
    if (doublePipeLogProdWrap) {
      doublePipeLogProdWrap.innerHTML = "";
      doublePipeLogProdWrap.hidden = true;
    }
    if (doublePipeLogMaintWrap) {
      doublePipeLogMaintWrap.innerHTML = "";
      doublePipeLogMaintWrap.hidden = true;
    }
    doublePipeLogGranularity = "month";
    doublePipeLogSelectedMonths = new Set();
    doublePipeLogSelectedDays = new Set();
    doublePipeLogTimelineYear = new Date().getFullYear();
    if (doublePipeLogTimeline) doublePipeLogTimeline.hidden = true;
    if (doublePipeLogMonthStrip) doublePipeLogMonthStrip.innerHTML = "";
    if (doublePipeLogDayStrip) {
      doublePipeLogDayStrip.innerHTML = "";
      doublePipeLogDayStrip.hidden = true;
    }
    if (doublePipeLogSlicerSelection) doublePipeLogSlicerSelection.textContent = "";
    if (doublePipeLogGranularityEl) doublePipeLogGranularityEl.value = "month";
    doublePipeLogProcessMach = true;
    doublePipeLogProcessForm = true;
    if (doublePipeLogFilterMachEl) doublePipeLogFilterMachEl.checked = true;
    if (doublePipeLogFilterFormEl) doublePipeLogFilterFormEl.checked = true;
    doublePipeLogKpiScope = "process";
    if (doublePipeLogKpiScopeEl) doublePipeLogKpiScopeEl.value = "process";
  }

  function clearHwaseungDoublePipeLogUi() {
    if (hwaseungDoublePipeLogEmpty) hwaseungDoublePipeLogEmpty.hidden = false;
    if (hwaseungDoublePipeLogContent) hwaseungDoublePipeLogContent.hidden = true;
    if (kpiUtilHwaseungDoublePipe) kpiUtilHwaseungDoublePipe.textContent = "—";
    if (kpiOeeHwaseungDoublePipe) kpiOeeHwaseungDoublePipe.textContent = "—";
    if (kpiDefectHwaseungDoublePipe) kpiDefectHwaseungDoublePipe.textContent = "—";
    if (kpiDefectLabelHwaseungDoublePipe) kpiDefectLabelHwaseungDoublePipe.textContent = "불량율";
    if (hwaseungDoublePipeLogMeta) hwaseungDoublePipeLogMeta.textContent = "";
    if (hwaseungDoublePipeLogOps) {
      hwaseungDoublePipeLogOps.textContent = "";
      hwaseungDoublePipeLogOps.hidden = true;
    }
    if (hwaseungDoublePipeLogTablesWrap) {
      hwaseungDoublePipeLogTablesWrap.innerHTML = "";
      hwaseungDoublePipeLogTablesWrap.hidden = true;
    }
    if (hwaseungDoublePipeLogProdWrap) {
      hwaseungDoublePipeLogProdWrap.innerHTML = "";
      hwaseungDoublePipeLogProdWrap.hidden = true;
    }
    if (hwaseungDoublePipeLogMaintWrap) {
      hwaseungDoublePipeLogMaintWrap.innerHTML = "";
      hwaseungDoublePipeLogMaintWrap.hidden = true;
    }
    hwaseungDoublePipeLogGranularity = "month";
    hwaseungDoublePipeLogSelectedMonths = new Set();
    hwaseungDoublePipeLogSelectedDays = new Set();
    hwaseungDoublePipeLogTimelineYear = new Date().getFullYear();
    if (hwaseungDoublePipeLogTimeline) hwaseungDoublePipeLogTimeline.hidden = true;
    if (hwaseungDoublePipeLogMonthStrip) hwaseungDoublePipeLogMonthStrip.innerHTML = "";
    if (hwaseungDoublePipeLogDayStrip) {
      hwaseungDoublePipeLogDayStrip.innerHTML = "";
      hwaseungDoublePipeLogDayStrip.hidden = true;
    }
    if (hwaseungDoublePipeLogSlicerSelection) hwaseungDoublePipeLogSlicerSelection.textContent = "";
    if (hwaseungDoublePipeLogGranularityEl) hwaseungDoublePipeLogGranularityEl.value = "month";
    hwaseungDoublePipeLogProcessSelected = new Set(HWASEUNG_DOUBLE_PIPE_ALL_PROC_KEYS);
    if (hwaseungDoublePipeLogProcessFilters) {
      hwaseungDoublePipeLogProcessFilters.querySelectorAll('input[type="checkbox"][data-hwaseung-proc]').forEach((el) => {
        el.checked = true;
      });
    }
    hwaseungDoublePipeLogKpiScope = "process";
    if (hwaseungDoublePipeLogKpiScopeEl) hwaseungDoublePipeLogKpiScopeEl.value = "process";
  }

  function clearMufflerLogUi() {
    if (mufflerLogEmpty) mufflerLogEmpty.hidden = false;
    if (mufflerLogContent) mufflerLogContent.hidden = true;
    if (mufflerLogMeta) mufflerLogMeta.textContent = "";
    if (mufflerLogSegmentsHost) mufflerLogSegmentsHost.innerHTML = "";
    if (mufflerLogMaintWrap) {
      mufflerLogMaintWrap.innerHTML = "";
      mufflerLogMaintWrap.hidden = true;
    }
    mufflerLogSegStates = [];
  }

  function formatDrawingDefectRateCell(st) {
    if (Number.isFinite(st.defectAvg)) return formatPercent(st.defectAvg);
    if (Number.isFinite(st.sumProdQty) && st.sumProdQty > 0 && Number.isFinite(st.sumDefectQty)) {
      return formatPercent((st.sumDefectQty / st.sumProdQty) * 100);
    }
    return "—";
  }

  /**
   * @param {string} title
   * @param {string} firstColTitle
   * @param {any[]} rows
   * @param {boolean} [splitStopByKind] Q~T(교환·수리·소재·계획정지) 네 열 표시
   * @param {string[]} [stopKindLabels] 네 열 헤더(이중관 S~V 등). 없으면 드로잉/페럴 라벨 사용
   */
  function buildDrawingBreakdownTable(title, firstColTitle, rows, splitStopByKind, stopKindLabels) {
    const block = document.createElement("div");
    block.className = "drawing-log-block";
    const h3 = document.createElement("h3");
    h3.className = "drawing-log-block__title";
    h3.textContent = title;
    block.appendChild(h3);
    const wrap = document.createElement("div");
    wrap.className = "drawing-log-preview-wrap";
    const tbl = document.createElement("table");
    const thead = document.createElement("thead");
    const trh = document.createElement("tr");
    const headBase = [
      firstColTitle,
      "건수",
      "가동율",
      "설비효율",
      "불량율·비율",
      "생산수량 합",
      "불량 합",
      "작업시간 합",
    ];
    const headStop = (() => {
      if (!splitStopByKind) return ["정지 합"];
      if (Array.isArray(stopKindLabels) && stopKindLabels.length === 5) return stopKindLabels.slice(0, 5);
      if (Array.isArray(stopKindLabels) && stopKindLabels.length >= 4) return stopKindLabels.slice(0, 4);
      return [...DRAWING_STOP_KIND_LABELS];
    })();
    [...headBase, ...headStop].forEach((text) => {
      const th = document.createElement("th");
      th.textContent = text;
      trh.appendChild(th);
    });
    thead.appendChild(trh);
    tbl.appendChild(thead);
    const tbody = document.createElement("tbody");
    const fmtQtyCell = (v) => (Number.isFinite(v) ? formatQty(v) : "—");
    for (const st of rows) {
      const tr = document.createElement("tr");
      const cells = [
        st.label || st.key,
        String(st.dataRows),
        formatPercent(st.utilAvg ?? NaN),
        formatPercent(st.oeeAvg ?? NaN),
        formatDrawingDefectRateCell(st),
        fmtQtyCell(st.sumProdQty),
        fmtQtyCell(st.sumDefectQty),
        fmtQtyCell(st.sumWorkTime),
      ];
      if (splitStopByKind) {
        const useFive = Array.isArray(stopKindLabels) && stopKindLabels.length === 5;
        if (useFive) {
          cells.push(
            fmtQtyCell(st.sumStopExchange),
            fmtQtyCell(st.sumStopRepair),
            fmtQtyCell(st.sumStopMaterial),
            fmtQtyCell(st.sumStopPlanned),
            fmtQtyCell(st.sumStopFifth)
          );
        } else {
          cells.push(
            fmtQtyCell(st.sumStopExchange),
            fmtQtyCell(st.sumStopRepair),
            fmtQtyCell(st.sumStopMaterial),
            fmtQtyCell(st.sumStopPlanned)
          );
        }
      } else {
        cells.push(fmtQtyCell(st.sumStopTime));
      }
      cells.forEach((text, i) => {
        const td = document.createElement("td");
        td.textContent = text;
        if (i > 0) td.classList.add("num");
        tr.appendChild(td);
      });
      tbody.appendChild(tr);
    }
    tbl.appendChild(tbody);
    wrap.appendChild(tbl);
    block.appendChild(wrap);
    return block;
  }

  function renderDrawingLogTables() {
    if (!drawingLogTablesWrap || !lastDrawingLog) return;
    drawingLogTablesWrap.innerHTML = "";
    const sub = reaggregateDrawingLogFiltered();
    const slicerActive = !!sub;
    const months = slicerActive ? [] : lastDrawingLog.monthlyStats || [];
    const eqs = getDrawingLogEquipmentStatsForTables();
    if (months.length === 0 && eqs.length === 0) {
      drawingLogTablesWrap.hidden = false;
      const p = document.createElement("p");
      p.className = "empty-state empty-state--flat";
      p.textContent = slicerActive
        ? "선택한 기간에 표시할 설비별 집계가 없습니다."
        : "표시할 집계가 없습니다.";
      drawingLogTablesWrap.appendChild(p);
      return;
    }
    const splitStop =
      Array.isArray(lastDrawingLog.iStopTimeCols) && lastDrawingLog.iStopTimeCols.length >= 4;
    if (months.length) {
      drawingLogTablesWrap.appendChild(buildDrawingBreakdownTable("월별 집계", "월", months, splitStop));
    }
    if (eqs.length) {
      let eqTitle = lastDrawingLog.hasEquipmentColumn ? "설비별 집계" : "설비별 집계 (설비 열 없음 · 전체)";
      if (slicerActive) {
        const hint =
          drawingLogGranularity === "month"
            ? [...drawingLogSelectedMonths].sort().join(", ")
            : [...drawingLogSelectedDays].sort().join(", ");
        eqTitle += ` · 선택: ${hint}`;
      }
      drawingLogTablesWrap.appendChild(buildDrawingBreakdownTable(eqTitle, "설비·구분", eqs, splitStop));
    }
    drawingLogTablesWrap.hidden = false;
  }

  /** 날짜 필터 적용 시 선택 월/일별 생산수량 표 */
  function renderDrawingLogProductionQtyTable() {
    if (!drawingLogProdWrap || !lastDrawingLog) return;
    drawingLogProdWrap.innerHTML = "";
    drawingLogProdWrap.hidden = true;
    const subDraw = reaggregateDrawingLogFiltered();
    if (!subDraw) return;

    const monthMode = drawingLogGranularity === "month" && drawingLogSelectedMonths.size > 0;
    const dayMode = drawingLogGranularity === "day" && drawingLogSelectedDays.size > 0;
    if (!monthMode && !dayMode) return;

    const monthPool = subDraw.pivotMonthlyStats || [];
    const dayPool = subDraw.pivotDailyStats || [];
    /** @type {any[]} */
    let statRows = [];
    if (monthMode) {
      for (const ym of [...drawingLogSelectedMonths].sort()) {
        const st = monthPool.find((m) => m.key === ym);
        if (st) statRows.push(st);
      }
    } else {
      for (const dk of [...drawingLogSelectedDays].sort()) {
        const st = dayPool.find((d) => d.key === dk);
        if (st) statRows.push(st);
      }
    }
    if (!statRows.length) return;

    const hasProdCol = lastDrawingLog.iProdQty >= 0;
    drawingLogProdWrap.hidden = false;
    const block = document.createElement("div");
    block.className = "drawing-log-block";
    const h3 = document.createElement("h3");
    h3.className = "drawing-log-block__title";
    h3.textContent = monthMode ? "선택 월 · 생산수량" : "선택 일 · 생산수량";
    block.appendChild(h3);
    if (!hasProdCol) {
      const note = document.createElement("p");
      note.className = "drawing-log-meta";
      note.textContent =
        "시트에 생산수량·생산량 열이 인식되지 않으면 생산수량 행은 표시되지 않습니다. 엑셀 헤더를 확인해 주세요.";
      block.appendChild(note);
    }
    const wrap = document.createElement("div");
    wrap.className = "drawing-log-preview-wrap drawing-log-prod-pivot-wrap";
    const tbl = document.createElement("table");
    tbl.className = "drawing-log-prod-pivot";

    const thead = document.createElement("thead");
    const trh = document.createElement("tr");
    const thCorner = document.createElement("th");
    thCorner.textContent = "";
    thCorner.classList.add("drawing-log-prod-pivot__corner");
    thCorner.setAttribute("aria-label", monthMode ? "월 열" : "날짜 열");
    trh.appendChild(thCorner);
    statRows.forEach((st) => {
      const th = document.createElement("th");
      th.textContent = String(st.label || st.key);
      th.classList.add("num");
      th.title = String(st.key || "");
      trh.appendChild(th);
    });
    const thTot = document.createElement("th");
    thTot.textContent = "합계";
    thTot.classList.add("num");
    trh.appendChild(thTot);
    thead.appendChild(trh);
    tbl.appendChild(thead);

    const tbody = document.createElement("tbody");
    /**
     * @param {string} label
     * @param {(st: any) => number} getter
     * @param {{ asInt?: boolean }} [opts]
     */
    function addPivotRow(label, getter, opts) {
      const asInt = opts && opts.asInt;
      const tr = document.createElement("tr");
      const td0 = document.createElement("td");
      td0.textContent = label;
      td0.classList.add("drawing-log-prod-pivot__rowhead");
      tr.appendChild(td0);
      let sum = 0;
      let n = 0;
      for (const st of statRows) {
        const v = getter(st);
        const td = document.createElement("td");
        td.classList.add("num");
        if (Number.isFinite(v)) {
          td.textContent = asInt ? String(Math.round(v)) : formatQty(v);
          sum += v;
          n++;
        } else {
          td.textContent = "—";
        }
        tr.appendChild(td);
      }
      const tdSum = document.createElement("td");
      tdSum.classList.add("num");
      tdSum.style.fontWeight = "700";
      if (n && Number.isFinite(sum)) {
        tdSum.textContent = asInt ? String(Math.round(sum)) : formatQty(sum);
      } else {
        tdSum.textContent = "—";
      }
      tr.appendChild(tdSum);
      tbody.appendChild(tr);
    }

    addPivotRow("생산수량 합", (st) => (hasProdCol && Number.isFinite(st.sumProdQty) ? st.sumProdQty : NaN));
    addPivotRow("건수", (st) => (Number.isFinite(st.dataRows) ? st.dataRows : NaN), { asInt: true });
    addPivotRow("작업시간 합", (st) => (Number.isFinite(st.sumWorkTime) ? st.sumWorkTime : NaN));

    tbl.appendChild(tbody);
    wrap.appendChild(tbl);
    block.appendChild(wrap);
    drawingLogProdWrap.appendChild(block);
  }

  function drawingLogSlicerSummaryText() {
    const sub = reaggregateDrawingLogFiltered();
    if (drawingLogGranularity === "month" && drawingLogSelectedMonths.size && sub) {
      const parts = [...drawingLogSelectedMonths].sort().map((ym) => {
        const p = ym.split("-");
        return p.length === 2 ? `${p[0]}년 ${parseInt(p[1], 10)}월` : ym;
      });
      return `${parts.join(", ")} · ${sub.dataRows}행`;
    }
    if (drawingLogGranularity === "day" && drawingLogSelectedDays.size && sub) {
      return `${[...drawingLogSelectedDays].sort().join(", ")} · ${sub.dataRows}행`;
    }
    return "";
  }

  function pruneDrawingLogSelectedDaysInvalid() {
    if (!lastDrawingLog || drawingLogGranularity !== "day") return;
    const pool = new Set(
      (lastDrawingLog.dailyStats || [])
        .filter((d) => {
          if (d.key === "일자미상") return false;
          if (!d.key.startsWith(`${drawingLogTimelineYear}-`)) return false;
          if (drawingLogSelectedMonths.size) {
            const ym = d.key.slice(0, 7);
            return drawingLogSelectedMonths.has(ym);
          }
          return true;
        })
        .map((d) => d.key)
    );
    [...drawingLogSelectedDays].forEach((k) => {
      if (!pool.has(k)) drawingLogSelectedDays.delete(k);
    });
  }

  function renderDrawingLogTimeline() {
    if (!drawingLogTimeline || !lastDrawingLog || !lastDrawingLog.monthlyStats) {
      if (drawingLogTimeline) drawingLogTimeline.hidden = true;
      return;
    }
    drawingLogTimeline.hidden = false;
    if (drawingLogGranularityEl) drawingLogGranularityEl.value = drawingLogGranularity;

    const ys = drawingLogYearsFromStats(lastDrawingLog);
    if (ys.length) {
      if (drawingLogTimelineYear < ys[0]) drawingLogTimelineYear = ys[0];
      if (drawingLogTimelineYear > ys[ys.length - 1]) drawingLogTimelineYear = ys[ys.length - 1];
    }
    if (drawingLogYearLabel) drawingLogYearLabel.textContent = `${drawingLogTimelineYear}년`;
    if (drawingLogYearPrev) drawingLogYearPrev.disabled = ys.length && drawingLogTimelineYear <= ys[0];
    if (drawingLogYearNext) drawingLogYearNext.disabled = ys.length && drawingLogTimelineYear >= ys[ys.length - 1];

    const monthKeys = new Set((lastDrawingLog.monthlyStats || []).map((m) => m.key));
    if (drawingLogMonthStrip) {
      drawingLogMonthStrip.innerHTML = "";
      for (let m = 1; m <= 12; m++) {
        const ym = `${drawingLogTimelineYear}-${String(m).padStart(2, "0")}`;
        const btn = document.createElement("button");
        btn.type = "button";
        btn.className = "drawing-log-timeline__chip";
        btn.dataset.ym = ym;
        btn.textContent = String(m);
        const has = monthKeys.has(ym);
        if (!has) {
          btn.classList.add("is-empty");
          btn.disabled = true;
        }
        if (drawingLogSelectedMonths.has(ym)) btn.classList.add("is-selected");
        drawingLogMonthStrip.appendChild(btn);
      }
    }

    const dayMode = drawingLogGranularity === "day";
    if (drawingLogDayStrip) {
      drawingLogDayStrip.hidden = !dayMode;
      drawingLogDayStrip.innerHTML = "";
      if (dayMode) {
        const pool = (lastDrawingLog.dailyStats || []).filter((d) => {
          if (d.key === "일자미상") return false;
          if (!d.key.startsWith(`${drawingLogTimelineYear}-`)) return false;
          if (drawingLogSelectedMonths.size) {
            const ym = d.key.length >= 7 ? d.key.slice(0, 7) : "";
            return drawingLogSelectedMonths.has(ym);
          }
          return true;
        });
        pool.forEach((d) => {
          const btn = document.createElement("button");
          btn.type = "button";
          btn.className = "drawing-log-timeline__chip";
          btn.dataset.dayKey = d.key;
          btn.textContent = d.key.slice(5);
          if (drawingLogSelectedDays.has(d.key)) btn.classList.add("is-selected");
          drawingLogDayStrip.appendChild(btn);
        });
      }
    }

    if (drawingLogSlicerSelection) {
      const t = drawingLogSlicerSummaryText();
      drawingLogSlicerSelection.textContent = t || "전체 기간 (필터 없음)";
    }
  }

  /** @returns {any} KPI·요약에 쓸 객체 */
  function getDrawingLogKpiSource() {
    if (!lastDrawingLog) return null;
    const sub = reaggregateDrawingLogFiltered();
    if (!sub) return lastDrawingLog;
    return { ...lastDrawingLog, ...sub };
  }

  function renderDrawingLogPanel() {
    if (!lastDrawingLog || !drawingLogContent) {
      clearDrawingLogUi();
      return;
    }
    const ksrc = getDrawingLogKpiSource();
    if (drawingLogEmpty) drawingLogEmpty.hidden = true;
    drawingLogContent.hidden = false;
    if (kpiUtil) kpiUtil.textContent = formatPercent(ksrc.utilAvg ?? NaN);
    if (kpiOee) kpiOee.textContent = formatPercent(ksrc.oeeAvg ?? NaN);
    if (kpiDefectLabel && kpiDefect) {
      if (Number.isFinite(ksrc.defectAvg)) {
        kpiDefectLabel.textContent = "불량율";
        kpiDefect.textContent = formatPercent(ksrc.defectAvg);
      } else if (Number.isFinite(ksrc.sumDefectQty)) {
        kpiDefectLabel.textContent = "불량(합)";
        kpiDefect.textContent = `${formatQty(ksrc.sumDefectQty)}건`;
      } else {
        kpiDefectLabel.textContent = "불량율";
        kpiDefect.textContent = "—";
      }
    }
    if (drawingLogMeta) {
      const dr =
        lastDrawingLog.dateFrom && lastDrawingLog.dateTo
          ? ` · 작업일자 ${lastDrawingLog.dateFrom} ~ ${lastDrawingLog.dateTo}`
          : "";
      const kpiNote =
        ksrc.utilIsDerived || ksrc.oeeIsDerived
          ? " · 가동율·설비효율: 열이 없을 때 작업시간·정지·투입·생산수량·불량(율)·생산성으로 행별 산출 후 평균"
          : "";
      const fn = reaggregateDrawingLogFiltered();
      const dayNote = fn
        ? ` · 보기: 필터 적용 (${fn.dataRows}행 / 전체 ${lastDrawingLog.dataRows}행)`
        : "";
      drawingLogMeta.textContent = `파일: ${lastDrawingLog.fileLabel} · 시트 「${lastDrawingLog.sheetName}」만 사용 · 유효 ${lastDrawingLog.dataRows}행(전체 행 집계)${dr}${dayNote}${kpiNote}`;
    }
    if (drawingLogOps) {
      const bits = [];
      if (Number.isFinite(ksrc.sumProdQty)) bits.push(`생산수량 합 ${formatQty(ksrc.sumProdQty)}`);
      if (Number.isFinite(ksrc.sumWorkTime)) bits.push(`작업시간 합 ${formatQty(ksrc.sumWorkTime)}`);
      if (Number.isFinite(ksrc.sumInputTime)) bits.push(`투입시간 합 ${formatQty(ksrc.sumInputTime)}`);
      if (Number.isFinite(ksrc.productivityAvg)) bits.push(`생산성 평균 ${formatPercent(ksrc.productivityAvg)}`);
      if (Number.isFinite(ksrc.sumDefectQty)) bits.push(`불량 합 ${formatQty(ksrc.sumDefectQty)}건`);
      const qrst =
        Array.isArray(lastDrawingLog.iStopTimeCols) && lastDrawingLog.iStopTimeCols.length >= 4;
      if (qrst) {
        if (Number.isFinite(ksrc.sumStopExchange)) bits.push(`교환 ${formatQty(ksrc.sumStopExchange)}`);
        if (Number.isFinite(ksrc.sumStopRepair)) bits.push(`수리 ${formatQty(ksrc.sumStopRepair)}`);
        if (Number.isFinite(ksrc.sumStopMaterial)) bits.push(`소재 ${formatQty(ksrc.sumStopMaterial)}`);
        if (Number.isFinite(ksrc.sumStopPlanned)) bits.push(`계획정지 ${formatQty(ksrc.sumStopPlanned)}`);
        if (Number.isFinite(ksrc.sumStopTime)) bits.push(`정지 합 ${formatQty(ksrc.sumStopTime)}`);
      } else if (Number.isFinite(ksrc.sumStopTime)) {
        bits.push(`정지 시간 합 ${formatQty(ksrc.sumStopTime)}`);
      }
      drawingLogOps.textContent = bits.join(" · ");
      drawingLogOps.hidden = bits.length === 0;
    }
    renderDrawingLogTimeline();
    renderDrawingLogTables();
    renderDrawingLogProductionQtyTable();
    if (drawingLogMaintWrap) {
      if (lastDrawingLog.maintenance) {
        drawingLogMaintWrap.hidden = false;
        renderWorkbookMaintenanceInto(drawingLogMaintWrap, lastDrawingLog.maintenance);
      } else {
        drawingLogMaintWrap.innerHTML = "";
        drawingLogMaintWrap.hidden = true;
      }
    }
  }

  function resetParallelLogSlicerState() {
    parallelLogGranularity = "month";
    parallelLogSelectedMonths = new Set();
    parallelLogSelectedDays = new Set();
    parallelLogTimelineYear = new Date().getFullYear();
  }

  function parallelLogYearsFromStats(log) {
    if (!log) return [];
    const ys = new Set();
    (log.monthlyStats || []).forEach((m) => {
      if (!m || m.key === "일자미상") return;
      const y = parseInt(String(m.key).slice(0, 4), 10);
      if (Number.isFinite(y)) ys.add(y);
    });
    return [...ys].sort((a, b) => a - b);
  }

  function initParallelLogSlicerFromData() {
    resetParallelLogSlicerState();
    parallelLogProcessCut = true;
    parallelLogProcessBend = true;
    if (parallelLogFilterCutEl) parallelLogFilterCutEl.checked = true;
    if (parallelLogFilterBendEl) parallelLogFilterBendEl.checked = true;
    if (!lastParallelLog) return;
    const ys = parallelLogYearsFromStats(lastParallelLog);
    if (ys.length) parallelLogTimelineYear = ys[ys.length - 1];
    else if (lastParallelLog.dateTo && String(lastParallelLog.dateTo).length >= 4) {
      const y = parseInt(String(lastParallelLog.dateTo).slice(0, 4), 10);
      if (Number.isFinite(y)) parallelLogTimelineYear = y;
    }
  }

  /** 절단·벤딩 중 하나만 선택해 행을 좁힐 때 true (둘 다 해제면 전체 보기라 false) */
  function parallelLogProcessFilterRestrictsRows() {
    if (!parallelLogProcessCut && !parallelLogProcessBend) return false;
    if (parallelLogProcessCut && parallelLogProcessBend) return false;
    return true;
  }

  /**
   * @param {any[]} row
   * @param {number} iProcess
   */
  function parallelLogRowMatchesProcessFilter(row, iProcess) {
    if (!parallelLogProcessCut && !parallelLogProcessBend) return true;
    if (parallelLogProcessCut && parallelLogProcessBend) return true;
    const pk = normalizeParallelProcess(iProcess >= 0 ? row[iProcess] : "");
    if (pk === "절단") return parallelLogProcessCut;
    if (pk === "벤딩") return parallelLogProcessBend;
    return true;
  }

  function parallelLogProcessFilterSummaryFragment() {
    if (!parallelLogProcessFilterRestrictsRows()) return "";
    const bits = [];
    if (parallelLogProcessCut) bits.push("절단");
    if (parallelLogProcessBend) bits.push("벤딩");
    return bits.length ? `공정: ${bits.join("+")}` : "";
  }

  function reaggregateParallelLogFiltered(options = {}) {
    const ignoreProcessForRowFilter = options.ignoreProcessForRowFilter === true;
    if (!lastParallelLog || !lastParallelLog._filterMatrix) return null;
    const hasMonth =
      parallelLogGranularity === "month" && parallelLogSelectedMonths.size > 0;
    const hasDay = parallelLogGranularity === "day" && parallelLogSelectedDays.size > 0;
    const procRestricts = parallelLogProcessFilterRestrictsRows();
    const hasProcF = procRestricts && !ignoreProcessForRowFilter;
    if (!hasMonth && !hasDay && !hasProcF) return null;

    const matrix = lastParallelLog._filterMatrix;
    const hi = lastParallelLog._filterHi;
    const idx = lastParallelLog._filterIdx;
    const iWorkDate = idx.iWorkDate;
    const iEq = lastParallelLog.iEquipment;
    const iProcess = lastParallelLog.iProcess;

    const rowIncluded = (row) => {
      if (!row || !row.length) return false;
      if (!row.some((c) => String(c ?? "").trim() !== "")) return false;
      if (!ignoreProcessForRowFilter && !parallelLogRowMatchesProcessFilter(row, iProcess)) return false;
      const ymd = iWorkDate >= 0 ? parseExcelDate(row[iWorkDate]) : "";
      const ym = ymd && ymd.length >= 7 ? ymd.slice(0, 7) : "일자미상";
      const dayKey = ymd || "일자미상";
      if (hasDay) return parallelLogSelectedDays.has(dayKey);
      if (hasMonth) return parallelLogSelectedMonths.has(ym);
      if (hasProcF) return true;
      return false;
    };

    const aggGlobal = makeDrawingAgg();
    /** @type {Map<string, ReturnType<typeof makeDrawingAgg>>} */
    const byEq = new Map();
    /** @type {Map<string, Map<string, ReturnType<typeof makeDrawingAgg>>>} */
    const byProcMonth = new Map();
    /** @type {Map<string, Map<string, ReturnType<typeof makeDrawingAgg>>>} */
    const byProcEq = new Map();
    /** @type {Map<string, ReturnType<typeof makeDrawingAgg>>} */
    const byMonthPivot = new Map();
    /** @type {Map<string, ReturnType<typeof makeDrawingAgg>>} */
    const byDayPivot = new Map();

    for (let r = hi + 1; r < matrix.length; r++) {
      const row = matrix[r];
      if (!rowIncluded(row)) continue;
      addDrawingRowToAgg(aggGlobal, row, idx, false);
      const ymdPV = iWorkDate >= 0 ? parseExcelDate(row[iWorkDate]) : "";
      const ymPV = ymdPV && ymdPV.length >= 7 ? ymdPV.slice(0, 7) : "일자미상";
      const dayKeyPV = ymdPV || "일자미상";
      if (!byMonthPivot.has(ymPV)) byMonthPivot.set(ymPV, makeDrawingAgg());
      addDrawingRowToAgg(byMonthPivot.get(ymPV), row, idx, false);
      if (!byDayPivot.has(dayKeyPV)) byDayPivot.set(dayKeyPV, makeDrawingAgg());
      addDrawingRowToAgg(byDayPivot.get(dayKeyPV), row, idx, false);
      const eqKey = iEq >= 0 ? String(row[iEq] ?? "").trim() || "미지정" : "전체";
      if (!byEq.has(eqKey)) byEq.set(eqKey, makeDrawingAgg());
      addDrawingRowToAgg(byEq.get(eqKey), row, idx, false);
      const ymd = iWorkDate >= 0 ? parseExcelDate(row[iWorkDate]) : "";
      const ym = ymd && ymd.length >= 7 ? ymd.slice(0, 7) : "일자미상";
      const procKey = normalizeParallelProcess(iProcess >= 0 ? row[iProcess] : "");
      addDrawingRowToAgg(ensureAggMapNested2(byProcMonth, procKey, ym), row, idx, false);
      addDrawingRowToAgg(ensureAggMapNested2(byProcEq, procKey, eqKey), row, idx, false);
    }
    const g = finalizeDrawingAgg(aggGlobal, idx);
    const equipmentStats = [...byEq.entries()]
      .sort((a, b) => compareDrawingEquipmentKeys(a[0], b[0]))
      .map(([key, agg]) => ({ key, label: key, ...finalizeDrawingAgg(agg, idx) }));

    const { pivotMonthlyStats, pivotDailyStats } = buildPivotMonthlyDailyStatsFromAggMaps(byMonthPivot, byDayPivot, idx);

    const monthLabel = (ymKey) => {
      if (ymKey === "일자미상") return ymKey;
      const p = ymKey.split("-");
      if (p.length !== 2) return ymKey;
      return `${p[0]}년 ${parseInt(p[1], 10)}월`;
    };
    const sortProcKeys = (a, b) => {
      const rank = (k) => (k === "절단" ? 0 : k === "벤딩" ? 1 : k === "기타" ? 2 : 99);
      const ra = rank(a);
      const rb = rank(b);
      if (ra !== rb) return ra - rb;
      return String(a).localeCompare(String(b));
    };
    const procKeySet = new Set([...byProcMonth.keys(), ...byProcEq.keys()]);
    const processBlocks = [...procKeySet].sort(sortProcKeys).map((pk) => {
      const pm = byProcMonth.get(pk) || new Map();
      const peq = byProcEq.get(pk) || new Map();
      const monthlyStatsP = [...pm.entries()]
        .sort((a, b) => {
          if (a[0] === "일자미상") return 1;
          if (b[0] === "일자미상") return -1;
          return a[0].localeCompare(b[0]);
        })
        .map(([key, agg]) => ({ key, label: monthLabel(key), ...finalizeDrawingAgg(agg, idx) }));
      const equipmentStatsP = [...peq.entries()]
        .sort((a, b) => compareDrawingEquipmentKeys(a[0], b[0]))
        .map(([key, agg]) => ({ key, label: key, ...finalizeDrawingAgg(agg, idx) }));
      return {
        process: pk,
        monthlyStats: monthlyStatsP,
        equipmentStats: equipmentStatsP,
        dailyStats: [],
        dayEquipmentStats: {},
      };
    });

    return { ...g, equipmentStats, dataRows: aggGlobal.dataRows, processBlocks, pivotMonthlyStats, pivotDailyStats };
  }

  function getParallelLogEquipmentStatsForTables() {
    const sub = reaggregateParallelLogFiltered();
    if (sub && Array.isArray(sub.equipmentStats)) return sub.equipmentStats;
    return lastParallelLog && lastParallelLog.equipmentStats ? lastParallelLog.equipmentStats : [];
  }

  function renderParallelLogTables() {
    if (!parallelLogTablesWrap || !lastParallelLog) return;
    parallelLogTablesWrap.innerHTML = "";
    const sub = reaggregateParallelLogFiltered();
    const slicerActive = !!sub;
    const months = slicerActive ? [] : lastParallelLog.monthlyStats || [];
    const eqs = getParallelLogEquipmentStatsForTables();
    const procBlocks =
      sub && Array.isArray(sub.processBlocks) ? sub.processBlocks : lastParallelLog.processBlocks || [];
    const splitStop =
      Array.isArray(lastParallelLog.iStopTimeCols) && lastParallelLog.iStopTimeCols.length >= 4;

    const procHasRows = procBlocks.some(
      (pb) => (pb.monthlyStats && pb.monthlyStats.length) || (pb.equipmentStats && pb.equipmentStats.length)
    );
    if (months.length === 0 && eqs.length === 0 && !procHasRows) {
      parallelLogTablesWrap.hidden = false;
      const p = document.createElement("p");
      p.className = "empty-state empty-state--flat";
      p.textContent = slicerActive
        ? "선택한 기간에 표시할 집계가 없습니다."
        : "표시할 집계가 없습니다.";
      parallelLogTablesWrap.appendChild(p);
      return;
    }
    if (months.length) {
      parallelLogTablesWrap.appendChild(buildDrawingBreakdownTable("월별 집계", "월", months, splitStop));
    }
    if (eqs.length) {
      let eqTitle = lastParallelLog.hasEquipmentColumn ? "설비별 집계" : "설비별 집계 (설비 열 없음 · 전체)";
      if (slicerActive) {
        const hint =
          parallelLogGranularity === "month"
            ? [...parallelLogSelectedMonths].sort().join(", ")
            : [...parallelLogSelectedDays].sort().join(", ");
        eqTitle += ` · 선택: ${hint}`;
      }
      parallelLogTablesWrap.appendChild(buildDrawingBreakdownTable(eqTitle, "설비·구분", eqs, splitStop));
    }
    for (const blk of procBlocks) {
      const m = blk.monthlyStats || [];
      const e = blk.equipmentStats || [];
      if (!m.length && !e.length) continue;
      const sec = document.createElement("div");
      sec.className = "drawing-log-block";
      const h3 = document.createElement("h3");
      h3.className = "drawing-log-block__title";
      h3.textContent = `공정 · ${blk.process}`;
      sec.appendChild(h3);
      if (m.length) sec.appendChild(buildDrawingBreakdownTable("월별 집계", "월", m, splitStop));
      if (e.length) {
        let eqTitle = lastParallelLog.hasEquipmentColumn ? "설비별 집계" : "설비별 집계 (설비 열 없음 · 전체)";
        if (slicerActive) {
          const hint =
            parallelLogGranularity === "month"
              ? [...parallelLogSelectedMonths].sort().join(", ")
              : [...parallelLogSelectedDays].sort().join(", ");
          eqTitle += ` · 선택: ${hint}`;
        }
        sec.appendChild(buildDrawingBreakdownTable(eqTitle, "설비·구분", e, splitStop));
      }
      parallelLogTablesWrap.appendChild(sec);
    }
    parallelLogTablesWrap.hidden = false;
  }

  function renderParallelLogProductionQtyTable() {
    if (!parallelLogProdWrap || !lastParallelLog) return;
    parallelLogProdWrap.innerHTML = "";
    parallelLogProdWrap.hidden = true;
    const subPar = reaggregateParallelLogFiltered();
    if (!subPar) return;

    const monthMode = parallelLogGranularity === "month" && parallelLogSelectedMonths.size > 0;
    const dayMode = parallelLogGranularity === "day" && parallelLogSelectedDays.size > 0;
    if (!monthMode && !dayMode) return;

    const monthPool = subPar.pivotMonthlyStats || [];
    const dayPool = subPar.pivotDailyStats || [];
    /** @type {any[]} */
    let statRows = [];
    if (monthMode) {
      for (const ym of [...parallelLogSelectedMonths].sort()) {
        const st = monthPool.find((m) => m.key === ym);
        if (st) statRows.push(st);
      }
    } else {
      for (const dk of [...parallelLogSelectedDays].sort()) {
        const st = dayPool.find((d) => d.key === dk);
        if (st) statRows.push(st);
      }
    }
    if (!statRows.length) return;

    const hasProdCol = lastParallelLog.iProdQty >= 0;
    parallelLogProdWrap.hidden = false;
    const block = document.createElement("div");
    block.className = "drawing-log-block";
    const h3 = document.createElement("h3");
    h3.className = "drawing-log-block__title";
    h3.textContent = monthMode ? "선택 월 · 생산수량" : "선택 일 · 생산수량";
    block.appendChild(h3);
    if (!hasProdCol) {
      const note = document.createElement("p");
      note.className = "drawing-log-meta";
      note.textContent =
        "시트에 생산량(ERP)·생산수량 열이 인식되지 않으면 생산수량 행은 표시되지 않습니다. 엑셀 헤더를 확인해 주세요.";
      block.appendChild(note);
    }
    const wrap = document.createElement("div");
    wrap.className = "drawing-log-preview-wrap drawing-log-prod-pivot-wrap";
    const tbl = document.createElement("table");
    tbl.className = "drawing-log-prod-pivot";

    const thead = document.createElement("thead");
    const trh = document.createElement("tr");
    const thCorner = document.createElement("th");
    thCorner.textContent = "";
    thCorner.classList.add("drawing-log-prod-pivot__corner");
    thCorner.setAttribute("aria-label", monthMode ? "월 열" : "날짜 열");
    trh.appendChild(thCorner);
    statRows.forEach((st) => {
      const th = document.createElement("th");
      th.textContent = String(st.label || st.key);
      th.classList.add("num");
      th.title = String(st.key || "");
      trh.appendChild(th);
    });
    const thTot = document.createElement("th");
    thTot.textContent = "합계";
    thTot.classList.add("num");
    trh.appendChild(thTot);
    thead.appendChild(trh);
    tbl.appendChild(thead);

    const tbody = document.createElement("tbody");
    /**
     * @param {string} label
     * @param {(st: any) => number} getter
     * @param {{ asInt?: boolean }} [opts]
     */
    function addPivotRow(label, getter, opts) {
      const asInt = opts && opts.asInt;
      const tr = document.createElement("tr");
      const td0 = document.createElement("td");
      td0.textContent = label;
      td0.classList.add("drawing-log-prod-pivot__rowhead");
      tr.appendChild(td0);
      let sum = 0;
      let n = 0;
      for (const st of statRows) {
        const v = getter(st);
        const td = document.createElement("td");
        td.classList.add("num");
        if (Number.isFinite(v)) {
          td.textContent = asInt ? String(Math.round(v)) : formatQty(v);
          sum += v;
          n++;
        } else {
          td.textContent = "—";
        }
        tr.appendChild(td);
      }
      const tdSum = document.createElement("td");
      tdSum.classList.add("num");
      tdSum.style.fontWeight = "700";
      if (n && Number.isFinite(sum)) {
        tdSum.textContent = asInt ? String(Math.round(sum)) : formatQty(sum);
      } else {
        tdSum.textContent = "—";
      }
      tr.appendChild(tdSum);
      tbody.appendChild(tr);
    }

    addPivotRow("생산수량 합", (st) => (hasProdCol && Number.isFinite(st.sumProdQty) ? st.sumProdQty : NaN));
    addPivotRow("건수", (st) => (Number.isFinite(st.dataRows) ? st.dataRows : NaN), { asInt: true });
    addPivotRow("작업시간 합", (st) => (Number.isFinite(st.sumWorkTime) ? st.sumWorkTime : NaN));

    tbl.appendChild(tbody);
    wrap.appendChild(tbl);
    block.appendChild(wrap);
    parallelLogProdWrap.appendChild(block);
  }

  function parallelLogSlicerSummaryText() {
    const sub = reaggregateParallelLogFiltered();
    const procFrag = parallelLogProcessFilterSummaryFragment();
    if (parallelLogGranularity === "month" && parallelLogSelectedMonths.size && sub) {
      const parts = [...parallelLogSelectedMonths].sort().map((ym) => {
        const p = ym.split("-");
        return p.length === 2 ? `${p[0]}년 ${parseInt(p[1], 10)}월` : ym;
      });
      let t = `${parts.join(", ")} · ${sub.dataRows}행`;
      if (procFrag) t += ` · ${procFrag}`;
      return t;
    }
    if (parallelLogGranularity === "day" && parallelLogSelectedDays.size && sub) {
      let t = `${[...parallelLogSelectedDays].sort().join(", ")} · ${sub.dataRows}행`;
      if (procFrag) t += ` · ${procFrag}`;
      return t;
    }
    if (procFrag && sub) return `${procFrag} · ${sub.dataRows}행`;
    return "";
  }

  function pruneParallelLogSelectedDaysInvalid() {
    if (!lastParallelLog || parallelLogGranularity !== "day") return;
    const pool = new Set(
      (lastParallelLog.dailyStats || [])
        .filter((d) => {
          if (d.key === "일자미상") return false;
          if (!d.key.startsWith(`${parallelLogTimelineYear}-`)) return false;
          if (parallelLogSelectedMonths.size) {
            const ym = d.key.slice(0, 7);
            return parallelLogSelectedMonths.has(ym);
          }
          return true;
        })
        .map((d) => d.key)
    );
    [...parallelLogSelectedDays].forEach((k) => {
      if (!pool.has(k)) parallelLogSelectedDays.delete(k);
    });
  }

  function renderParallelLogTimeline() {
    if (!parallelLogTimeline || !lastParallelLog || !lastParallelLog.monthlyStats) {
      if (parallelLogTimeline) parallelLogTimeline.hidden = true;
      return;
    }
    parallelLogTimeline.hidden = false;
    if (parallelLogGranularityEl) parallelLogGranularityEl.value = parallelLogGranularity;
    if (parallelLogKpiScopeEl)
      parallelLogKpiScopeEl.value = parallelLogKpiScope === "overall" ? "overall" : "process";

    const ys = parallelLogYearsFromStats(lastParallelLog);
    if (ys.length) {
      if (parallelLogTimelineYear < ys[0]) parallelLogTimelineYear = ys[0];
      if (parallelLogTimelineYear > ys[ys.length - 1]) parallelLogTimelineYear = ys[ys.length - 1];
    }
    if (parallelLogYearLabel) parallelLogYearLabel.textContent = `${parallelLogTimelineYear}년`;
    if (parallelLogYearPrev) parallelLogYearPrev.disabled = ys.length && parallelLogTimelineYear <= ys[0];
    if (parallelLogYearNext) parallelLogYearNext.disabled = ys.length && parallelLogTimelineYear >= ys[ys.length - 1];

    const monthKeys = new Set((lastParallelLog.monthlyStats || []).map((m) => m.key));
    if (parallelLogMonthStrip) {
      parallelLogMonthStrip.innerHTML = "";
      for (let m = 1; m <= 12; m++) {
        const ym = `${parallelLogTimelineYear}-${String(m).padStart(2, "0")}`;
        const btn = document.createElement("button");
        btn.type = "button";
        btn.className = "drawing-log-timeline__chip";
        btn.dataset.ym = ym;
        btn.textContent = String(m);
        const has = monthKeys.has(ym);
        if (!has) {
          btn.classList.add("is-empty");
          btn.disabled = true;
        }
        if (parallelLogSelectedMonths.has(ym)) btn.classList.add("is-selected");
        parallelLogMonthStrip.appendChild(btn);
      }
    }

    const dayMode = parallelLogGranularity === "day";
    if (parallelLogDayStrip) {
      parallelLogDayStrip.hidden = !dayMode;
      parallelLogDayStrip.innerHTML = "";
      if (dayMode) {
        const pool = (lastParallelLog.dailyStats || []).filter((d) => {
          if (d.key === "일자미상") return false;
          if (!d.key.startsWith(`${parallelLogTimelineYear}-`)) return false;
          if (parallelLogSelectedMonths.size) {
            const ym = d.key.length >= 7 ? d.key.slice(0, 7) : "";
            return parallelLogSelectedMonths.has(ym);
          }
          return true;
        });
        pool.forEach((d) => {
          const btn = document.createElement("button");
          btn.type = "button";
          btn.className = "drawing-log-timeline__chip";
          btn.dataset.dayKey = d.key;
          btn.textContent = d.key.slice(5);
          if (parallelLogSelectedDays.has(d.key)) btn.classList.add("is-selected");
          parallelLogDayStrip.appendChild(btn);
        });
      }
    }

    if (parallelLogSlicerSelection) {
      const t = parallelLogSlicerSummaryText();
      parallelLogSlicerSelection.textContent = t || "전체 기간 (필터 없음)";
    }
  }

  function getParallelLogKpiSource() {
    if (!lastParallelLog) return null;
    const sub = reaggregateParallelLogFiltered({
      ignoreProcessForRowFilter: parallelLogKpiScope === "overall",
    });
    if (!sub) return lastParallelLog;
    return { ...lastParallelLog, ...sub };
  }

  function renderParallelLogPanel() {
    if (!lastParallelLog || !parallelLogContent) {
      clearParallelLogUi();
      return;
    }
    if (parallelLogKpiScopeEl) {
      parallelLogKpiScope = parallelLogKpiScopeEl.value === "overall" ? "overall" : "process";
    }
    const ksrc = getParallelLogKpiSource();
    if (parallelLogEmpty) parallelLogEmpty.hidden = true;
    parallelLogContent.hidden = false;
    if (kpiUtilParallel) kpiUtilParallel.textContent = formatPercent(ksrc.utilAvg ?? NaN);
    if (kpiOeeParallel) kpiOeeParallel.textContent = formatPercent(ksrc.oeeAvg ?? NaN);
    if (kpiDefectLabelParallel && kpiDefectParallel) {
      if (Number.isFinite(ksrc.defectAvg)) {
        kpiDefectLabelParallel.textContent = "불량율";
        kpiDefectParallel.textContent = formatPercent(ksrc.defectAvg);
      } else if (Number.isFinite(ksrc.sumDefectQty)) {
        kpiDefectLabelParallel.textContent = "불량(합)";
        kpiDefectParallel.textContent = `${formatQty(ksrc.sumDefectQty)}건`;
      } else {
        kpiDefectLabelParallel.textContent = "불량율";
        kpiDefectParallel.textContent = "—";
      }
    }
    if (parallelLogMeta) {
      const dr =
        lastParallelLog.dateFrom && lastParallelLog.dateTo
          ? ` · 작업일자 ${lastParallelLog.dateFrom} ~ ${lastParallelLog.dateTo}`
          : "";
      const kpiNote =
        ksrc.utilIsDerived || ksrc.oeeIsDerived
          ? " · 가동율·설비효율: 열이 없을 때 작업시간·정지·투입·생산수량·불량(율)·생산성으로 행별 산출 후 평균"
          : "";
      const fn = reaggregateParallelLogFiltered();
      const dayNote = fn
        ? ` · 보기: 필터 적용 (${fn.dataRows}행 / 전체 ${lastParallelLog.dataRows}행)`
        : "";
      parallelLogMeta.textContent = `파일: ${lastParallelLog.fileLabel} · 시트 「${lastParallelLog.sheetName}」만 사용 · 유효 ${lastParallelLog.dataRows}행(전체 행 집계)${dr}${dayNote}${kpiNote}`;
    }
    if (parallelLogOps) {
      const bits = [];
      if (Number.isFinite(ksrc.sumProdQty)) bits.push(`생산수량 합 ${formatQty(ksrc.sumProdQty)}`);
      if (Number.isFinite(ksrc.sumWorkTime)) bits.push(`작업시간 합 ${formatQty(ksrc.sumWorkTime)}`);
      if (Number.isFinite(ksrc.sumInputTime)) bits.push(`투입시간 합 ${formatQty(ksrc.sumInputTime)}`);
      if (Number.isFinite(ksrc.productivityAvg)) bits.push(`생산성 평균 ${formatPercent(ksrc.productivityAvg)}`);
      if (Number.isFinite(ksrc.sumDefectQty)) bits.push(`불량 합 ${formatQty(ksrc.sumDefectQty)}건`);
      const qrst =
        Array.isArray(lastParallelLog.iStopTimeCols) && lastParallelLog.iStopTimeCols.length >= 4;
      if (qrst) {
        if (Number.isFinite(ksrc.sumStopExchange)) bits.push(`교환 ${formatQty(ksrc.sumStopExchange)}`);
        if (Number.isFinite(ksrc.sumStopRepair)) bits.push(`수리 ${formatQty(ksrc.sumStopRepair)}`);
        if (Number.isFinite(ksrc.sumStopMaterial)) bits.push(`소재 ${formatQty(ksrc.sumStopMaterial)}`);
        if (Number.isFinite(ksrc.sumStopPlanned)) bits.push(`계획정지 ${formatQty(ksrc.sumStopPlanned)}`);
        if (Number.isFinite(ksrc.sumStopTime)) bits.push(`정지 합 ${formatQty(ksrc.sumStopTime)}`);
      } else if (Number.isFinite(ksrc.sumStopTime)) {
        bits.push(`정지 시간 합 ${formatQty(ksrc.sumStopTime)}`);
      }
      parallelLogOps.textContent = bits.join(" · ");
      parallelLogOps.hidden = bits.length === 0;
    }
    renderParallelLogTimeline();
    renderParallelLogTables();
    renderParallelLogProductionQtyTable();
    if (parallelLogMaintWrap) {
      if (lastParallelLog.maintenance) {
        parallelLogMaintWrap.hidden = false;
        renderWorkbookMaintenanceInto(parallelLogMaintWrap, lastParallelLog.maintenance);
      } else {
        parallelLogMaintWrap.innerHTML = "";
        parallelLogMaintWrap.hidden = true;
      }
    }
  }

  function resetDoublePipeLogSlicerState() {
    doublePipeLogGranularity = "month";
    doublePipeLogSelectedMonths = new Set();
    doublePipeLogSelectedDays = new Set();
    doublePipeLogTimelineYear = new Date().getFullYear();
  }

  function doublePipeLogYearsFromStats(log) {
    if (!log) return [];
    const ys = new Set();
    (log.monthlyStats || []).forEach((m) => {
      if (!m || m.key === "일자미상") return;
      const y = parseInt(String(m.key).slice(0, 4), 10);
      if (Number.isFinite(y)) ys.add(y);
    });
    return [...ys].sort((a, b) => a - b);
  }

  function initDoublePipeLogSlicerFromData() {
    resetDoublePipeLogSlicerState();
    doublePipeLogProcessMach = true;
    doublePipeLogProcessForm = true;
    if (doublePipeLogFilterMachEl) doublePipeLogFilterMachEl.checked = true;
    if (doublePipeLogFilterFormEl) doublePipeLogFilterFormEl.checked = true;
    if (!lastDoublePipeLog) return;
    const ys = doublePipeLogYearsFromStats(lastDoublePipeLog);
    if (ys.length) doublePipeLogTimelineYear = ys[ys.length - 1];
    else if (lastDoublePipeLog.dateTo && String(lastDoublePipeLog.dateTo).length >= 4) {
      const y = parseInt(String(lastDoublePipeLog.dateTo).slice(0, 4), 10);
      if (Number.isFinite(y)) doublePipeLogTimelineYear = y;
    }
  }

  function doublePipeLogProcessFilterRestrictsRows() {
    if (!doublePipeLogProcessMach && !doublePipeLogProcessForm) return false;
    if (doublePipeLogProcessMach && doublePipeLogProcessForm) return false;
    return true;
  }

  function doublePipeLogRowMatchesProcessFilter(row, iProcess) {
    if (!doublePipeLogProcessMach && !doublePipeLogProcessForm) return true;
    if (doublePipeLogProcessMach && doublePipeLogProcessForm) return true;
    const pk = normalizeDoublePipeProcess(iProcess >= 0 ? row[iProcess] : "");
    if (pk === "가공") return doublePipeLogProcessMach;
    if (pk === "성형") return doublePipeLogProcessForm;
    return true;
  }

  function doublePipeLogProcessFilterSummaryFragment() {
    if (!doublePipeLogProcessFilterRestrictsRows()) return "";
    const bits = [];
    if (doublePipeLogProcessMach) bits.push("가공");
    if (doublePipeLogProcessForm) bits.push("성형");
    return bits.length ? `공정: ${bits.join("+")}` : "";
  }

  function reaggregateDoublePipeLogFiltered(options = {}) {
    const ignoreProcessForRowFilter = options.ignoreProcessForRowFilter === true;
    if (!lastDoublePipeLog || !lastDoublePipeLog._filterMatrix) return null;
    const hasMonth = doublePipeLogGranularity === "month" && doublePipeLogSelectedMonths.size > 0;
    const hasDay = doublePipeLogGranularity === "day" && doublePipeLogSelectedDays.size > 0;
    const procRestricts = doublePipeLogProcessFilterRestrictsRows();
    const hasProcF = procRestricts && !ignoreProcessForRowFilter;
    if (!hasMonth && !hasDay && !hasProcF) return null;

    const matrix = lastDoublePipeLog._filterMatrix;
    const hi = lastDoublePipeLog._filterHi;
    const idx = lastDoublePipeLog._filterIdx;
    const iWorkDate = idx.iWorkDate;
    const iEq = lastDoublePipeLog.iEquipment;
    const iProcess = lastDoublePipeLog.iProcess;

    const rowIncluded = (row) => {
      if (!row || !row.length) return false;
      if (!row.some((c) => String(c ?? "").trim() !== "")) return false;
      if (!ignoreProcessForRowFilter && !doublePipeLogRowMatchesProcessFilter(row, iProcess)) return false;
      const ymd = iWorkDate >= 0 ? parseExcelDate(row[iWorkDate]) : "";
      const ym = ymd && ymd.length >= 7 ? ymd.slice(0, 7) : "일자미상";
      const dayKey = ymd || "일자미상";
      if (hasDay) return doublePipeLogSelectedDays.has(dayKey);
      if (hasMonth) return doublePipeLogSelectedMonths.has(ym);
      if (hasProcF) return true;
      return false;
    };

    const aggGlobal = makeDrawingAgg();
    const byEq = new Map();
    const byProcMonth = new Map();
    const byProcEq = new Map();
    const byMonthPivot = new Map();
    const byDayPivot = new Map();

    for (let r = hi + 1; r < matrix.length; r++) {
      const row = matrix[r];
      if (!rowIncluded(row)) continue;
      addDrawingRowToAgg(aggGlobal, row, idx, false);
      const ymdPV = iWorkDate >= 0 ? parseExcelDate(row[iWorkDate]) : "";
      const ymPV = ymdPV && ymdPV.length >= 7 ? ymdPV.slice(0, 7) : "일자미상";
      const dayKeyPV = ymdPV || "일자미상";
      if (!byMonthPivot.has(ymPV)) byMonthPivot.set(ymPV, makeDrawingAgg());
      addDrawingRowToAgg(byMonthPivot.get(ymPV), row, idx, false);
      if (!byDayPivot.has(dayKeyPV)) byDayPivot.set(dayKeyPV, makeDrawingAgg());
      addDrawingRowToAgg(byDayPivot.get(dayKeyPV), row, idx, false);
      const eqKey = iEq >= 0 ? String(row[iEq] ?? "").trim() || "미지정" : "전체";
      if (!byEq.has(eqKey)) byEq.set(eqKey, makeDrawingAgg());
      addDrawingRowToAgg(byEq.get(eqKey), row, idx, false);
      const ymd = iWorkDate >= 0 ? parseExcelDate(row[iWorkDate]) : "";
      const ym = ymd && ymd.length >= 7 ? ymd.slice(0, 7) : "일자미상";
      const procKey = normalizeDoublePipeProcess(iProcess >= 0 ? row[iProcess] : "");
      addDrawingRowToAgg(ensureAggMapNested2(byProcMonth, procKey, ym), row, idx, false);
      addDrawingRowToAgg(ensureAggMapNested2(byProcEq, procKey, eqKey), row, idx, false);
    }
    const g = finalizeDrawingAgg(aggGlobal, idx);
    const equipmentStats = [...byEq.entries()]
      .sort((a, b) => compareDrawingEquipmentKeys(a[0], b[0]))
      .map(([key, agg]) => ({ key, label: key, ...finalizeDrawingAgg(agg, idx) }));

    const { pivotMonthlyStats, pivotDailyStats } = buildPivotMonthlyDailyStatsFromAggMaps(byMonthPivot, byDayPivot, idx);

    const monthLabel = (ymKey) => {
      if (ymKey === "일자미상") return ymKey;
      const p = ymKey.split("-");
      if (p.length !== 2) return ymKey;
      return `${p[0]}년 ${parseInt(p[1], 10)}월`;
    };
    const sortProcKeys = (a, b) => {
      const rank = (k) => (k === "가공" ? 0 : k === "성형" ? 1 : k === "기타" ? 2 : 99);
      const ra = rank(a);
      const rb = rank(b);
      if (ra !== rb) return ra - rb;
      return String(a).localeCompare(String(b));
    };
    const procKeySet = new Set([...byProcMonth.keys(), ...byProcEq.keys()]);
    const processBlocks = [...procKeySet].sort(sortProcKeys).map((pk) => {
      const pm = byProcMonth.get(pk) || new Map();
      const peq = byProcEq.get(pk) || new Map();
      const monthlyStatsP = [...pm.entries()]
        .sort((a, b) => {
          if (a[0] === "일자미상") return 1;
          if (b[0] === "일자미상") return -1;
          return a[0].localeCompare(b[0]);
        })
        .map(([key, agg]) => ({ key, label: monthLabel(key), ...finalizeDrawingAgg(agg, idx) }));
      const equipmentStatsP = [...peq.entries()]
        .sort((a, b) => compareDrawingEquipmentKeys(a[0], b[0]))
        .map(([key, agg]) => ({ key, label: key, ...finalizeDrawingAgg(agg, idx) }));
      return {
        process: pk,
        monthlyStats: monthlyStatsP,
        equipmentStats: equipmentStatsP,
        dailyStats: [],
        dayEquipmentStats: {},
      };
    });

    return { ...g, equipmentStats, dataRows: aggGlobal.dataRows, processBlocks, pivotMonthlyStats, pivotDailyStats };
  }

  function getDoublePipeLogEquipmentStatsForTables() {
    const sub = reaggregateDoublePipeLogFiltered();
    if (sub && Array.isArray(sub.equipmentStats)) return sub.equipmentStats;
    return lastDoublePipeLog && lastDoublePipeLog.equipmentStats ? lastDoublePipeLog.equipmentStats : [];
  }

  function renderDoublePipeLogTables() {
    if (!doublePipeLogTablesWrap || !lastDoublePipeLog) return;
    doublePipeLogTablesWrap.innerHTML = "";
    const sub = reaggregateDoublePipeLogFiltered();
    const slicerActive = !!sub;
    const months = slicerActive ? [] : lastDoublePipeLog.monthlyStats || [];
    const eqs = getDoublePipeLogEquipmentStatsForTables();
    const procBlocks =
      sub && Array.isArray(sub.processBlocks) ? sub.processBlocks : lastDoublePipeLog.processBlocks || [];
    const splitStop =
      Array.isArray(lastDoublePipeLog.iStopTimeCols) && lastDoublePipeLog.iStopTimeCols.length >= 4;

    const procHasRows = procBlocks.some(
      (pb) => (pb.monthlyStats && pb.monthlyStats.length) || (pb.equipmentStats && pb.equipmentStats.length)
    );
    if (months.length === 0 && eqs.length === 0 && !procHasRows) {
      doublePipeLogTablesWrap.hidden = false;
      const p = document.createElement("p");
      p.className = "empty-state empty-state--flat";
      p.textContent = slicerActive
        ? "선택한 기간에 표시할 집계가 없습니다."
        : "표시할 집계가 없습니다.";
      doublePipeLogTablesWrap.appendChild(p);
      return;
    }
    const stopLabels = splitStop ? DOUBLE_PIPE_STOP_LABELS : undefined;
    if (months.length) {
      doublePipeLogTablesWrap.appendChild(
        buildDrawingBreakdownTable("월별 집계", "월", months, splitStop, stopLabels)
      );
    }
    if (eqs.length) {
      let eqTitle = lastDoublePipeLog.hasEquipmentColumn ? "설비별 집계" : "설비별 집계 (설비 열 없음 · 전체)";
      if (slicerActive) {
        const hint =
          doublePipeLogGranularity === "month"
            ? [...doublePipeLogSelectedMonths].sort().join(", ")
            : [...doublePipeLogSelectedDays].sort().join(", ");
        eqTitle += ` · 선택: ${hint}`;
      }
      doublePipeLogTablesWrap.appendChild(
        buildDrawingBreakdownTable(eqTitle, "설비·구분", eqs, splitStop, stopLabels)
      );
    }
    for (const blk of procBlocks) {
      const m = blk.monthlyStats || [];
      const e = blk.equipmentStats || [];
      if (!m.length && !e.length) continue;
      const sec = document.createElement("div");
      sec.className = "drawing-log-block";
      const h3 = document.createElement("h3");
      h3.className = "drawing-log-block__title";
      h3.textContent = `공정 · ${blk.process}`;
      sec.appendChild(h3);
      if (m.length) sec.appendChild(buildDrawingBreakdownTable("월별 집계", "월", m, splitStop, stopLabels));
      if (e.length) {
        let eqTitle = lastDoublePipeLog.hasEquipmentColumn ? "설비별 집계" : "설비별 집계 (설비 열 없음 · 전체)";
        if (slicerActive) {
          const hint =
            doublePipeLogGranularity === "month"
              ? [...doublePipeLogSelectedMonths].sort().join(", ")
              : [...doublePipeLogSelectedDays].sort().join(", ");
          eqTitle += ` · 선택: ${hint}`;
        }
        sec.appendChild(buildDrawingBreakdownTable(eqTitle, "설비·구분", e, splitStop, stopLabels));
      }
      doublePipeLogTablesWrap.appendChild(sec);
    }
    doublePipeLogTablesWrap.hidden = false;
  }

  function renderDoublePipeLogProductionQtyTable() {
    if (!doublePipeLogProdWrap || !lastDoublePipeLog) return;
    doublePipeLogProdWrap.innerHTML = "";
    doublePipeLogProdWrap.hidden = true;
    const subDp = reaggregateDoublePipeLogFiltered();
    if (!subDp) return;

    const monthMode = doublePipeLogGranularity === "month" && doublePipeLogSelectedMonths.size > 0;
    const dayMode = doublePipeLogGranularity === "day" && doublePipeLogSelectedDays.size > 0;
    if (!monthMode && !dayMode) return;

    const monthPool = subDp.pivotMonthlyStats || [];
    const dayPool = subDp.pivotDailyStats || [];
    let statRows = [];
    if (monthMode) {
      for (const ym of [...doublePipeLogSelectedMonths].sort()) {
        const st = monthPool.find((m) => m.key === ym);
        if (st) statRows.push(st);
      }
    } else {
      for (const dk of [...doublePipeLogSelectedDays].sort()) {
        const st = dayPool.find((d) => d.key === dk);
        if (st) statRows.push(st);
      }
    }
    if (!statRows.length) return;

    const hasProdCol = lastDoublePipeLog.iProdQty >= 0;
    doublePipeLogProdWrap.hidden = false;
    const block = document.createElement("div");
    block.className = "drawing-log-block";
    const h3 = document.createElement("h3");
    h3.className = "drawing-log-block__title";
    h3.textContent = monthMode ? "선택 월 · 생산수량" : "선택 일 · 생산수량";
    block.appendChild(h3);
    if (!hasProdCol) {
      const note = document.createElement("p");
      note.className = "drawing-log-meta";
      note.textContent =
        "시트에 생산량 열(N열 또는 헤더 인식)이 없으면 생산수량 표가 비어 있습니다. 엑셀을 확인해 주세요.";
      block.appendChild(note);
    }
    const wrap = document.createElement("div");
    wrap.className = "drawing-log-preview-wrap drawing-log-prod-pivot-wrap";
    const tbl = document.createElement("table");
    tbl.className = "drawing-log-prod-pivot";

    const thead = document.createElement("thead");
    const trh = document.createElement("tr");
    const thCorner = document.createElement("th");
    thCorner.textContent = "";
    thCorner.classList.add("drawing-log-prod-pivot__corner");
    thCorner.setAttribute("aria-label", monthMode ? "월 열" : "날짜 열");
    trh.appendChild(thCorner);
    statRows.forEach((st) => {
      const th = document.createElement("th");
      th.textContent = String(st.label || st.key);
      th.classList.add("num");
      th.title = String(st.key || "");
      trh.appendChild(th);
    });
    const thTot = document.createElement("th");
    thTot.textContent = "합계";
    thTot.classList.add("num");
    trh.appendChild(thTot);
    thead.appendChild(trh);
    tbl.appendChild(thead);

    const tbody = document.createElement("tbody");
    function addPivotRow(label, getter, opts) {
      const asInt = opts && opts.asInt;
      const tr = document.createElement("tr");
      const td0 = document.createElement("td");
      td0.textContent = label;
      td0.classList.add("drawing-log-prod-pivot__rowhead");
      tr.appendChild(td0);
      let sum = 0;
      let n = 0;
      for (const st of statRows) {
        const v = getter(st);
        const td = document.createElement("td");
        td.classList.add("num");
        if (Number.isFinite(v)) {
          td.textContent = asInt ? String(Math.round(v)) : formatQty(v);
          sum += v;
          n++;
        } else {
          td.textContent = "—";
        }
        tr.appendChild(td);
      }
      const tdSum = document.createElement("td");
      tdSum.classList.add("num");
      tdSum.style.fontWeight = "700";
      if (n && Number.isFinite(sum)) {
        tdSum.textContent = asInt ? String(Math.round(sum)) : formatQty(sum);
      } else {
        tdSum.textContent = "—";
      }
      tr.appendChild(tdSum);
      tbody.appendChild(tr);
    }

    addPivotRow("생산수량 합", (st) => (hasProdCol && Number.isFinite(st.sumProdQty) ? st.sumProdQty : NaN));
    addPivotRow("건수", (st) => (Number.isFinite(st.dataRows) ? st.dataRows : NaN), { asInt: true });
    addPivotRow("작업시간 합", (st) => (Number.isFinite(st.sumWorkTime) ? st.sumWorkTime : NaN));

    tbl.appendChild(tbody);
    wrap.appendChild(tbl);
    block.appendChild(wrap);
    doublePipeLogProdWrap.appendChild(block);
  }

  function doublePipeLogSlicerSummaryText() {
    const sub = reaggregateDoublePipeLogFiltered();
    const procFrag = doublePipeLogProcessFilterSummaryFragment();
    if (doublePipeLogGranularity === "month" && doublePipeLogSelectedMonths.size && sub) {
      const parts = [...doublePipeLogSelectedMonths].sort().map((ym) => {
        const p = ym.split("-");
        return p.length === 2 ? `${p[0]}년 ${parseInt(p[1], 10)}월` : ym;
      });
      let t = `${parts.join(", ")} · ${sub.dataRows}행`;
      if (procFrag) t += ` · ${procFrag}`;
      return t;
    }
    if (doublePipeLogGranularity === "day" && doublePipeLogSelectedDays.size && sub) {
      let t = `${[...doublePipeLogSelectedDays].sort().join(", ")} · ${sub.dataRows}행`;
      if (procFrag) t += ` · ${procFrag}`;
      return t;
    }
    if (procFrag && sub) return `${procFrag} · ${sub.dataRows}행`;
    return "";
  }

  function pruneDoublePipeLogSelectedDaysInvalid() {
    if (!lastDoublePipeLog || doublePipeLogGranularity !== "day") return;
    const pool = new Set(
      (lastDoublePipeLog.dailyStats || [])
        .filter((d) => {
          if (d.key === "일자미상") return false;
          if (!d.key.startsWith(`${doublePipeLogTimelineYear}-`)) return false;
          if (doublePipeLogSelectedMonths.size) {
            const ym = d.key.slice(0, 7);
            return doublePipeLogSelectedMonths.has(ym);
          }
          return true;
        })
        .map((d) => d.key)
    );
    [...doublePipeLogSelectedDays].forEach((k) => {
      if (!pool.has(k)) doublePipeLogSelectedDays.delete(k);
    });
  }

  function renderDoublePipeLogTimeline() {
    if (!doublePipeLogTimeline || !lastDoublePipeLog || !lastDoublePipeLog.monthlyStats) {
      if (doublePipeLogTimeline) doublePipeLogTimeline.hidden = true;
      return;
    }
    doublePipeLogTimeline.hidden = false;
    if (doublePipeLogGranularityEl) doublePipeLogGranularityEl.value = doublePipeLogGranularity;
    if (doublePipeLogKpiScopeEl)
      doublePipeLogKpiScopeEl.value = doublePipeLogKpiScope === "overall" ? "overall" : "process";

    const ys = doublePipeLogYearsFromStats(lastDoublePipeLog);
    if (ys.length) {
      if (doublePipeLogTimelineYear < ys[0]) doublePipeLogTimelineYear = ys[0];
      if (doublePipeLogTimelineYear > ys[ys.length - 1]) doublePipeLogTimelineYear = ys[ys.length - 1];
    }
    if (doublePipeLogYearLabel) doublePipeLogYearLabel.textContent = `${doublePipeLogTimelineYear}년`;
    if (doublePipeLogYearPrev) doublePipeLogYearPrev.disabled = ys.length && doublePipeLogTimelineYear <= ys[0];
    if (doublePipeLogYearNext)
      doublePipeLogYearNext.disabled = ys.length && doublePipeLogTimelineYear >= ys[ys.length - 1];

    const monthKeys = new Set((lastDoublePipeLog.monthlyStats || []).map((m) => m.key));
    if (doublePipeLogMonthStrip) {
      doublePipeLogMonthStrip.innerHTML = "";
      for (let m = 1; m <= 12; m++) {
        const ym = `${doublePipeLogTimelineYear}-${String(m).padStart(2, "0")}`;
        const btn = document.createElement("button");
        btn.type = "button";
        btn.className = "drawing-log-timeline__chip";
        btn.dataset.ym = ym;
        btn.textContent = String(m);
        const has = monthKeys.has(ym);
        if (!has) {
          btn.classList.add("is-empty");
          btn.disabled = true;
        }
        if (doublePipeLogSelectedMonths.has(ym)) btn.classList.add("is-selected");
        doublePipeLogMonthStrip.appendChild(btn);
      }
    }

    const dayMode = doublePipeLogGranularity === "day";
    if (doublePipeLogDayStrip) {
      doublePipeLogDayStrip.hidden = !dayMode;
      doublePipeLogDayStrip.innerHTML = "";
      if (dayMode) {
        const pool = (lastDoublePipeLog.dailyStats || []).filter((d) => {
          if (d.key === "일자미상") return false;
          if (!d.key.startsWith(`${doublePipeLogTimelineYear}-`)) return false;
          if (doublePipeLogSelectedMonths.size) {
            const ym = d.key.length >= 7 ? d.key.slice(0, 7) : "";
            return doublePipeLogSelectedMonths.has(ym);
          }
          return true;
        });
        pool.forEach((d) => {
          const btn = document.createElement("button");
          btn.type = "button";
          btn.className = "drawing-log-timeline__chip";
          btn.dataset.dayKey = d.key;
          btn.textContent = d.key.slice(5);
          if (doublePipeLogSelectedDays.has(d.key)) btn.classList.add("is-selected");
          doublePipeLogDayStrip.appendChild(btn);
        });
      }
    }

    if (doublePipeLogSlicerSelection) {
      const t = doublePipeLogSlicerSummaryText();
      doublePipeLogSlicerSelection.textContent = t || "전체 기간 (필터 없음)";
    }
  }

  function getDoublePipeLogKpiSource() {
    if (!lastDoublePipeLog) return null;
    const sub = reaggregateDoublePipeLogFiltered({
      ignoreProcessForRowFilter: doublePipeLogKpiScope === "overall",
    });
    if (!sub) return lastDoublePipeLog;
    return { ...lastDoublePipeLog, ...sub };
  }

  function renderDoublePipeLogPanel() {
    if (!lastDoublePipeLog || !doublePipeLogContent) {
      clearDoublePipeLogUi();
      return;
    }
    if (doublePipeLogKpiScopeEl) {
      doublePipeLogKpiScope = doublePipeLogKpiScopeEl.value === "overall" ? "overall" : "process";
    }
    const ksrc = getDoublePipeLogKpiSource();
    if (doublePipeLogEmpty) doublePipeLogEmpty.hidden = true;
    doublePipeLogContent.hidden = false;
    if (kpiUtilDoublePipe) kpiUtilDoublePipe.textContent = formatPercent(ksrc.utilAvg ?? NaN);
    if (kpiOeeDoublePipe) kpiOeeDoublePipe.textContent = formatPercent(ksrc.oeeAvg ?? NaN);
    if (kpiDefectLabelDoublePipe && kpiDefectDoublePipe) {
      if (Number.isFinite(ksrc.defectAvg)) {
        kpiDefectLabelDoublePipe.textContent = "불량율";
        kpiDefectDoublePipe.textContent = formatPercent(ksrc.defectAvg);
      } else if (Number.isFinite(ksrc.sumDefectQty)) {
        kpiDefectLabelDoublePipe.textContent = "불량(합)";
        kpiDefectDoublePipe.textContent = `${formatQty(ksrc.sumDefectQty)}건`;
      } else {
        kpiDefectLabelDoublePipe.textContent = "불량율";
        kpiDefectDoublePipe.textContent = "—";
      }
    }
    if (doublePipeLogMeta) {
      const dr =
        lastDoublePipeLog.dateFrom && lastDoublePipeLog.dateTo
          ? ` · 작업일자 ${lastDoublePipeLog.dateFrom} ~ ${lastDoublePipeLog.dateTo}`
          : "";
      const kpiNote =
        ksrc.utilIsDerived || ksrc.oeeIsDerived
          ? " · 가동율·설비효율: 열이 없을 때 작업시간·정지·투입·생산수량·불량(율)·생산성으로 행별 산출 후 평균"
          : "";
      const fn = reaggregateDoublePipeLogFiltered();
      const dayNote = fn
        ? ` · 보기: 필터 적용 (${fn.dataRows}행 / 전체 ${lastDoublePipeLog.dataRows}행)`
        : "";
      doublePipeLogMeta.textContent = `파일: ${lastDoublePipeLog.fileLabel} · 시트 「${lastDoublePipeLog.sheetName}」 · 유효 ${lastDoublePipeLog.dataRows}행 · 생산 N열 · 비가동 S~V · 불량 W+AJ${dr}${dayNote}${kpiNote}`;
    }
    if (doublePipeLogOps) {
      const bits = [];
      if (Number.isFinite(ksrc.sumProdQty)) bits.push(`생산수량 합 ${formatQty(ksrc.sumProdQty)}`);
      if (Number.isFinite(ksrc.sumWorkTime)) bits.push(`작업시간 합 ${formatQty(ksrc.sumWorkTime)}`);
      if (Number.isFinite(ksrc.sumInputTime)) bits.push(`투입시간 합 ${formatQty(ksrc.sumInputTime)}`);
      if (Number.isFinite(ksrc.productivityAvg)) bits.push(`생산성 평균 ${formatPercent(ksrc.productivityAvg)}`);
      if (Number.isFinite(ksrc.sumDefectQty)) bits.push(`불량 합 ${formatQty(ksrc.sumDefectQty)}건`);
      const qrst =
        Array.isArray(lastDoublePipeLog.iStopTimeCols) && lastDoublePipeLog.iStopTimeCols.length >= 4;
      if (qrst) {
        if (Number.isFinite(ksrc.sumStopExchange)) bits.push(`S ${formatQty(ksrc.sumStopExchange)}`);
        if (Number.isFinite(ksrc.sumStopRepair)) bits.push(`T ${formatQty(ksrc.sumStopRepair)}`);
        if (Number.isFinite(ksrc.sumStopMaterial)) bits.push(`U ${formatQty(ksrc.sumStopMaterial)}`);
        if (Number.isFinite(ksrc.sumStopPlanned)) bits.push(`V ${formatQty(ksrc.sumStopPlanned)}`);
        if (Number.isFinite(ksrc.sumStopTime)) bits.push(`비가동 합 ${formatQty(ksrc.sumStopTime)}`);
      } else if (Number.isFinite(ksrc.sumStopTime)) {
        bits.push(`정지 시간 합 ${formatQty(ksrc.sumStopTime)}`);
      }
      doublePipeLogOps.textContent = bits.join(" · ");
      doublePipeLogOps.hidden = bits.length === 0;
    }
    renderDoublePipeLogTimeline();
    renderDoublePipeLogTables();
    renderDoublePipeLogProductionQtyTable();
    if (doublePipeLogMaintWrap) {
      if (lastDoublePipeLog.maintenance) {
        doublePipeLogMaintWrap.hidden = false;
        renderWorkbookMaintenanceInto(doublePipeLogMaintWrap, lastDoublePipeLog.maintenance);
      } else {
        doublePipeLogMaintWrap.innerHTML = "";
        doublePipeLogMaintWrap.hidden = true;
      }
    }
  }

  function resetHwaseungDoublePipeLogSlicerState() {
    hwaseungDoublePipeLogGranularity = "month";
    hwaseungDoublePipeLogSelectedMonths = new Set();
    hwaseungDoublePipeLogSelectedDays = new Set();
    hwaseungDoublePipeLogTimelineYear = new Date().getFullYear();
  }

  function initHwaseungDoublePipeLogSlicerFromData() {
    resetHwaseungDoublePipeLogSlicerState();
    hwaseungDoublePipeLogProcessSelected = new Set(HWASEUNG_DOUBLE_PIPE_ALL_PROC_KEYS);
    if (hwaseungDoublePipeLogProcessFilters) {
      hwaseungDoublePipeLogProcessFilters.querySelectorAll('input[type="checkbox"][data-hwaseung-proc]').forEach((el) => {
        el.checked = true;
      });
    }
    if (!lastHwaseungDoublePipeLog) return;
    const ys = doublePipeLogYearsFromStats(lastHwaseungDoublePipeLog);
    if (ys.length) hwaseungDoublePipeLogTimelineYear = ys[ys.length - 1];
    else if (lastHwaseungDoublePipeLog.dateTo && String(lastHwaseungDoublePipeLog.dateTo).length >= 4) {
      const y = parseInt(String(lastHwaseungDoublePipeLog.dateTo).slice(0, 4), 10);
      if (Number.isFinite(y)) hwaseungDoublePipeLogTimelineYear = y;
    }
  }

  function readHwaseungDoublePipeLogProcessSelectionFromDom() {
    const s = new Set();
    if (hwaseungDoublePipeLogProcessFilters) {
      hwaseungDoublePipeLogProcessFilters.querySelectorAll('input[type="checkbox"][data-hwaseung-proc]').forEach((el) => {
        const k = el.getAttribute("data-hwaseung-proc");
        if (k && el.checked) s.add(k);
      });
    }
    if (s.size === 0) {
      HWASEUNG_DOUBLE_PIPE_ALL_PROC_KEYS.forEach((k) => s.add(k));
      if (hwaseungDoublePipeLogProcessFilters) {
        hwaseungDoublePipeLogProcessFilters
          .querySelectorAll('input[type="checkbox"][data-hwaseung-proc]')
          .forEach((el) => {
            el.checked = true;
          });
      }
    }
    hwaseungDoublePipeLogProcessSelected = s;
  }

  function readHwaseungDoublePipeLogKpiScopeFromDom() {
    if (!hwaseungDoublePipeLogKpiScopeEl) return;
    hwaseungDoublePipeLogKpiScope = hwaseungDoublePipeLogKpiScopeEl.value === "overall" ? "overall" : "process";
  }

  function hwaseungDoublePipeLogProcessFilterRestrictsRows() {
    if (hwaseungDoublePipeLogProcessSelected.size !== HWASEUNG_DOUBLE_PIPE_ALL_PROC_KEYS.length) return true;
    return !HWASEUNG_DOUBLE_PIPE_ALL_PROC_KEYS.every((k) => hwaseungDoublePipeLogProcessSelected.has(k));
  }

  function hwaseungDoublePipeLogRowMatchesProcessFilter(row, iProcess) {
    const pk = normalizeHwaseungDoublePipeProcess(iProcess >= 0 ? row[iProcess] : "");
    return hwaseungDoublePipeLogProcessSelected.has(pk);
  }

  function hwaseungDoublePipeLogProcessFilterSummaryFragment() {
    if (!hwaseungDoublePipeLogProcessFilterRestrictsRows()) return "";
    const ordered = HWASEUNG_DOUBLE_PIPE_ALL_PROC_KEYS.filter((k) => hwaseungDoublePipeLogProcessSelected.has(k));
    return ordered.length ? `공정: ${ordered.join("·")}` : "";
  }

  function reaggregateHwaseungDoublePipeLogFiltered(options = {}) {
    const ignoreProcessForRowFilter = options.ignoreProcessForRowFilter === true;
    if (!lastHwaseungDoublePipeLog || !lastHwaseungDoublePipeLog._filterMatrix) return null;
    const hasMonth = hwaseungDoublePipeLogGranularity === "month" && hwaseungDoublePipeLogSelectedMonths.size > 0;
    const hasDay = hwaseungDoublePipeLogGranularity === "day" && hwaseungDoublePipeLogSelectedDays.size > 0;
    const procRestricts = hwaseungDoublePipeLogProcessFilterRestrictsRows();
    const hasProcF = procRestricts && !ignoreProcessForRowFilter;
    if (!hasMonth && !hasDay && !hasProcF) return null;

    const matrix = lastHwaseungDoublePipeLog._filterMatrix;
    const hi = lastHwaseungDoublePipeLog._filterHi;
    const idx = lastHwaseungDoublePipeLog._filterIdx;
    const iWorkDate = idx.iWorkDate;
    const iEq = lastHwaseungDoublePipeLog.iEquipment;
    const iProcess = lastHwaseungDoublePipeLog.iProcess;

    const rowIncluded = (row) => {
      if (!row || !row.length) return false;
      if (!row.some((c) => String(c ?? "").trim() !== "")) return false;
      if (!ignoreProcessForRowFilter && !hwaseungDoublePipeLogRowMatchesProcessFilter(row, iProcess)) return false;
      const ymd = iWorkDate >= 0 ? parseExcelDate(row[iWorkDate]) : "";
      const ym = ymd && ymd.length >= 7 ? ymd.slice(0, 7) : "일자미상";
      const dayKey = ymd || "일자미상";
      if (hasDay) return hwaseungDoublePipeLogSelectedDays.has(dayKey);
      if (hasMonth) return hwaseungDoublePipeLogSelectedMonths.has(ym);
      if (hasProcF) return true;
      return false;
    };

    const aggGlobal = makeDrawingAgg();
    const byEq = new Map();
    const byProcMonth = new Map();
    const byProcEq = new Map();
    const byMonthPivot = new Map();
    const byDayPivot = new Map();

    for (let r = hi + 1; r < matrix.length; r++) {
      const row = matrix[r];
      if (!rowIncluded(row)) continue;
      addDrawingRowToAgg(aggGlobal, row, idx, false);
      const ymdPV = iWorkDate >= 0 ? parseExcelDate(row[iWorkDate]) : "";
      const ymPV = ymdPV && ymdPV.length >= 7 ? ymdPV.slice(0, 7) : "일자미상";
      const dayKeyPV = ymdPV || "일자미상";
      if (!byMonthPivot.has(ymPV)) byMonthPivot.set(ymPV, makeDrawingAgg());
      addDrawingRowToAgg(byMonthPivot.get(ymPV), row, idx, false);
      if (!byDayPivot.has(dayKeyPV)) byDayPivot.set(dayKeyPV, makeDrawingAgg());
      addDrawingRowToAgg(byDayPivot.get(dayKeyPV), row, idx, false);
      const eqKey = iEq >= 0 ? String(row[iEq] ?? "").trim() || "미지정" : "전체";
      if (!byEq.has(eqKey)) byEq.set(eqKey, makeDrawingAgg());
      addDrawingRowToAgg(byEq.get(eqKey), row, idx, false);
      const ymd = iWorkDate >= 0 ? parseExcelDate(row[iWorkDate]) : "";
      const ym = ymd && ymd.length >= 7 ? ymd.slice(0, 7) : "일자미상";
      const procKey = normalizeHwaseungDoublePipeProcess(iProcess >= 0 ? row[iProcess] : "");
      addDrawingRowToAgg(ensureAggMapNested2(byProcMonth, procKey, ym), row, idx, false);
      addDrawingRowToAgg(ensureAggMapNested2(byProcEq, procKey, eqKey), row, idx, false);
    }
    const g = finalizeDrawingAgg(aggGlobal, idx);
    const equipmentStats = [...byEq.entries()]
      .sort((a, b) => compareDrawingEquipmentKeys(a[0], b[0]))
      .map(([key, agg]) => ({ key, label: key, ...finalizeDrawingAgg(agg, idx) }));

    const { pivotMonthlyStats, pivotDailyStats } = buildPivotMonthlyDailyStatsFromAggMaps(byMonthPivot, byDayPivot, idx);

    const monthLabel = (ymKey) => {
      if (ymKey === "일자미상") return ymKey;
      const p = ymKey.split("-");
      if (p.length !== 2) return ymKey;
      return `${p[0]}년 ${parseInt(p[1], 10)}월`;
    };
    const procKeySet = new Set([...byProcMonth.keys(), ...byProcEq.keys()]);
    const processBlocks = [...procKeySet].sort(sortHwaseungDoublePipeProcessKeys).map((pk) => {
      const pm = byProcMonth.get(pk) || new Map();
      const peq = byProcEq.get(pk) || new Map();
      const monthlyStatsP = [...pm.entries()]
        .sort((a, b) => {
          if (a[0] === "일자미상") return 1;
          if (b[0] === "일자미상") return -1;
          return a[0].localeCompare(b[0]);
        })
        .map(([key, agg]) => ({ key, label: monthLabel(key), ...finalizeDrawingAgg(agg, idx) }));
      const equipmentStatsP = [...peq.entries()]
        .sort((a, b) => compareDrawingEquipmentKeys(a[0], b[0]))
        .map(([key, agg]) => ({ key, label: key, ...finalizeDrawingAgg(agg, idx) }));
      return {
        process: pk,
        monthlyStats: monthlyStatsP,
        equipmentStats: equipmentStatsP,
        dailyStats: [],
        dayEquipmentStats: {},
      };
    });

    return { ...g, equipmentStats, dataRows: aggGlobal.dataRows, processBlocks, pivotMonthlyStats, pivotDailyStats };
  }

  function getHwaseungDoublePipeLogEquipmentStatsForTables() {
    const sub = reaggregateHwaseungDoublePipeLogFiltered();
    if (sub && Array.isArray(sub.equipmentStats)) return sub.equipmentStats;
    return lastHwaseungDoublePipeLog && lastHwaseungDoublePipeLog.equipmentStats
      ? lastHwaseungDoublePipeLog.equipmentStats
      : [];
  }

  function renderHwaseungDoublePipeLogTables() {
    if (!hwaseungDoublePipeLogTablesWrap || !lastHwaseungDoublePipeLog) return;
    hwaseungDoublePipeLogTablesWrap.innerHTML = "";
    const sub = reaggregateHwaseungDoublePipeLogFiltered();
    const slicerActive = !!sub;
    const months = slicerActive ? [] : lastHwaseungDoublePipeLog.monthlyStats || [];
    const eqs = getHwaseungDoublePipeLogEquipmentStatsForTables();
    const procBlocks =
      sub && Array.isArray(sub.processBlocks) ? sub.processBlocks : lastHwaseungDoublePipeLog.processBlocks || [];
    const splitStop =
      Array.isArray(lastHwaseungDoublePipeLog.iStopTimeCols) && lastHwaseungDoublePipeLog.iStopTimeCols.length >= 4;

    const procHasRows = procBlocks.some(
      (pb) => (pb.monthlyStats && pb.monthlyStats.length) || (pb.equipmentStats && pb.equipmentStats.length)
    );
    if (months.length === 0 && eqs.length === 0 && !procHasRows) {
      hwaseungDoublePipeLogTablesWrap.hidden = false;
      const p = document.createElement("p");
      p.className = "empty-state empty-state--flat";
      p.textContent = slicerActive
        ? "선택한 기간에 표시할 집계가 없습니다."
        : "표시할 집계가 없습니다.";
      hwaseungDoublePipeLogTablesWrap.appendChild(p);
      return;
    }
    const stopLabels = splitStop ? HWASEUNG_STOP_LABELS : undefined;
    if (months.length) {
      hwaseungDoublePipeLogTablesWrap.appendChild(
        buildDrawingBreakdownTable("월별 집계", "월", months, splitStop, stopLabels)
      );
    }
    if (eqs.length) {
      let eqTitle = lastHwaseungDoublePipeLog.hasEquipmentColumn ? "설비별 집계" : "설비별 집계 (설비 열 없음 · 전체)";
      if (slicerActive) {
        const hint =
          hwaseungDoublePipeLogGranularity === "month"
            ? [...hwaseungDoublePipeLogSelectedMonths].sort().join(", ")
            : [...hwaseungDoublePipeLogSelectedDays].sort().join(", ");
        eqTitle += ` · 선택: ${hint}`;
      }
      hwaseungDoublePipeLogTablesWrap.appendChild(
        buildDrawingBreakdownTable(eqTitle, "설비·구분", eqs, splitStop, stopLabels)
      );
    }
    for (const blk of procBlocks) {
      const m = blk.monthlyStats || [];
      const e = blk.equipmentStats || [];
      if (!m.length && !e.length) continue;
      const sec = document.createElement("div");
      sec.className = "drawing-log-block";
      const h3 = document.createElement("h3");
      h3.className = "drawing-log-block__title";
      h3.textContent = `공정 · ${blk.process}`;
      sec.appendChild(h3);
      if (m.length) sec.appendChild(buildDrawingBreakdownTable("월별 집계", "월", m, splitStop, stopLabels));
      if (e.length) {
        let eqTitle = lastHwaseungDoublePipeLog.hasEquipmentColumn ? "설비별 집계" : "설비별 집계 (설비 열 없음 · 전체)";
        if (slicerActive) {
          const hint =
            hwaseungDoublePipeLogGranularity === "month"
              ? [...hwaseungDoublePipeLogSelectedMonths].sort().join(", ")
              : [...hwaseungDoublePipeLogSelectedDays].sort().join(", ");
          eqTitle += ` · 선택: ${hint}`;
        }
        sec.appendChild(buildDrawingBreakdownTable(eqTitle, "설비·구분", e, splitStop, stopLabels));
      }
      hwaseungDoublePipeLogTablesWrap.appendChild(sec);
    }
    hwaseungDoublePipeLogTablesWrap.hidden = false;
  }

  function renderHwaseungDoublePipeLogProductionQtyTable() {
    if (!hwaseungDoublePipeLogProdWrap || !lastHwaseungDoublePipeLog) return;
    hwaseungDoublePipeLogProdWrap.innerHTML = "";
    hwaseungDoublePipeLogProdWrap.hidden = true;
    const subHw = reaggregateHwaseungDoublePipeLogFiltered();
    if (!subHw) return;

    const monthMode = hwaseungDoublePipeLogGranularity === "month" && hwaseungDoublePipeLogSelectedMonths.size > 0;
    const dayMode = hwaseungDoublePipeLogGranularity === "day" && hwaseungDoublePipeLogSelectedDays.size > 0;
    if (!monthMode && !dayMode) return;

    const monthPool = subHw.pivotMonthlyStats || [];
    const dayPool = subHw.pivotDailyStats || [];
    let statRows = [];
    if (monthMode) {
      for (const ym of [...hwaseungDoublePipeLogSelectedMonths].sort()) {
        const st = monthPool.find((m) => m.key === ym);
        if (st) statRows.push(st);
      }
    } else {
      for (const dk of [...hwaseungDoublePipeLogSelectedDays].sort()) {
        const st = dayPool.find((d) => d.key === dk);
        if (st) statRows.push(st);
      }
    }
    if (!statRows.length) return;

    const hasProdCol = lastHwaseungDoublePipeLog.iProdQty >= 0;
    hwaseungDoublePipeLogProdWrap.hidden = false;
    const block = document.createElement("div");
    block.className = "drawing-log-block";
    const h3 = document.createElement("h3");
    h3.className = "drawing-log-block__title";
    h3.textContent = monthMode ? "선택 월 · 생산수량" : "선택 일 · 생산수량";
    block.appendChild(h3);
    if (!hasProdCol) {
      const note = document.createElement("p");
      note.className = "drawing-log-meta";
      note.textContent =
        "시트에 생산량 열(L열 또는 헤더 인식)이 없으면 생산수량 표가 비어 있습니다. 엑셀을 확인해 주세요.";
      block.appendChild(note);
    }
    const wrap = document.createElement("div");
    wrap.className = "drawing-log-preview-wrap drawing-log-prod-pivot-wrap";
    const tbl = document.createElement("table");
    tbl.className = "drawing-log-prod-pivot";

    const thead = document.createElement("thead");
    const trh = document.createElement("tr");
    const thCorner = document.createElement("th");
    thCorner.textContent = "";
    thCorner.classList.add("drawing-log-prod-pivot__corner");
    thCorner.setAttribute("aria-label", monthMode ? "월 열" : "날짜 열");
    trh.appendChild(thCorner);
    statRows.forEach((st) => {
      const th = document.createElement("th");
      th.textContent = String(st.label || st.key);
      th.classList.add("num");
      th.title = String(st.key || "");
      trh.appendChild(th);
    });
    const thTot = document.createElement("th");
    thTot.textContent = "합계";
    thTot.classList.add("num");
    trh.appendChild(thTot);
    thead.appendChild(trh);
    tbl.appendChild(thead);

    const tbody = document.createElement("tbody");
    function addPivotRow(label, getter, opts) {
      const asInt = opts && opts.asInt;
      const tr = document.createElement("tr");
      const td0 = document.createElement("td");
      td0.textContent = label;
      td0.classList.add("drawing-log-prod-pivot__rowhead");
      tr.appendChild(td0);
      let sum = 0;
      let n = 0;
      for (const st of statRows) {
        const v = getter(st);
        const td = document.createElement("td");
        td.classList.add("num");
        if (Number.isFinite(v)) {
          td.textContent = asInt ? String(Math.round(v)) : formatQty(v);
          sum += v;
          n++;
        } else {
          td.textContent = "—";
        }
        tr.appendChild(td);
      }
      const tdSum = document.createElement("td");
      tdSum.classList.add("num");
      tdSum.style.fontWeight = "700";
      if (n && Number.isFinite(sum)) {
        tdSum.textContent = asInt ? String(Math.round(sum)) : formatQty(sum);
      } else {
        tdSum.textContent = "—";
      }
      tr.appendChild(tdSum);
      tbody.appendChild(tr);
    }

    addPivotRow("생산수량 합", (st) => (hasProdCol && Number.isFinite(st.sumProdQty) ? st.sumProdQty : NaN));
    addPivotRow("건수", (st) => (Number.isFinite(st.dataRows) ? st.dataRows : NaN), { asInt: true });
    addPivotRow("작업시간 합", (st) => (Number.isFinite(st.sumWorkTime) ? st.sumWorkTime : NaN));

    tbl.appendChild(tbody);
    wrap.appendChild(tbl);
    block.appendChild(wrap);
    hwaseungDoublePipeLogProdWrap.appendChild(block);
  }

  function hwaseungDoublePipeLogSlicerSummaryText() {
    const sub = reaggregateHwaseungDoublePipeLogFiltered();
    const procFrag = hwaseungDoublePipeLogProcessFilterSummaryFragment();
    if (hwaseungDoublePipeLogGranularity === "month" && hwaseungDoublePipeLogSelectedMonths.size && sub) {
      const parts = [...hwaseungDoublePipeLogSelectedMonths].sort().map((ym) => {
        const p = ym.split("-");
        return p.length === 2 ? `${p[0]}년 ${parseInt(p[1], 10)}월` : ym;
      });
      let t = `${parts.join(", ")} · ${sub.dataRows}행`;
      if (procFrag) t += ` · ${procFrag}`;
      return t;
    }
    if (hwaseungDoublePipeLogGranularity === "day" && hwaseungDoublePipeLogSelectedDays.size && sub) {
      let t = `${[...hwaseungDoublePipeLogSelectedDays].sort().join(", ")} · ${sub.dataRows}행`;
      if (procFrag) t += ` · ${procFrag}`;
      return t;
    }
    if (procFrag && sub) return `${procFrag} · ${sub.dataRows}행`;
    return "";
  }

  function pruneHwaseungDoublePipeLogSelectedDaysInvalid() {
    if (!lastHwaseungDoublePipeLog || hwaseungDoublePipeLogGranularity !== "day") return;
    const pool = new Set(
      (lastHwaseungDoublePipeLog.dailyStats || [])
        .filter((d) => {
          if (d.key === "일자미상") return false;
          if (!d.key.startsWith(`${hwaseungDoublePipeLogTimelineYear}-`)) return false;
          if (hwaseungDoublePipeLogSelectedMonths.size) {
            const ym = d.key.slice(0, 7);
            return hwaseungDoublePipeLogSelectedMonths.has(ym);
          }
          return true;
        })
        .map((d) => d.key)
    );
    [...hwaseungDoublePipeLogSelectedDays].forEach((k) => {
      if (!pool.has(k)) hwaseungDoublePipeLogSelectedDays.delete(k);
    });
  }

  function renderHwaseungDoublePipeLogTimeline() {
    if (!hwaseungDoublePipeLogTimeline || !lastHwaseungDoublePipeLog || !lastHwaseungDoublePipeLog.monthlyStats) {
      if (hwaseungDoublePipeLogTimeline) hwaseungDoublePipeLogTimeline.hidden = true;
      return;
    }
    hwaseungDoublePipeLogTimeline.hidden = false;
    if (hwaseungDoublePipeLogGranularityEl) hwaseungDoublePipeLogGranularityEl.value = hwaseungDoublePipeLogGranularity;
    if (hwaseungDoublePipeLogKpiScopeEl)
      hwaseungDoublePipeLogKpiScopeEl.value =
        hwaseungDoublePipeLogKpiScope === "overall" ? "overall" : "process";

    const ys = doublePipeLogYearsFromStats(lastHwaseungDoublePipeLog);
    if (ys.length) {
      if (hwaseungDoublePipeLogTimelineYear < ys[0]) hwaseungDoublePipeLogTimelineYear = ys[0];
      if (hwaseungDoublePipeLogTimelineYear > ys[ys.length - 1]) hwaseungDoublePipeLogTimelineYear = ys[ys.length - 1];
    }
    if (hwaseungDoublePipeLogYearLabel) hwaseungDoublePipeLogYearLabel.textContent = `${hwaseungDoublePipeLogTimelineYear}년`;
    if (hwaseungDoublePipeLogYearPrev)
      hwaseungDoublePipeLogYearPrev.disabled = ys.length && hwaseungDoublePipeLogTimelineYear <= ys[0];
    if (hwaseungDoublePipeLogYearNext)
      hwaseungDoublePipeLogYearNext.disabled =
        ys.length && hwaseungDoublePipeLogTimelineYear >= ys[ys.length - 1];

    const monthKeys = new Set((lastHwaseungDoublePipeLog.monthlyStats || []).map((m) => m.key));
    if (hwaseungDoublePipeLogMonthStrip) {
      hwaseungDoublePipeLogMonthStrip.innerHTML = "";
      for (let m = 1; m <= 12; m++) {
        const ym = `${hwaseungDoublePipeLogTimelineYear}-${String(m).padStart(2, "0")}`;
        const btn = document.createElement("button");
        btn.type = "button";
        btn.className = "drawing-log-timeline__chip";
        btn.dataset.ym = ym;
        btn.textContent = String(m);
        const has = monthKeys.has(ym);
        if (!has) {
          btn.classList.add("is-empty");
          btn.disabled = true;
        }
        if (hwaseungDoublePipeLogSelectedMonths.has(ym)) btn.classList.add("is-selected");
        hwaseungDoublePipeLogMonthStrip.appendChild(btn);
      }
    }

    const dayMode = hwaseungDoublePipeLogGranularity === "day";
    if (hwaseungDoublePipeLogDayStrip) {
      hwaseungDoublePipeLogDayStrip.hidden = !dayMode;
      hwaseungDoublePipeLogDayStrip.innerHTML = "";
      if (dayMode) {
        const pool = (lastHwaseungDoublePipeLog.dailyStats || []).filter((d) => {
          if (d.key === "일자미상") return false;
          if (!d.key.startsWith(`${hwaseungDoublePipeLogTimelineYear}-`)) return false;
          if (hwaseungDoublePipeLogSelectedMonths.size) {
            const ym = d.key.length >= 7 ? d.key.slice(0, 7) : "";
            return hwaseungDoublePipeLogSelectedMonths.has(ym);
          }
          return true;
        });
        pool.forEach((d) => {
          const btn = document.createElement("button");
          btn.type = "button";
          btn.className = "drawing-log-timeline__chip";
          btn.dataset.dayKey = d.key;
          btn.textContent = d.key.slice(5);
          if (hwaseungDoublePipeLogSelectedDays.has(d.key)) btn.classList.add("is-selected");
          hwaseungDoublePipeLogDayStrip.appendChild(btn);
        });
      }
    }

    if (hwaseungDoublePipeLogSlicerSelection) {
      const t = hwaseungDoublePipeLogSlicerSummaryText();
      hwaseungDoublePipeLogSlicerSelection.textContent = t || "전체 기간 (필터 없음)";
    }
  }

  function getHwaseungDoublePipeLogKpiSource() {
    if (!lastHwaseungDoublePipeLog) return null;
    const sub = reaggregateHwaseungDoublePipeLogFiltered({
      ignoreProcessForRowFilter: hwaseungDoublePipeLogKpiScope === "overall",
    });
    if (!sub) return lastHwaseungDoublePipeLog;
    return { ...lastHwaseungDoublePipeLog, ...sub };
  }

  function renderHwaseungDoublePipeLogPanel() {
    if (!lastHwaseungDoublePipeLog || !hwaseungDoublePipeLogContent) {
      clearHwaseungDoublePipeLogUi();
      return;
    }
    readHwaseungDoublePipeLogProcessSelectionFromDom();
    readHwaseungDoublePipeLogKpiScopeFromDom();
    const ksrc = getHwaseungDoublePipeLogKpiSource();
    if (hwaseungDoublePipeLogEmpty) hwaseungDoublePipeLogEmpty.hidden = true;
    hwaseungDoublePipeLogContent.hidden = false;
    if (kpiUtilHwaseungDoublePipe) kpiUtilHwaseungDoublePipe.textContent = formatPercent(ksrc.utilAvg ?? NaN);
    if (kpiOeeHwaseungDoublePipe) kpiOeeHwaseungDoublePipe.textContent = formatPercent(ksrc.oeeAvg ?? NaN);
    if (kpiDefectLabelHwaseungDoublePipe && kpiDefectHwaseungDoublePipe) {
      if (Number.isFinite(ksrc.defectAvg)) {
        kpiDefectLabelHwaseungDoublePipe.textContent = "불량율";
        kpiDefectHwaseungDoublePipe.textContent = formatPercent(ksrc.defectAvg);
      } else if (Number.isFinite(ksrc.sumDefectQty)) {
        kpiDefectLabelHwaseungDoublePipe.textContent = "불량(합)";
        kpiDefectHwaseungDoublePipe.textContent = `${formatQty(ksrc.sumDefectQty)}건`;
      } else {
        kpiDefectLabelHwaseungDoublePipe.textContent = "불량율";
        kpiDefectHwaseungDoublePipe.textContent = "—";
      }
    }
    if (hwaseungDoublePipeLogMeta) {
      const dr =
        lastHwaseungDoublePipeLog.dateFrom && lastHwaseungDoublePipeLog.dateTo
          ? ` · 작업일자 ${lastHwaseungDoublePipeLog.dateFrom} ~ ${lastHwaseungDoublePipeLog.dateTo}`
          : "";
      const kpiNote =
        ksrc.utilIsDerived || ksrc.oeeIsDerived
          ? " · 가동율·설비효율: 열이 없을 때 작업시간·정지·투입·생산수량·불량(율)·생산성으로 행별 산출 후 평균"
          : "";
      const fn = reaggregateHwaseungDoublePipeLogFiltered();
      const dayNote = fn
        ? ` · 보기: 필터 적용 (${fn.dataRows}행 / 전체 ${lastHwaseungDoublePipeLog.dataRows}행)`
        : "";
      hwaseungDoublePipeLogMeta.textContent =
        `파일: ${lastHwaseungDoublePipeLog.fileLabel} · 시트 「${lastHwaseungDoublePipeLog.sheetName}」 · 유효 ${lastHwaseungDoublePipeLog.dataRows}행 · 생산 L열 · 비가동 R~V · 불량 소재(X+Y+Z)+공정(AK+AL+AR) — W·AJ 합계 열은 제외(이중 계산 방지)${dr}${dayNote}${kpiNote}`;
    }
    if (hwaseungDoublePipeLogOps) {
      const bits = [];
      if (Number.isFinite(ksrc.sumProdQty)) bits.push(`생산수량 합 ${formatQty(ksrc.sumProdQty)}`);
      if (Number.isFinite(ksrc.sumWorkTime)) bits.push(`작업시간 합 ${formatQty(ksrc.sumWorkTime)}`);
      if (Number.isFinite(ksrc.sumInputTime)) bits.push(`투입시간 합 ${formatQty(ksrc.sumInputTime)}`);
      if (Number.isFinite(ksrc.productivityAvg)) bits.push(`생산성 평균 ${formatPercent(ksrc.productivityAvg)}`);
      if (Number.isFinite(ksrc.sumDefectQty)) bits.push(`불량 합 ${formatQty(ksrc.sumDefectQty)}건`);
      const stopSplit =
        Array.isArray(lastHwaseungDoublePipeLog.iStopTimeCols) && lastHwaseungDoublePipeLog.iStopTimeCols.length >= 4;
      if (stopSplit) {
        if (Number.isFinite(ksrc.sumStopExchange)) bits.push(`R ${formatQty(ksrc.sumStopExchange)}`);
        if (Number.isFinite(ksrc.sumStopRepair)) bits.push(`S ${formatQty(ksrc.sumStopRepair)}`);
        if (Number.isFinite(ksrc.sumStopMaterial)) bits.push(`T ${formatQty(ksrc.sumStopMaterial)}`);
        if (Number.isFinite(ksrc.sumStopPlanned)) bits.push(`U ${formatQty(ksrc.sumStopPlanned)}`);
        if (lastHwaseungDoublePipeLog.iStopTimeCols.length === 5 && Number.isFinite(ksrc.sumStopFifth))
          bits.push(`V ${formatQty(ksrc.sumStopFifth)}`);
        if (Number.isFinite(ksrc.sumStopTime)) bits.push(`비가동 합 ${formatQty(ksrc.sumStopTime)}`);
      } else if (Number.isFinite(ksrc.sumStopTime)) {
        bits.push(`정지 시간 합 ${formatQty(ksrc.sumStopTime)}`);
      }
      hwaseungDoublePipeLogOps.textContent = bits.join(" · ");
      hwaseungDoublePipeLogOps.hidden = bits.length === 0;
    }
    renderHwaseungDoublePipeLogTimeline();
    renderHwaseungDoublePipeLogTables();
    renderHwaseungDoublePipeLogProductionQtyTable();
    if (hwaseungDoublePipeLogMaintWrap) {
      if (lastHwaseungDoublePipeLog.maintenance) {
        hwaseungDoublePipeLogMaintWrap.hidden = false;
        renderWorkbookMaintenanceInto(hwaseungDoublePipeLogMaintWrap, lastHwaseungDoublePipeLog.maintenance);
      } else {
        hwaseungDoublePipeLogMaintWrap.innerHTML = "";
        hwaseungDoublePipeLogMaintWrap.hidden = true;
      }
    }
  }

  function defaultMufflerSegUiState() {
    return {
      granularity: "month",
      selectedMonths: new Set(),
      selectedDays: new Set(),
      timelineYear: new Date().getFullYear(),
      processCut: true,
      processForming: true,
      processMachine: true,
      kpiScope: "process",
    };
  }

  function resetMufflerLogSegStatesFromSegments() {
    mufflerLogSegStates = (lastMufflerLog && lastMufflerLog.segments ? lastMufflerLog.segments : []).map(() =>
      defaultMufflerSegUiState()
    );
  }

  function getMufflerSegState(segIdx) {
    return mufflerLogSegStates[segIdx] || defaultMufflerSegUiState();
  }

  /** @returns {null | (ReturnType<typeof parseMufflerLogFromMatrix> & { sheetName: string })} */
  function mufflerSegmentLog(segIdx) {
    if (!lastMufflerLog || !lastMufflerLog.segments) return null;
    return lastMufflerLog.segments[segIdx] || null;
  }

  function initMufflerLogSlicerFromData() {
    resetMufflerLogSegStatesFromSegments();
    if (!lastMufflerLog || !lastMufflerLog.segments) return;
    lastMufflerLog.segments.forEach((seg, i) => {
      const ys = mufflerLogYearsFromStats(seg);
      if (ys.length) mufflerLogSegStates[i].timelineYear = ys[ys.length - 1];
      else if (seg.dateTo && String(seg.dateTo).length >= 4) {
        const y = parseInt(String(seg.dateTo).slice(0, 4), 10);
        if (Number.isFinite(y)) mufflerLogSegStates[i].timelineYear = y;
      }
    });
  }

  function mufflerLogYearsFromStats(log) {
    if (!log) return [];
    const ys = new Set();
    (log.monthlyStats || []).forEach((m) => {
      if (!m || m.key === "일자미상") return;
      const y = parseInt(String(m.key).slice(0, 4), 10);
      if (Number.isFinite(y)) ys.add(y);
    });
    return [...ys].sort((a, b) => a - b);
  }

  function mufflerLogHasProcessTristateFilter(segIdx) {
    const st = getMufflerSegState(segIdx);
    if (!st.processCut && !st.processForming && !st.processMachine) return false;
    if (st.processCut && st.processForming && st.processMachine) return false;
    return true;
  }

  function mufflerLogRowMatchesProcessFilter(segIdx, row, iProcess) {
    if (!mufflerLogHasProcessTristateFilter(segIdx)) return true;
    const st = getMufflerSegState(segIdx);
    const pk = normalizeMufflerProcess(iProcess >= 0 ? row[iProcess] : "");
    if (pk === "절단") return st.processCut;
    if (pk === "포밍") return st.processForming;
    if (pk === "가공") return st.processMachine;
    return true;
  }

  function mufflerLogProcessFilterSummaryFragment(segIdx) {
    if (!mufflerLogHasProcessTristateFilter(segIdx)) return "";
    const st = getMufflerSegState(segIdx);
    const bits = [];
    if (st.processCut) bits.push("절단");
    if (st.processForming) bits.push("포밍");
    if (st.processMachine) bits.push("가공");
    return bits.length ? `공정: ${bits.join("+")}` : "";
  }

  function reaggregateMufflerLogFiltered(segIdx, options = {}) {
    const ignoreProcessForRowFilter = options.ignoreProcessForRowFilter === true;
    const log = mufflerSegmentLog(segIdx);
    const st = getMufflerSegState(segIdx);
    if (!log || !log._filterMatrix) return null;
    const hasMonth = st.granularity === "month" && st.selectedMonths.size > 0;
    const hasDay = st.granularity === "day" && st.selectedDays.size > 0;
    const procRestricts = mufflerLogHasProcessTristateFilter(segIdx);
    const hasProcF = procRestricts && !ignoreProcessForRowFilter;
    if (!hasMonth && !hasDay && !hasProcF) return null;

    const matrix = log._filterMatrix;
    const hi = log._filterHi;
    const idx = log._filterIdx;
    const iWorkDate = idx.iWorkDate;
    const iEq = log.iEquipment;
    const iProcess = log.iProcess;

    const rowIncluded = (row) => {
      if (!row || !row.length) return false;
      if (!row.some((c) => String(c ?? "").trim() !== "")) return false;
      if (!ignoreProcessForRowFilter && !mufflerLogRowMatchesProcessFilter(segIdx, row, iProcess)) return false;
      const ymd = iWorkDate >= 0 ? parseExcelDate(row[iWorkDate]) : "";
      const ym = ymd && ymd.length >= 7 ? ymd.slice(0, 7) : "일자미상";
      const dayKey = ymd || "일자미상";
      if (hasDay) return st.selectedDays.has(dayKey);
      if (hasMonth) return st.selectedMonths.has(ym);
      if (hasProcF) return true;
      return false;
    };

    const aggGlobal = makeDrawingAgg();
    const byEq = new Map();
    const byProcMonth = new Map();
    const byProcEq = new Map();
    const byMonthPivot = new Map();
    const byDayPivot = new Map();

    for (let r = hi + 1; r < matrix.length; r++) {
      const row = matrix[r];
      if (!rowIncluded(row)) continue;
      addDrawingRowToAgg(aggGlobal, row, idx, false);
      const ymdPV = iWorkDate >= 0 ? parseExcelDate(row[iWorkDate]) : "";
      const ymPV = ymdPV && ymdPV.length >= 7 ? ymdPV.slice(0, 7) : "일자미상";
      const dayKeyPV = ymdPV || "일자미상";
      if (!byMonthPivot.has(ymPV)) byMonthPivot.set(ymPV, makeDrawingAgg());
      addDrawingRowToAgg(byMonthPivot.get(ymPV), row, idx, false);
      if (!byDayPivot.has(dayKeyPV)) byDayPivot.set(dayKeyPV, makeDrawingAgg());
      addDrawingRowToAgg(byDayPivot.get(dayKeyPV), row, idx, false);
      const eqKey = iEq >= 0 ? String(row[iEq] ?? "").trim() || "미지정" : "전체";
      if (!byEq.has(eqKey)) byEq.set(eqKey, makeDrawingAgg());
      addDrawingRowToAgg(byEq.get(eqKey), row, idx, false);
      const ymd = iWorkDate >= 0 ? parseExcelDate(row[iWorkDate]) : "";
      const ym = ymd && ymd.length >= 7 ? ymd.slice(0, 7) : "일자미상";
      const procKey = normalizeMufflerProcess(iProcess >= 0 ? row[iProcess] : "");
      addDrawingRowToAgg(ensureAggMapNested2(byProcMonth, procKey, ym), row, idx, false);
      addDrawingRowToAgg(ensureAggMapNested2(byProcEq, procKey, eqKey), row, idx, false);
    }
    const g = finalizeDrawingAgg(aggGlobal, idx);
    const equipmentStats = [...byEq.entries()]
      .sort((a, b) => compareDrawingEquipmentKeys(a[0], b[0]))
      .map(([key, agg]) => ({ key, label: key, ...finalizeDrawingAgg(agg, idx) }));

    const { pivotMonthlyStats, pivotDailyStats } = buildPivotMonthlyDailyStatsFromAggMaps(byMonthPivot, byDayPivot, idx);

    const monthLabel = (ymKey) => {
      if (ymKey === "일자미상") return ymKey;
      const p = ymKey.split("-");
      if (p.length !== 2) return ymKey;
      return `${p[0]}년 ${parseInt(p[1], 10)}월`;
    };
    const sortProcKeys = (a, b) => {
      const rank = (k) => (k === "절단" ? 0 : k === "포밍" ? 1 : k === "가공" ? 2 : k === "기타" ? 3 : 99);
      const ra = rank(a);
      const rb = rank(b);
      if (ra !== rb) return ra - rb;
      return String(a).localeCompare(String(b));
    };
    const procKeySet = new Set([...byProcMonth.keys(), ...byProcEq.keys()]);
    const processBlocks = [...procKeySet].sort(sortProcKeys).map((pk) => {
      const pm = byProcMonth.get(pk) || new Map();
      const peq = byProcEq.get(pk) || new Map();
      const monthlyStatsP = [...pm.entries()]
        .sort((a, b) => {
          if (a[0] === "일자미상") return 1;
          if (b[0] === "일자미상") return -1;
          return a[0].localeCompare(b[0]);
        })
        .map(([key, agg]) => ({ key, label: monthLabel(key), ...finalizeDrawingAgg(agg, idx) }));
      const equipmentStatsP = [...peq.entries()]
        .sort((a, b) => compareDrawingEquipmentKeys(a[0], b[0]))
        .map(([key, agg]) => ({ key, label: key, ...finalizeDrawingAgg(agg, idx) }));
      return {
        process: pk,
        monthlyStats: monthlyStatsP,
        equipmentStats: equipmentStatsP,
        dailyStats: [],
        dayEquipmentStats: {},
      };
    });

    return { ...g, equipmentStats, dataRows: aggGlobal.dataRows, processBlocks, pivotMonthlyStats, pivotDailyStats };
  }

  function getMufflerLogEquipmentStatsForTables(segIdx) {
    const sub = reaggregateMufflerLogFiltered(segIdx);
    const log = mufflerSegmentLog(segIdx);
    if (sub && Array.isArray(sub.equipmentStats)) return sub.equipmentStats;
    return log && log.equipmentStats ? log.equipmentStats : [];
  }

  /** @param {HTMLElement} tablesWrap */
  function renderMufflerLogTablesInto(tablesWrap, segIdx) {
    const log = mufflerSegmentLog(segIdx);
    if (!tablesWrap || !log) return;
    tablesWrap.innerHTML = "";
    const st = getMufflerSegState(segIdx);
    const sub = reaggregateMufflerLogFiltered(segIdx);
    const slicerActive = !!sub;
    const months = slicerActive ? [] : log.monthlyStats || [];
    const eqs = getMufflerLogEquipmentStatsForTables(segIdx);
    const procBlocks =
      sub && Array.isArray(sub.processBlocks) ? sub.processBlocks : log.processBlocks || [];
    const splitStop = Array.isArray(log.iStopTimeCols) && log.iStopTimeCols.length >= 4;

    const procHasRows = procBlocks.some(
      (pb) => (pb.monthlyStats && pb.monthlyStats.length) || (pb.equipmentStats && pb.equipmentStats.length)
    );
    if (months.length === 0 && eqs.length === 0 && !procHasRows) {
      tablesWrap.hidden = false;
      const p = document.createElement("p");
      p.className = "empty-state empty-state--flat";
      p.textContent = slicerActive ? "선택한 기간에 표시할 집계가 없습니다." : "표시할 집계가 없습니다.";
      tablesWrap.appendChild(p);
      return;
    }
    if (months.length) {
      tablesWrap.appendChild(buildDrawingBreakdownTable("월별 집계", "월", months, splitStop));
    }
    if (eqs.length) {
      let eqTitle = log.hasEquipmentColumn ? "설비별 집계" : "설비별 집계 (설비 열 없음 · 전체)";
      if (slicerActive) {
        const hint =
          st.granularity === "month"
            ? [...st.selectedMonths].sort().join(", ")
            : [...st.selectedDays].sort().join(", ");
        eqTitle += ` · 선택: ${hint}`;
      }
      tablesWrap.appendChild(buildDrawingBreakdownTable(eqTitle, "설비·구분", eqs, splitStop));
    }
    for (const blk of procBlocks) {
      const m = blk.monthlyStats || [];
      const e = blk.equipmentStats || [];
      if (!m.length && !e.length) continue;
      const sec = document.createElement("div");
      sec.className = "drawing-log-block";
      const h3 = document.createElement("h3");
      h3.className = "drawing-log-block__title";
      h3.textContent = `공정 · ${blk.process}`;
      sec.appendChild(h3);
      if (m.length) sec.appendChild(buildDrawingBreakdownTable("월별 집계", "월", m, splitStop));
      if (e.length) {
        let eqTitle = log.hasEquipmentColumn ? "설비별 집계" : "설비별 집계 (설비 열 없음 · 전체)";
        if (slicerActive) {
          const hint =
            st.granularity === "month"
              ? [...st.selectedMonths].sort().join(", ")
              : [...st.selectedDays].sort().join(", ");
          eqTitle += ` · 선택: ${hint}`;
        }
        sec.appendChild(buildDrawingBreakdownTable(eqTitle, "설비·구분", e, splitStop));
      }
      tablesWrap.appendChild(sec);
    }
    tablesWrap.hidden = false;
  }

  /** @param {HTMLElement} prodWrap */
  function renderMufflerLogProductionQtyInto(prodWrap, segIdx) {
    const log = mufflerSegmentLog(segIdx);
    const st = getMufflerSegState(segIdx);
    if (!prodWrap || !log) return;
    prodWrap.innerHTML = "";
    prodWrap.hidden = true;
    const subMu = reaggregateMufflerLogFiltered(segIdx);
    if (!subMu) return;

    const monthMode = st.granularity === "month" && st.selectedMonths.size > 0;
    const dayMode = st.granularity === "day" && st.selectedDays.size > 0;
    if (!monthMode && !dayMode) return;

    const monthPool = subMu.pivotMonthlyStats || [];
    const dayPool = subMu.pivotDailyStats || [];
    /** @type {any[]} */
    let statRows = [];
    if (monthMode) {
      for (const ym of [...st.selectedMonths].sort()) {
        const row = monthPool.find((m) => m.key === ym);
        if (row) statRows.push(row);
      }
    } else {
      for (const dk of [...st.selectedDays].sort()) {
        const row = dayPool.find((d) => d.key === dk);
        if (row) statRows.push(row);
      }
    }
    if (!statRows.length) return;

    const hasProdCol = log.iProdQty >= 0;
    prodWrap.hidden = false;
    const block = document.createElement("div");
    block.className = "drawing-log-block";
    const h3 = document.createElement("h3");
    h3.className = "drawing-log-block__title";
    h3.textContent = monthMode ? "선택 월 · 생산수량" : "선택 일 · 생산수량";
    block.appendChild(h3);
    if (!hasProdCol) {
      const note = document.createElement("p");
      note.className = "drawing-log-meta";
      note.textContent =
        "시트에 생산수량·생산량 열이 인식되지 않으면 생산수량 행은 표시되지 않습니다. 엑셀 헤더를 확인해 주세요.";
      block.appendChild(note);
    }
    const wrap = document.createElement("div");
    wrap.className = "drawing-log-preview-wrap drawing-log-prod-pivot-wrap";
    const tbl = document.createElement("table");
    tbl.className = "drawing-log-prod-pivot";

    const thead = document.createElement("thead");
    const trh = document.createElement("tr");
    const thCorner = document.createElement("th");
    thCorner.textContent = "";
    thCorner.classList.add("drawing-log-prod-pivot__corner");
    thCorner.setAttribute("aria-label", monthMode ? "월 열" : "날짜 열");
    trh.appendChild(thCorner);
    statRows.forEach((row) => {
      const th = document.createElement("th");
      th.textContent = String(row.label || row.key);
      th.classList.add("num");
      th.title = String(row.key || "");
      trh.appendChild(th);
    });
    const thTot = document.createElement("th");
    thTot.textContent = "합계";
    thTot.classList.add("num");
    trh.appendChild(thTot);
    thead.appendChild(trh);
    tbl.appendChild(thead);

    const tbody = document.createElement("tbody");
    function addPivotRow(label, getter, opts) {
      const asInt = opts && opts.asInt;
      const tr = document.createElement("tr");
      const td0 = document.createElement("td");
      td0.textContent = label;
      td0.classList.add("drawing-log-prod-pivot__rowhead");
      tr.appendChild(td0);
      let sum = 0;
      let n = 0;
      for (const row of statRows) {
        const v = getter(row);
        const td = document.createElement("td");
        td.classList.add("num");
        if (Number.isFinite(v)) {
          td.textContent = asInt ? String(Math.round(v)) : formatQty(v);
          sum += v;
          n++;
        } else {
          td.textContent = "—";
        }
        tr.appendChild(td);
      }
      const tdSum = document.createElement("td");
      tdSum.classList.add("num");
      tdSum.style.fontWeight = "700";
      if (n && Number.isFinite(sum)) {
        tdSum.textContent = asInt ? String(Math.round(sum)) : formatQty(sum);
      } else {
        tdSum.textContent = "—";
      }
      tr.appendChild(tdSum);
      tbody.appendChild(tr);
    }

    addPivotRow("생산수량 합", (row) => (hasProdCol && Number.isFinite(row.sumProdQty) ? row.sumProdQty : NaN));
    addPivotRow("건수", (row) => (Number.isFinite(row.dataRows) ? row.dataRows : NaN), { asInt: true });
    addPivotRow("작업시간 합", (row) => (Number.isFinite(row.sumWorkTime) ? row.sumWorkTime : NaN));

    tbl.appendChild(tbody);
    wrap.appendChild(tbl);
    block.appendChild(wrap);
    prodWrap.appendChild(block);
  }

  function mufflerLogSlicerSummaryText(segIdx) {
    const sub = reaggregateMufflerLogFiltered(segIdx);
    const st = getMufflerSegState(segIdx);
    const procFrag = mufflerLogProcessFilterSummaryFragment(segIdx);
    if (st.granularity === "month" && st.selectedMonths.size && sub) {
      const parts = [...st.selectedMonths].sort().map((ym) => {
        const p = ym.split("-");
        return p.length === 2 ? `${p[0]}년 ${parseInt(p[1], 10)}월` : ym;
      });
      let t = `${parts.join(", ")} · ${sub.dataRows}행`;
      if (procFrag) t += ` · ${procFrag}`;
      return t;
    }
    if (st.granularity === "day" && st.selectedDays.size && sub) {
      let t = `${[...st.selectedDays].sort().join(", ")} · ${sub.dataRows}행`;
      if (procFrag) t += ` · ${procFrag}`;
      return t;
    }
    if (procFrag && sub) return `${procFrag} · ${sub.dataRows}행`;
    return "";
  }

  function pruneMufflerLogSelectedDaysInvalid(segIdx) {
    const log = mufflerSegmentLog(segIdx);
    const st = getMufflerSegState(segIdx);
    if (!log || st.granularity !== "day") return;
    const pool = new Set(
      (log.dailyStats || [])
        .filter((d) => {
          if (d.key === "일자미상") return false;
          if (!d.key.startsWith(`${st.timelineYear}-`)) return false;
          if (st.selectedMonths.size) {
            const ym = d.key.slice(0, 7);
            return st.selectedMonths.has(ym);
          }
          return true;
        })
        .map((d) => d.key)
    );
    [...st.selectedDays].forEach((k) => {
      if (!pool.has(k)) st.selectedDays.delete(k);
    });
  }

  /** @param {HTMLElement} timelineRoot */
  function renderMufflerLogTimelineInto(timelineRoot, segIdx) {
    const log = mufflerSegmentLog(segIdx);
    const st = getMufflerSegState(segIdx);
    if (!timelineRoot || !log || !log.monthlyStats) {
      if (timelineRoot) timelineRoot.hidden = true;
      return;
    }
    timelineRoot.hidden = false;
    const granEl = timelineRoot.querySelector(".muffler-log-seg-granularity");
    if (granEl) granEl.value = st.granularity;

    const ys = mufflerLogYearsFromStats(log);
    if (ys.length) {
      if (st.timelineYear < ys[0]) st.timelineYear = ys[0];
      if (st.timelineYear > ys[ys.length - 1]) st.timelineYear = ys[ys.length - 1];
    }
    const yearLabel = timelineRoot.querySelector(".muffler-log-seg-year-label");
    if (yearLabel) yearLabel.textContent = `${st.timelineYear}년`;
    const yearPrev = timelineRoot.querySelector(".muffler-log-seg-year-prev");
    const yearNext = timelineRoot.querySelector(".muffler-log-seg-year-next");
    if (yearPrev) yearPrev.disabled = ys.length && st.timelineYear <= ys[0];
    if (yearNext) yearNext.disabled = ys.length && st.timelineYear >= ys[ys.length - 1];

    const monthStrip = timelineRoot.querySelector(".muffler-log-seg-month-strip");
    const monthKeys = new Set((log.monthlyStats || []).map((m) => m.key));
    if (monthStrip) {
      monthStrip.innerHTML = "";
      for (let m = 1; m <= 12; m++) {
        const ym = `${st.timelineYear}-${String(m).padStart(2, "0")}`;
        const btn = document.createElement("button");
        btn.type = "button";
        btn.className = "drawing-log-timeline__chip";
        btn.dataset.ym = ym;
        btn.textContent = String(m);
        const has = monthKeys.has(ym);
        if (!has) {
          btn.classList.add("is-empty");
          btn.disabled = true;
        }
        if (st.selectedMonths.has(ym)) btn.classList.add("is-selected");
        monthStrip.appendChild(btn);
      }
    }

    const dayStrip = timelineRoot.querySelector(".muffler-log-seg-day-strip");
    const dayMode = st.granularity === "day";
    if (dayStrip) {
      dayStrip.hidden = !dayMode;
      dayStrip.innerHTML = "";
      if (dayMode) {
        const pool = (log.dailyStats || []).filter((d) => {
          if (d.key === "일자미상") return false;
          if (!d.key.startsWith(`${st.timelineYear}-`)) return false;
          if (st.selectedMonths.size) {
            const ym = d.key.length >= 7 ? d.key.slice(0, 7) : "";
            return st.selectedMonths.has(ym);
          }
          return true;
        });
        pool.forEach((d) => {
          const btn = document.createElement("button");
          btn.type = "button";
          btn.className = "drawing-log-timeline__chip";
          btn.dataset.dayKey = d.key;
          btn.textContent = d.key.slice(5);
          if (st.selectedDays.has(d.key)) btn.classList.add("is-selected");
          dayStrip.appendChild(btn);
        });
      }
    }

    const selEl = timelineRoot.querySelector(".muffler-log-seg-slicer-summary");
    if (selEl) {
      const t = mufflerLogSlicerSummaryText(segIdx);
      selEl.textContent = t || "전체 기간 (필터 없음)";
    }
  }

  function getMufflerLogKpiSource(segIdx) {
    const log = mufflerSegmentLog(segIdx);
    if (!log) return null;
    const st = getMufflerSegState(segIdx);
    const kpiIgn = (st.kpiScope || "process") === "overall";
    const sub = reaggregateMufflerLogFiltered(segIdx, { ignoreProcessForRowFilter: kpiIgn });
    if (!sub) return log;
    return { ...log, ...sub };
  }

  function updateMufflerSegmentKpiRow(segRoot, segIdx) {
    const ksrc = getMufflerLogKpiSource(segIdx);
    if (!ksrc) return;
    const utilEl = segRoot.querySelector(".muffler-log-seg-kpi-util");
    const oeeEl = segRoot.querySelector(".muffler-log-seg-kpi-oee");
    const defLabel = segRoot.querySelector(".muffler-log-seg-kpi-def-label");
    const defVal = segRoot.querySelector(".muffler-log-seg-kpi-def");
    if (utilEl) utilEl.textContent = formatPercent(ksrc.utilAvg ?? NaN);
    if (oeeEl) oeeEl.textContent = formatPercent(ksrc.oeeAvg ?? NaN);
    if (defLabel && defVal) {
      if (Number.isFinite(ksrc.defectAvg)) {
        defLabel.textContent = "불량율";
        defVal.textContent = formatPercent(ksrc.defectAvg);
      } else if (Number.isFinite(ksrc.sumDefectQty)) {
        defLabel.textContent = "불량(합)";
        defVal.textContent = `${formatQty(ksrc.sumDefectQty)}건`;
      } else {
        defLabel.textContent = "불량율";
        defVal.textContent = "—";
      }
    }
  }

  function updateMufflerSegmentOpsRow(segRoot, segIdx) {
    const log = mufflerSegmentLog(segIdx);
    const ksrc = getMufflerLogKpiSource(segIdx);
    const ops = segRoot.querySelector(".muffler-log-seg-ops");
    if (!ops || !log || !ksrc) return;
    const bits = [];
    if (Number.isFinite(ksrc.sumProdQty)) bits.push(`생산수량 합 ${formatQty(ksrc.sumProdQty)}`);
    if (Number.isFinite(ksrc.sumWorkTime)) bits.push(`작업시간 합 ${formatQty(ksrc.sumWorkTime)}`);
    if (Number.isFinite(ksrc.sumInputTime)) bits.push(`투입시간 합 ${formatQty(ksrc.sumInputTime)}`);
    if (Number.isFinite(ksrc.productivityAvg)) bits.push(`생산성 평균 ${formatPercent(ksrc.productivityAvg)}`);
    if (Number.isFinite(ksrc.sumDefectQty)) bits.push(`불량 합 ${formatQty(ksrc.sumDefectQty)}건`);
    const qrst = Array.isArray(log.iStopTimeCols) && log.iStopTimeCols.length >= 4;
    if (qrst) {
      if (Number.isFinite(ksrc.sumStopExchange)) bits.push(`교환 ${formatQty(ksrc.sumStopExchange)}`);
      if (Number.isFinite(ksrc.sumStopRepair)) bits.push(`수리 ${formatQty(ksrc.sumStopRepair)}`);
      if (Number.isFinite(ksrc.sumStopMaterial)) bits.push(`소재 ${formatQty(ksrc.sumStopMaterial)}`);
      if (Number.isFinite(ksrc.sumStopPlanned)) bits.push(`계획정지 ${formatQty(ksrc.sumStopPlanned)}`);
      if (Number.isFinite(ksrc.sumStopTime)) bits.push(`정지 합 ${formatQty(ksrc.sumStopTime)}`);
    } else if (Number.isFinite(ksrc.sumStopTime)) {
      bits.push(`정지 시간 합 ${formatQty(ksrc.sumStopTime)}`);
    }
    ops.textContent = bits.join(" · ");
    ops.hidden = bits.length === 0;
  }

  function buildMufflerLogSegmentSection(segIdx) {
    const seg = mufflerSegmentLog(segIdx);
    if (!seg) return document.createDocumentFragment();
    const st = getMufflerSegState(segIdx);

    const root = document.createElement("section");
    root.className = "muffler-log-segment";
    root.dataset.segIndex = String(segIdx);

    const h3 = document.createElement("h3");
    h3.className = "drawing-log-block__title";
    h3.textContent = `시트 · ${seg.sheetName}`;
    root.appendChild(h3);

    const subMeta = document.createElement("p");
    subMeta.className = "drawing-log-meta";
    const dr =
      seg.dateFrom && seg.dateTo ? `작업일자 ${seg.dateFrom} ~ ${seg.dateTo}` : "작업일자 범위 없음";
    const fn = reaggregateMufflerLogFiltered(segIdx);
    const dayNote = fn ? ` · 보기: 필터 적용 (${fn.dataRows}행 / 전체 ${seg.dataRows}행)` : "";
    const kpiNote =
      (() => {
        const k = getMufflerLogKpiSource(segIdx);
        return k && (k.utilIsDerived || k.oeeIsDerived)
          ? " · 가동율·설비효율: 열이 없을 때 작업시간·정지·투입·생산수량·불량(율)·생산성으로 행별 산출 후 평균"
          : "";
      })();
    subMeta.textContent = `${dr} · 유효 ${seg.dataRows}행(전체 행 집계)${dayNote}${kpiNote}`;
    root.appendChild(subMeta);

    const timeline = document.createElement("div");
    timeline.className = "drawing-log-timeline muffler-log-segment-timeline";
    timeline.setAttribute("aria-label", "작업일자 필터");
    const granSel = st.granularity === "day" ? "day" : "month";
    timeline.innerHTML = `
      <div class="drawing-log-timeline__toolbar">
        <span class="drawing-log-timeline__label">작업일자</span>
        <div class="drawing-log-timeline__process-filters" role="group" aria-label="공정 필터">
          <label class="drawing-log-timeline__chk"><input type="checkbox" class="muffler-seg-filter" data-proc="cut" ${
            st.processCut ? "checked" : ""
          }/> 절단</label>
          <label class="drawing-log-timeline__chk"><input type="checkbox" class="muffler-seg-filter" data-proc="forming" ${
            st.processForming ? "checked" : ""
          }/> 포밍</label>
          <label class="drawing-log-timeline__chk"><input type="checkbox" class="muffler-seg-filter" data-proc="machine" ${
            st.processMachine ? "checked" : ""
          }/> 가공</label>
        </div>
        <button type="button" class="drawing-log-timeline__clear muffler-log-seg-clear" title="선택 해제">필터 지우기</button>
        <label class="drawing-log-timeline__gran">
          단위
          <select class="muffler-log-seg-granularity" aria-label="집계 단위">
            <option value="month"${granSel === "month" ? " selected" : ""}>월</option>
            <option value="day"${granSel === "day" ? " selected" : ""}>일</option>
          </select>
        </label>
        <label class="drawing-log-timeline__gran">
          상단 KPI
          <select class="muffler-log-seg-kpi-scope" aria-label="상단 KPI 집계 범위">
            <option value="overall"${st.kpiScope === "overall" ? " selected" : ""}>종합</option>
            <option value="process"${st.kpiScope !== "overall" ? " selected" : ""}>공정별</option>
          </select>
        </label>
      </div>
      <div class="drawing-log-timeline__selection muffler-log-seg-slicer-summary" aria-live="polite"></div>
      <div class="drawing-log-timeline__year">
        <button type="button" class="drawing-log-timeline__nav muffler-log-seg-year-prev" aria-label="이전 연도">‹</button>
        <span class="drawing-log-timeline__year-label muffler-log-seg-year-label">—</span>
        <button type="button" class="drawing-log-timeline__nav muffler-log-seg-year-next" aria-label="다음 연도">›</button>
      </div>
      <div class="drawing-log-timeline__months muffler-log-seg-month-strip" role="group" aria-label="월 선택"></div>
      <div class="drawing-log-timeline__days muffler-log-seg-day-strip" role="group" aria-label="일 선택" hidden></div>
    `;
    root.appendChild(timeline);

    const kpi = document.createElement("div");
    kpi.className = "kpi-grid";
    kpi.innerHTML = `
      <div class="kpi-card"><div class="kpi-card__label">가동율</div><div class="kpi-card__value muffler-log-seg-kpi-util">—</div></div>
      <div class="kpi-card"><div class="kpi-card__label">설비효율</div><div class="kpi-card__value muffler-log-seg-kpi-oee">—</div></div>
      <div class="kpi-card"><div class="kpi-card__label muffler-log-seg-kpi-def-label">불량율</div><div class="kpi-card__value muffler-log-seg-kpi-def">—</div></div>
    `;
    root.appendChild(kpi);

    const ops = document.createElement("p");
    ops.className = "drawing-log-ops muffler-log-seg-ops";
    ops.hidden = true;
    root.appendChild(ops);

    const tablesWrap = document.createElement("div");
    tablesWrap.className = "drawing-log-tables";
    root.appendChild(tablesWrap);

    const prodWrap = document.createElement("div");
    prodWrap.className = "drawing-log-prod-wrap";
    root.appendChild(prodWrap);

    renderMufflerLogTimelineInto(timeline, segIdx);
    updateMufflerSegmentKpiRow(root, segIdx);
    updateMufflerSegmentOpsRow(root, segIdx);
    renderMufflerLogTablesInto(tablesWrap, segIdx);
    renderMufflerLogProductionQtyInto(prodWrap, segIdx);

    return root;
  }

  function renderMufflerLogSegmentsHost() {
    if (!mufflerLogSegmentsHost || !lastMufflerLog || !lastMufflerLog.segments.length) return;
    mufflerLogSegmentsHost.innerHTML = "";
    lastMufflerLog.segments.forEach((_, segIdx) => {
      mufflerLogSegmentsHost.appendChild(buildMufflerLogSegmentSection(segIdx));
    });
  }

  function renderMufflerLogPanel() {
    if (!lastMufflerLog || !mufflerLogContent || !lastMufflerLog.segments || lastMufflerLog.segments.length === 0) {
      clearMufflerLogUi();
      return;
    }
    if (mufflerLogEmpty) mufflerLogEmpty.hidden = true;
    mufflerLogContent.hidden = false;

    const totalRows = lastMufflerLog.segments.reduce((a, s) => a + (s.dataRows || 0), 0);
    const names = lastMufflerLog.segments.map((s) => s.sheetName).join(", ");
    if (mufflerLogMeta) {
      const anyDerived = lastMufflerLog.segments.some((_, i) => {
        const k = getMufflerLogKpiSource(i);
        return k && (k.utilIsDerived || k.oeeIsDerived);
      });
      const kpiNote = anyDerived
        ? " · 가동율·설비효율: 열이 없을 때 작업시간·정지·투입·생산수량·불량(율)·생산성으로 행별 산출 후 평균"
        : "";
      mufflerLogMeta.textContent = `파일: ${lastMufflerLog.fileLabel} · 시트별 집계: ${names} · 유효 행 합계 ${totalRows}행${kpiNote}`;
    }

    renderMufflerLogSegmentsHost();

    if (mufflerLogMaintWrap) {
      if (lastMufflerLog.maintenance) {
        mufflerLogMaintWrap.hidden = false;
        renderWorkbookMaintenanceInto(mufflerLogMaintWrap, lastMufflerLog.maintenance);
      } else {
        mufflerLogMaintWrap.innerHTML = "";
        mufflerLogMaintWrap.hidden = true;
      }
    }
  }

  /**
   * @param {any} wb
   * @param {string} fileLabel
   * @returns {boolean}
   */
  function tryConsumeWorkbookAsMufflerLog(wb, fileLabel) {
    const hit = findMufflerLogSheetInWorkbook(wb);
    if (!hit || !hit.segments.length) return false;
    const skipNames = hit.segments.map((s) => s.sheetName);
    const maintParsed = parseWorkbookMaintenance(wb, fileLabel, skipNames);
    const maintenance =
      maintParsed.failures.length || maintParsed.pmRows.length ? maintParsed : null;
    lastMufflerLog = {
      fileLabel,
      maintenance,
      segments: hit.segments.map((s) => ({ sheetName: s.sheetName, ...s.parsed })),
    };
    initMufflerLogSlicerFromData();
    renderMufflerLogPanel();
    setView("mufflerLog");
    return true;
  }

  async function handleMufflerLogFile(file) {
    if (!file) return;
    try {
      const buf = await loadArrayBuffer(file);
      const wb = XLSX.read(buf, { type: "array", cellDates: true });
      const ok = tryConsumeWorkbookAsMufflerLog(wb, file.name);
      if (!ok) {
        alert(
          "「머플러」또는「단조머플러」시트를 찾을 수 없거나, 작업일자·생산·작업시간 등 알 수 있는 열이 없습니다.\n시트 이름·헤더를 확인해 주세요."
        );
        lastMufflerLog = null;
        clearMufflerLogUi();
      }
    } catch (e) {
      console.error(e);
      alert("파일을 읽는 중 오류가 났습니다. 엑셀 형식인지 확인해 주세요.");
    }
  }

  /**
   * @param {any} wb
   * @param {string} fileLabel
   * @returns {boolean}
   */
  function tryConsumeWorkbookAsDrawingLog(wb, fileLabel) {
    const hit = findDrawingLogSheetInWorkbook(wb);
    if (!hit) return false;
    const p = hit.parsed;
    const maintParsed = parseWorkbookMaintenance(wb, fileLabel, [hit.sheetName]);
    const maintenance =
      maintParsed.failures.length || maintParsed.pmRows.length ? maintParsed : null;
    lastDrawingLog = {
      ...p,
      sheetName: hit.sheetName,
      fileLabel,
      maintenance,
    };
    initDrawingLogSlicerFromData();
    renderDrawingLogPanel();
    setView("drawingLog");
    return true;
  }

  async function handleDrawingLogFile(file) {
    if (!file) return;
    try {
      const buf = await loadArrayBuffer(file);
      const wb = XLSX.read(buf, { type: "array", cellDates: true });
      const ok = tryConsumeWorkbookAsDrawingLog(wb, file.name);
      if (!ok) {
        alert(
          "「드로잉」시트가 없거나, 작업일자·생산수량·작업시간 등 알 수 있는 열이 없습니다.\n시트 이름에 드로잉이 들어가는지, 헤더 행을 확인해 주세요."
        );
        lastDrawingLog = null;
        clearDrawingLogUi();
      }
    } catch (e) {
      console.error(e);
      alert("파일을 읽는 중 오류가 났습니다. 엑셀 형식인지 확인해 주세요.");
    }
  }

  /**
   * @param {any} wb
   * @param {string} fileLabel
   * @returns {boolean}
   */
  function tryConsumeWorkbookAsParallelLog(wb, fileLabel) {
    const hit = findParallelLogSheetInWorkbook(wb);
    if (!hit) return false;
    const p = hit.parsed;
    const maintParsed = parseWorkbookMaintenance(wb, fileLabel, [hit.sheetName]);
    const maintenance =
      maintParsed.failures.length || maintParsed.pmRows.length ? maintParsed : null;
    lastParallelLog = {
      ...p,
      sheetName: hit.sheetName,
      fileLabel,
      maintenance,
    };
    initParallelLogSlicerFromData();
    renderParallelLogPanel();
    setView("parallelLog");
    return true;
  }

  async function handleParallelLogFile(file) {
    if (!file) return;
    try {
      const buf = await loadArrayBuffer(file);
      const wb = XLSX.read(buf, { type: "array", cellDates: true });
      const ok = tryConsumeWorkbookAsParallelLog(wb, file.name);
      if (!ok) {
        alert(
          "「페럴」시트가 없거나, 작업일자·생산량(ERP)·작업시간 등 알 수 있는 열이 없습니다.\n시트 이름이 페럴인지, 헤더 행을 확인해 주세요."
        );
        lastParallelLog = null;
        clearParallelLogUi();
      }
    } catch (e) {
      console.error(e);
      alert("파일을 읽는 중 오류가 났습니다. 엑셀 형식인지 확인해 주세요.");
    }
  }

  function tryConsumeWorkbookAsDoublePipeLog(wb, fileLabel) {
    const hit = findDoublePipeLogSheetInWorkbook(wb);
    if (!hit) return false;
    const p = hit.parsed;
    const maintParsed = parseWorkbookMaintenance(wb, fileLabel, [hit.sheetName]);
    const maintenance =
      maintParsed.failures.length || maintParsed.pmRows.length ? maintParsed : null;
    lastDoublePipeLog = {
      ...p,
      sheetName: hit.sheetName,
      fileLabel,
      maintenance,
    };
    initDoublePipeLogSlicerFromData();
    renderDoublePipeLogPanel();
    setView("doublePipeLog");
    return true;
  }

  async function handleDoublePipeLogFile(file) {
    if (!file) return;
    try {
      const buf = await loadArrayBuffer(file);
      const wb = XLSX.read(buf, { type: "array", cellDates: true });
      const ok = tryConsumeWorkbookAsDoublePipeLog(wb, file.name);
      if (!ok) {
        alert(
          "「이중관」시트가 없거나, 작업일자·생산·작업시간 등 읽을 수 있는 열이 없습니다.\n시트 이름과 헤더 행을 확인해 주세요.\n(비가동 S~V열, 불량 W+AJ열, 생산 N열 기준)"
        );
        lastDoublePipeLog = null;
        clearDoublePipeLogUi();
      }
    } catch (e) {
      console.error(e);
      alert("파일을 읽는 중 오류가 났습니다. 엑셀 형식인지 확인해 주세요.");
    }
  }

  function tryConsumeWorkbookAsHwaseungDoublePipeLog(wb, fileLabel) {
    const hit = findHwaseungDoublePipeLogSheetInWorkbook(wb);
    if (!hit) return false;
    const p = hit.parsed;
    const maintParsed = parseWorkbookMaintenance(wb, fileLabel, [hit.sheetName]);
    const maintenance =
      maintParsed.failures.length || maintParsed.pmRows.length ? maintParsed : null;
    lastHwaseungDoublePipeLog = {
      ...p,
      sheetName: hit.sheetName,
      fileLabel,
      maintenance,
    };
    initHwaseungDoublePipeLogSlicerFromData();
    renderHwaseungDoublePipeLogPanel();
    setView("hwaseungDoublePipeLog");
    return true;
  }

  async function handleHwaseungDoublePipeLogFile(file) {
    if (!file) return;
    try {
      const buf = await loadArrayBuffer(file);
      const wb = XLSX.read(buf, { type: "array", cellDates: true });
      const ok = tryConsumeWorkbookAsHwaseungDoublePipeLog(wb, file.name);
      if (!ok) {
        alert(
          "「화승이중관」시트가 없거나, 작업일자·생산·작업시간 등 읽을 수 있는 열이 없습니다.\n시트 이름과 헤더 행을 확인해 주세요.\n(비가동 R~V열, 불량은 X+Y+Z 및 AK+AL+AR 세부만 합산, 생산 L열 기준)"
        );
        lastHwaseungDoublePipeLog = null;
        clearHwaseungDoublePipeLogUi();
      }
    } catch (e) {
      console.error(e);
      alert("파일을 읽는 중 오류가 났습니다. 엑셀 형식인지 확인해 주세요.");
    }
  }

  const MAINT_FAILURE_DATE_KEYS = [
    "고장일",
    "고장날짜",
    "발생일",
    "장애일",
    "장애발생일",
    "접수일",
    "수리요청일",
    "날짜",
    "일자",
    "date",
  ];
  const MAINT_EQUIP_KEYS = ["설비", "설비명", "호기", "호기명", "기계", "기계명", "장비", "라인", "line", "equipment"];
  const MAINT_CAUSE_KEYS = ["원인", "고장원인", "사유", "내용", "상세", "조치내용", "비고"];
  const MAINT_REPAIR_KEYS = ["수리시간", "조치시간", "복구시간", "다운타임", "수리분", "정비시간", "mttr"];
  const MAINT_PM_NEXT_KEYS = ["다음정기", "다음점검", "예정일", "예방일", "차기점검", "차기일", "다음일자", "점검예정"];
  const MAINT_PM_CYCLE_KEYS = ["주기", "점검주기", "pm주기", "cycle", "정비주기"];

  function ymdToMsLocal(ymd) {
    const p = String(ymd || "").split("-");
    if (p.length !== 3) return NaN;
    const d = new Date(parseInt(p[0], 10), parseInt(p[1], 10) - 1, parseInt(p[2], 10));
    return Number.isNaN(d.getTime()) ? NaN : d.getTime();
  }

  function medianFinite(values) {
    const v = values.filter((x) => Number.isFinite(x)).sort((a, b) => a - b);
    if (v.length === 0) return null;
    const n = v.length;
    const mid = Math.floor(n / 2);
    if (n % 2 === 1) return v[mid];
    return (v[mid - 1] + v[mid]) / 2;
  }

  function describeFailureGapPattern(medianDays) {
    if (!Number.isFinite(medianDays) || medianDays < 0.5) return "";
    const d = medianDays;
    if (d >= 5 && d <= 11) return "약 1주 간격으로 고장이 반복되는 편입니다.";
    if (d >= 22 && d <= 40) return "약 1개월 간격으로 고장이 반복되는 편입니다.";
    if (d >= 55 && d <= 105) return "약 2~3개월 간격으로 고장이 반복되는 편입니다.";
    if (d >= 115 && d <= 210) return "약 4~6개월 간격으로 고장이 반복되는 편입니다.";
    if (d >= 250 && d <= 400) return "약 9~12개월 간격으로 고장이 반복되는 편입니다.";
    if (d < 5) return "짧은 주기(일주 이내)로 이벤트가 잦습니다.";
    if (d > 400) return "비교적 긴 간격입니다.";
    return `고장 간격 중앙값 약 ${Math.round(d)}일입니다.`;
  }

  /**
   * @param {any[][]} matrix
   * @param {string} sheetName
   * @returns {null | { rows: { ymd: string, t: number, equipment: string, cause: string, repairHours: number }[] }}
   */
  function parseFailureSheet(matrix, sheetName) {
    if (!matrix || matrix.length < 2) return null;
    const hi = detectHeaderRowIndex(matrix);
    const h = matrix[hi] || [];
    const iDate = findHeaderIndex(h, MAINT_FAILURE_DATE_KEYS);
    const iEq = findHeaderIndex(h, MAINT_EQUIP_KEYS);
    const iCause = findHeaderIndex(h, MAINT_CAUSE_KEYS);
    const iRepair = findHeaderIndex(h, MAINT_REPAIR_KEYS);
    if (iDate < 0) return null;
    const tnm = norm(sheetName || "");
    const nameLikely = /고장|수리|이력|장애|break|fail|down|장비고장|수리이력/.test(tnm);
    if (!nameLikely && iEq < 0 && iRepair < 0 && iCause < 0) return null;
    /** @type {any[]} */
    const rows = [];
    for (let r = hi + 1; r < matrix.length; r++) {
      const row = matrix[r];
      if (!row || !row.some((c) => String(c ?? "").trim() !== "")) continue;
      const ymd = parseExcelDate(row[iDate]);
      if (!ymd) continue;
      const t = ymdToMsLocal(ymd);
      if (!Number.isFinite(t)) continue;
      const equipment = iEq >= 0 ? String(row[iEq] ?? "").trim() || "미지정" : "전체";
      const cause = iCause >= 0 ? String(row[iCause] ?? "").trim() : "";
      const repairHours = iRepair >= 0 ? parseNumber(row[iRepair]) : NaN;
      rows.push({ ymd, t, equipment, cause, repairHours });
    }
    return rows.length ? { rows } : null;
  }

  /**
   * @param {any[][]} matrix
   * @param {string} sheetName
   * @returns {null | { rows: { equipment: string, nextYmd: string, cycle: string, note: string }[] }}
   */
  function parsePmSheet(matrix, sheetName) {
    if (!matrix || matrix.length < 2) return null;
    const hi = detectHeaderRowIndex(matrix);
    const h = matrix[hi] || [];
    const iNext = findHeaderIndex(h, MAINT_PM_NEXT_KEYS);
    const iEq = findHeaderIndex(h, MAINT_EQUIP_KEYS);
    if (iNext < 0 || iEq < 0) return null;
    const iCycle = findHeaderIndex(h, MAINT_PM_CYCLE_KEYS);
    const tnm = norm(sheetName || "");
    const nameLikely = /예방|정기|pm|점검|정비|스케줄|maintenance|계획/.test(tnm);
    if (!nameLikely && iCycle < 0) return null;
    const iNote = findHeaderIndex(h, ["비고", "메모", "비고사항", "내용"]);
    /** @type {any[]} */
    const rows = [];
    for (let r = hi + 1; r < matrix.length; r++) {
      const row = matrix[r];
      if (!row || !row.some((c) => String(c ?? "").trim() !== "")) continue;
      const nextYmd = parseExcelDate(row[iNext]);
      if (!nextYmd) continue;
      const equipment = String(row[iEq] ?? "").trim() || "미지정";
      const cycle = iCycle >= 0 ? String(row[iCycle] ?? "").trim() : "";
      const note = iNote >= 0 ? String(row[iNote] ?? "").trim() : "";
      rows.push({ equipment, nextYmd, cycle, note });
    }
    return rows.length ? { rows } : null;
  }

  /**
   * @param {{ ymd: string, t: number, equipment: string, cause: string, repairHours: number }[]} failures
   */
  function buildMaintEquipmentStats(failures) {
    /** @type {Map<string, any[]>} */
    const byEq = new Map();
    for (const f of failures) {
      const k = f.equipment || "미지정";
      if (!byEq.has(k)) byEq.set(k, []);
      byEq.get(k).push(f);
    }
    /** @type {any[]} */
    const out = [];
    for (const [equipment, arr] of byEq) {
      arr.sort((a, b) => a.t - b.t);
      /** @type {number[]} */
      const gapsDays = [];
      for (let i = 1; i < arr.length; i++) {
        gapsDays.push((arr[i].t - arr[i - 1].t) / 86400000);
      }
      const mtbfDays = gapsDays.length ? averageFinite(gapsDays) : null;
      const medGap = medianFinite(gapsDays);
      const mttrHours = averageFinite(arr.map((x) => x.repairHours));
      const hintLine = gapsDays.length
        ? describeFailureGapPattern(medGap ?? NaN) ||
          (Number.isFinite(medGap) ? `고장 간격 중앙값 약 ${Math.round(medGap)}일입니다.` : "")
        : "고장 1건 이하로 간격(MTBF) 분석 없음";
      out.push({
        equipment,
        failureCount: arr.length,
        mtbfDays,
        mttrHours,
        patternHint: hintLine,
      });
    }
    out.sort((a, b) => b.failureCount - a.failureCount || a.equipment.localeCompare(b.equipment, "ko"));
    return out;
  }

  /**
   * @param {any} wb
   * @param {string} fileLabel
   * @param {string[]} [skipSheetNames] 드로잉 등 생산 시트는 고장/PM 파싱에서 제외
   */
  function parseWorkbookMaintenance(wb, fileLabel, skipSheetNames) {
    const names = wb.SheetNames || [];
    const skip = new Set(Array.isArray(skipSheetNames) ? skipSheetNames : []);
    /** @type {any[]} */
    const failures = [];
    /** @type {string[]} */
    const failSheets = [];
    /** @type {any[]} */
    const pmRows = [];
    /** @type {string[]} */
    const pmSheets = [];
    for (const name of names) {
      if (skip.has(name)) continue;
      const matrix = sheetToMatrix(wb, name);
      const fr = parseFailureSheet(matrix, name);
      if (fr && fr.rows.length) {
        failures.push(...fr.rows.map((row) => ({ ...row, sheet: name })));
        failSheets.push(name);
        continue;
      }
      const pr = parsePmSheet(matrix, name);
      if (pr && pr.rows.length) {
        pmRows.push(...pr.rows.map((row) => ({ ...row, sheet: name })));
        pmSheets.push(name);
      }
    }
    return {
      fileLabel,
      failures,
      failSheets: [...new Set(failSheets)],
      pmRows,
      pmSheets: [...new Set(pmSheets)],
      equipmentStats: buildMaintEquipmentStats(failures),
    };
  }

  /**
   * @param {HTMLElement} host
   * @param {ReturnType<typeof parseWorkbookMaintenance>} data
   */
  function renderWorkbookMaintenanceInto(host, data) {
    if (!host || !data) return;
    const { failures, failSheets, pmRows, pmSheets, equipmentStats, fileLabel } = data;
    host.innerHTML = "";
    const blockTitle = document.createElement("h3");
    blockTitle.className = "drawing-log-block__title";
    blockTitle.textContent = "같은 파일 · 설비 고장 이력 및 예방정비";
    host.appendChild(blockTitle);
    const meta = document.createElement("p");
    meta.className = "drawing-log-meta";
    const fs = failSheets.length ? failSheets.join(", ") : "없음";
    const ps = pmSheets.length ? pmSheets.join(", ") : "없음";
    meta.textContent = `파일: ${fileLabel} · 고장 시트: ${fs} · 예방정비 시트: ${ps} · 고장 행 ${failures.length}건`;
    host.appendChild(meta);

    const hKpi = document.createElement("h3");
    hKpi.className = "drawing-log-block__title";
    hKpi.textContent = "설비별 MTBF · MTTR";
    host.appendChild(hKpi);
    const kpiWrap = document.createElement("div");
    kpiWrap.className = "drawing-log-preview-wrap";
    const tblK = document.createElement("table");
    const theadK = document.createElement("thead");
    const trhK = document.createElement("tr");
    ["설비", "고장 건수", "MTBF(일)", "MTTR(시간)", "간격 패턴"].forEach((tx) => {
      const th = document.createElement("th");
      th.textContent = tx;
      trhK.appendChild(th);
    });
    theadK.appendChild(trhK);
    tblK.appendChild(theadK);
    const tbodyK = document.createElement("tbody");
    if (!equipmentStats.length) {
      const tr = document.createElement("tr");
      const td = document.createElement("td");
      td.colSpan = 5;
      td.textContent = "집계할 고장 이력이 없습니다.";
      tr.appendChild(td);
      tbodyK.appendChild(tr);
    } else {
      for (const st of equipmentStats) {
        const tr = document.createElement("tr");
        const cells = [
          st.equipment,
          String(st.failureCount),
          Number.isFinite(st.mtbfDays) ? `${st.mtbfDays.toFixed(1)}일` : "—",
          Number.isFinite(st.mttrHours) ? `${formatQty(st.mttrHours)}시간` : "—",
          st.patternHint || "—",
        ];
        cells.forEach((text, i) => {
          const td = document.createElement("td");
          td.textContent = text;
          if (i > 0) td.classList.add("num");
          tr.appendChild(td);
        });
        tbodyK.appendChild(tr);
      }
    }
    tblK.appendChild(tbodyK);
    kpiWrap.appendChild(tblK);
    host.appendChild(kpiWrap);

    const hFail = document.createElement("h3");
    hFail.className = "drawing-log-block__title";
    hFail.textContent = "고장 이력";
    host.appendChild(hFail);
    const failWrap = document.createElement("div");
    failWrap.className = "drawing-log-preview-wrap";
    if (failures.length) {
      const sorted = failures.slice().sort((a, b) => b.t - a.t);
      const show = sorted.slice(0, 100);
      const tbl = document.createElement("table");
      const thead = document.createElement("thead");
      const trh = document.createElement("tr");
      ["고장일", "설비", "원인", "수리시간", "시트"].forEach((tx) => {
        const th = document.createElement("th");
        th.textContent = tx;
        trh.appendChild(th);
      });
      thead.appendChild(trh);
      tbl.appendChild(thead);
      const tbody = document.createElement("tbody");
      for (const f of show) {
        const tr = document.createElement("tr");
        [
          f.ymd,
          f.equipment,
          f.cause || "—",
          Number.isFinite(f.repairHours) ? formatQty(f.repairHours) : "—",
          f.sheet || "—",
        ].forEach((text, i) => {
          const td = document.createElement("td");
          td.textContent = text;
          if (i === 3) td.classList.add("num");
          tr.appendChild(td);
        });
        tbody.appendChild(tr);
      }
      tbl.appendChild(tbody);
      failWrap.appendChild(tbl);
      if (sorted.length > show.length) {
        const p = document.createElement("p");
        p.className = "drawing-log-meta";
        p.textContent = `최근 ${show.length}건만 표시 (전체 ${sorted.length}건)`;
        failWrap.appendChild(p);
      }
    } else {
      const p = document.createElement("p");
      p.className = "drawing-log-meta";
      p.textContent = "고장 이력 시트는 없습니다.";
      failWrap.appendChild(p);
    }
    host.appendChild(failWrap);

    const hPm = document.createElement("h3");
    hPm.className = "drawing-log-block__title";
    hPm.textContent = "예방정비 스케줄";
    host.appendChild(hPm);
    const pmWrap = document.createElement("div");
    pmWrap.className = "drawing-log-preview-wrap";
    if (!pmRows.length) {
      const p = document.createElement("p");
      p.className = "drawing-log-meta";
      p.textContent = failures.length
        ? "예방정비 형식의 시트는 이 파일에서 찾지 못했습니다."
        : "예방정비 시트를 찾지 못했습니다. 시트 이름에 예방·정기·점검 등을 넣거나, 열에 「다음정기일·설비」(또는 예정일·호기)를 맞춰 주세요.";
      pmWrap.appendChild(p);
    } else {
      const tbl = document.createElement("table");
      const thead = document.createElement("thead");
      const trh = document.createElement("tr");
      ["설비", "다음 점검일", "주기", "비고", "시트"].forEach((tx) => {
        const th = document.createElement("th");
        th.textContent = tx;
        trh.appendChild(th);
      });
      thead.appendChild(trh);
      tbl.appendChild(thead);
      const tbody = document.createElement("tbody");
      const sortedPm = pmRows.slice().sort((a, b) => String(a.nextYmd).localeCompare(String(b.nextYmd)));
      for (const r of sortedPm) {
        const tr = document.createElement("tr");
        [r.equipment, r.nextYmd, r.cycle || "—", r.note || "—", r.sheet || "—"].forEach((text) => {
          const td = document.createElement("td");
          td.textContent = text || "—";
          tr.appendChild(td);
        });
        tbody.appendChild(tr);
      }
      tbl.appendChild(tbody);
      pmWrap.appendChild(tbl);
    }
    host.appendChild(pmWrap);
  }

  /**
   * @param {any[][]} matrix
   * @param {any} [workbook]
   */
  function processMatrix(matrix, workbook) {
    if (!matrix || matrix.length < 2) {
      alert("데이터가 너무 적습니다. 「발주서 입력」 시트에 헤더와 데이터가 있는지 확인하세요.");
      return;
    }
    rawRows = matrix;
    headers = matrix[0].map((c, i) => String(c ?? "").trim() || `열 ${i + 1}`);
    const wideSpec = analyzeWideHeader(matrix[0]);

    if (wideSpec.wide) {
      dataMode = "wide";
      if (mappingBlock) mappingBlock.hidden = true;
      let board = buildBoardFromWide(matrix, wideSpec);
      if (board.products.length === 0) {
        alert("「발주서 입력」 시트에서 구분·품목코드·제품·날짜 열이 있는지 확인하세요.");
        clearOutputs();
        return;
      }
      if (workbook) board = mergePastOrderSheetIntoBoard(workbook, board);
      if (lastStockData) board = mergeStockDataIntoBoard(board, lastStockData);
      alignBoardDates(board);
      lastBoard = board;
      onDataReady();
      return;
    }

    dataMode = "long";
    if (mappingBlock) mappingBlock.hidden = true;
    const guessed = guessColumnIndex(headers);
    fillSelect(colDate, headers, guessed.date);
    fillSelect(colName, headers, guessed.name);
    fillSelect(colCode, headers, guessed.code);
    fillSelect(colType, headers, guessed.type);
    fillSelect(colQty, headers, guessed.qty);
    fillSelect(colStock, headers, guessed.stock);
    applyMapping();
  }

  function readIndices() {
    return {
      date: parseInt(colDate.value, 10),
      name: parseInt(colName.value, 10),
      code: parseInt(colCode.value, 10),
      type: parseInt(colType.value, 10),
      qty: parseInt(colQty.value, 10),
      stock: parseInt(colStock.value, 10),
    };
  }

  function applyMapping() {
    if (!rawRows || dataMode !== "long") return;
    const idx = readIndices();
    const pivot = buildPivot(rawRows, idx);
    if (!pivot || pivot.dates.length === 0) {
      alert("날짜·제품명·발주량 열을 올바르게 지정했는지 확인하세요.");
      if (mappingBlock) mappingBlock.hidden = false;
      lastBoard = null;
      orderCalendarNeedsSync = true;
      const bt = tableWrap.querySelector("table.board-table");
      if (bt) bt.remove();
      emptyState.hidden = false;
      const st = simpleTableWrap.querySelector("table.simple-table");
      if (st) st.remove();
      emptyStateSimple.hidden = false;
      dailyFilters.classList.add("filter-bar--hidden");
      clearFilterSelections();
      renderSummary(null);
      rowCountEl.textContent = "0건";
      btnExport.disabled = true;
      renderOrderCalendar(null);
      return;
    }
    if (mappingBlock) mappingBlock.hidden = true;
    lastBoard = buildBoardFromLong(pivot, idx);
    if (lastWorkbook) lastBoard = mergePastOrderSheetIntoBoard(lastWorkbook, lastBoard);
    if (lastStockData) lastBoard = mergeStockDataIntoBoard(lastBoard, lastStockData);
    alignBoardDates(lastBoard);
    onDataReady();
  }

  async function handleFile(file) {
    if (!file) return;
    lastFileName = file.name;
    fileNameEl.textContent = file.name;

    try {
      const buf = await loadArrayBuffer(file);
      const wb = XLSX.read(buf, { type: "array", cellDates: true });
      lastWorkbook = wb;
      const sheetName = findSheetName(wb);
      if (!sheetName) {
        if (tryConsumeWorkbookAsDrawingLog(wb, file.name)) {
          fileNameEl.textContent = `${file.name} (작업일지)`;
          sheetInfoEl.textContent =
            "「발주서 입력」 시트는 없고, 드로잉 작업일지로 인식했습니다. 좌측 「드로잉」에서 언제든 다시 볼 수 있습니다.";
          lastWorkbook = null;
          return;
        }
        if (tryConsumeWorkbookAsParallelLog(wb, file.name)) {
          fileNameEl.textContent = `${file.name} (페럴 작업일지)`;
          sheetInfoEl.textContent =
            "「발주서 입력」 시트는 없고, 페럴 작업일지로 인식했습니다. 좌측 「페럴」에서 언제든 다시 볼 수 있습니다.";
          lastWorkbook = null;
          return;
        }
        if (tryConsumeWorkbookAsDoublePipeLog(wb, file.name)) {
          fileNameEl.textContent = `${file.name} (이중관 작업일지)`;
          sheetInfoEl.textContent =
            "「발주서 입력」 시트는 없고, 이중관 작업일지로 인식했습니다. 좌측 「이중관」작업일지에서 언제든 다시 볼 수 있습니다.";
          lastWorkbook = null;
          return;
        }
        if (tryConsumeWorkbookAsHwaseungDoublePipeLog(wb, file.name)) {
          fileNameEl.textContent = `${file.name} (화승 이중관 작업일지)`;
          sheetInfoEl.textContent =
            "「발주서 입력」 시트는 없고, 화승 이중관 작업일지로 인식했습니다. 좌측 「화승이중관」에서 언제든 다시 볼 수 있습니다.";
          lastWorkbook = null;
          return;
        }
        if (tryConsumeWorkbookAsMufflerLog(wb, file.name)) {
          fileNameEl.textContent = `${file.name} (머플러 작업일지)`;
          sheetInfoEl.textContent =
            "「발주서 입력」 시트는 없고, 머플러 작업일지로 인식했습니다. 좌측 「머플러」에서 언제든 다시 볼 수 있습니다.";
          lastWorkbook = null;
          return;
        }
        const avail = (wb.SheetNames || []).join(", ");
        alert(
          `「${SHEET_TARGET}」 시트를 찾을 수 없습니다.\n` +
            (avail ? `파일의 시트: ${avail}` : "시트가 없습니다.") +
            "\n\n작업일지라면 좌측 「드로잉」「페럴」「이중관」「화승이중관」「머플러」화면에서 올려 주세요."
        );
        sheetInfoEl.textContent = "";
        lastWorkbook = null;
        return;
      }
      const pastNm = findPastOrderSheetName(wb);
      sheetInfoEl.textContent =
        `발주서 입력: ${sheetName}` + (pastNm ? ` · 지난발주서: ${pastNm}` : " (지난발주서 시트 없음)");
      const matrix = sheetToMatrix(wb, sheetName);
      processMatrix(matrix, wb);
    } catch (e) {
      console.error(e);
      alert("파일을 읽는 중 오류가 났습니다. 엑셀 형식인지 확인해 주세요.");
    }
  }

  /**
   * @param {File} file
   */
  async function handleStockFile(file) {
    if (!file) return;
    if (stockFileNameEl) stockFileNameEl.textContent = `재고파일: ${file.name}`;
    try {
      const buf = await loadArrayBuffer(file);
      const wb = XLSX.read(buf, { type: "array", cellDates: true });
      const sheetName = findStockSheetName(wb);
      if (!sheetName) {
        alert("재고파일에서 시트를 찾지 못했습니다.");
        if (stockSheetInfoEl) stockSheetInfoEl.textContent = "";
        return;
      }
      const matrix = sheetToMatrix(wb, sheetName);
      const parsed = extractStockData(matrix);
      if (parsed.stockByKey.size === 0 && parsed.stockByName.size === 0) {
        alert("재고파일에서 품목코드/제품명/재고수량 데이터를 읽지 못했습니다. 헤더를 확인해 주세요.");
        if (stockSheetInfoEl) stockSheetInfoEl.textContent = "";
        renderStockTableView();
        return;
      }
      lastStockData = parsed;
      const infoLine = `재고 시트: ${sheetName} · ${parsed.rowCount}행 · 품목 ${parsed.stockByKey.size || parsed.stockByName.size}건`;
      if (stockSheetInfoEl) stockSheetInfoEl.textContent = infoLine;
      populateStockTableFiltersFromData();
      renderStockTableView();

      if (lastBoard) {
        applyStockDataToCurrentBoard();
        onDataReady();
      }
    } catch (e) {
      console.error(e);
      alert("재고파일을 읽는 중 오류가 났습니다. 엑셀 형식인지 확인해 주세요.");
    }
  }

  async function handleStockUnitPriceFile(file) {
    if (!file) return;
    hideStockUnitPriceError();
    if (stockUnitPriceFileName) stockUnitPriceFileName.textContent = file.name;
    try {
      const buf = await loadArrayBuffer(file);
      const wb = XLSX.read(buf, { type: "array", cellDates: true });
      const { sheetName, parsed } = buildStockUnitPriceDataFromWorkbook(wb);
      if (!sheetName || !parsed || parsed.rowCount === 0) {
        lastStockUnitPriceData = null;
        clearStockUnitPricePreviewTable();
        resetStockUnitPriceCodeFilter();
        showStockUnitPriceError(
          "파일에서 재고_단가 품목명/현재고를 읽지 못했습니다. 재고 시트에 품목명·현재고 헤더가 있는지 확인해 주세요."
        );
        renderStockUnitPricePanel();
        return;
      }
      const uploadDateLabel = formatKoreanUploadDate(new Date());
      lastStockUnitPriceData = {
        byKey: new Map(parsed.byKey),
        byName: new Map(parsed.byName),
        byNormName: new Map(parsed.byNormName || []),
        byNameRows: new Map(parsed.byNameRows || []),
        previewRows: parsed.previewRows.slice(),
        fileLabel: file.name,
        sheetName,
        rowCount: parsed.rowCount,
        priceMode: parsed.priceMode,
        uploadDateLabel,
      };
      hideStockUnitPriceError();
      populateStockUnitPriceCodeFilterFromPreview();
      renderStockUnitPricePanel();
      if (lastBoard) {
        renderBoard(lastBoard);
        if (currentView === "simple") renderSimpleTable(lastBoard);
        if (currentView === "orderCalendar") renderOrderCalendar(lastBoard);
      }
    } catch (e) {
      console.error(e);
      lastStockUnitPriceData = null;
      clearStockUnitPricePreviewTable();
      resetStockUnitPriceCodeFilter();
      showStockUnitPriceError("파일을 읽는 중 오류가 났습니다. .xlsx / .xls 형식인지 확인해 주세요.");
      renderStockUnitPricePanel();
    }
  }

  fileInput.addEventListener("change", () => {
    const f = fileInput.files && fileInput.files[0];
    if (f) handleFile(f);
    fileInput.value = "";
  });
  stockFileInput.addEventListener("change", () => {
    const f = stockFileInput.files && stockFileInput.files[0];
    if (f) handleStockFile(f);
    stockFileInput.value = "";
  });

  if (stockUnitPriceFileInput) {
    stockUnitPriceFileInput.addEventListener("change", () => {
      const f = stockUnitPriceFileInput.files && stockUnitPriceFileInput.files[0];
      if (f) handleStockUnitPriceFile(f);
      stockUnitPriceFileInput.value = "";
    });
  }
  if (stockUnitPriceDropzone) {
    ["dragenter", "dragover"].forEach((ev) => {
      stockUnitPriceDropzone.addEventListener(ev, (e) => {
        e.preventDefault();
        stockUnitPriceDropzone.classList.add("dragover");
      });
    });
    ["dragleave", "drop"].forEach((ev) => {
      stockUnitPriceDropzone.addEventListener(ev, (e) => {
        e.preventDefault();
        stockUnitPriceDropzone.classList.remove("dragover");
      });
    });
    stockUnitPriceDropzone.addEventListener("drop", (e) => {
      const f = e.dataTransfer.files && e.dataTransfer.files[0];
      if (f) handleStockUnitPriceFile(f);
    });
  }

  if (btnStockUnitPriceFilterCode && panelStockUnitPriceFilterCode) {
    btnStockUnitPriceFilterCode.addEventListener("click", (e) => {
      e.stopPropagation();
      const willOpen = panelStockUnitPriceFilterCode.hidden;
      closeAllPickers("");
      panelStockUnitPriceFilterCode.hidden = !willOpen;
      if (willOpen && searchStockUnitPriceFilterCode) searchStockUnitPriceFilterCode.focus();
    });
    if (searchStockUnitPriceFilterCode) {
      searchStockUnitPriceFilterCode.addEventListener("input", () => renderStockUnitPriceCodePickerList());
    }
    panelStockUnitPriceFilterCode.addEventListener("click", (e) => e.stopPropagation());
  }
  if (btnStockUnitPriceFilterReset) {
    btnStockUnitPriceFilterReset.addEventListener("click", (e) => {
      e.stopPropagation();
      if (!lastStockUnitPriceData?.previewRows?.length) return;
      populateStockUnitPriceCodeFilterFromPreview();
      renderStockUnitPricePreviewTable();
    });
  }

  /**
   * 생산 엑셀 업로드 허브 — 팀별 화면과 동일 핸들러 연결
   * @param {HTMLElement | null} dropzone
   * @param {HTMLInputElement | null} input
   * @param {(file: File) => void | Promise<void>} handler
   */
  function bindTeamProdHubPair(dropzone, input, handler) {
    if (!input) return;
    input.addEventListener("change", () => {
      const f = input.files && input.files[0];
      if (f) handler(f);
      input.value = "";
    });
    if (!dropzone) return;
    ["dragenter", "dragover"].forEach((ev) => {
      dropzone.addEventListener(ev, (e) => {
        e.preventDefault();
        dropzone.classList.add("dragover");
      });
    });
    ["dragleave", "drop"].forEach((ev) => {
      dropzone.addEventListener(ev, (e) => {
        e.preventDefault();
        dropzone.classList.remove("dragover");
      });
    });
    dropzone.addEventListener("drop", (e) => {
      const f = e.dataTransfer.files && e.dataTransfer.files[0];
      if (f) handler(f);
    });
  }

  bindTeamProdHubPair(teamProdDrawingDropzone, teamProdDrawingFileInput, handleDrawingLogFile);
  bindTeamProdHubPair(teamProdParallelDropzone, teamProdParallelFileInput, handleParallelLogFile);
  bindTeamProdHubPair(teamProdMufflerDropzone, teamProdMufflerFileInput, handleMufflerLogFile);
  bindTeamProdHubPair(teamProdDoublePipeDropzone, teamProdDoublePipeFileInput, handleDoublePipeLogFile);
  bindTeamProdHubPair(teamProdHwaseungDropzone, teamProdHwaseungFileInput, handleHwaseungDoublePipeLogFile);

  if (drawingLogSlicerClear) {
    drawingLogSlicerClear.addEventListener("click", () => {
      drawingLogSelectedMonths.clear();
      drawingLogSelectedDays.clear();
      renderDrawingLogPanel();
    });
  }
  if (drawingLogGranularityEl) {
    drawingLogGranularityEl.addEventListener("change", () => {
      drawingLogGranularity = drawingLogGranularityEl.value === "day" ? "day" : "month";
      if (drawingLogGranularity === "month") drawingLogSelectedDays.clear();
      renderDrawingLogPanel();
    });
  }
  if (drawingLogYearPrev) {
    drawingLogYearPrev.addEventListener("click", () => {
      drawingLogTimelineYear -= 1;
      pruneDrawingLogSelectedDaysInvalid();
      renderDrawingLogPanel();
    });
  }
  if (drawingLogYearNext) {
    drawingLogYearNext.addEventListener("click", () => {
      drawingLogTimelineYear += 1;
      pruneDrawingLogSelectedDaysInvalid();
      renderDrawingLogPanel();
    });
  }
  if (drawingLogTimeline) {
    drawingLogTimeline.addEventListener("click", (e) => {
      if (!lastDrawingLog) return;
      const ymChip = e.target.closest("[data-ym]");
      if (ymChip && ymChip.dataset.ym) {
        if (ymChip.disabled) return;
        const ym = ymChip.dataset.ym;
        if (drawingLogSelectedMonths.has(ym)) drawingLogSelectedMonths.delete(ym);
        else drawingLogSelectedMonths.add(ym);
        if (drawingLogGranularity === "day") pruneDrawingLogSelectedDaysInvalid();
        renderDrawingLogPanel();
        return;
      }
      const dayChip = e.target.closest("[data-day-key]");
      if (dayChip && dayChip.dataset.dayKey) {
        const dk = dayChip.dataset.dayKey;
        if (drawingLogSelectedDays.has(dk)) drawingLogSelectedDays.delete(dk);
        else drawingLogSelectedDays.add(dk);
        renderDrawingLogPanel();
      }
    });
  }

  if (parallelLogSlicerClear) {
    parallelLogSlicerClear.addEventListener("click", () => {
      parallelLogSelectedMonths.clear();
      parallelLogSelectedDays.clear();
      parallelLogProcessCut = true;
      parallelLogProcessBend = true;
      if (parallelLogFilterCutEl) parallelLogFilterCutEl.checked = true;
      if (parallelLogFilterBendEl) parallelLogFilterBendEl.checked = true;
      renderParallelLogPanel();
    });
  }
  if (parallelLogGranularityEl) {
    parallelLogGranularityEl.addEventListener("change", () => {
      parallelLogGranularity = parallelLogGranularityEl.value === "day" ? "day" : "month";
      if (parallelLogGranularity === "month") parallelLogSelectedDays.clear();
      renderParallelLogPanel();
    });
  }
  function readParallelLogProcessFromCheckboxes() {
    parallelLogProcessCut = parallelLogFilterCutEl ? parallelLogFilterCutEl.checked : true;
    parallelLogProcessBend = parallelLogFilterBendEl ? parallelLogFilterBendEl.checked : true;
  }
  function onParallelLogProcessCheckboxChange() {
    if (parallelLogFilterCutEl && parallelLogFilterBendEl) {
      if (!parallelLogFilterCutEl.checked && !parallelLogFilterBendEl.checked) {
        parallelLogFilterCutEl.checked = true;
        parallelLogFilterBendEl.checked = true;
      }
    }
    readParallelLogProcessFromCheckboxes();
    renderParallelLogPanel();
  }
  if (parallelLogFilterCutEl) {
    parallelLogFilterCutEl.addEventListener("change", onParallelLogProcessCheckboxChange);
  }
  if (parallelLogFilterBendEl) {
    parallelLogFilterBendEl.addEventListener("change", onParallelLogProcessCheckboxChange);
  }
  if (parallelLogKpiScopeEl) {
    parallelLogKpiScopeEl.addEventListener("change", () => {
      parallelLogKpiScope = parallelLogKpiScopeEl.value === "overall" ? "overall" : "process";
      renderParallelLogPanel();
    });
  }
  if (parallelLogYearPrev) {
    parallelLogYearPrev.addEventListener("click", () => {
      parallelLogTimelineYear -= 1;
      pruneParallelLogSelectedDaysInvalid();
      renderParallelLogPanel();
    });
  }
  if (parallelLogYearNext) {
    parallelLogYearNext.addEventListener("click", () => {
      parallelLogTimelineYear += 1;
      pruneParallelLogSelectedDaysInvalid();
      renderParallelLogPanel();
    });
  }
  if (parallelLogTimeline) {
    parallelLogTimeline.addEventListener("click", (e) => {
      if (!lastParallelLog) return;
      const ymChip = e.target.closest("[data-ym]");
      if (ymChip && ymChip.dataset.ym) {
        if (ymChip.disabled) return;
        const ym = ymChip.dataset.ym;
        if (parallelLogSelectedMonths.has(ym)) parallelLogSelectedMonths.delete(ym);
        else parallelLogSelectedMonths.add(ym);
        if (parallelLogGranularity === "day") pruneParallelLogSelectedDaysInvalid();
        renderParallelLogPanel();
        return;
      }
      const dayChip = e.target.closest("[data-day-key]");
      if (dayChip && dayChip.dataset.dayKey) {
        const dk = dayChip.dataset.dayKey;
        if (parallelLogSelectedDays.has(dk)) parallelLogSelectedDays.delete(dk);
        else parallelLogSelectedDays.add(dk);
        renderParallelLogPanel();
      }
    });
  }

  if (doublePipeLogSlicerClear) {
    doublePipeLogSlicerClear.addEventListener("click", () => {
      doublePipeLogSelectedMonths.clear();
      doublePipeLogSelectedDays.clear();
      doublePipeLogProcessMach = true;
      doublePipeLogProcessForm = true;
      if (doublePipeLogFilterMachEl) doublePipeLogFilterMachEl.checked = true;
      if (doublePipeLogFilterFormEl) doublePipeLogFilterFormEl.checked = true;
      renderDoublePipeLogPanel();
    });
  }
  if (doublePipeLogGranularityEl) {
    doublePipeLogGranularityEl.addEventListener("change", () => {
      doublePipeLogGranularity = doublePipeLogGranularityEl.value === "day" ? "day" : "month";
      if (doublePipeLogGranularity === "month") doublePipeLogSelectedDays.clear();
      renderDoublePipeLogPanel();
    });
  }
  function readDoublePipeLogProcessFromCheckboxes() {
    doublePipeLogProcessMach = doublePipeLogFilterMachEl ? doublePipeLogFilterMachEl.checked : true;
    doublePipeLogProcessForm = doublePipeLogFilterFormEl ? doublePipeLogFilterFormEl.checked : true;
  }
  function onDoublePipeLogProcessCheckboxChange() {
    if (doublePipeLogFilterMachEl && doublePipeLogFilterFormEl) {
      if (!doublePipeLogFilterMachEl.checked && !doublePipeLogFilterFormEl.checked) {
        doublePipeLogFilterMachEl.checked = true;
        doublePipeLogFilterFormEl.checked = true;
      }
    }
    readDoublePipeLogProcessFromCheckboxes();
    renderDoublePipeLogPanel();
  }
  if (doublePipeLogFilterMachEl) {
    doublePipeLogFilterMachEl.addEventListener("change", onDoublePipeLogProcessCheckboxChange);
  }
  if (doublePipeLogFilterFormEl) {
    doublePipeLogFilterFormEl.addEventListener("change", onDoublePipeLogProcessCheckboxChange);
  }
  if (doublePipeLogKpiScopeEl) {
    doublePipeLogKpiScopeEl.addEventListener("change", () => {
      doublePipeLogKpiScope = doublePipeLogKpiScopeEl.value === "overall" ? "overall" : "process";
      renderDoublePipeLogPanel();
    });
  }
  if (doublePipeLogYearPrev) {
    doublePipeLogYearPrev.addEventListener("click", () => {
      doublePipeLogTimelineYear -= 1;
      pruneDoublePipeLogSelectedDaysInvalid();
      renderDoublePipeLogPanel();
    });
  }
  if (doublePipeLogYearNext) {
    doublePipeLogYearNext.addEventListener("click", () => {
      doublePipeLogTimelineYear += 1;
      pruneDoublePipeLogSelectedDaysInvalid();
      renderDoublePipeLogPanel();
    });
  }
  if (doublePipeLogTimeline) {
    doublePipeLogTimeline.addEventListener("click", (e) => {
      if (!lastDoublePipeLog) return;
      const ymChip = e.target.closest("[data-ym]");
      if (ymChip && ymChip.dataset.ym) {
        if (ymChip.disabled) return;
        const ym = ymChip.dataset.ym;
        if (doublePipeLogSelectedMonths.has(ym)) doublePipeLogSelectedMonths.delete(ym);
        else doublePipeLogSelectedMonths.add(ym);
        if (doublePipeLogGranularity === "day") pruneDoublePipeLogSelectedDaysInvalid();
        renderDoublePipeLogPanel();
        return;
      }
      const dayChip = e.target.closest("[data-day-key]");
      if (dayChip && dayChip.dataset.dayKey) {
        const dk = dayChip.dataset.dayKey;
        if (doublePipeLogSelectedDays.has(dk)) doublePipeLogSelectedDays.delete(dk);
        else doublePipeLogSelectedDays.add(dk);
        renderDoublePipeLogPanel();
      }
    });
  }

  if (hwaseungDoublePipeLogSlicerClear) {
    hwaseungDoublePipeLogSlicerClear.addEventListener("click", () => {
      hwaseungDoublePipeLogSelectedMonths.clear();
      hwaseungDoublePipeLogSelectedDays.clear();
      hwaseungDoublePipeLogProcessSelected = new Set(HWASEUNG_DOUBLE_PIPE_ALL_PROC_KEYS);
      if (hwaseungDoublePipeLogProcessFilters) {
        hwaseungDoublePipeLogProcessFilters.querySelectorAll('input[type="checkbox"][data-hwaseung-proc]').forEach((el) => {
          el.checked = true;
        });
      }
      renderHwaseungDoublePipeLogPanel();
    });
  }
  if (hwaseungDoublePipeLogGranularityEl) {
    hwaseungDoublePipeLogGranularityEl.addEventListener("change", () => {
      hwaseungDoublePipeLogGranularity =
        hwaseungDoublePipeLogGranularityEl.value === "day" ? "day" : "month";
      if (hwaseungDoublePipeLogGranularity === "month") hwaseungDoublePipeLogSelectedDays.clear();
      renderHwaseungDoublePipeLogPanel();
    });
  }
  if (hwaseungDoublePipeLogKpiScopeEl) {
    hwaseungDoublePipeLogKpiScopeEl.addEventListener("change", () => {
      hwaseungDoublePipeLogKpiScope = hwaseungDoublePipeLogKpiScopeEl.value === "overall" ? "overall" : "process";
      renderHwaseungDoublePipeLogPanel();
    });
  }
  function onHwaseungDoublePipeLogProcessCheckboxChange() {
    readHwaseungDoublePipeLogProcessSelectionFromDom();
    renderHwaseungDoublePipeLogPanel();
  }
  if (hwaseungDoublePipeLogProcessFilters) {
    hwaseungDoublePipeLogProcessFilters.addEventListener("change", (e) => {
      if (e.target && e.target.matches && e.target.matches('input[type="checkbox"][data-hwaseung-proc]')) {
        onHwaseungDoublePipeLogProcessCheckboxChange();
      }
    });
  }
  if (hwaseungDoublePipeLogYearPrev) {
    hwaseungDoublePipeLogYearPrev.addEventListener("click", () => {
      hwaseungDoublePipeLogTimelineYear -= 1;
      pruneHwaseungDoublePipeLogSelectedDaysInvalid();
      renderHwaseungDoublePipeLogPanel();
    });
  }
  if (hwaseungDoublePipeLogYearNext) {
    hwaseungDoublePipeLogYearNext.addEventListener("click", () => {
      hwaseungDoublePipeLogTimelineYear += 1;
      pruneHwaseungDoublePipeLogSelectedDaysInvalid();
      renderHwaseungDoublePipeLogPanel();
    });
  }
  if (hwaseungDoublePipeLogTimeline) {
    hwaseungDoublePipeLogTimeline.addEventListener("click", (e) => {
      if (!lastHwaseungDoublePipeLog) return;
      const ymChip = e.target.closest("[data-ym]");
      if (ymChip && ymChip.dataset.ym) {
        if (ymChip.disabled) return;
        const ym = ymChip.dataset.ym;
        if (hwaseungDoublePipeLogSelectedMonths.has(ym)) hwaseungDoublePipeLogSelectedMonths.delete(ym);
        else hwaseungDoublePipeLogSelectedMonths.add(ym);
        if (hwaseungDoublePipeLogGranularity === "day") pruneHwaseungDoublePipeLogSelectedDaysInvalid();
        renderHwaseungDoublePipeLogPanel();
        return;
      }
      const dayChip = e.target.closest("[data-day-key]");
      if (dayChip && dayChip.dataset.dayKey) {
        const dk = dayChip.dataset.dayKey;
        if (hwaseungDoublePipeLogSelectedDays.has(dk)) hwaseungDoublePipeLogSelectedDays.delete(dk);
        else hwaseungDoublePipeLogSelectedDays.add(dk);
        renderHwaseungDoublePipeLogPanel();
      }
    });
  }

  function mufflerLogSegIndexFromEventTarget(t) {
    const root = t && t.closest && t.closest(".muffler-log-segment");
    if (!root || root.dataset.segIndex == null) return NaN;
    const n = parseInt(root.dataset.segIndex, 10);
    return Number.isFinite(n) ? n : NaN;
  }

  if (mufflerLogSegmentsHost) {
    mufflerLogSegmentsHost.addEventListener("click", (e) => {
      if (!lastMufflerLog || !lastMufflerLog.segments) return;
      const segIdx = mufflerLogSegIndexFromEventTarget(e.target);
      if (!Number.isFinite(segIdx)) return;
      const st = getMufflerSegState(segIdx);

      if (e.target.closest(".muffler-log-seg-clear")) {
        st.selectedMonths.clear();
        st.selectedDays.clear();
        st.processCut = true;
        st.processForming = true;
        st.processMachine = true;
        const root = e.target.closest(".muffler-log-segment");
        if (root) {
          root.querySelectorAll(".muffler-seg-filter").forEach((el) => {
            el.checked = true;
          });
        }
        renderMufflerLogPanel();
        return;
      }
      if (e.target.closest(".muffler-log-seg-year-prev")) {
        st.timelineYear -= 1;
        pruneMufflerLogSelectedDaysInvalid(segIdx);
        renderMufflerLogPanel();
        return;
      }
      if (e.target.closest(".muffler-log-seg-year-next")) {
        st.timelineYear += 1;
        pruneMufflerLogSelectedDaysInvalid(segIdx);
        renderMufflerLogPanel();
        return;
      }
      const ymChip = e.target.closest("[data-ym]");
      if (ymChip && ymChip.dataset.ym) {
        if (ymChip.disabled) return;
        const ym = ymChip.dataset.ym;
        if (st.selectedMonths.has(ym)) st.selectedMonths.delete(ym);
        else st.selectedMonths.add(ym);
        if (st.granularity === "day") pruneMufflerLogSelectedDaysInvalid(segIdx);
        renderMufflerLogPanel();
        return;
      }
      const dayChip = e.target.closest("[data-day-key]");
      if (dayChip && dayChip.dataset.dayKey) {
        const dk = dayChip.dataset.dayKey;
        if (st.selectedDays.has(dk)) st.selectedDays.delete(dk);
        else st.selectedDays.add(dk);
        renderMufflerLogPanel();
      }
    });

    mufflerLogSegmentsHost.addEventListener("change", (e) => {
      if (!lastMufflerLog || !lastMufflerLog.segments) return;
      const segIdx = mufflerLogSegIndexFromEventTarget(e.target);
      if (!Number.isFinite(segIdx)) return;
      const st = getMufflerSegState(segIdx);

      const gran = e.target.closest(".muffler-log-seg-granularity");
      if (gran && e.target === gran) {
        st.granularity = gran.value === "day" ? "day" : "month";
        if (st.granularity === "month") st.selectedDays.clear();
        renderMufflerLogPanel();
        return;
      }

      const kpiScopeSel = e.target.closest(".muffler-log-seg-kpi-scope");
      if (kpiScopeSel && e.target === kpiScopeSel) {
        st.kpiScope = kpiScopeSel.value === "overall" ? "overall" : "process";
        renderMufflerLogPanel();
        return;
      }

      const procChk = e.target.closest(".muffler-seg-filter");
      if (procChk && procChk.matches("input[type=checkbox]")) {
        const root = procChk.closest(".muffler-log-segment");
        const cCut = root && root.querySelector('.muffler-seg-filter[data-proc="cut"]');
        const cF = root && root.querySelector('.muffler-seg-filter[data-proc="forming"]');
        const cM = root && root.querySelector('.muffler-seg-filter[data-proc="machine"]');
        if (cCut && cF && cM && !cCut.checked && !cF.checked && !cM.checked) {
          cCut.checked = true;
          cF.checked = true;
          cM.checked = true;
        }
        st.processCut = cCut ? cCut.checked : true;
        st.processForming = cF ? cF.checked : true;
        st.processMachine = cM ? cM.checked : true;
        renderMufflerLogPanel();
      }
    });
  }

  ["dragenter", "dragover"].forEach((ev) => {
    dropzone.addEventListener(ev, (e) => {
      e.preventDefault();
      dropzone.classList.add("dragover");
    });
  });
  ["dragleave", "drop"].forEach((ev) => {
    dropzone.addEventListener(ev, (e) => {
      e.preventDefault();
      dropzone.classList.remove("dragover");
    });
  });
  dropzone.addEventListener("drop", (e) => {
    const f = e.dataTransfer.files && e.dataTransfer.files[0];
    if (f) handleFile(f);
  });

  btnApply.addEventListener("click", applyMapping);
  btnExport.addEventListener("click", () => {
    if (lastBoard) exportBoard(lastBoard);
  });

  function closeAllPickers(exceptKey) {
    Object.keys(filterState).forEach((k) => {
      if (k === exceptKey) return;
      filterState[k].panel.hidden = true;
    });
    closeAllStockUnitPricePicker();
    closeStockTablePickerPanels();
  }

  function bindPicker(key) {
    const st = filterState[key];
    st.button.addEventListener("click", (e) => {
      e.stopPropagation();
      const willOpen = st.panel.hidden;
      closeAllPickers(key);
      st.panel.hidden = !willOpen;
      if (willOpen) st.search.focus();
    });
    st.search.addEventListener("input", () => {
      renderPickerList(key);
    });
    st.panel.addEventListener("click", (e) => e.stopPropagation());
  }

  initializePickers();
  bindPicker("gubun");
  bindPicker("code");
  bindPicker("name");
  bindPicker("type");
  bindPicker("export");

  document.addEventListener("click", () => {
    closeAllPickers("");
  });

  filterReset.addEventListener("click", () => {
    clearFilterSelections();
    if (lastBoard) {
      populateAllFilters(lastBoard, false);
      renderBoard(lastBoard);
      renderSimpleTable(lastBoard);
    }
  });

  /**
   * @param {"category"|"code"|"name"|"stock"} key
   */
  function bindStockTableFilterKey(key) {
    const map = {
      category: { btn: btnFilterStockCategory, panel: panelFilterStockCategory, search: searchFilterStockCategory },
      code: { btn: btnFilterStockCode, panel: panelFilterStockCode, search: searchFilterStockCode },
      name: { btn: btnFilterStockName, panel: panelFilterStockName, search: searchFilterStockName },
      stock: { btn: btnFilterStockQty, panel: panelFilterStockQty, search: searchFilterStockQty },
    };
    const u = map[key];
    if (!u || !u.btn || !u.panel || !u.search) return;
    u.btn.addEventListener("click", (e) => {
      e.stopPropagation();
      const willOpen = u.panel.hidden;
      closeAllPickers("");
      closeStockTablePickerPanels();
      u.panel.hidden = !willOpen;
      if (willOpen) {
        renderStockFilterList(key);
        u.search.focus();
      }
    });
    u.search.addEventListener("input", () => {
      if (key === "code") applyStockCodeSearchToSelection();
      else renderStockFilterList(key);
    });
    u.panel.addEventListener("click", (e) => e.stopPropagation());
  }

  ["category", "code", "name", "stock"].forEach((k) =>
    bindStockTableFilterKey(/** @type {"category"|"code"|"name"|"stock"} */ (k))
  );

  if (btnStockCodeSelectBySearch) {
    btnStockCodeSelectBySearch.addEventListener("click", (e) => {
      e.stopPropagation();
      stockFilterBulkBySearchQuery("code", "select");
    });
  }
  if (btnStockCodeClearBySearch) {
    btnStockCodeClearBySearch.addEventListener("click", (e) => {
      e.stopPropagation();
      stockFilterBulkBySearchQuery("code", "clear");
    });
  }
  if (btnStockNameSelectBySearch) {
    btnStockNameSelectBySearch.addEventListener("click", (e) => {
      e.stopPropagation();
      stockFilterBulkBySearchQuery("name", "select");
    });
  }
  if (btnStockNameClearBySearch) {
    btnStockNameClearBySearch.addEventListener("click", (e) => {
      e.stopPropagation();
      stockFilterBulkBySearchQuery("name", "clear");
    });
  }

  if (stockTableFilterReset) {
    stockTableFilterReset.addEventListener("click", (e) => {
      e.stopPropagation();
      if (!lastStockData || !lastStockData.rows || lastStockData.rows.length === 0) return;
      populateStockTableFiltersFromData();
      renderStockTableView();
    });
  }

  if (orderCalPrev) {
    orderCalPrev.addEventListener("click", () => {
      orderCalendarCursor.setMonth(orderCalendarCursor.getMonth() - 1);
      orderCalendarNeedsSync = false;
      orderCalendarSelectedYmd = "";
      if (lastBoard) renderOrderCalendar(lastBoard);
    });
  }
  if (orderCalNext) {
    orderCalNext.addEventListener("click", () => {
      orderCalendarCursor.setMonth(orderCalendarCursor.getMonth() + 1);
      orderCalendarNeedsSync = false;
      orderCalendarSelectedYmd = "";
      if (lastBoard) renderOrderCalendar(lastBoard);
    });
  }

  if (orderCalendarTypeSelect) {
    orderCalendarTypeSelect.addEventListener("change", () => {
      applyOrderCalendarTypeFromSelect();
    });
  }

  if (orderCalendarGrid) {
    orderCalendarGrid.addEventListener("click", (e) => {
      const cell = e.target && e.target.closest && e.target.closest(".order-calendar-cell[data-ymd]");
      if (!cell) return;
      const y = cell.dataset.ymd;
      orderCalendarSelectedYmd = orderCalendarSelectedYmd === y ? "" : y;
      if (lastBoard) renderOrderCalendar(lastBoard);
    });
  }

  function clearOrderCalendarPrintFit() {
    if (!orderCalendarPanel) return;
    orderCalendarPanel.style.removeProperty("zoom");
  }

  /** 1mm → px (@page margin 과 동일한 mm 기준) */
  let orderCalendarCssMmToPxCache = null;
  function orderCalendarCssMmToPx(mm) {
    if (orderCalendarCssMmToPxCache == null) {
      const probe = document.createElement("div");
      probe.style.cssText =
        "position:absolute;left:-9999px;top:0;width:1mm;height:1mm;margin:0;padding:0;border:0;visibility:hidden;pointer-events:none";
      document.body.appendChild(probe);
      orderCalendarCssMmToPxCache = probe.getBoundingClientRect().width;
      document.body.removeChild(probe);
    }
    return mm * orderCalendarCssMmToPxCache;
  }

  /**
   * A3 가로 인쇄 영역(420−10)mm × (297−10)mm 안에 #orderCalendarPanel 전체가 들어가도록,
   * 넘칠 때만 zoom 을 1 이하로 적용. 확대는 하지 않음.
   */
  function applyOrderCalendarPrintFit() {
    clearOrderCalendarPrintFit();
    if (!document.body.classList.contains("app-body--print-calendar-only")) return;
    if (!orderCalendarPanel || orderCalendarPanel.hidden) return;

    const marginMm = 5;
    /* A3 landscape: 가로 420mm × 세로 297mm (ISO 216), @page margin 5mm × 2 */
    const targetW = orderCalendarCssMmToPx(420 - 2 * marginMm);
    const targetH = orderCalendarCssMmToPx(297 - 2 * marginMm);
    if (targetW < 2 || targetH < 2) return;

    const w = orderCalendarPanel.scrollWidth;
    const h = orderCalendarPanel.scrollHeight;
    if (w < 1 || h < 1) return;

    const scale = Math.min(targetW / w, targetH / h, 1);
    if (scale >= 1) return;

    if (typeof CSS !== "undefined" && CSS.supports && CSS.supports("zoom", "1")) {
      orderCalendarPanel.style.zoom = String(scale);
    }
  }

  function clearStockTablePrintFit() {
    if (!stockTablePanel) return;
    stockTablePanel.style.removeProperty("zoom");
  }

  /** A4 세로 인쇄 영역(210−16)mm × (297−16)mm — 넘칠 때만 zoom */
  function applyStockTablePrintFit() {
    clearStockTablePrintFit();
    if (!document.body.classList.contains("app-body--print-stock-only")) return;
    if (!stockTablePanel || (viewStockTable && viewStockTable.hidden)) return;

    const marginMm = 8;
    const targetW = orderCalendarCssMmToPx(210 - 2 * marginMm);
    const targetH = orderCalendarCssMmToPx(297 - 2 * marginMm);
    if (targetW < 2 || targetH < 2) return;

    const w = stockTablePanel.scrollWidth;
    const h = stockTablePanel.scrollHeight;
    if (w < 1 || h < 1) return;

    const scale = Math.min(targetW / w, targetH / h, 1);
    if (scale >= 1) return;

    if (typeof CSS !== "undefined" && CSS.supports && CSS.supports("zoom", "1")) {
      stockTablePanel.style.zoom = String(scale);
    }
  }

  window.addEventListener("beforeprint", () => {
    requestAnimationFrame(() => {
      requestAnimationFrame(() => {
        applyOrderCalendarPrintFit();
        applyStockTablePrintFit();
      });
    });
  });
  window.addEventListener("afterprint", () => {
    clearOrderCalendarPrintFit();
    clearStockTablePrintFit();
  });

  navItems.forEach((btn) => {
    btn.addEventListener("click", () => {
      applyMainViewFromElement(btn);
    });
  });

  if (sidebarBrandHome) {
    sidebarBrandHome.addEventListener("click", () => {
      setView("home");
    });
  }

  if (viewHome) {
    viewHome.addEventListener("click", (e) => {
      const t = e.target;
      const el = t && t.closest && t.closest("[data-view]");
      if (!el || !viewHome.contains(el)) return;
      applyMainViewFromElement(el);
    });
  }

  renderSummary(null);
  setView("home");
})();
