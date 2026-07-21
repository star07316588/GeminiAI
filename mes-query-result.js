/**
 * mes-query-result.js
 * MES Query Result Panel Component  —  Vue 2 Global Component
 * v1.0.0 · 2026-04-21
 *
 * 相依：Tabulator 6.x（需在此 js 之前載入）
 *
 * ════════════════════════════════════════════════════════════════
 * 使用方式
 * ════════════════════════════════════════════════════════════════
 *
 *   <mes-query-result
 *     ref="result"
 *     title="Wip Information"
 *     :columns="tableColumns"
 *     :data="tableData"
 *     :summary="summaryItems"
 *     :tabulator-options="tabulatorOptions"
 *     @row-click="onRowClick"
 *     @row-dblclick="onRowDblClick"
 *     @selection-change="onSelectionChange"
 *   ></mes-query-result>
 *
 *   // 從父層呼叫 Tabulator 方法：
 *   this.$refs.result.print()
 *   this.$refs.result.download('csv', 'export.csv')
 *   this.$refs.result.setFilter(filterText)
 *   this.$refs.result.clearFilter()
 *   this.$refs.result.setData(newData)
 *
 * ════════════════════════════════════════════════════════════════
 * Props
 * ════════════════════════════════════════════════════════════════
 *
 *   title             String   結果面板標題            預設 'Query Result'
 *   columns           Array    Tabulator columns 定義  必填
 *   data              Array    初始資料               預設 []
 *   summary           Array    pane-header 右側摘要項  預設 []
 *                              格式：[{ label:'WQTY', value:computed/ref }]
 *                              value 為響應式時，父層傳入 computed 字串 key 搭配
 *                              :summary="[{label:'WQTY', value: totalWQty}]"
 *   tabulatorOptions  Object   額外 Tabulator 建構選項 預設 {}
 *                              會與元件預設值 merge（頁面設定優先）
 *   selectable        Boolean  Tabulator selectableRows 預設 false
 *
 * ════════════════════════════════════════════════════════════════
 * Events
 * ════════════════════════════════════════════════════════════════
 *
 *   row-click(row, data, e)          點擊列
 *   row-dblclick(row, data, e)       雙擊列
 *   selection-change(selectedRows)   勾選變動（selectable 模式）
 *   table-built()                    Tabulator 初始化完成
 *
 * ════════════════════════════════════════════════════════════════
 * 公開方法（透過 ref 呼叫）
 * ════════════════════════════════════════════════════════════════
 *
 *   setData(data)              更新全部資料
 *   setFilter(text, fields)    快速全文篩選
 *                              fields 預設 ['LotId','IPN','LotStatus','LotOwner','Step']
 *   clearFilter()              清除篩選
 *   download(type, filename)   下載 ('csv'|'xlsx'|'json'|'pdf'|'html', 'filename.ext')
 *   print(active, style)       Tabulator 原生列印
 *   getSelectedData()          取得已選取列資料陣列
 *   redraw(force)              重繪（視窗 resize 後呼叫）
 *
 * ════════════════════════════════════════════════════════════════
 * Summary 使用範例
 * ════════════════════════════════════════════════════════════════
 *
 *   // 在父層 computed：
 *   totalWQty() { return this.tableData.reduce((s,r) => s+r.WaferQty, 0) },
 *   totalCQty() { return this.tableData.reduce((s,r) => s+r.ChipQty,  0) },
 *
 *   // template：
 *   :summary="[
 *     { label:'WQTY', value: totalWQty },
 *     { label:'CQTY', value: totalCQty },
 *   ]"
 *
 * ════════════════════════════════════════════════════════════════
 * Tabulator columns 快速範例
 * ════════════════════════════════════════════════════════════════
 *
 *   tableColumns: [
 *     { title:'LotId',     field:'LotId',     width:90,  frozen:true,
 *       formatter: c => '<b style="color:var(--accent)">'+c.getValue()+'</b>' },
 *     { title:'LotStatus', field:'LotStatus', width:85,
 *       formatter: c => {
 *         const v=c.getValue(), cls=v==='Hold'?'hold':v==='Active'?'active':'normal'
 *         return '<span class="tab-badge '+cls+'">'+v+'</span>'
 *       }
 *     },
 *     { title:'WaferQty',  field:'WaferQty',  width:80,  hozAlign:'right' },
 *   ]
 */

;(function () {
  'use strict'

  /* ── 預設篩選欄位（setFilter 快速全文時用）── */
  const DEFAULT_FILTER_FIELDS = [
    'LotId', 'IPN', 'LotStatus', 'LotOwner', 'Step',
    'Stage', 'Location', 'ErunTickNo',
  ]

  /* ── 元件預設 Tabulator 設定 ── */
  const DEFAULT_TAB_OPTIONS = {
    layout:         'fitDataFill',
    height:         '100%',   /* 繼承掛載 div 的高度（由 CSS position:absolute;inset:0 確保）*/
    movableColumns:  true,
    columnDefaults:  { headerSort: true, resizable: true },
    rowHeader:{resizable: false, frozen: true, hozAlign:"center", headerSort:false, formatter: "rownum", cssClass:"range-header-col", width: "20"},
    selectableRange:true,
    selectableRangeColumns:true,
    selectableRangeRows:true,
    clipboard:true,
    clipboardCopyStyled:false,
    clipboardCopyRowRange:"range",
    clipboardCopyConfig:{
        rowHeaders:false,
        columnHeaders:false,
    },
  }

  Vue.component('mes-query-result', {
    name: 'MesQueryResult',

    props: {
      title:            { type: String,          default: 'Query Result' },
      columns:          { type: Array,           required: true },
      data:             { type: Array,           default: () => [] },
      summary:          { type: Array,           default: () => [] },
      tabulatorOptions: { type: Object,          default: () => ({}) },
      selectable:       { type: Boolean,         default: false },
    },

    data () {
      return {
        _tab: null,           // Tabulator 實例
        _destroyed: false,  // 元件已銷毀旗標，防止非同步 callback 操作已卸載的 DOM
        rowCount: 0,          // 當前列數（由 dataLoaded 事件同步）
      }
    },

    computed: {
      /* 合併 options：元件預設 < 頁面傳入 */
      mergedOptions () {
        return Object.assign({}, DEFAULT_TAB_OPTIONS, this.tabulatorOptions, {
          columns:    this.columns,
          data:       this.data,
//          selectableRows: this.selectable,
        })
      },
    },

    watch: {
      /* data prop 變動時同步 Tabulator（父層直接 setData 效能更好，此為備援） */
      data (newData) {
        if (this._destroyed || !this._tab) return
          this._tab.setData(newData).then(() => {
          if (!this._destroyed) this.rowCount = newData.length
          })
      },
    },

    mounted () {
      this._initTabulator()
    },

    beforeDestroy () {
      this._destroyed = true
      const tab = this._tab
      this._tab = null
      this.$nextTick(() => {
        try { if (tab) tab.destroy() } catch {}
      })
    },

    methods: {
      /* ── 初始化 Tabulator ── */
      _initTabulator () {
        if (typeof Tabulator === 'undefined') {
          console.error('[mes-query-result] Tabulator is not loaded.')
          return
        }

        const vm   = this
        const opts = Object.assign({}, this.mergedOptions)
        if (this.selectable) { // 設定有 checkbox ，關閉 range selection
            delete opts.selectableRange;
            delete opts.selectableRangeColumns;
            delete opts.selectableRangeRows;
            opts.clipboardCopyRowRange = 'selected';
            delete opts.clipboardCopyConfig;
            const chkcol = { formatter: "rowSelection", titleFormatter: "rowSelection", hozAlign: "center", width: 50, headerSort: false, headerHozAlign: 'center',
                            cellClick: function (e, cell) {
                                cell.getRow().toggleSelect();
                            }
                        };
            opts.columns.unshift(chkcol);
        }

        /* 事件 hook */
        
        opts.rowClick = function (e, row) {
          vm.$emit('row-click', row, row.getData(), e)
        }
        opts.rowDblClick = function (e, row) {
          vm.$emit('row-dblclick', row, row.getData(), e)
        }
        opts.rowSelectionChanged = function (data, rows) {
          vm.$emit('selection-change', data)
        }
        opts.dataLoaded = function (data) {
          if (vm._destroyed) return
          vm.rowCount = data.length
          vm.$emit('table-built')
        }

        console.log('opts', opts);
        /* 先不傳 height，等 DOM 穩定後再設定 */
        this._tab = new Tabulator(this.$refs.tableEl, opts)

        if (!this.selectable) {
            this._tab.on("rangeChanged", function (range) {
                //range - range component for the selected range
                try {
                    // 修正：只過濾出「可見 (isVisible)」的欄位來計算數量
                    let visibleColsCount = this.getColumns().filter(c => c.isVisible()).length;

                    if (visibleColsCount == range.getColumns().length + 1) {
                        this.options.clipboardCopyConfig.columnHeaders = true;
                    } else {
                        this.options.clipboardCopyConfig.columnHeaders = false;
                    }
                } catch { }
            });
        }
        this.rowCount = this.data.length

        if (this.selectable) {
            this._tab.on("rowSelected", function(row){
                row.update({chkbox: true})
                row.reformat(); 
            });
            this._tab.on("rowDeselected", function(row){
                row.update({chkbox: false})
                row.reformat();
            });  
        }
      },

      /* ════════ 公開方法 ════════ */

      /** 更新全部資料 */
      setData (data) {
        if (this._destroyed || !this._tab) return Promise.resolve()
        return this._tab.setData(data).then(() => {
          if (!this._destroyed) this.rowCount = data.length
        }).catch(() => {})
      },

      /**
       * 快速全文篩選
       * @param {string} text   - 篩選字串，空字串則清除
       * @param {string[]} [fields] - 要搜尋的欄位，省略使用預設清單
       */
      setFilter (text, fields) {
        if (this._destroyed || !this._tab) return
        const t = (text || '').trim()
        if (!t) { this._tab.clearFilter(); return }
        const cols = (fields && fields.length) ? fields : DEFAULT_FILTER_FIELDS
        this._tab.setFilter(
          cols.map(f => ({ field: f, type: 'like', value: t }))
        )
      },

      /** 清除所有篩選 */
      clearFilter () {
        if (this._destroyed || !this._tab) return
        this._tab.clearFilter()
      },

      /**
       * 下載
       * @param {'csv'|'json'|'xlsx'|'pdf'|'html'} type
       * @param {string} filename
       */
      download (type, filename) {
        if (this._destroyed || !this._tab) return
        if (type == 'csv') {
            const data = this._tab.getData();
            const headers = this.columns.map(x=>x.field);
            const titles = this.columns.map(x=>x.title);
            const rows = data.map(obj => 
              headers.map(header => JSON.stringify(obj[header] ?? "")).join(",")
            );
            const csvContent = [titles.join(","), ...rows].join("\n");
            var link = window.document.createElement("a");
            link.setAttribute("href", "data:text/csv;charset=utf-8,%EF%BB%BF" + encodeURI(csvContent));
            link.setAttribute("download", "upload_data.csv");
            link.click();
        }
        else 
          this._tab.download(type, filename || `export.${type}`)
      },

      /**
       * 列印
       * @param {boolean} active  - true: 僅列印勾選列
       * @param {boolean} style   - true: 帶入目前樣式
       */
      print (active = false, style = true) {
        if (this._destroyed || !this._tab) return
        /* Tabulator print() 的第二參數為 true 時，會複製頁面的
           <link>/<style> 到列印視窗，但動態注入的 <style> 不一定
           能完整帶過去，因此額外注入一段專用列印修正 CSS。        */
        this._injectPrintStyle()
        this._tab.print(active, style)
      },

      /* 在 <head> 補上列印修正樣式（只注入一次）*/
      _injectPrintStyle () {
        if (document.getElementById('__mqr_print_style__')) return
        const el = document.createElement('style')
        el.id = '__mqr_print_style__'
        el.textContent = MQR_PRINT_CSS
        document.head.appendChild(el)
      },

      /** 取得已選取列的資料陣列 */
      getSelectedData () {
        if (this._destroyed || !this._tab) return []
        return this._tab.getSelectedData()
      },
      getSelectedRows() {
        if (this._destroyed || !this._tab) return []
        return this._tab.getSelectedRows()
      },

        getTabulator() {
            return this._tab;
        },

        deselectRow() {
            this._tab.deselectRow()
        },

      /** 重繪並同步高度（視窗 resize 後呼叫） */
      redraw (force = false) {
        if (this._destroyed || !this._tab) return
        try { this._tab.redraw(force) } catch {}
      },
    },

    template: `
<div class="mqr-panel">
  <!-- pane-header -->
  <div class="pane-header mqr-header">
    <i class="mdi mdi-table" style="color:var(--accent,#1565C0);font-size:14px;"></i>
    <span class="pane-title">{{ title }}</span>
    <span class="pane-badge">{{ rowCount }}</span>
    <div class="pane-spacer"></div>
    <!-- 摘要項目 -->
    <template v-if="summary.length">
      <div class="mqr-summary">
        <template v-for="(item, i) in summary">
          <div :key="'si-'+i" class="mqr-summary-item">
            <span class="mqr-s-lbl">{{ item.label }}</span>
            <span class="mqr-s-val">{{ item.value }}</span>
          </div>
          <div v-if="i < summary.length - 1" :key="'ss-'+i" class="mqr-summary-sep"></div>
        </template>
      </div>
    </template>
  </div>

  <!-- Tabulator mount point -->
  <div class="mqr-table-wrap">
    <div ref="tableEl"></div>
  </div>
</div>
    `,
  })

  /* ════════════════════════════════════════════════════════════
     列印修正 CSS（透過 _injectPrintStyle 注入，讓 Tabulator
     print() 複製到列印視窗時能帶到正確的邊框樣式）
     ════════════════════════════════════════════════════════════ */
  const MQR_PRINT_CSS = `
@media print {
  /* Tabulator 列印輸出的根容器 */
  .tabulator-print-fullscreen,
  .tabulator-print-fullscreen-hide {
    font-family: 'Titillium Web', 'Segoe UI', sans-serif;
    font-size: 11px;
  }

  /* 整個表格外框 */
  .tabulator-print-table {
    border-collapse: collapse;
    width: 100%;
  }

  /* header 整行：補上頂部邊線 */
  .tabulator-print-table thead tr:first-child th {
    border-top: 1.5px solid #666 !important;
    border-bottom: 1.5px solid #666 !important;
  }

  /* 所有 th */
  .tabulator-print-table thead th {
    background: #f0f0f0 !important;
    color: #333 !important;
    font-size: 10px;
    font-weight: 700;
    letter-spacing: 0.5px;
    text-transform: uppercase;
    padding: 5px 8px !important;
    border-right: 1px solid #ccc !important;
    text-align: left;
  }
  .tabulator-print-table thead th:last-child {
    border-right: none !important;
  }

  /* 資料列 */
  .tabulator-print-table tbody tr {
    page-break-inside: avoid;
  }
  .tabulator-print-table tbody tr:nth-child(even) {
    background: #f8f8f8 !important;
  }

  /* td */
  .tabulator-print-table tbody td {
    padding: 4px 8px !important;
    border-bottom: 1px solid #e0e0e0 !important;
    border-right: 1px solid #e0e0e0 !important;
    color: #111 !important;
    font-size: 11px;
  }
  .tabulator-print-table tbody td:last-child {
    border-right: none !important;
  }

  /* 最後一列底部加線 */
  .tabulator-print-table tbody tr:last-child td {
    border-bottom: 1.5px solid #666 !important;
  }
}
`

  /* ════════════════════════════════════════════════════════════
     CSS
     ════════════════════════════════════════════════════════════ */
  const CSS = `
/* ── mes-query-result ─────────────────────────────────────── */
.mqr-panel {
  flex: 1;
  display: flex; flex-direction: column;
  overflow: hidden;
  background: var(--bg-pane, #f8f8f8);
  border: 1px solid var(--border);
  border-radius: var(--radius-lg);
}
.mqr-header { flex-shrink: 0; }

.mqr-table-wrap {
  flex: 1;
  overflow: hidden;
  /* 高度由 JS 在 mounted 後以 offsetHeight 讀取並傳給 Tabulator，
     確保 Tabulator 拿到真實像素值而非百分比。 */
}

/* Summary */
.mqr-summary {
  display: flex; align-items: center;
  gap: 8px;
  font-family: var(--font-mono, 'Share Tech Mono', monospace);
  font-size: 12px;
}
.mqr-summary-item {
  display: flex; align-items: center; gap: 4px;
}
.mqr-s-lbl { color: var(--accent, blue); }
.mqr-s-val { color: var(--text-main, #1a1a1a); font-weight: 600; }
.mqr-summary-sep {
  width: 1px; height: 12px;
  background: var(--border, #b8b8b8);
  flex-shrink: 0;
}
.cell-editable {
  background-color: white;
}
`

  if (!document.getElementById('__mqr_style__')) {
    const el = document.createElement('style')
    el.id = '__mqr_style__'
    el.textContent = CSS
    document.head.appendChild(el)
  }

})()
